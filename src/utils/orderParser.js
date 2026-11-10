/**
 * 客戶訂單與交期明細表 解析器
 * 支援「客戶訂單與交期明細表.xlsx」：一年一 sheet（例：2026年、2027年、空白）。
 *
 * 資料品質處理（皆為實檔實測出現過的形態）：
 * - 日期混雜：真 Excel Date、"2026/3/6(2026/02/25)"（括號內為後續異動日）、
 *   "26/3/6(26-03-20)(26-03-19)"（多次修改）、"交期不變"
 * - 數量字串："400000(200000/200000)"、"100000(改40000)"
 * - 通知單號："260224(260225)" 之 yymmdd 編號，非日期欄
 * - 客戶名尾隨空格
 *
 * 欄位以「表頭名稱」比對定位（考量日後欄序異動，例如 2027 年新增主管確認欄）。
 * 輸出為 JSON 可序列化結構（日期一律 ISO 字串），供 localStorage 快取。
 */
import * as XLSX from 'xlsx';

// 表頭名稱 → 欄位語意（正規化後比對：去空白）
const HEADERS = {
  originalDue: '原始訂單交期',
  customer: '客戶名',
  orderNo: '訂單號碼',
  product: '品名',
  qty: '數量',
  lotNo: '批號',
  fixIssueDate: '修正訂單發出日期',
  notifyNo: '通知單號',
  updatedDue: '更新的交期',
  shipReportDone: '檢查出貨報表',
  remark: '備註或說明',
  labelDone: '標籤',
  detailDone: '明細',
};

const pad2 = (n) => String(n).padStart(2, '0');

const toISO = (y, m, d) => `${y}-${pad2(m)}-${pad2(d)}`;

/**
 * 將 "26/3/6"、"2026-03-20"、"2026/02/25" 之字串片段解析為 ISO 日期字串。
 * 兩位數年份 → 20xx（此表僅追溯 2026 年內資料，安全）。
 * 回傳 null 表示無法解析。
 */
const parseDateLoose = (str) => {
  const m = String(str).trim().match(/^(\d{2,4})[/\-.](\d{1,2})[/\-.](\d{1,2})/);
  if (!m) return null;
  let y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  const d = parseInt(m[3], 10);
  if (y < 100) y += 2000;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return toISO(y, mo, d);
};

/**
 * 從單一儲存格值抽取「所有」日期（原始日期在前、括號內異動日在後）。
 * 儲存格若為真 Excel Date，回傳 [ISO]。
 */
const extractDates = (cell) => {
  if (cell === null || cell === undefined) return [];
  if (cell instanceof Date && !isNaN(cell.getTime())) {
    return [toISO(cell.getFullYear(), cell.getMonth() + 1, cell.getDate())];
  }
  const s = String(cell).trim();
  if (!s) return [];
  // 可能直接是 ISO（快取還原後資料）
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return [s];
  const dates = [];
  const re = /\d{2,4}[/\-.]\d{1,2}[/\-.]\d{1,2}/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    const iso = parseDateLoose(m[0]);
    if (iso) dates.push(iso);
  }
  return dates;
};

/**
 * yymmdd 通知單號 → ISO 日期；"260224(260225)" 取括號前主編號。
 */
const parseNotifyNo = (cell) => {
  if (cell === null || cell === undefined) return null;
  const m = String(cell).trim().match(/^(\d{2})(\d{2})(\d{2})/);
  if (!m) return null;
  const y = 2000 + parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  const d = parseInt(m[3], 10);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return toISO(y, mo, d);
};

/**
 * 數量正規化："400000(200000/200000)" → 400000、"100000(改40000)" → 100000。
 * 取括號前主數值；無數字則 0。
 */
const parseQty = (cell) => {
  if (typeof cell === 'number' && isFinite(cell)) return Math.round(cell);
  const m = String(cell || '').replace(/,/g, '').match(/\d+/);
  return m ? parseInt(m[0], 10) : 0;
};

const isDone = (cell) => cell !== null && cell !== undefined && String(cell).includes('V');

const diffDays = (isoA, isoB) => {
  if (!isoA || !isoB) return null;
  return Math.round((new Date(isoA) - new Date(isoB)) / 86400000);
};

const buildHeaderMap = (headerRow) => {
  const map = {};
  (headerRow || []).forEach((h, idx) => {
    if (h === null || h === undefined) return;
    const key = String(h).trim();
    for (const [field, name] of Object.entries(HEADERS)) {
      if (key === name && map[field] === undefined) {
        map[field] = idx;
        break;
      }
    }
  });
  return map;
};

/**
 * 解析單一 sheet 的資料列。
 * @returns {{ rows: Array, skipped: number }}
 */
const parseSheetRows = (rawRows) => {
  const header = rawRows[0];
  const col = buildHeaderMap(header);
  if (col.originalDue === undefined || col.customer === undefined) {
    return { rows: [], skipped: 0 };
  }
  const rows = [];
  let skipped = 0;
  for (let i = 1; i < rawRows.length; i++) {
    const r = rawRows[i];
    if (!r || r.every((c) => c === null || c === undefined || String(c).trim() === '')) continue;
    const originalDueDates = extractDates(r[col.originalDue]);
    if (originalDueDates.length === 0) { skipped++; continue; }
    const originalDue = originalDueDates[0];

    // 客戶名 / 品名 / 訂單號 / 批號
    const customer = String(r[col.customer] ?? '').trim();
    if (!customer) { skipped++; continue; }
    const product = String(r[col.product] ?? '').trim();
    const orderNo = String(r[col.orderNo] ?? '').trim();
    const lotNo = col.lotNo !== undefined ? String(r[col.lotNo] ?? '').trim() : '';

    // 更新的交期："交期不變" → 沿用原始交期；含日期 → 主值為最新交期、括號內為歷史異動鏈
    const rawUpdated = r[col.updatedDue];
    const updatedStr = rawUpdated === null || rawUpdated === undefined ? '' : String(rawUpdated).trim();
    const updatedDates = extractDates(rawUpdated);
    const isUnchanged = /交期不變/.test(updatedStr);
    const latestDue = !updatedStr ? null
      : updatedDates.length > 0 ? updatedDates[0]
      : isUnchanged ? originalDue
      : null;

    // 修正訂單發出日期：取第一個日期（發出日）
    const fixDates = col.fixIssueDate !== undefined ? extractDates(r[col.fixIssueDate]) : [];

    // 修改次數：括號內歷史日期個數（每次異動累積）；交期不變 = 0
    let revisedCount = !updatedStr ? 0
      : updatedDates.length > 1 ? updatedDates.length - 1
      : isUnchanged ? 0
      : updatedDates.length === 1 && updatedDates[0] !== originalDue ? 1
      : 0;

    // 通知單號 → 通知發出日
    const notifyDate = col.notifyNo !== undefined ? parseNotifyNo(r[col.notifyNo]) : null;

    const row = {
      originalDue,            // ISO 字串
      latestDue,              // 最終確認交期（ISO | null）
      dueChanged: isUnchanged ? false : (updatedDates.length > 0 ? updatedDates[0] !== originalDue : false),
      revisedCount,
      fixIssueDate: fixDates.length > 0 ? fixDates[0] : null,
      notifyDate,
      delayDays: latestDue ? diffDays(latestDue, originalDue) : null,
      customer,
      product,
      orderNo,
      lotNo,
      qty: parseQty(r[col.qty]),
      shipDone: isDone(col.shipReportDone !== undefined ? r[col.shipReportDone] : null),
      labelDone: col.labelDone !== undefined ? isDone(r[col.labelDone]) : null,
      detailDone: col.detailDone !== undefined ? isDone(r[col.detailDone]) : null,
      remark: col.remark !== undefined ? String(r[col.remark] ?? '').trim() : '',
    };
    rows.push(row);
  }
  return { rows, skipped };
};

/**
 * 解析整本「客戶訂單與交期明細表.xlsx」工作簿。
 * @param {File} file - 使用者選取之 .xlsx 檔案
 * @returns {Promise<{ years: { [year: string]: Array }, skipped: number, sheetNames: string[] }>}
 */
export const parseOrderWorkbook = async (file) => {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { cellDates: true });
  const years = {};
  let skipped = 0;
  const sheetNames = [];
  for (const sn of wb.SheetNames) {
    const m = sn.match(/(\d{4})\s*年?/);
    if (!m) continue; // 跳過「空白」等非年度 sheet
    const year = m[1];
    sheetNames.push(sn);
    const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null });
    const { rows, skipped: sk } = parseSheetRows(rawRows);
    if (rows.length > 0 || !years[year]) {
      years[year] = rows;
      skipped += sk;
    }
  }
  return { years, skipped, sheetNames };
};

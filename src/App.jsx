import { useState, useEffect, useMemo, useRef } from 'react';
import { getMappings, saveMappings, resetMappings } from './utils/db';
import { parseExcelFile, parseSummaryExcel, parseSummaryJSON } from './utils/excelParser';
import { runETLInBrowser, exportSummaryExcelInBrowser, exportSummaryJSONInBrowser, isUUID } from './utils/browserETL';
import { generateMultiYearLLMPrompt, downloadFile, copyTextToClipboard } from './utils/llmExport';
import * as XLSX from 'xlsx';

const MONTH_LABELS = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];

// Inset Focus Soft Palette — 與 --med-cobalt 主色同源的柔霧色階
const CHART_PALETTE = [
  '#5172CD', // Periwinkle (主色)
  '#4E9E92', // Soft Teal
  '#C68A2E', // Warm Amber
  '#C4574E', // Dusty Red
  '#7C7BD1', // Soft Violet
  '#5FA8D3', // Powder Sky
  '#A0688F', // Muted Plum
  '#3E7FB8', // Steel Blue
  '#8A9A4B', // Sage Olive
];

// Clean Technical SVG Icons
const Icons = {
  Folder: () => (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
    </svg>
  ),
  Chart: () => (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="18" y1="20" x2="18" y2="10"></line>
      <line x1="12" y1="20" x2="12" y2="4"></line>
      <line x1="6" y1="20" x2="6" y2="14"></line>
    </svg>
  ),
  FileText: () => (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
      <polyline points="14 2 14 8 20 8"></polyline>
      <line x1="16" y1="13" x2="8" y2="13"></line>
      <line x1="16" y1="17" x2="8" y2="17"></line>
    </svg>
  ),
  Download: () => (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
      <polyline points="7 10 12 15 17 10"></polyline>
      <line x1="12" y1="15" x2="12" y2="3"></line>
    </svg>
  ),
  Sparkles: () => (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z"></path>
    </svg>
  ),
  Settings: () => (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="12" cy="12" r="3"></circle>
      <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
    </svg>
  ),
  AlertTriangle: () => (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"></path>
      <line x1="12" y1="9" x2="12" y2="13"></line>
      <line x1="12" y1="17" x2="12.01" y2="17"></line>
    </svg>
  ),
  Check: () => (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="20 6 9 17 4 12"></polyline>
    </svg>
  ),
  Copy: () => (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect width="14" height="14" x="8" y="8" rx="2" ry="2"></rect>
      <path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"></path>
    </svg>
  ),
  Trash: () => (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="3 6 5 6 21 6"></polyline>
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
    </svg>
  ),
  Refresh: () => (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"></path>
      <path d="M21 3v5h-5"></path>
      <path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"></path>
      <path d="M8 16H3v5"></path>
    </svg>
  )
};

function App() {
  // Workflow Step: "step1_scan" | "step2_etl" | "step3_analytics"
  const [activeStep, setActiveStep] = useState("step1_scan");

  // Step 1: Scan & Verify State (Lazy Initialized)
  const [mappings, setMappings] = useState(() => getMappings());
  const [scannedRows, setScannedRows] = useState(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      if (params.get('mock') === 'true') {
        return [
          {
            fileName: "裝配C-2026A.xlsx",
            filePath: "RawData/2026/零組件入庫-2026/裝配C-2026",
            sheetName: "260102",
            foundCode: "QC10006-R02",
            foundName: "半成品品檢表",
            status: "matched",
            etlStatus: "已納入",
            etlReason: "已納入 (一般品檢數據 · QC10006-R02 · 裝配C · 1月)",
            etlTimestamp: "2026-07-06 14:00:00"
          },
          {
            fileName: "Vivus-2026A.xlsx",
            filePath: "RawData/2026/裝配檢驗-2026/Vivus-2026",
            sheetName: "260103",
            foundCode: "QC10006-R02",
            foundName: "半成品品檢表",
            status: "matched",
            etlStatus: "已納入",
            etlReason: "已納入 (一般品檢數據 · QC10006-R02 · Vivus · 1月)",
            etlTimestamp: "2026-07-06 14:05:00"
          },
          {
            fileName: "射出D-2026B.xlsx",
            filePath: "RawData/2026/零組件入庫-2026/射出D-2026",
            sheetName: "260215",
            foundCode: "QC10007-R03",
            foundName: "零組件入庫品檢表",
            status: "matched",
            etlStatus: "已納入",
            etlReason: "已納入 (一般品檢數據 · QC10007-R03 · 射出D(組件) · 2月)",
            etlTimestamp: "2026-07-06 14:10:00"
          },
          {
            fileName: "QIP-2025-03.xlsx",
            filePath: "RawData/2025/QIP尺寸檢驗-2025/QIP-2025(1~10)",
            sheetName: "250311",
            foundCode: "QC10004-R02",
            foundName: "QIP",
            status: "matched",
            etlStatus: "已納入",
            etlReason: "已納入 (QIP射出巡檢數據)",
            etlTimestamp: "2026-07-06 14:15:00"
          },
          {
            fileName: "出貨檢驗-2025.xlsx",
            filePath: "RawData/2025",
            sheetName: "250401",
            foundCode: "QC10008-R02",
            foundName: "出貨檢驗報告",
            status: "matched",
            etlStatus: "已納入",
            etlReason: "已納入 (一般品檢數據 · QC10008-R02 · 出貨檢驗 · 4月)",
            etlTimestamp: "2026-07-06 14:20:00"
          }
        ];
      }
    }
    return [];
  });

  const [isScanning, setIsScanning] = useState(false);
  const [folderName, setFolderName] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [scanProgress, setScanProgress] = useState(null);
  const [columnFilters, setColumnFilters] = useState({
    fileName: [],
    filePath: [],
    sheetName: [],
    foundCode: [],
    foundName: [],
    status: [],
    etlStatus: [],
    etlReason: [],
    etlTimestamp: []
  });
  const [activeFilterPopover, setActiveFilterPopover] = useState(null);

  // Step 2: ETL Pipeline State
  const [etlYear, setEtlYear] = useState(2025);
  const [isProcessingETL, setIsProcessingETL] = useState(false);
  const [etlProgress, setEtlProgress] = useState(null);
  const [cachedCounts, setCachedCounts] = useState(null);
  const isETLCancelledRef = useRef(false);

  // Step 3: Multi-Year Analytics State
  const [summaryFiles, setSummaryFiles] = useState({});
  const [activeYear, setActiveYear] = useState("2025");
  const [activeSheet, setActiveSheet] = useState("");
  const [selectedItems, setSelectedItems] = useState([]);
  const [selectedMonth, setSelectedMonth] = useState(0);
  const [compareYearSelection, setCompareYearSelection] = useState({});

  // Modals & UI States
  const [showMappingsModal, setShowMappingsModal] = useState(false);
  const [showLLMModal, setShowLLMModal] = useState(false);
  const [llmExportContent, setLlmExportContent] = useState("");
  const [newCode, setNewCode] = useState("");
  const [newName, setNewName] = useState("");
  const [toastMessage, setToastMessage] = useState(null);

  const chartRef = useRef(null);
  const chartInstance = useRef(null);
  const folderInputRef = useRef(null);

  const showToast = (msg) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3000);
  };

  // Derive active summaryData directly
  const summaryData = useMemo(() => {
    if (activeYear === "compare") {
      const years = Object.keys(summaryFiles).sort();
      return summaryFiles[years[0]] || null;
    }
    return summaryFiles[activeYear] || null;
  }, [activeYear, summaryFiles]);

  // Derive effective selected items for Dashboard
  const availableItems = useMemo(() => {
    if (!summaryData || !activeSheet) return [];
    const rows = summaryData[activeSheet];
    if (!rows || rows.length < 2) return [];
    const headerRow = rows[1] || [];
    const items = [];
    headerRow.forEach((h, idx) => {
      const name = String(h || '').trim();
      if (name && name !== '月份' && name !== '小計' && name !== 'NCA') {
        items.push({ name, idx });
      }
    });
    return items;
  }, [activeSheet, summaryData]);

  const effectiveSelectedItems = useMemo(() => {
    if (selectedItems.length > 0) return selectedItems;
    return availableItems.slice(0, 4);
  }, [selectedItems, availableItems]);

  const updateMappings = (newMappings) => {
    setMappings(newMappings);
    saveMappings(newMappings);
    // Sync scanned rows with new mappings directly
    if (scannedRows.length > 0) {
      const updated = scannedRows.map(row => {
        const code = row.foundCode;
        if (!/^QC\d{5}-R\d{2}$/.test(code)) return row;
        return {
          ...row,
          foundName: newMappings[code] || "無對照編碼",
          status: newMappings[code] ? "matched" : "unmatched"
        };
      });
      setScannedRows(updated);
    }
  };

  const handleAddMapping = (e) => {
    e.preventDefault();
    if (!newCode || !newName) return;
    const formattedCode = newCode.trim().toUpperCase();
    if (!/^QC\d{5}-R\d{2}$/i.test(formattedCode)) {
      alert("表單編碼格式必須為 QCxxxxx-Rxx (例如 QC10001-R01)！");
      return;
    }
    const updated = { ...mappings, [formattedCode]: newName.trim() };
    updateMappings(updated);
    setNewCode("");
    setNewName("");
    showToast("已成功新增 QC 表單對照關係");
  };

  const handleDeleteMapping = (code) => {
    const updated = { ...mappings };
    delete updated[code];
    updateMappings(updated);
    showToast(`已刪除 ${code} 對照`);
  };

  const handleResetMappings = () => {
    if (window.confirm("確定要將對照表恢復為系統預設值嗎？自訂對照將會遺失。")) {
      const reset = resetMappings();
      updateMappings(reset);
      showToast("對照表已恢復為系統預設值");
    }
  };

  const handleImportMappings = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const json = JSON.parse(event.target.result);
        const merged = { ...mappings, ...json };
        updateMappings(merged);
        showToast("匯入對照表成功！");
      } catch {
        alert("檔案格式錯誤，請確保匯入正確的 JSON 檔案！");
      }
    };
    reader.readAsText(file);
    e.target.value = null;
  };

  const handleExportMappings = () => {
    downloadFile(JSON.stringify(mappings, null, 2), "qc_mappings_export.json", "application/json");
    showToast("已下載對照表 JSON 檔案");
  };

  const processFilesList = async (files, nameOfFolder) => {
    if (files.length === 0) return;
    setUploadedFiles(files);
    setCachedCounts(null);
    setColumnFilters({
      fileName: [],
      filePath: [],
      sheetName: [],
      foundCode: [],
      foundName: [],
      status: [],
      etlStatus: [],
      etlReason: [],
      etlTimestamp: []
    });
    setIsScanning(true);
    setFolderName(nameOfFolder || "個別檔案統計");

    let detected = null;
    const folderMatch = (nameOfFolder || "").match(/20\d{2}/);
    if (folderMatch) {
      const parsed = parseInt(folderMatch[0], 10);
      if (parsed >= 2010 && parsed <= 2040) detected = parsed;
    }
    if (!detected) {
      for (let i = 0; i < Math.min(files.length, 50); i++) {
        const path = files[i].webkitRelativePath || files[i].name || "";
        const m = path.match(/20\d{2}/);
        if (m) {
          const parsed = parseInt(m[0], 10);
          if (parsed >= 2010 && parsed <= 2040) {
            detected = parsed;
            break;
          }
        }
      }
    }
    if (detected) {
      setEtlYear(detected);
    } else {
      setEtlYear(new Date().getFullYear());
    }

    const excelFiles = Array.from(files).filter(f =>
      f.name.endsWith('.xlsx') || f.name.endsWith('.xls') || f.name.endsWith('.xlsm')
    );

    if (excelFiles.length === 0) {
      alert("選取的對象中沒有可支援的 Excel 檔案！");
      setIsScanning(false);
      return;
    }

    const BATCH_SIZE = 8;
    const results = [];
    const yearForParse = detected || etlYear;
    setScanProgress({ current: 0, total: excelFiles.length });

    const processOneFile = async (file) => {
      const filePath = file.webkitRelativePath || file.name;
      const pathParts = filePath.replace(/\\/g, '/').split('/').filter(p => !isUUID(p));
      const dirPath = pathParts.length > 1 ? pathParts.slice(0, pathParts.length - 1).join('/') : "";

      try {
        const fileRes = await parseExcelFile(file, mappings, yearForParse);
        if (fileRes.success) {
          return fileRes.sheets.map(sheet => ({
            ...sheet,
            filePath: dirPath
          }));
        } else {
          return [{
            fileName: file.name,
            filePath: dirPath,
            sheetName: "N/A",
            foundCode: "錯誤",
            foundName: fileRes.error || "解析失敗",
            status: "error",
            etlStatus: "狀態異常",
            etlReason: `讀取錯誤 (${fileRes.error || "檔案結構損毀"})`,
            etlTimestamp: new Date().toLocaleString('zh-TW', { hour12: false })
          }];
        }
      } catch (e) {
        return [{
          fileName: file.name,
          filePath: dirPath,
          sheetName: "N/A",
          foundCode: "錯誤",
          foundName: "開啟失敗",
          status: "error",
          etlStatus: "狀態異常",
          etlReason: `開啟失敗 (${e.message || "未知錯誤"})`,
          etlTimestamp: new Date().toLocaleString('zh-TW', { hour12: false })
        }];
      }
    };

    for (let i = 0; i < excelFiles.length; i += BATCH_SIZE) {
      const batch = excelFiles.slice(i, i + BATCH_SIZE);
      const batchResults = await Promise.all(batch.map(processOneFile));
      batchResults.forEach(r => results.push(...r));
      setScanProgress({ current: Math.min(i + BATCH_SIZE, excelFiles.length), total: excelFiles.length });
      await new Promise(r => setTimeout(r, 0));
    }

    setScannedRows(results);
    setScanProgress(null);
    setIsScanning(false);
    showToast(`掃描完成！共驗證 ${results.length} 個 QC 檢驗工作表`);
  };

  const handleRunBrowserETL = async (year) => {
    if (!uploadedFiles || uploadedFiles.length === 0) return;

    if (cachedCounts && cachedCounts.year === year && cachedCounts.filesCount === uploadedFiles.length) {
      exportSummaryJSONInBrowser(cachedCounts.data, year);
      showToast(`${year} 年度品檢報表統計.json 已下載！`);
      return;
    }

    setIsProcessingETL(true);
    isETLCancelledRef.current = false;
    setEtlProgress({ current: 0, total: uploadedFiles.length, filename: "初始化中..." });

    try {
      const counts = await runETLInBrowser(uploadedFiles, year, (current, total, filename) => {
        if (isETLCancelledRef.current) {
          setIsProcessingETL(false);
          throw new Error('ETL cancelled by user');
        }
        setEtlProgress({ current, total, filename });
      });

      setCachedCounts({
        data: counts,
        year: year,
        filesCount: uploadedFiles.length
      });

      exportSummaryJSONInBrowser(counts, year);
      showToast(`${year} 年度品檢報表統計.json 已下載！`);
    } catch (e) {
      if (e.message !== 'ETL cancelled by user') {
        console.error(e);
        alert("ETL 運算失敗，請確保選取的是正確的年度品檢原始資料夾。");
      }
    } finally {
      setIsProcessingETL(false);
    }
  };

  const handleExportExcel = (year) => {
    if (!cachedCounts || cachedCounts.year !== year) {
      alert("請先執行一次「輸出 JSON 統計」以產生快取資料。");
      return;
    }
    exportSummaryExcelInBrowser(cachedCounts.data, year);
    showToast(`已下載 ${year} 年度品檢報表統計 Excel`);
  };

  const handleSyncToDashboard = async (year) => {
    let counts = cachedCounts?.year === year ? cachedCounts.data : null;
    if (!counts) {
      if (!uploadedFiles || uploadedFiles.length === 0) {
        alert("請先在步驟 1 選取原始品檢檔案夾！");
        return;
      }
      setIsProcessingETL(true);
      try {
        counts = await runETLInBrowser(uploadedFiles, year, (current, total, filename) => {
          setEtlProgress({ current, total, filename });
        });
        setCachedCounts({ data: counts, year, filesCount: uploadedFiles.length });
      } catch (e) {
        alert("ETL 計算失敗: " + e.message);
        setIsProcessingETL(false);
        return;
      }
      setIsProcessingETL(false);
    }

    const MONTHS_ARR = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    const EXCEL_SHEET_NAMES = {
      'QC10002-R02': '原物料品檢(QC10002-R02)',
      'QC10004-R02': 'QIP(QC10004-R02)',
      'QC10006-R01': '裝配對樣巡檢(QC10006-R01)',
      'QC10006-R02': '半成品品檢(QC10006-R02)',
      'QC10007-R01': '完成品品檢(QC10007-R01 R02)',
      'QC10007-R03': '零組件入庫品檢(QC10007-R03)',
      'QC10008-R02': '出貨檢驗(QC10008-R02)'
    };

    const sheetData = {};
    Object.keys(EXCEL_SHEET_NAMES).forEach(qcCode => {
      const sheetName = EXCEL_SHEET_NAMES[qcCode];
      const qcCounts = counts[qcCode] || {};
      const subCats = Object.keys(qcCounts).filter(k => k !== '未分類').sort();
      if (qcCounts['未分類']) subCats.push('未分類');

      const rows = [];
      rows.push([sheetName]);
      const hasTotal = qcCode !== 'QC10006-R01';
      const headers = ['月份', ...subCats];
      if (hasTotal) headers.push('小計');
      rows.push(headers);

      MONTHS_ARR.forEach(m => {
        const r = [`${m}月`];
        let mSum = 0;
        subCats.forEach(s => {
          const val = qcCounts[s]?.[m] || 0;
          r.push(val);
          mSum += val;
        });
        if (hasTotal) r.push(mSum);
        rows.push(r);
      });

      const totRow = ['小計'];
      let gSum = 0;
      subCats.forEach(s => {
        let colTot = 0;
        MONTHS_ARR.forEach(m => { colTot += (qcCounts[s]?.[m] || 0); });
        totRow.push(colTot);
        gSum += colTot;
      });
      if (hasTotal) totRow.push(gSum);
      rows.push(totRow);

      sheetData[sheetName] = rows;
    });

    const strYear = String(year);
    setSummaryFiles(prev => ({
      ...prev,
      [strYear]: sheetData
    }));
    setActiveYear(strYear);
    const initialSheet = Object.keys(sheetData)[0] || "";
    setActiveSheet(initialSheet);
    setActiveStep("step3_analytics");
    showToast(`已將 ${year} 年度 QMS 數據同步至階段 03 儀表板！`);
  };

  const handleExportIndividualReports = async (year) => {
    if (!uploadedFiles || uploadedFiles.length === 0) {
      alert("請先在步驟 1 選取原始品檢檔案夾！");
      return;
    }

    let counts = cachedCounts?.year === year ? cachedCounts.data : null;
    if (!counts) {
      setIsProcessingETL(true);
      try {
        counts = await runETLInBrowser(uploadedFiles, year, (current, total, filename) => {
          setEtlProgress({ current, total, filename });
        });
        setCachedCounts({ data: counts, year, filesCount: uploadedFiles.length });
      } catch (e) {
        alert("ETL 運算失敗: " + e.message);
        setIsProcessingETL(false);
        return;
      }
      setIsProcessingETL(false);
    }

    try {
      const outputDirHandle = window.showDirectoryPicker ? await window.showDirectoryPicker().catch(() => null) : null;
      await exportIndividualReportsInternal(counts, year, outputDirHandle);
      showToast(`${year} 年度 7 大類獨立 QC 分冊匯出完成！`);
    } catch (e) {
      console.error(e);
      alert("匯出失敗: " + e.message);
    }
  };

  const exportIndividualReportsInternal = async (counts, year, outputDirHandle = null) => {
    const MONTHS_ARR = ["1","2","3","4","5","6","7","8","9","10","11","12"];
    const baseReports = [
      { qcCode: 'QC10002-R02', title: '原物料/配件進料品檢', namePrefix: '進料檢驗' },
      { qcCode: 'QC10004-R02', title: 'QIP尺寸檢驗', namePrefix: 'QIP尺寸檢驗', fixed: ['QIP-Setup','QIP-Patrol','押出-Setup','押出-Patrol'] },
      { qcCode: 'QC10006-R01', title: '裝配對樣巡檢', namePrefix: '裝配巡檢' },
      { qcCode: 'QC10006-R02', title: '半成品品檢', namePrefix: '裝配檢驗' },
      { qcCode: 'QC10007-R01', title: '完成品品檢', namePrefix: '完成品品檢' },
      { qcCode: 'QC10007-R03', title: '零組件入庫品檢', namePrefix: '零組件入庫', splitSubCats: true },
      { qcCode: 'QC10008-R02', title: '出貨檢驗', namePrefix: '出貨檢驗' }
    ];

    const generatedReports = [];
    for (let base of baseReports) {
      const qcCounts = counts[base.qcCode] || {};
      const subCats = Object.keys(qcCounts).filter(k => k !== '未分類').sort();
      if (qcCounts['未分類']) subCats.push('未分類');

      if (base.fixed) {
        generatedReports.push({
          name: `${base.namePrefix}-${year}`,
          title: base.title,
          qcCode: base.qcCode,
          categories: base.fixed,
          qcCounts: qcCounts
        });
      } else if (base.splitSubCats) {
        subCats.forEach(sub => {
          generatedReports.push({
            name: `${base.namePrefix}-${year}_${sub}`,
            title: base.title,
            qcCode: base.qcCode,
            categories: [sub],
            qcCounts: qcCounts
          });
        });
      } else {
        if (subCats.length > 0) {
          generatedReports.push({
            name: `${base.namePrefix}-${year}`,
            title: base.title,
            qcCode: base.qcCode,
            categories: subCats,
            qcCounts: qcCounts
          });
        }
      }
    }

    for (let report of generatedReports) {
      const { name, categories, title, qcCounts, qcCode } = report;
      const colData = categories.map(() => ({}));

      for (let subCat in qcCounts) {
        const colIdx = categories.indexOf(subCat);
        if (colIdx === -1) continue;
        const monthly = qcCounts[subCat];
        for (let m = 1; m <= 12; m++) {
          if (monthly[m]) {
            colData[colIdx][m] = (colData[colIdx][m] || 0) + monthly[m];
          }
        }
      }

      const rows = [[title], ['月份', ...categories]];
      if (qcCode !== 'QC10006-R01') rows[1].push('小計');

      for (let m = 1; m <= 12; m++) {
        const row = [`${MONTHS_ARR[m-1]}月`];
        let total = 0;
        categories.forEach((_, ci) => {
          const v = colData[ci][m] || 0;
          row.push(v);
          total += v;
        });
        if (qcCode !== 'QC10006-R01') row.push(total);
        rows.push(row);
      }

      const totalRow = ['小計'];
      let grandTotal = 0;
      categories.forEach((_, ci) => {
        let t = 0;
        for (let m = 1; m <= 12; m++) t += (colData[ci][m] || 0);
        totalRow.push(t);
        grandTotal += t;
      });
      if (qcCode !== 'QC10006-R01') totalRow.push(grandTotal);
      rows.push(totalRow);

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
      const fileName = `${name}.xlsx`;

      if (outputDirHandle) {
        try {
          const fileHandle = await outputDirHandle.getFileHandle(fileName, { create: true });
          const writable = await fileHandle.createWritable();
          const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          await writable.write(wbout);
          await writable.close();
        } catch {
          XLSX.writeFile(wb, fileName);
        }
      } else {
        XLSX.writeFile(wb, fileName);
      }
    }
  };

  const exportToCSV = (data, name = "QMS_檢驗工作表提取清單") => {
    if (!data || data.length === 0) return;
    const headers = ["檔案名稱", "檔案路徑", "工作表名稱", "表單編碼", "表單對照名稱", "狀態", "是否納入ETL計算", "原因說明", "更新時間"];
    const keys = ["fileName", "filePath", "sheetName", "foundCode", "foundName", "status", "etlStatus", "etlReason", "etlTimestamp"];
    const csvRows = [headers.join(",")];

    for (const row of data) {
      const values = keys.map(key => {
        let val = row[key] || "";
        if (key === 'status') val = getStatusLabel(val);
        const escaped = String(val).replace(/"/g, '""');
        if (escaped.includes(",") || escaped.includes("\n") || escaped.includes('"')) {
          return `"${escaped}"`;
        }
        return escaped;
      });
      csvRows.push(values.join(","));
    }

    const csvContent = "\uFEFF" + csvRows.join("\n");
    downloadFile(csvContent, `${name}.csv`, 'text/csv;charset=utf-8;');
    showToast("已匯出 QMS CSV 提取清單");
  };

  const handleLoadSummaryFile = async (file) => {
    try {
      const isJSON = file.name.endsWith('.json');
      const data = isJSON ? await parseSummaryJSON(file) : await parseSummaryExcel(file);
      const yearMatch = isJSON
        ? (data._meta?.year ? String(data._meta.year) : null)
        : (file.name.match(/(20\d{2})/)?.[1] || null);
      const year = yearMatch || "未知年度";
      if (data._meta) delete data._meta;

      setSummaryFiles(prev => {
        const next = { ...prev, [year]: data };
        const sheets = Object.keys(data).filter(s => s !== "品檢地圖");
        setActiveSheet(prevSheet => prevSheet || sheets[0] || "");
        return next;
      });
      setActiveYear(year);
      showToast(`已載入 ${year} 年度 QMS 統計檔 (${file.name})`);
    } catch {
      alert(`解析檔案失敗: ${file.name}`);
    }
  };

  const handleSummaryFileChange = (e) => {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
      Promise.all(files.map(file => handleLoadSummaryFile(file)));
    }
    e.target.value = null;
  };

  // Derived Dashboard Insights calculation via useMemo
  const dashboardInsights = useMemo(() => {
    if (!summaryData || !activeSheet || effectiveSelectedItems.length === 0) {
      return { total: 0, peakMonth: "N/A", peakVal: 0, avg: 0 };
    }
    const allYears = Object.keys(summaryFiles).sort();
    const selectedCompareYears = activeYear === 'compare'
      ? Object.keys(compareYearSelection).filter(k => compareYearSelection[k]).sort()
      : [];
    const yearsToAggregate = activeYear === 'compare' && selectedCompareYears.length > 0
      ? selectedCompareYears
      : (activeYear === 'compare' ? allYears : [activeYear]);

    let totalSum = 0;
    let peakValue = 0;
    let peakM = "";

    for (let m = 0; m < 12; m++) {
      if (selectedMonth > 0 && m + 1 !== selectedMonth) continue;
      let monthlySum = 0;

      yearsToAggregate.forEach(yr => {
        const yrRows = summaryFiles[yr] ? summaryFiles[yr][activeSheet] : null;
        if (!yrRows) return;
        effectiveSelectedItems.forEach(item => {
          const val = yrRows[m + 2] ? Number(yrRows[m + 2][item.idx]) || 0 : 0;
          monthlySum += val;
        });
      });

      totalSum += monthlySum;
      if (monthlySum > peakValue) {
        peakValue = monthlySum;
        peakM = MONTH_LABELS[m];
      }
    }

    const filteredMonths = selectedMonth > 0 ? 1 : 12;
    const yearCount = activeYear === 'compare' ? Math.max(1, yearsToAggregate.length) : 1;
    return {
      total: totalSum,
      peakMonth: peakM || "N/A",
      peakVal: peakValue,
      avg: filteredMonths > 0 ? Math.round(totalSum / (filteredMonths * yearCount)) : 0
    };
  }, [effectiveSelectedItems, activeSheet, summaryData, selectedMonth, activeYear, summaryFiles, compareYearSelection]);

  // Chart Rendering
  useEffect(() => {
    if (!chartRef.current || !summaryData || !activeSheet || effectiveSelectedItems.length === 0) return;
    const ctx = chartRef.current.getContext('2d');
    const rows = summaryData[activeSheet];
    if (!rows || rows.length < 2) return;

    const allYears = Object.keys(summaryFiles).sort();
    const selectedCompareYears = activeYear === 'compare'
      ? Object.keys(compareYearSelection).filter(k => compareYearSelection[k]).sort()
      : [];
    const years = selectedCompareYears.length > 0 ? selectedCompareYears : allYears;

    let datasets;
    if (activeYear === 'compare' && years.length > 1) {
      datasets = effectiveSelectedItems.map((item, idx) => {
        const dataPoints = years.map(yr => {
          const yrRows = summaryFiles[yr] ? summaryFiles[yr][activeSheet] : null;
          if (!yrRows) return 0;
          if (selectedMonth > 0 && selectedMonth <= 12) {
            const mRow = yrRows[selectedMonth + 1];
            return mRow ? Number(mRow[item.idx]) || 0 : 0;
          }
          const totalRow = yrRows[14];
          return totalRow ? Number(totalRow[item.idx]) || 0 : 0;
        });

        return {
          label: item.name,
          data: dataPoints,
          backgroundColor: CHART_PALETTE[idx % CHART_PALETTE.length],
          borderRadius: 2,
        };
      });
    } else {
      datasets = effectiveSelectedItems.map((item, idx) => {
        const dataPoints = [];
        for (let m = 2; m <= 13; m++) {
          if (selectedMonth > 0 && m - 1 !== selectedMonth) {
            dataPoints.push(0);
            continue;
          }
          const val = rows[m] ? Number(rows[m][item.idx]) || 0 : 0;
          dataPoints.push(val);
        }

        return {
          label: item.name,
          data: dataPoints,
          backgroundColor: CHART_PALETTE[idx % CHART_PALETTE.length],
          borderRadius: 2,
        };
      });
    }

    if (chartInstance.current) {
      chartInstance.current.destroy();
    }

    const chartLabels = (activeYear === 'compare' && years.length > 1)
      ? years
      : MONTH_LABELS;

    chartInstance.current = new window.Chart(ctx, {
      type: 'bar',
      data: {
        labels: chartLabels,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'top',
            labels: {
              font: { family: 'Outfit, sans-serif', size: 11, weight: '600' },
              usePointStyle: true,
              boxWidth: 8
            }
          },
          tooltip: {
            padding: 8,
            backgroundColor: '#2E3440',
            titleFont: { family: 'Outfit, sans-serif', weight: 'bold', size: 12 },
            bodyFont: { family: 'JetBrains Mono, monospace', size: 11 }
          }
        },
        scales: {
          x: {
            stacked: true,
            grid: { display: false },
            ticks: { color: '#7A8394', font: { family: 'Outfit, sans-serif', size: 11 } }
          },
          y: {
            stacked: true,
            beginAtZero: true,
            grid: { color: '#E1E5EC', borderDash: [2, 2] },
            ticks: { color: '#7A8394', font: { family: 'JetBrains Mono, monospace', size: 11 } }
          }
        }
      }
    });

  }, [effectiveSelectedItems, activeSheet, summaryData, selectedMonth, activeYear, summaryFiles, compareYearSelection]);

  const handleOpenLLMModal = () => {
    let filesForLLM = { ...summaryFiles };
    if (Object.keys(filesForLLM).length === 0 && cachedCounts) {
      handleSyncToDashboard(cachedCounts.year);
    }
    const mdPrompt = generateMultiYearLLMPrompt(filesForLLM);
    setLlmExportContent(mdPrompt);
    setShowLLMModal(true);
  };

  const handleCopyPrompt = async () => {
    const success = await copyTextToClipboard(llmExportContent);
    if (success) {
      showToast("QMS 大模型分析指令與數據已複製至剪貼簿！");
    } else {
      alert("複製失敗，請手動選取複製");
    }
  };

  const handleDownloadPrompt = () => {
    const filename = `ISO13485_QC_MultiYear_LLM_Prompt_${new Date().toISOString().split('T')[0]}.md`;
    downloadFile(llmExportContent, filename, 'text/markdown;charset=utf-8');
    showToast("已下載 QMS 診斷指令 .md 檔案");
  };

  const filteredRows = scannedRows.filter(row => {
    const matchesSearch =
      row.fileName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (row.filePath || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.sheetName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.foundCode.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.foundName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (row.etlStatus || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
      (row.etlReason || '').toLowerCase().includes(searchQuery.toLowerCase());

    const matchesStatusSelect = statusFilter === "all" ? true : row.status === statusFilter;
    if (!matchesSearch || !matchesStatusSelect) return false;

    for (const key in columnFilters) {
      const selected = columnFilters[key];
      if (selected && selected.length > 0) {
        if (!selected.includes(row[key])) return false;
      }
    }
    return true;
  });

  const getUniqueColumnValues = (fieldKey) => {
    const tempFiltered = scannedRows.filter(row => {
      const matchesSearch =
        row.fileName.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (row.filePath || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
        row.sheetName.toLowerCase().includes(searchQuery.toLowerCase()) ||
        row.foundCode.toLowerCase().includes(searchQuery.toLowerCase()) ||
        row.foundName.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (row.etlStatus || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
        (row.etlReason || '').toLowerCase().includes(searchQuery.toLowerCase());

      const matchesStatusSelect = statusFilter === "all" ? true : row.status === statusFilter;
      if (!matchesSearch || !matchesStatusSelect) return false;

      for (const key in columnFilters) {
        if (key === fieldKey) continue;
        const selected = columnFilters[key];
        if (selected && selected.length > 0 && !selected.includes(row[key])) return false;
      }
      return true;
    });

    const vals = tempFiltered.map(row => row[fieldKey]);
    const unique = Array.from(new Set(vals.map(v => v === undefined || v === null ? "" : v)));
    return unique.sort((a, b) => String(a).localeCompare(String(b), 'zh-Hant'));
  };

  const getStatusLabel = (status) => {
    switch (status) {
      case 'matched': return '✓ 識別符合';
      case 'unmatched': return '⚠ 缺對照編碼';
      case 'none': return '無 QC 編碼';
      case 'error': return '✗ 讀取錯誤';
      default: return status;
    }
  };

  const renderFilterableHeader = (label, fieldKey) => {
    const isFiltered = columnFilters[fieldKey] && columnFilters[fieldKey].length > 0;
    return (
      <th style={{ paddingRight: '8px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '4px' }}>
          <span>{label}</span>
          <button 
            type="button"
            className={`filter-trigger-btn ${isFiltered ? 'active' : ''}`}
            onClick={(e) => {
              e.stopPropagation();
              setActiveFilterPopover(activeFilterPopover === fieldKey ? null : fieldKey);
            }}
            title={`篩選${label}`}
          >
            ▼
          </button>
        </div>
        {activeFilterPopover === fieldKey && (
          <ColumnFilterPopover 
            fieldKey={fieldKey}
            allValues={getUniqueColumnValues(fieldKey)}
            columnFilters={columnFilters}
            setColumnFilters={setColumnFilters}
            onClose={() => setActiveFilterPopover(null)}
            getStatusLabel={getStatusLabel}
          />
        )}
      </th>
    );
  };

  return (
    <div className="app-container">
      {/* Toast Notification */}
      {toastMessage && (
        <div className="toast">
          <Icons.Check />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* MedTech Header */}
      <header className="app-header">
        <div className="app-title-group">
          <div className="app-main-title">
            <Icons.Chart />
            <span>Mouldex QMS · 醫療器材品管檢驗數據提取與跨年度確效系統</span>
            <span className="medtech-badge-inline">ISO 13485 & GMP</span>
          </div>
          <p className="app-subtitle">MEDICAL DEVICE QUALITY CONTROL WORKBENCH & MULTI-YEAR ETL ENGINE</p>
        </div>
        <div className="header-actions">
          <div className="system-status-indicator">
            <span className="pulse-led"></span>
            <span>系統確效合格</span>
          </div>
          <button 
            className="btn btn-purple btn-sm"
            onClick={handleOpenLLMModal}
            title="一鍵匯出供大模型解析之 ISO 13485 診斷指令與全年度數據"
          >
            <Icons.Sparkles />
            <span>一鍵匯出大模型數據</span>
          </button>
          <button 
            className="btn btn-secondary btn-sm"
            onClick={() => setShowMappingsModal(true)}
            title="設定表單編碼與名稱對照關係"
          >
            <Icons.Settings />
            <span>QC 編碼對照表</span>
          </button>
        </div>
      </header>

      {/* 3-Stage Bench Instrument Stepper */}
      <nav className="workflow-stepper" aria-label="醫療器材品管流程">
        <button 
          className={`step-card ${activeStep === 'step1_scan' ? 'active' : ''} ${scannedRows.length > 0 ? 'completed' : ''}`}
          onClick={() => setActiveStep('step1_scan')}
        >
          <div className="step-number">01</div>
          <div className="step-content">
            <span className="step-tag">STAGE 01</span>
            <span className="step-title">原始檢驗表單掃描與批次校驗</span>
            <span className="step-desc">{scannedRows.length > 0 ? `已驗證 ${scannedRows.length} 工作表` : '上傳批次原始 Excel 報表'}</span>
          </div>
        </button>

        <button 
          className={`step-card ${activeStep === 'step2_etl' ? 'active' : ''} ${cachedCounts ? 'completed' : ''}`}
          onClick={() => setActiveStep('step2_etl')}
        >
          <div className="step-number">02</div>
          <div className="step-content">
            <span className="step-tag">STAGE 02</span>
            <span className="step-title">QMS ETL 清洗與法規報表產出</span>
            <span className="step-desc">{cachedCounts ? `已確效快取 ${cachedCounts.year} 年數據` : '雙引擎自動分類與空白防禦'}</span>
          </div>
        </button>

        <button 
          className={`step-card ${activeStep === 'step3_analytics' ? 'active' : ''} ${Object.keys(summaryFiles).length > 0 ? 'completed' : ''}`}
          onClick={() => setActiveStep('step3_analytics')}
        >
          <div className="step-number">03</div>
          <div className="step-content">
            <span className="step-tag">STAGE 03</span>
            <span className="step-title">跨年度品質趨勢與大模型診斷</span>
            <span className="step-desc">{Object.keys(summaryFiles).length > 0 ? `已載入 ${Object.keys(summaryFiles).length} 個年度` : '製程能力分析與 AI 審查'}</span>
          </div>
        </button>
      </nav>

      {/* ========================================================================= */}
      {/* STAGE 01: RAW DATA SCAN & LOT VERIFICATION */}
      {/* ========================================================================= */}
      {activeStep === 'step1_scan' && (
        <div className="panel-card">
          <div className="panel-header">
            <div>
              <h2 className="panel-title">
                <Icons.Folder />
                <span>階段 01：原始品管 Excel 表單掃描與批次校驗</span>
              </h2>
              <p className="panel-subtitle">批次讀取原始品管資料夾，自動檢驗 7 大類 QC 編碼、Date Code 格式及空白樣板防呆守衛</p>
            </div>
            {scannedRows.length > 0 && (
              <div style={{ display: 'flex', gap: '8px' }}>
                <button 
                  className="btn btn-secondary btn-sm"
                  onClick={() => exportToCSV(filteredRows, folderName || "QMS_品管提取清單")}
                >
                  <Icons.Download />
                  <span>匯出 CSV 檢驗清單</span>
                </button>
                <button 
                  className="btn btn-danger btn-sm"
                  onClick={() => {
                    setScannedRows([]);
                    setUploadedFiles([]);
                    setFolderName("");
                    showToast("已重設檢驗工作表資料");
                  }}
                >
                  <Icons.Trash />
                  <span>清除結果</span>
                </button>
              </div>
            )}
          </div>

          {scannedRows.length === 0 && (
            <div 
              className="upload-zone"
              onClick={() => folderInputRef.current?.click()}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                if (e.dataTransfer.files) processFilesList(e.dataTransfer.files, "拖曳檔案統計");
              }}
            >
              <Icons.Folder />
              <div style={{ fontWeight: 700, fontSize: '16px', color: 'var(--text-primary)' }}>選取或拖曳醫療器材品檢原始資料夾</div>
              <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>支援批次讀取多層子目錄及射出、押出、裝配、完成品等各類 QC 原始檔案</div>
              <button 
                type="button" 
                className="btn btn-primary btn-sm" 
                style={{ marginTop: '6px' }}
                onClick={(e) => {
                  e.stopPropagation();
                  folderInputRef.current?.click();
                }}
              >
                <Icons.Folder />
                <span>選取品檢資料夾</span>
              </button>
            </div>
          )}

          <input 
            type="file" 
            ref={folderInputRef} 
            webkitdirectory="true" 
            directory="true" 
            multiple 
            onChange={(e) => {
              if (e.target.files.length === 0) return;
              const firstFile = e.target.files[0];
              const path = firstFile.webkitRelativePath || "";
              const parts = path.split(/[\\/]/).filter(p => !isUUID(p));
              const folder = parts[0] || "資料夾檔案統計";
              processFilesList(e.target.files, folder);
            }} 
            style={{ display: 'none' }} 
          />

          {isScanning && (
            <div style={{ padding: '36px 0', textAlign: 'center' }}>
              <div style={{ fontSize: '15px', fontWeight: 600, color: 'var(--med-cobalt)', marginBottom: '8px' }}>
                正在掃描校驗原始 Excel 表單... {scanProgress ? `(${scanProgress.current}/${scanProgress.total})` : ''}
              </div>
              {scanProgress && (
                <div className="progress-track" style={{ width: '320px', margin: '0 auto' }}>
                  <div className="progress-fill" style={{ width: `${(scanProgress.current / scanProgress.total) * 100}%` }}></div>
                </div>
              )}
            </div>
          )}

          {scannedRows.some(row => row.etlStatus === '狀態異常') && (
            <div className="alert-banner alert-error" style={{ marginTop: '14px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Icons.AlertTriangle />
                <span style={{ fontWeight: 600, fontSize: '13px' }}>
                  偏差警示：偵測到有部分工作表在確效中被判定為「狀態異常」，請點擊右側按鈕立即過濾排查。
                </span>
              </div>
              <button 
                className="btn btn-danger btn-sm"
                onClick={() => setColumnFilters(prev => ({ ...prev, etlStatus: ['狀態異常'] }))}
              >
                過濾異常項目
              </button>
            </div>
          )}

          {scannedRows.length > 0 && !isScanning && (
            <div style={{ marginTop: '14px' }}>
              <div className="table-controls">
                <div className="filter-group">
                  <input 
                    type="text" 
                    className="search-input" 
                    placeholder="搜尋表單編碼、檢驗批號或檔名..." 
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                  />
                  <select 
                    className="filter-select"
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value)}
                  >
                    <option value="all">所有識別狀態</option>
                    <option value="matched">識別符合</option>
                    <option value="unmatched">缺對照編碼</option>
                    <option value="none">無 QC 編碼</option>
                    <option value="error">讀取錯誤</option>
                  </select>

                  {Object.values(columnFilters).some(arr => arr.length > 0) && (
                    <button 
                      className="btn btn-secondary btn-sm"
                      onClick={() => {
                        setColumnFilters({
                          fileName: [],
                          filePath: [],
                          sheetName: [],
                          foundCode: [],
                          foundName: [],
                          status: [],
                          etlStatus: [],
                          etlReason: [],
                          etlTimestamp: []
                        });
                        setActiveFilterPopover(null);
                      }}
                    >
                      <Icons.Refresh />
                      <span>重設篩選</span>
                    </button>
                  )}
                </div>

                <div style={{ fontSize: '13px', color: 'var(--text-secondary)', fontFamily: 'var(--font-mono)' }}>
                  RECORDS: <strong>{filteredRows.length}</strong> / TOTAL: <strong>{scannedRows.length}</strong>
                </div>
              </div>

              <div className="table-wrapper">
                {activeFilterPopover && (
                  <div className="filter-backdrop" onClick={() => setActiveFilterPopover(null)} />
                )}
                <table className="data-table">
                  <colgroup>
                    <col style={{ width: '40px' }} />
                    <col style={{ width: '160px' }} />
                    <col />
                    <col />
                    <col style={{ width: '120px' }} />
                    <col />
                    <col />
                    <col />
                    <col />
                    <col />
                  </colgroup>
                  <thead>
                    <tr>
                      <th style={{ width: '40px' }}>#</th>
                      {renderFilterableHeader('檔案名稱', 'fileName')}
                      {renderFilterableHeader('路徑資料夾', 'filePath')}
                      {renderFilterableHeader('工作表名稱', 'sheetName')}
                      {renderFilterableHeader('表單編碼', 'foundCode')}
                      {renderFilterableHeader('表單對照名稱', 'foundName')}
                      {renderFilterableHeader('狀態', 'status')}
                      {renderFilterableHeader('是否納入ETL', 'etlStatus')}
                      {renderFilterableHeader('原因說明', 'etlReason')}
                      {renderFilterableHeader('時間戳記', 'etlTimestamp')}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.map((row, idx) => (
                      <tr key={idx}>
                        <td style={{ color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}>{idx + 1}</td>
                        <td className="filename-cell" style={{ fontWeight: 600 }} title={row.fileName}>{row.fileName}</td>
                        <td className="filepath-cell" title={row.filePath}>{row.filePath || '-'}</td>
                        <td style={{ fontFamily: 'var(--font-mono)' }}>{row.sheetName}</td>
                        <td style={{ fontFamily: 'var(--font-mono)', fontWeight: 700, color: 'var(--med-cobalt)', whiteSpace: 'nowrap' }}>{row.foundCode}</td>
                        <td>{row.foundName}</td>
                        <td>
                          {row.status === 'matched' && <span className="status-badge status-matched">✓ 符合</span>}
                          {row.status === 'unmatched' && <span className="status-badge status-unmatched">⚠ 缺對照</span>}
                          {row.status === 'none' && <span className="status-badge status-none">無編碼</span>}
                          {row.status === 'error' && <span className="status-badge status-error">✗ 錯誤</span>}
                        </td>
                        <td>
                          {row.etlStatus === '已納入' && <span className="status-badge status-matched">已納入</span>}
                          {row.etlStatus === '未納入' && <span className="status-badge status-none">未納入</span>}
                          {row.etlStatus === '狀態異常' && <span className="status-badge status-error">狀態異常</span>}
                        </td>
                        <td style={{ fontSize: '13px', color: row.etlStatus === '狀態異常' ? 'var(--med-alert-text)' : 'var(--text-secondary)' }}>
                          {row.etlReason || '-'}
                        </td>
                        <td style={{ color: 'var(--text-muted)', fontSize: '13px', fontFamily: 'var(--font-mono)' }}>{row.etlTimestamp || 'N/A'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '14px' }}>
                <button 
                  className="btn btn-primary"
                  onClick={() => setActiveStep('step2_etl')}
                >
                  <span>進入階段 02：執行 QMS ETL 數據清洗與產出</span>
                  <span>➔</span>
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ========================================================================= */}
      {/* STAGE 02: QMS ETL CLEANSING & REGULATORY REPORT EXPORT */}
      {/* ========================================================================= */}
      {activeStep === 'step2_etl' && (
        <div className="panel-card">
          <div className="panel-header">
            <div>
              <h2 className="panel-title">
                <Icons.FileText />
                <span>階段 02：QMS ETL 數據清洗與法規報表產出</span>
              </h2>
              <p className="panel-subtitle">嚴格遵循 ISO 13485 品質追溯規範，進行多維度月份自動歸併、同檔後綴去重與防呆過濾</p>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-secondary)' }}>確效年度：</span>
              <select 
                className="filter-select"
                value={etlYear}
                onChange={(e) => {
                  setEtlYear(parseInt(e.target.value, 10));
                  setCachedCounts(null);
                }}
              >
                {Array.from({ length: 31 }, (_, i) => 2010 + i).map(year => (
                  <option key={year} value={year}>{year} 年度</option>
                ))}
              </select>
            </div>
          </div>

          {isProcessingETL && (
            <div className="panel-card" style={{ background: 'var(--bg-surface-subtle)', marginBottom: '14px' }}>
              <div style={{ fontSize: '15px', fontWeight: 600, color: 'var(--med-cobalt)', marginBottom: '4px' }}>
                正在執行 {etlYear} 年度 QMS ETL 運算... ({etlProgress?.current}/{etlProgress?.total})
              </div>
              <div style={{ fontSize: '13px', color: 'var(--text-secondary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontFamily: 'var(--font-mono)' }}>
                FILE: {etlProgress?.filename}
              </div>
              <div className="progress-track" style={{ marginTop: '8px' }}>
                <div className="progress-fill" style={{ width: `${((etlProgress?.current || 0) / (etlProgress?.total || 1)) * 100}%` }}></div>
              </div>
              <button 
                className="btn btn-danger btn-sm"
                onClick={() => {
                  isETLCancelledRef.current = true;
                }}
                style={{ marginTop: '10px' }}
              >
                取消確效
              </button>
            </div>
          )}

          <div className="etl-grid">
            <div className="etl-action-card">
              <div>
                <div className="etl-card-title">
                  <Icons.FileText />
                  <span>QMS 年度統計報表 (JSON)</span>
                </div>
                <p className="etl-card-desc" style={{ marginTop: '4px' }}>
                  輸出機器可讀之標準化 QC 年度統計 JSON 資料結構，包含 7 大類品檢與每月細項數據。
                </p>
              </div>
              <button 
                className="btn btn-primary"
                onClick={() => handleRunBrowserETL(etlYear)}
                disabled={isProcessingETL || isScanning || uploadedFiles.length === 0}
              >
                <Icons.Download />
                <span>輸出 {etlYear} QMS 統計 (JSON)</span>
              </button>
            </div>

            <div className="etl-action-card">
              <div>
                <div className="etl-card-title">
                  <Icons.FileText />
                  <span>法規稽核統計報表 (Excel)</span>
                </div>
                <p className="etl-card-desc" style={{ marginTop: '4px' }}>
                  產出符合 ISO 13485 管理審查格式之 Excel 彙整報表，適合法規主管簽核與存檔。
                </p>
              </div>
              <button 
                className="btn btn-secondary"
                onClick={() => handleExportExcel(etlYear)}
                disabled={isProcessingETL || isScanning || !cachedCounts}
              >
                <Icons.Download />
                <span>下載 {etlYear} Excel 報表</span>
              </button>
            </div>

            <div className="etl-action-card">
              <div>
                <div className="etl-card-title">
                  <Icons.Folder />
                  <span>7 大類獨立 QC 分冊 (Excel)</span>
                </div>
                <p className="etl-card-desc" style={{ marginTop: '4px' }}>
                  按進料、QIP、裝配、半成品、完成品、零組件與出貨分別輸出獨立 Excel 驗證分冊。
                </p>
              </div>
              <button 
                className="btn btn-secondary"
                onClick={() => handleExportIndividualReports(etlYear)}
                disabled={isProcessingETL || isScanning || uploadedFiles.length === 0}
              >
                <Icons.Download />
                <span>輸出 7 大類分冊 (Excel)</span>
              </button>
            </div>

            <div className="etl-action-card" style={{ borderColor: 'var(--med-cobalt-border)', background: 'var(--med-cobalt-light)' }}>
              <div>
                <div className="etl-card-title" style={{ color: 'var(--med-cobalt)' }}>
                  <Icons.Chart />
                  <span>同步至階段 03 儀表板</span>
                </div>
                <p className="etl-card-desc" style={{ marginTop: '4px' }}>
                  將當前 ETL 數據直接同步至可視化儀表板與大模型診斷引擎，免去手動重複上傳。
                </p>
              </div>
              <button 
                className="btn btn-primary"
                onClick={() => handleSyncToDashboard(etlYear)}
                disabled={isProcessingETL || isScanning || uploadedFiles.length === 0}
              >
                <span>立即同步並檢視趨勢 ➔</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* STAGE 03: MULTI-YEAR ANALYTICS DASHBOARD & LLM DIAGNOSTICS */}
      {/* ========================================================================= */}
      {activeStep === 'step3_analytics' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
          {/* Top Control Bar */}
          <div className="panel-card" style={{ padding: '12px 18px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <input 
                  type="file" 
                  accept=".xlsx,.json" 
                  onChange={handleSummaryFileChange} 
                  id="dashboard-file-input" 
                  multiple
                  style={{ display: 'none' }} 
                />
                <button 
                  className="btn btn-primary btn-sm"
                  onClick={() => document.getElementById('dashboard-file-input')?.click()}
                >
                  <Icons.Folder />
                  <span>匯入統計檔 (.json / .xlsx)</span>
                </button>

                {Object.keys(summaryFiles).length > 0 && (
                  <button 
                    className="btn btn-danger btn-sm"
                    onClick={() => {
                      if (window.confirm("確定要清空所有已載入的報表檔案嗎？")) {
                        setSummaryFiles({});
                        setActiveSheet("");
                        setSelectedItems([]);
                        showToast("已清空所有報表");
                      }
                    }}
                  >
                    <Icons.Trash />
                    <span>清空報表</span>
                  </button>
                )}

                {Object.keys(summaryFiles).length > 0 && (
                  <div style={{ display: 'flex', gap: '4px', alignItems: 'center', marginLeft: '6px' }}>
                    <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-secondary)' }}>年度切換：</span>
                    {Object.keys(summaryFiles).sort().map(year => (
                      <button 
                        key={year}
                        className={`btn btn-sm ${activeYear === year ? 'btn-primary' : 'btn-secondary'}`}
                        onClick={() => setActiveYear(year)}
                      >
                        {year} 年
                      </button>
                    ))}
                    <button 
                      className={`btn btn-sm ${activeYear === 'compare' ? 'btn-primary' : 'btn-secondary'}`}
                      onClick={() => setActiveYear('compare')}
                      disabled={Object.keys(summaryFiles).length < 2}
                    >
                      跨年度對比
                    </button>
                  </div>
                )}
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-secondary)' }}>月份篩選：</span>
                <select 
                  className="filter-select"
                  value={selectedMonth}
                  onChange={(e) => setSelectedMonth(parseInt(e.target.value, 10))}
                >
                  <option value={0}>全年度 (1-12月)</option>
                  {MONTH_LABELS.map((label, idx) => (
                    <option key={idx + 1} value={idx + 1}>{label}</option>
                  ))}
                </select>
              </div>
            </div>

            {activeYear === 'compare' && (
              <div style={{ marginTop: '10px', paddingTop: '8px', borderTop: '1px solid var(--border-subtle)', display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-secondary)', fontFamily: 'var(--font-mono)' }}>COMPARE YEARS:</span>
                {Object.keys(summaryFiles).sort().map(year => {
                  const isSelected = compareYearSelection[year];
                  return (
                    <button
                      key={year}
                      className={`btn btn-sm ${isSelected ? 'btn-primary' : 'btn-secondary'}`}
                      onClick={() => {
                        setCompareYearSelection(prev => ({
                          ...prev,
                          [year]: !prev[year]
                        }));
                      }}
                    >
                      {isSelected ? '✓ ' : '+ '} {year} 年
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          {summaryData ? (
            <div className="dashboard-layout">
              {/* Left QC Sheet Selector */}
              <aside className="sheet-sidebar">
                <div style={{ fontSize: '13px', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.6px', padding: '4px 6px', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}>
                  QMS STAGES
                </div>
                {Object.keys(summaryData).map((sheetName) => (
                  <button
                    key={sheetName}
                    className={`sheet-tab-btn ${activeSheet === sheetName ? 'active' : ''}`}
                    onClick={() => setActiveSheet(sheetName)}
                  >
                    <span>{sheetName}</span>
                  </button>
                ))}
              </aside>

              {/* Right Analysis Area */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                {/* Technical KPI Cards */}
                <div className="kpi-grid">
                  <div className="kpi-card">
                    <span className="kpi-label">TOTAL INSPECTIONS (全期檢驗總數)</span>
                    <span className="kpi-val">{dashboardInsights.total.toLocaleString()}</span>
                  </div>
                  <div className="kpi-card">
                    <span className="kpi-label">PEAK PERIOD (最高峰月份)</span>
                    <span className="kpi-val" style={{ color: 'var(--med-cobalt)' }}>{dashboardInsights.peakMonth}</span>
                  </div>
                  <div className="kpi-card">
                    <span className="kpi-label">PEAK VOLUME (單月最高檢驗量)</span>
                    <span className="kpi-val">{dashboardInsights.peakVal.toLocaleString()}</span>
                  </div>
                  <div className="kpi-card">
                    <span className="kpi-label">MONTHLY AVERAGE (月均檢驗負荷)</span>
                    <span className="kpi-val">{dashboardInsights.avg.toLocaleString()}</span>
                  </div>
                </div>

                {/* Subcategory Filter Pills */}
                <div className="panel-card">
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <h3 style={{ fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>自訂品檢項目分析篩選</h3>
                      <p style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>勾選細項組件以動態生成製程堆疊分析柱狀圖</p>
                    </div>
                    {availableItems.length > 0 && (
                      <div style={{ display: 'flex', gap: '6px' }}>
                        <button 
                          className="btn btn-secondary btn-sm"
                          onClick={() => setSelectedItems(availableItems)}
                        >
                          全選
                        </button>
                        <button 
                          className="btn btn-secondary btn-sm"
                          onClick={() => setSelectedItems([])}
                        >
                          清空
                        </button>
                      </div>
                    )}
                  </div>

                  <div className="pills-container">
                    {availableItems.length === 0 ? (
                      <p style={{ color: 'var(--text-muted)', fontSize: '13px' }}>無可選項目</p>
                    ) : (
                      availableItems.map((item) => {
                        const active = effectiveSelectedItems.some(i => i.idx === item.idx);
                        return (
                          <div 
                            key={item.idx} 
                            className={`pill-item ${active ? 'active' : ''}`}
                            onClick={() => {
                              if (active) setSelectedItems(effectiveSelectedItems.filter(i => i.idx !== item.idx));
                              else setSelectedItems([...effectiveSelectedItems, item]);
                            }}
                          >
                            <span className="pill-check"></span>
                            <span>{item.name}</span>
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>

                {/* Stacked Bar Chart */}
                <div className="panel-card">
                  <div className="panel-header">
                    <h3 className="panel-title">
                      <Icons.Chart />
                      <span>{activeSheet} - 品質檢驗趨勢堆疊分析 (QMS Stage Trend Analysis)</span>
                    </h3>
                  </div>
                  <div className="chart-card">
                    <canvas ref={chartRef} />
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="panel-card" style={{ padding: '70px 0', textAlign: 'center' }}>
              <Icons.Chart />
              <h3 style={{ fontSize: '17px', fontWeight: 700, marginTop: '10px' }}>尚未載入任何 QMS 統計數據</h3>
              <p style={{ color: 'var(--text-secondary)', fontSize: '13px', marginTop: '4px' }}>
                請點擊上方「匯入統計檔」載入現有報表，或由「階段 02」一鍵同步 ETL 結果。
              </p>
            </div>
          )}
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODAL: LLM ANALYSIS PROMPT & DATASET EXPORT */}
      {/* ========================================================================= */}
      {showLLMModal && (
        <div className="modal-overlay" onClick={() => setShowLLMModal(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">
                <Icons.Sparkles />
                <span>ISO 13485 醫療器材大模型 (LLM) 診斷指令與全年度數據包</span>
              </div>
              <button className="modal-close-btn" onClick={() => setShowLLMModal(false)}>✕</button>
            </div>

            <div className="modal-body">
              <div style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                此數據包已自動整合系統中載入的所有年份品管數據，依據<strong>ISO 13485 醫療器材主任品質稽核員 (Lead Quality Auditor)</strong>規範格式化為 Markdown 統計表與機器可讀 JSON 數據包，可直接複製餵給 ChatGPT, Claude, Gemini, DeepSeek 執行專業品質深度審查與 CAPA 建議。
              </div>

              <div className="llm-preview-box">
                {llmExportContent}
              </div>
            </div>

            <div className="modal-footer">
              <button className="btn btn-secondary btn-sm" onClick={() => setShowLLMModal(false)}>
                關閉
              </button>
              <button className="btn btn-secondary btn-sm" onClick={handleDownloadPrompt}>
                <Icons.Download />
                <span>下載 .md 指令檔</span>
              </button>
              <button className="btn btn-purple btn-sm" onClick={handleCopyPrompt}>
                <Icons.Copy />
                <span>一鍵複製 Prompt</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODAL: QC CODE MAPPINGS MANAGER */}
      {/* ========================================================================= */}
      {showMappingsModal && (
        <div className="modal-overlay" onClick={() => setShowMappingsModal(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()} style={{ maxWidth: '640px' }}>
            <div className="modal-header">
              <div className="modal-title">
                <Icons.Settings />
                <span>表單編碼與名稱對照表 (QC Code Mappings)</span>
              </div>
              <button className="modal-close-btn" onClick={() => setShowMappingsModal(false)}>✕</button>
            </div>

            <div className="modal-body">
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: '8px' }}>
                <button className="btn btn-secondary btn-sm" onClick={handleExportMappings}>
                  <Icons.Download />
                  <span>匯出 JSON</span>
                </button>
                <label className="btn btn-secondary btn-sm" style={{ cursor: 'pointer' }}>
                  <Icons.Folder />
                  <span>匯入 JSON</span>
                  <input type="file" accept=".json" onChange={handleImportMappings} style={{ display: 'none' }} />
                </label>
                <button className="btn btn-danger btn-sm" onClick={handleResetMappings}>
                  <Icons.Refresh />
                  <span>恢復系統預設</span>
                </button>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '300px', overflowY: 'auto' }}>
                {Object.keys(mappings).map((code) => (
                  <div key={code} style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                    <span style={{ width: '120px', fontFamily: 'var(--font-mono)', fontWeight: 700, fontSize: '13px', color: 'var(--med-cobalt)' }}>{code}</span>
                    <input 
                      type="text" 
                      className="search-input" 
                      style={{ flex: 1 }}
                      value={mappings[code]} 
                      onChange={(e) => {
                        const updated = { ...mappings, [code]: e.target.value };
                        updateMappings(updated);
                      }} 
                    />
                    <button 
                      className="btn btn-danger btn-sm" 
                      onClick={() => handleDeleteMapping(code)}
                      title="刪除對照"
                    >
                      ✕
                    </button>
                  </div>
                ))}
              </div>

              <form onSubmit={handleAddMapping} style={{ borderTop: '1px solid var(--border-subtle)', paddingTop: '10px', display: 'flex', gap: '8px' }}>
                <input 
                  type="text" 
                  className="search-input" 
                  placeholder="QC10006-R02" 
                  value={newCode} 
                  onChange={(e) => setNewCode(e.target.value)} 
                  style={{ width: '120px' }}
                />
                <input 
                  type="text" 
                  className="search-input" 
                  placeholder="對應表單名稱" 
                  value={newName} 
                  onChange={(e) => setNewName(e.target.value)} 
                  style={{ flex: 1 }}
                />
                <button type="submit" className="btn btn-primary btn-sm">
                  新增
                </button>
              </form>
            </div>

            <div className="modal-footer">
              <button className="btn btn-primary btn-sm" onClick={() => setShowMappingsModal(false)}>
                完成
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer style={{
        marginTop: 'auto',
        paddingTop: '14px',
        borderTop: '1px solid var(--border-precision)',
        textAlign: 'center',
        color: 'var(--text-muted)',
        fontSize: '13px',
        fontFamily: 'var(--font-mono)'
      }}>
        MOULDEX MEDICAL DEVICE QUALITY MANAGEMENT SYSTEM · ISO 13485 & GMP COMPLIANT · v2026.08
      </footer>
    </div>
  );
}

export default App;

const ColumnFilterPopover = ({ 
  fieldKey, 
  allValues, 
  columnFilters, 
  setColumnFilters, 
  onClose, 
  getStatusLabel 
}) => {
  const [searchVal, setSearchVal] = useState("");
  const currentFilter = columnFilters[fieldKey] || [];
  
  const [tempSelected, setTempSelected] = useState(
    currentFilter.length === 0 ? [...allValues] : [...currentFilter]
  );

  const filteredValues = allValues.filter(val => {
    const displayLabel = fieldKey === 'status' ? getStatusLabel(val) : String(val);
    return displayLabel.toLowerCase().includes(searchVal.toLowerCase());
  });

  const handleToggleValue = (val) => {
    if (tempSelected.includes(val)) {
      setTempSelected(tempSelected.filter(v => v !== val));
    } else {
      setTempSelected([...tempSelected, val]);
    }
  };

  const handleToggleAll = () => {
    const allVisibleSelected = filteredValues.every(val => tempSelected.includes(val));
    if (allVisibleSelected) {
      setTempSelected(tempSelected.filter(val => !filteredValues.includes(val)));
    } else {
      const newSelected = new Set([...tempSelected, ...filteredValues]);
      setTempSelected(Array.from(newSelected));
    }
  };

  const handleApply = () => {
    if (tempSelected.length === allValues.length || tempSelected.length === 0) {
      setColumnFilters(prev => ({ ...prev, [fieldKey]: [] }));
    } else {
      setColumnFilters(prev => ({ ...prev, [fieldKey]: tempSelected }));
    }
    onClose();
  };

  const handleClear = () => {
    setColumnFilters(prev => ({ ...prev, [fieldKey]: [] }));
    onClose();
  };

  return (
    <div className="filter-popover">
      <div className="filter-popover-search">
        <input 
          type="text" 
          placeholder="搜尋篩選值..." 
          value={searchVal}
          onChange={(e) => setSearchVal(e.target.value)}
          className="filter-popover-input"
          autoFocus
        />
      </div>
      
      <div className="filter-popover-list">
        <label className="filter-popover-item select-all">
          <input 
            type="checkbox" 
            checked={filteredValues.length > 0 && filteredValues.every(val => tempSelected.includes(val))}
            ref={el => {
              if (el) {
                const someSelected = filteredValues.some(val => tempSelected.includes(val));
                const allSelected = filteredValues.every(val => tempSelected.includes(val));
                el.indeterminate = someSelected && !allSelected;
              }
            }}
            onChange={handleToggleAll}
          />
          <span style={{ fontWeight: 'bold' }}>(全選)</span>
        </label>
        
        {filteredValues.map(val => {
          const displayLabel = fieldKey === 'status' ? getStatusLabel(val) : String(val);
          return (
            <label key={val} className="filter-popover-item">
              <input 
                type="checkbox" 
                checked={tempSelected.includes(val)}
                onChange={() => handleToggleValue(val)}
              />
              <span title={displayLabel}>{displayLabel}</span>
            </label>
          );
        })}
      </div>
      
      <div className="filter-popover-actions">
        <button className="btn btn-secondary filter-popover-btn" onClick={handleClear}>
          清除
        </button>
        <div style={{ display: 'flex', gap: '4px' }}>
          <button className="btn btn-secondary filter-popover-btn" onClick={onClose}>
            取消
          </button>
          <button className="btn btn-primary filter-popover-btn" onClick={handleApply}>
            套用
          </button>
        </div>
      </div>
    </div>
  );
};

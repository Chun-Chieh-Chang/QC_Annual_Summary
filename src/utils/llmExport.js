/**
 * Utility for exporting multi-year QC summary data as an LLM-ready analysis prompt and structured dataset.
 * Specially tailored for Medical Device (ISO 13485 & GMP) Quality Management Systems (QMS).
 */

const MONTH_LABELS = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];

/**
 * Parses raw summaryFiles (2D array rows per sheet per year) into a clean structured multi-year object.
 * @param {Object} summaryFiles - Format: { [year]: { [sheetName]: rows } }
 * @returns {Object} Structured data: { years: [], categories: { [sheetName]: { columns: [], yearlyData: { [year]: { monthly: {}, total: 0 } } } }, overall: {} }
 */
export function extractMultiYearDataset(summaryFiles) {
  const years = Object.keys(summaryFiles || {}).filter(y => y !== 'compare').sort();
  if (years.length === 0) return null;

  const allSheets = new Set();
  years.forEach(y => {
    const data = summaryFiles[y] || {};
    Object.keys(data).forEach(sheet => {
      if (sheet !== '品檢地圖' && !sheet.startsWith('_')) {
        allSheets.add(sheet);
      }
    });
  });

  const categories = {};
  const overallSummary = {};

  years.forEach(y => {
    overallSummary[y] = { total: 0, bySheet: {} };
  });

  allSheets.forEach(sheetName => {
    categories[sheetName] = {
      sheetName,
      columns: [],
      yearly: {}
    };

    const colSet = new Set();
    years.forEach(y => {
      const rows = summaryFiles[y]?.[sheetName];
      if (rows && rows.length >= 2) {
        const headerRow = rows[1] || [];
        headerRow.forEach(h => {
          const name = String(h || '').trim();
          if (name && name !== '月份' && name !== '小計' && name !== 'NCA') {
            colSet.add(name);
          }
        });
      }
    });

    const columns = Array.from(colSet);
    categories[sheetName].columns = columns;

    years.forEach(y => {
      const rows = summaryFiles[y]?.[sheetName];
      const monthlyData = {};
      let sheetYearTotal = 0;

      for (let m = 1; m <= 12; m++) {
        monthlyData[m] = { total: 0 };
        columns.forEach(col => { monthlyData[m][col] = 0; });
      }

      if (rows && rows.length >= 2) {
        const headerRow = rows[1] || [];
        const colIndices = {};
        columns.forEach(col => {
          colIndices[col] = headerRow.findIndex(h => String(h || '').trim() === col);
        });

        for (let m = 1; m <= 12; m++) {
          const row = rows[m + 1] || [];
          let monthSum = 0;
          columns.forEach(col => {
            const idx = colIndices[col];
            const val = idx >= 0 && row[idx] !== undefined ? Number(row[idx]) || 0 : 0;
            monthlyData[m][col] = val;
            monthSum += val;
          });
          monthlyData[m].total = monthSum;
          sheetYearTotal += monthSum;
        }
      }

      categories[sheetName].yearly[y] = {
        monthly: monthlyData,
        total: sheetYearTotal
      };

      overallSummary[y].bySheet[sheetName] = sheetYearTotal;
      overallSummary[y].total += sheetYearTotal;
    });
  });

  return {
    years,
    categories,
    overallSummary
  };
}

/**
 * Generates the complete LLM Prompt & Data document in Markdown tailored for Medical Device QMS analysis.
 * @param {Object} summaryFiles - Format: { [year]: { [sheetName]: rows } }
 * @returns {string} Markdown text ready for LLM consumption
 */
export function generateMultiYearLLMPrompt(summaryFiles) {
  const dataset = extractMultiYearDataset(summaryFiles);
  if (!dataset) {
    return `# 錯誤：尚未載入任何品檢年度報表數據\n\n請先在系統中載入至少一個年份的品檢報表統計檔 (.xlsx / .json)。`;
  }

  const { years, categories, overallSummary } = dataset;
  const generatedDate = new Date().toISOString().split('T')[0];

  let md = '';

  // 1. Medical Device QMS Expert Role & Task Header
  md += `# 醫療器材品保檢驗數據跨年度綜合診斷報告指令 (Medical Device QMS Analytics Prompt)\n\n`;
  md += `> **體系依據**：ISO 13485 醫療器材品質管理系統 & GMP 醫療器材優良製造規範\n`;
  md += `> **數據來源**：Mouldex 醫療器材品保檢驗自動化 ETL Pipeline\n`;
  md += `> **涵蓋年度**：${years.join(', ')} 年（共 ${years.length} 個年度）\n`;
  md += `> **生成時間**：${generatedDate}\n\n`;

  md += `## 任務目標與分析指令 (Task Objectives & Lead Auditor Prompt)\n\n`;
  md += `你現在是一位具備 20 年以上高階醫療器材製造與法規品質經驗的**「ISO 13485 / FDA QSR 主任品質稽核員 (Lead Quality Auditor)」**與**「醫療器材精實品質架構師」**。\n`;
  md += `請依據下方由品保系統產出的**跨年度醫療器材品檢統計數據**（包含 7 大製程品檢階段、12 個月份分佈、細項組件與 Setup/巡檢頻次），進行全方位的品質風險評估與管理審查報告。請重點輸出以下 5 大核心維度：\n\n`;

  md += `1. **跨年度檢驗負荷與品質趨勢 (Multi-Year Inspection & Capacity Trends)**：\n`;
  md += `   - 分析總檢驗批數/次數在各年度間的變化（YoY 成長或衰退率）。\n`;
  md += `   - 評估各關鍵製程工序（進料、射出/押出、裝配巡檢、半成品、完成品、零組件入庫、出貨）檢驗負荷變化與產能匹配度。\n\n`;

  md += `2. **季節性與月度品質波動風險 (Seasonality & Monthly Anomaly Evaluation)**：\n`;
  md += `   - 辨識各月檢驗數是否存在異常峰值或銳減月份（如產線換模、新產品導入、旺季出貨衝刺）。\n`;
  md += `   - 評估檢驗高峰期是否可能引發檢驗人力不足導致漏檢風險，並提出預防性調度策略。\n\n`;

  md += `3. **高風險工序穩定度與 Setup / 巡檢控制結構 (Injection & Extrusion Process Control)**：\n`;
  md += `   - 深入分析 QIP (QC10004-R02) 射出與押出之「Setup (開模調校/首件檢驗)」與「Patrol (製程巡檢)」比例關係。\n`;
  md += `   - 評估調機頻率與巡檢密度是否符合醫材關鍵製程控制要求，巡檢抽樣頻率是否足以保證製程能力穩定 (Cpk/Ppk)。\n\n`;

  md += `4. **供應鏈與關鍵組件品質漏斗 (Medical Supply Chain & Component Risk Pareto)**：\n`;
  md += `   - 分析進料品檢 (QC10002-R02)、零組件入庫 (QC10007-R03) 至裝配完成品 (QC10007-R01) 之檢驗分佈漏斗。\n`;
  md += `   - 依據 80/20 原則（Pareto Analysis）標註出檢驗量佔比最高的關鍵核心品項與組件，給出供應商品質保證 (SQA) 與進料免檢/加嚴抽樣建議。\n\n`;

  md += `5. **CAPA 矯正與預防措施建議 (Actionable ISO 13485 CAPA Roadmap)**：\n`;
  md += `   - 提出可落地之具體矯正與預防措施 (CAPA)，涵蓋檢驗規範優化、防呆措施 (Poka-Yoke)、首件檢驗確效 (First Article Inspection) 及品檢自動化路徑。\n\n`;

  md += `---\n\n`;

  // 2. Multi-Year High-level Summary Table
  md += `## 1. 跨年度醫療器材品檢總量彙整總表 (Multi-Year QC Summary Table)\n\n`;
  md += `| 醫療器材品檢階段 (QMS Stage) | ` + years.map(y => `${y} 年檢驗總量`).join(' | ') + ` | 法規控制重點 |\n`;
  md += `|---|` + years.map(() => '---:').join('|') + `|---|\n`;

  Object.keys(categories).forEach(sheetName => {
    const rowValues = years.map(y => {
      const total = categories[sheetName].yearly[y]?.total || 0;
      return total.toLocaleString();
    });
    md += `| **${sheetName}** | ${rowValues.join(' | ')} | ISO 13485 關鍵檢驗點 |\n`;
  });

  const totalRow = years.map(y => (overallSummary[y]?.total || 0).toLocaleString());
  md += `| **全廠醫材品檢總計 (Grand Total)** | **${totalRow.join('** | **')}** | **全流程受控** |\n\n`;

  md += `---\n\n`;

  // 3. Detailed Monthly Tables for each QC Category
  md += `## 2. 各品檢階段月度數據明細 (Detailed Monthly Stage Breakdown)\n\n`;

  Object.keys(categories).forEach((sheetName, sIdx) => {
    const sheetInfo = categories[sheetName];
    md += `### 2.${sIdx + 1} ${sheetName}\n\n`;

    years.forEach(y => {
      const yData = sheetInfo.yearly[y];
      if (!yData) return;

      md += `#### 📅 ${y} 年度月度檢驗明細\n\n`;
      
      const headers = ['月份', ...sheetInfo.columns, '小計'];
      md += `| ` + headers.join(' | ') + ` |\n`;
      md += `| ` + headers.map((h, i) => i === 0 ? '---' : '---:').join(' | ') + ` |\n`;

      for (let m = 1; m <= 12; m++) {
        const mRow = [MONTH_LABELS[m - 1]];
        sheetInfo.columns.forEach(col => {
          mRow.push(yData.monthly[m]?.[col] ?? 0);
        });
        mRow.push(yData.monthly[m]?.total ?? 0);
        md += `| ` + mRow.join(' | ') + ` |\n`;
      }

      const catTotalRow = ['**小計**'];
      sheetInfo.columns.forEach(col => {
        let colSum = 0;
        for (let m = 1; m <= 12; m++) {
          colSum += (yData.monthly[m]?.[col] || 0);
        }
        catTotalRow.push(`**${colSum}**`);
      });
      catTotalRow.push(`**${yData.total}**`);
      md += `| ` + catTotalRow.join(' | ') + ` |\n\n`;
    });
  });

  md += `---\n\n`;

  // 4. Raw JSON Payload
  md += `## 3. 機器可讀結構化 QMS 數據包 (Raw JSON QMS Dataset Payload)\n\n`;
  md += `\`\`\`json\n`;
  md += JSON.stringify({
    metadata: {
      standard: "ISO 13485:2016 & GMP Compliant Format",
      generator: "Mouldex Medical Device QMS Pipeline v2026",
      generatedAt: new Date().toISOString(),
      coveredYears: years,
      totalRecordsAcrossYears: Object.values(overallSummary).reduce((acc, cur) => acc + cur.total, 0)
    },
    overallSummary,
    detailedCategories: categories
  }, null, 2);
  md += `\n\`\`\`\n`;

  return md;
}

/**
 * Downloads a text file (e.g. .md or .json) in the browser.
 */
export function downloadFile(content, fileName, mimeType = 'text/markdown;charset=utf-8') {
  const blob = new Blob([content], { type: mimeType });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', fileName);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Copies text to system clipboard.
 */
export async function copyTextToClipboard(text) {
  if (navigator.clipboard && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      console.warn("Clipboard API failed, falling back to execCommand", e);
    }
  }

  const textArea = document.createElement("textarea");
  textArea.value = text;
  textArea.style.position = "fixed";
  textArea.style.left = "-999999px";
  textArea.style.top = "-999999px";
  document.body.appendChild(textArea);
  textArea.focus();
  textArea.select();
  let successful = false;
  try {
    successful = document.execCommand('copy');
  } catch (err) {
    console.error('execCommand copy error', err);
  }
  document.body.removeChild(textArea);
  return successful;
}

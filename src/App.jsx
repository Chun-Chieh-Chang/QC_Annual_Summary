import React, { useState, useEffect, useRef } from 'react';
import { getMappings, saveMappings, resetMappings } from './utils/db';
import { parseExcelFile, parseSummaryExcel, parseSummaryJSON } from './utils/excelParser';
import { runETLInBrowser, exportSummaryExcelInBrowser, exportSummaryJSONInBrowser, isUUID } from './utils/browserETL';
import * as XLSX from 'xlsx';

const MONTH_LABELS = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];

const MCK_COLORS = [
  '#007AFF', // Vivid Blue
  '#FF3B30', // Vivid Red
  '#34C759', // Vivid Green
  '#FF9500', // Vivid Orange
  '#AF52DE', // Vivid Purple
  '#5AC8FA', // Vivid Cyan
  '#FF2D55', // Vivid Pink
  '#5856D6', // Vivid Indigo
  '#FFCC00', // Vivid Yellow
];

function App() {
  // Navigation State
  const [activeTab, setActiveTab] = useState("dashboard"); // "dashboard" or "extractor"

  // ==========================================
  // STATE FOR TAB 1: McKinsey Dashboard
  // ==========================================
  const [summaryFiles, setSummaryFiles] = useState({}); // { "2025": data2025, "2026": data2026 }
  const [activeYear, setActiveYear] = useState("2025"); // "2025", "2026", or "compare"
  const [summaryData, setSummaryData] = useState(null);
  const [activeSheet, setActiveSheet] = useState("");
  const [selectedItems, setSelectedItems] = useState([]); // Array of {name, idx}
  const [summaryFileName, setSummaryFileName] = useState("");
  const [dashboardInsights, setDashboardInsights] = useState({ total: 0, peakMonth: "", peakVal: 0, avg: 0 });
  const [selectedMonth, setSelectedMonth] = useState(0); // 0=all, 1-12=specific month
  const [compareYearSelection, setCompareYearSelection] = useState({}); // { "2023": true, "2024": false, ... }

  const chartRef = useRef(null);
  const chartInstance = useRef(null);
  const isETLCancelledRef = useRef(false);

  // ==========================================
  // STATE FOR TAB 2: Mappings & Extractor
  // ==========================================
  const [mappings, setMappings] = useState({});
  const [scannedRows, setScannedRows] = useState([]);
  const [isScanning, setIsScanning] = useState(false);
  const [folderName, setFolderName] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [newCode, setNewCode] = useState("");
  const [newName, setNewName] = useState("");
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [etlProgress, setEtlProgress] = useState(null);
  const [scanProgress, setScanProgress] = useState(null);
  const [isProcessingETL, setIsProcessingETL] = useState(false);
  const [etlYear, setEtlYear] = useState(2025);
  const [isETLCancelled, setIsETLCancelled] = useState(false);
  const [cachedCounts, setCachedCounts] = useState(null);
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

  // 當 scannedRows 改變時重設過濾器
  useEffect(() => {
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
  }, [scannedRows]);

  const folderInputRef = useRef(null);
  const leftPanelRef = useRef(null);
  const [leftPanelHeight, setLeftPanelHeight] = useState(null);

  useEffect(() => {
    if (!leftPanelRef.current) return;
    
    const resizeObserver = new ResizeObserver((entries) => {
      for (let entry of entries) {
        setLeftPanelHeight(entry.target.getBoundingClientRect().height);
      }
    });
    
    resizeObserver.observe(leftPanelRef.current);
    return () => resizeObserver.disconnect();
  }, [scannedRows, uploadedFiles]);
  const fileInputRef = useRef(null);

  // Load mappings on mount
  useEffect(() => {
    setMappings(getMappings());
    
    // Inject mock data if ?mock=true is present in URL query
    const params = new URLSearchParams(window.location.search);
    if (params.get('mock') === 'true') {
      setScannedRows([
        {
          fileName: "裝配C-2026A.xlsx",
          filePath: "RawData/2026/零組件入庫-2026/裝配C-2026",
          sheetName: "260102",
          foundCode: "QC10006-R02",
          foundName: "半成品品檢表",
          status: "matched",
          etlStatus: "已納入",
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
          etlTimestamp: "2026-07-06 14:20:00"
        },
        {
          fileName: "樣板測試檔.xlsx",
          filePath: "RawData/Test",
          sheetName: "Sheet1",
          foundCode: "QC99999-R99",
          foundName: "無對照編碼",
          status: "unmatched",
          etlStatus: "狀態異常",
          etlTimestamp: "2026-07-06 14:25:00"
        }
      ]);
    }
  }, []);

  // Update mappings helper
  const updateMappings = (newMappings) => {
    setMappings(newMappings);
    saveMappings(newMappings);
  };

  // Sync localStorage
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
  };

  const handleDeleteMapping = (code) => {
    const updated = { ...mappings };
    delete updated[code];
    updateMappings(updated);
  };

  const handleResetMappings = () => {
    if (window.confirm("確定要將對照表恢復為系統預設值嗎？自訂對照將會遺失。")) {
      const reset = resetMappings();
      setMappings(reset);
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
        alert("匯入對照表成功！");
      } catch (err) {
        alert("檔案格式錯誤，請確保匯入正確的 JSON 檔案！");
      }
    };
    reader.readAsText(file);
    e.target.value = null;
  };

  const handleExportMappings = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(mappings, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", "qc_mappings_export.json");
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
  };

  // Process folder extraction
  const processFilesList = async (files, nameOfFolder) => {
    if (files.length === 0) return;
    setUploadedFiles(files);
    setCachedCounts(null);
    setIsScanning(true);
    setFolderName(nameOfFolder || "個別檔案統計");
    
    // Auto-detect year (between 2010 and 2040)
    let detected = null;
    const folderMatch = (nameOfFolder || "").match(/20\d{2}/);
    if (folderMatch) {
      const parsed = parseInt(folderMatch[0], 10);
      if (parsed >= 2010 && parsed <= 2040) {
        detected = parsed;
      }
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
      // ponytail: yield to event loop between batches to prevent UI freeze / browser hang
      await new Promise(r => setTimeout(r, 0));
    }
    
    setScannedRows(results);
    setScanProgress(null);
    setIsScanning(false);
  };

  const handleRunBrowserETL = async (year) => {
    if (!uploadedFiles || uploadedFiles.length === 0) return;
    
    if (cachedCounts && cachedCounts.year === year && cachedCounts.filesCount === uploadedFiles.length) {
      exportSummaryJSONInBrowser(cachedCounts.data, year);
      alert(`🎉 ${year} 年度品檢報表統計.json 已下載！`);
      return;
    }

    setIsProcessingETL(true);
    setIsETLCancelled(false);
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
      alert(`🎉 ${year} 年度品檢報表統計.json 已下載！（亦可點選「📥 下載 Excel 報表」按鈕取得傳統 Excel 格式）`);
    } catch (e) {
      if (e.message === 'ETL cancelled by user') {
        console.log('ETL cancelled');
      } else {
        console.error(e);
        alert("ETL 運算失敗，請確保選取的是正確的年度品檢原始資料夾。");
      }
    } finally {
      setIsProcessingETL(false);
      setIsETLCancelled(false);
    }
  };

  // Re-export as Excel (legacy format for human reading)
  const handleExportExcel = (year) => {
    if (!cachedCounts || cachedCounts.year !== year) {
      alert("請先執行一次「輸出 JSON 統計」以產生快取資料。");
      return;
    }
    exportSummaryExcelInBrowser(cachedCounts.data, year);
  };

  // Export individual QC report sheets
  const handleExportIndividualReports = (year) => {
    if (!uploadedFiles || uploadedFiles.length === 0) {
      alert("請先選取或拖入原始品檢資料夾以執行 ETL。");
      return;
    }

    if (cachedCounts && cachedCounts.year === year && cachedCounts.filesCount === uploadedFiles.length) {
      (async () => {
        try {
          const outputPath = await promptExportDirectory();
          if (!outputPath) return;
          await exportIndividualReports(cachedCounts.data, year, outputPath);
          alert(`🎉 ${year} 年度獨立報表全部匯出成功！\n已儲存至您選擇的資料夾。`);
        } catch (e) {
          console.error(e);
          alert("匯出失敗，請確認選取的是正確的目標資料夾。");
        }
      })();
      return;
    }

    setIsProcessingETL(true);
    setIsETLCancelled(false);
    isETLCancelledRef.current = false;
    setEtlProgress({ current: 0, total: uploadedFiles.length, filename: "初始化中..." });
    
    (async () => {
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

        // Ask user for output directory
        const outputPath = await promptExportDirectory();
        if (!outputPath) return;
        await exportIndividualReports(counts, year, outputPath);
        alert(`🎉 ${year} 年度獨立報表全部匯出成功！\n已儲存至您選擇的資料夾。`);
      } catch (e) {
        if (e.message !== 'ETL cancelled by user') {
          console.error(e);
          alert("匯出失敗，請確認選取的是正確的年度品檢原始資料夾。");
        }
      } finally {
        setIsProcessingETL(false);
        setIsETLCancelled(false);
      }
    })();
  };

  // Prompt user to select export directory
  const promptExportDirectory = async () => {
    // Try File System Access API first (Chrome/Edge)
    if (window.showDirectoryPicker) {
      try {
        const handle = await window.showDirectoryPicker();
        return handle;
      } catch (e) {
        if (e.name === 'AbortError') return null; // User cancelled
        console.error("Directory picker error:", e);
      }
    }
    // Fallback: return null to use default download location
    return null;
  };

  // Helper: export individual QC reports as separate Excel files
  const exportIndividualReports = async (counts, year, outputDirHandle = null) => {
    const MONTHS_ARR = ["1","2","3","4","5","6","7","8","9","10","11","12"];
    
    const baseReports = [
      { qcCode: 'QC10002-R02', title: '原物料/配件進料品檢', namePrefix: '進料檢驗' },
      { qcCode: 'QC10004-R02', title: 'QIP尺寸檢驗', namePrefix: 'QIP尺寸檢驗', fixed: ['QIP-Setup','QIP-Patrol','押出-Setup','押出-Patrol'] },
      { qcCode: 'QC10006-R01', title: '裝配對樣巡檢', namePrefix: '裝配巡檢' },
      { qcCode: 'QC10006-R02', title: '半成品品檢', namePrefix: '裝配檢驗' },
      { qcCode: 'QC10007-R01', title: '完成品品檢', namePrefix: '完成品品檢' },
      { qcCode: 'QC10007-R03', title: '零組件入庫品檢', namePrefix: '零組件入庫', splitSubCats: true }, // Split subcats into separate files
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
        } catch (err) {
          console.error(`Failed to write file ${fileName} directly to directory:`, err);
          XLSX.writeFile(wb, fileName);
        }
      } else {
        XLSX.writeFile(wb, fileName);
      }
    }
  };

  const handleFolderChange = (e) => {
    if (e.target.files.length === 0) return;
    const firstFile = e.target.files[0];
    const path = firstFile.webkitRelativePath || "";
    const parts = path.split(/[\\/]/).filter(p => !isUUID(p));
    const folder = parts[0] || "資料夾檔案統計";
    processFilesList(e.target.files, folder);
  };

  const handleFilesChange = (e) => {
    processFilesList(e.target.files, "個別選取檔案");
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    if (e.dataTransfer.files) {
      processFilesList(e.dataTransfer.files, "拖曳檔案統計");
    }
  };

  /**
   * 匯出提取結果為 CSV 檔案
   */
  const exportToCSV = (data, name = "提取結果") => {
    if (!data || data.length === 0) return;
    
    const headers = ["檔案名稱", "檔案路徑", "工作表名稱", "表單編碼", "表單對照名稱", "狀態", "是否納入ETL計算", "原因說明", "更新時間"];
    const keys = ["fileName", "filePath", "sheetName", "foundCode", "foundName", "status", "etlStatus", "etlReason", "etlTimestamp"];
    
    const csvRows = [];
    csvRows.push(headers.join(","));
    
    for (const row of data) {
      const values = keys.map(key => {
        let val = row[key] || "";
        if (key === 'status') {
          val = getStatusLabel(val);
        }
        const escaped = String(val).replace(/"/g, '""');
        if (escaped.includes(",") || escaped.includes("\n") || escaped.includes('"')) {
          return `"${escaped}"`;
        }
        return escaped;
      });
      csvRows.push(values.join(","));
    }
    
    const csvContent = "\uFEFF" + csvRows.join("\n");
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    
    const link = document.createElement("a");
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", `${name}_提取結果.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  // ==========================================
  // MCKINSEY SUMMARY EXCEL PARSER HANDLERS
  // ==========================================
  const handleLoadSummaryFile = async (file) => {
    try {
      const isJSON = file.name.endsWith('.json');
      const data = isJSON ? await parseSummaryJSON(file) : await parseSummaryExcel(file);
      const yearMatch = isJSON
        ? (data._meta?.year ? String(data._meta.year) : null)
        : (file.name.match(/(20\d{2})/)?.[1] || null);
      const year = yearMatch || "未知年度";
      // Strip internal meta key before storing
      if (data._meta) delete data._meta;
      setSummaryFiles(prev => {
        const next = { ...prev, [year]: data };
        const sheets = Object.keys(data);
        const initialSheet = sheets.find(s => s !== "品檢地圖") || sheets[0] || "";
        setActiveSheet(prevSheet => prevSheet || initialSheet);
        return next;
      });
      setActiveYear(year);
      setSummaryFileName(file.name);
    } catch (err) {
      console.error(err);
      alert(`解析品檢彙總 ${file.name.endsWith('.json') ? 'JSON' : 'Excel'} 失敗，請確保上傳的是正確的報表統計檔。`);
    }
  };

  const handleSummaryFileChange = (e) => {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
      Promise.all(files.map(file => handleLoadSummaryFile(file))).then(() => {
        // After loading all files, set the active year to the most recently loaded one
        const years = Object.keys(summaryFiles);
        if (years.length > 0) {
          setActiveYear(years[years.length - 1]);
        }
      });
    }
    // Reset input so same files can be re-selected
    e.target.value = null;
  };

  // Sync summaryData when activeYear or summaryFiles changes
  useEffect(() => {
    if (activeYear === "compare") {
      const years = Object.keys(summaryFiles).sort();
      setSummaryData(summaryFiles[years[0]] || null);
    } else {
      setSummaryData(summaryFiles[activeYear]);
    }
  }, [activeYear, summaryFiles]);

  // Handle active sheet changing
  useEffect(() => {
    if (!summaryData || !activeSheet) return;
    
    const rows = summaryData[activeSheet];
    if (!rows || rows.length < 2) {
      setSelectedItems([]);
      return;
    }
    
    // Row 1 contains column headers
    const headerRow = rows[1] || [];
    
    // Exclude columns like Month (月份), Grand Total (小計), NCA or empty spacers
    const available = [];
    headerRow.forEach((h, idx) => {
      const name = String(h || '').trim();
      if (name && name !== '月份' && name !== '小計' && name !== 'NCA') {
        available.push({ name, idx });
      }
    });
    
    // Auto select first 3 items by default to initialize the plot beautifully
    setSelectedItems(available.slice(0, 3));
  }, [activeSheet, summaryData]);

  // Handle plotting Chart.js
  useEffect(() => {
    if (!summaryData || !activeSheet || selectedItems.length === 0 || !chartRef.current) {
      if (chartInstance.current) {
        chartInstance.current.destroy();
        chartInstance.current = null;
      }
      return;
    }

    const rows = summaryData[activeSheet];
    const ctx = chartRef.current.getContext('2d');

    // Create datasets
    let datasets = [];
    const allYears = Object.keys(summaryFiles).sort();
    const selectedCompareYears = activeYear === 'compare'
      ? Object.keys(compareYearSelection).filter(k => compareYearSelection[k]).sort()
      : [];
    const years = selectedCompareYears.length > 0 ? selectedCompareYears : allYears;
    
    if (activeYear === 'compare' || (selectedMonth > 0 && allYears.length > 1)) {
      // Cross-year mode: x-axis = selected years, each item = 1 bar per year
      datasets = selectedItems.map((item, idx) => {
        const dataPoints = years.map((yr) => {
          const rowsYr = summaryFiles[yr] ? summaryFiles[yr][activeSheet] : null;
          let sum = 0;
          let targetIdx = -1;
          
          if (rowsYr && rowsYr.length >= 2 && rowsYr[1]) {
            targetIdx = rowsYr[1].findIndex(h => String(h || '').trim() === item.name);
          }
          
          if (targetIdx !== -1) {
            for (let m = 2; m <= 13; m++) {
              // Respect month filter
              if (selectedMonth > 0 && m - 1 !== selectedMonth) continue;
              const val = rowsYr[m] ? Number(rowsYr[m][targetIdx]) || 0 : 0;
              sum += val;
            }
          }
          return sum;
        });
        return {
          label: item.name,
          data: dataPoints,
          borderColor: MCK_COLORS[idx % MCK_COLORS.length],
          backgroundColor: MCK_COLORS[idx % MCK_COLORS.length] + 'CC',
          borderWidth: 1,
          borderRadius: 4,
        };
      });
    } else {
      // Single year mode: x-axis = months
      datasets = selectedItems.map((item, idx) => {
        // Monthly data is at rows 2 to 13 (index 2-13)
        const dataPoints = [];
        for (let m = 2; m <= 13; m++) {
          // Respect month filter
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
          borderColor: MCK_COLORS[idx % MCK_COLORS.length],
          backgroundColor: MCK_COLORS[idx % MCK_COLORS.length] + 'CC', // 80% opacity fill
          borderWidth: 1,
          borderRadius: 4, // beautiful rounded top corners
        };
      });
    }

    if (chartInstance.current) {
      chartInstance.current.destroy();
    }

    // McK style: Minimal grid, clean legend and title
    const chartLabels = (activeYear === 'compare' || (selectedMonth > 0 && years.length > 1))
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
              font: {
                family: 'Outfit, sans-serif',
                size: 12
              },
              usePointStyle: true,
              boxWidth: 8
            }
          },
          tooltip: {
            padding: 12,
            backgroundColor: '#0A192F',
            titleFont: { family: 'Outfit, sans-serif', weight: 'bold' },
            bodyFont: { family: 'Outfit, sans-serif' }
          }
        },
        scales: {
          x: {
            stacked: activeYear !== 'compare',
            grid: {
              display: false // McKinsey style hides X vertical gridlines
            },
            ticks: {
              color: '#5A6A85',
              font: { family: 'Outfit, sans-serif', size: 11 }
            }
          },
          y: {
            stacked: activeYear !== 'compare',
            beginAtZero: true,
            grid: {
              color: '#F1F5F9' // light grey thin gridlines
            },
            ticks: {
              color: '#5A6A85',
              font: { family: 'Outfit, sans-serif', size: 11 }
            }
          }
        }
      }
    });

    // Calculate executive insights
    let totalSum = 0;
    let peakValue = 0;
    let peakM = "";
    
    // Sum month by month across all years (or single year)
    for (let m = 0; m < 12; m++) {
      // Respect month filter
      if (selectedMonth > 0 && m + 1 !== selectedMonth) continue;
      
      let monthlySum = 0;
      const selectedCompareYears = activeYear === 'compare'
        ? Object.keys(compareYearSelection).filter(k => compareYearSelection[k]).sort()
        : [];
      const yearsToAggregate = selectedCompareYears.length > 0 ? selectedCompareYears : allYears;
      
      yearsToAggregate.forEach(yr => {
        const yrRows = summaryFiles[yr] ? summaryFiles[yr][activeSheet] : null;
        if (!yrRows) return;
        selectedItems.forEach(item => {
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
    
    // Adjust average calculation based on filtered months and active mode
    const filteredMonths = selectedMonth > 0 ? 1 : 12;
    const yearMultiplier = activeYear === 'compare'
      ? (selectedCompareYears.length > 0 ? selectedCompareYears.length : allYears.length)
      : 1;
    setDashboardInsights({
      total: totalSum,
      peakMonth: peakM || "N/A",
      peakVal: peakValue,
      avg: filteredMonths > 0 ? Math.round(totalSum / (filteredMonths * yearMultiplier)) : 0
    });

  }, [selectedItems, activeSheet, summaryData, selectedMonth, activeYear, summaryFiles, compareYearSelection]);

  // Toggle dynamic selection pill
  const handleTogglePill = (item) => {
    const isSelected = selectedItems.some(i => i.idx === item.idx);
    if (isSelected) {
      setSelectedItems(selectedItems.filter(i => i.idx !== item.idx));
    } else {
      setSelectedItems([...selectedItems, item]);
    }
  };

  // Render Table content for Dashboard
  const getDashboardTableData = () => {
    if (!summaryData || !activeSheet || selectedItems.length === 0) return { columns: [], rows: [] };
    
    const allYears = Object.keys(summaryFiles).sort();
    const selectedCompareYears = activeYear === 'compare'
      ? Object.keys(compareYearSelection).filter(k => compareYearSelection[k]).sort()
      : [];
    const yearsToAggregate = selectedCompareYears.length > 0 ? selectedCompareYears : allYears;
    
    // Build columns based on mode
    const columns = ["月份", ...selectedItems.map(i => i.name)];
    const tableRows = [];
    
    // Monthly rows (indices 2 to 13)
    for (let m = 2; m <= 13; m++) {
      // Respect month filter
      if (selectedMonth > 0 && m - 1 !== selectedMonth) continue;
      
      const rowArr = [`${m-1}月`];
      let monthTotal = 0;
      
      selectedItems.forEach(item => {
        let sum = 0;
        yearsToAggregate.forEach(yr => {
          const yrRows = summaryFiles[yr] ? summaryFiles[yr][activeSheet] : null;
          if (!yrRows) return;
          const val = yrRows[m] ? Number(yrRows[m][item.idx]) || 0 : 0;
          sum += val;
        });
        rowArr.push(sum);
        monthTotal += sum;
      });
      tableRows.push(rowArr);
    }

    // Totals row - only show if month filter is active
    if (selectedMonth > 0 && selectedMonth <= 12) {
      const totalArr = [`${selectedMonth}月`];
      selectedItems.forEach(item => {
        let sum = 0;
        yearsToAggregate.forEach(yr => {
          const yrRows = summaryFiles[yr] ? summaryFiles[yr][activeSheet] : null;
          if (!yrRows) return;
          const totalsRow = yrRows[selectedMonth + 1];
          sum += totalsRow ? Number(totalsRow[item.idx]) || 0 : 0;
        });
        totalArr.push(sum);
      });
      tableRows.push(totalArr);
    } else {
      // Normal total row - aggregate across all years
      const totalArr = ["小計"];
      selectedItems.forEach(item => {
        let sum = 0;
        yearsToAggregate.forEach(yr => {
          const yrRows = summaryFiles[yr] ? summaryFiles[yr][activeSheet] : null;
          if (!yrRows) return;
          const tr = yrRows[14];
          sum += tr ? Number(tr[item.idx]) || 0 : 0;
        });
        totalArr.push(sum);
      });
      tableRows.push(totalArr);
    }

    return { columns, rows: tableRows };
  };

  const { columns: dbCols, rows: dbRows } = getDashboardTableData();

  // Search/Filters for Tab 2
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

    // Excel Column Filters
    for (const key in columnFilters) {
      const selected = columnFilters[key];
      if (selected && selected.length > 0) {
        if (!selected.includes(row[key])) {
          return false;
        }
      }
    }

    return true;
  });

  const getUniqueColumnValues = (fieldKey) => {
    // 依據其他所有欄位的篩選條件，動態收窄當前欄位候選值（連動過濾）
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

      // 檢查除了當前 fieldKey 以外的其他欄位過濾
      for (const key in columnFilters) {
        if (key === fieldKey) continue; // 忽略當前欄位本身的過濾
        const selected = columnFilters[key];
        if (selected && selected.length > 0) {
          if (!selected.includes(row[key])) {
            return false;
          }
        }
      }

      return true;
    });

    const vals = tempFiltered.map(row => row[fieldKey]);
    const unique = Array.from(new Set(vals.map(v => v === undefined || v === null ? "" : v)));
    return unique.sort((a, b) => String(a).localeCompare(String(b), 'zh-Hant'));
  };

  const getStatusLabel = (status) => {
    switch (status) {
      case 'matched': return '✓ 成功識別';
      case 'unmatched': return '⚠ 缺對照';
      case 'none': return '無編碼';
      case 'error': return '✗ 讀取錯誤';
      default: return status;
    }
  };

  const renderFilterableHeader = (label, fieldKey) => {
    const isFiltered = columnFilters[fieldKey] && columnFilters[fieldKey].length > 0;
    return (
      <th style={{ paddingRight: '8px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '6px' }}>
          <span style={{ whiteSpace: 'nowrap' }}>{label}</span>
          <button 
            type="button"
            className={`filter-trigger-btn ${isFiltered ? 'active' : ''}`}
            onClick={(e) => {
              e.stopPropagation();
              setActiveFilterPopover(activeFilterPopover === fieldKey ? null : fieldKey);
            }}
            title={`篩選${label}`}
            aria-label={`篩選${label}`}
            aria-expanded={activeFilterPopover === fieldKey}
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
      {/* App Header */}
      <header className="app-header">
        <div className="app-title-group">
          <h1 className="mck-serif-title">Mouldex 品管報表統計系統</h1>
          <p>Mouldex QC Record Extraction & Dynamic McKinsey Dashboard</p>
        </div>
        <span className="badge-version">v2026.06.28</span>
      </header>

      {/* Main Tab Mode Toggle */}
      <nav className="app-nav" role="tablist" aria-label="主要功能切換">
        <button 
          className={`nav-tab ${activeTab === 'extractor' ? 'active' : ''}`}
          onClick={() => setActiveTab('extractor')}
          role="tab"
          aria-selected={activeTab === 'extractor'}
          aria-controls="panel-extractor"
        >
          ⚙️ 品檢編碼對照與提取工具
        </button>
        <button 
          className={`nav-tab ${activeTab === 'dashboard' ? 'active' : ''}`}
          onClick={() => setActiveTab('dashboard')}
          role="tab"
          aria-selected={activeTab === 'dashboard'}
          aria-controls="panel-dashboard"
        >
          📊 McKinsey 品檢分析儀表板
        </button>
      </nav>

      {/* Tab 1: McKinsey Dashboard */}
      {activeTab === 'dashboard' && (
        <div id="panel-dashboard" role="tabpanel" className="mck-main-content">
          {/* Uploader Card */}
          <div className="mck-card mck-card--flat">
            <div className="mck-card-header mck-card-header--compact">
              <div>
                <h2 className="mck-card-title">匯入年度報表統計檔</h2>
                <div className="mck-card-subtitle">請上傳由 ETL Pipeline 產出的品檢報表統計檔案 (.xlsx / .json)</div>
              </div>
              {Object.keys(summaryFiles).length > 0 && (
                <div className="mck-status-success">
                  ✓ 已載入 {Object.keys(summaryFiles).length} 個檔案: {Object.keys(summaryFiles).sort().join(', ')}
                </div>
              )}
            </div>
            
            <div style={{ display: 'flex', gap: '16px', alignItems: 'center', flexWrap: 'wrap' }}>
              <input 
                type="file" 
                accept=".xlsx,.json" 
                onChange={handleSummaryFileChange} 
                id="summary-file-input" 
                multiple
                style={{ display: 'none' }}
              />
              <button 
                type="button"
                className="btn btn-primary"
                style={{ minHeight: '40px', padding: '8px 16px', fontSize: '13px' }}
                aria-label="選擇報表檔案"
                onClick={() => document.getElementById('summary-file-input')?.click()}
              >
                📁 選擇報表檔案 (xlsx/json)
              </button>
              {Object.keys(summaryFiles).length > 0 && (
                <button 
                  className="btn btn-danger"
                  onClick={() => {
                    if (window.confirm("確定要清空所有已載入的報表檔案嗎？")) {
                      setSummaryFiles({});
                      setSummaryData(null);
                      setActiveSheet("");
                      setSelectedItems([]);
                    }
                  }}
                  style={{ minHeight: '40px', padding: '8px 16px', fontSize: '13px' }}
                  aria-label="清空所有報表"
                >
                  🗑 清空所有報表
                </button>
              )}
                <span className="mck-text-slate">
                 支援 ETL Pipeline 產出的 .xlsx 報表統計檔 或 .json 結構化資料檔
                </span>
            </div>
          </div>

          {/* Year Switcher Selector */}
          {Object.keys(summaryFiles).length > 0 && (
            <div className="mck-card mck-card--compact mck-flex-center">
              <div>
                <h3 className="mck-card-title" style={{ fontSize: '15px' }}>📅 選擇分析年度與跨年對比</h3>
                <div className="mck-card-subtitle" style={{ marginTop: '2px' }}>切換單一年度數據，或啟動跨年度數據對比</div>
              </div>
              <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
                  <span className="mck-label-sm">月份篩選：</span>
                  <select 
                    value={selectedMonth}
                    onChange={(e) => setSelectedMonth(parseInt(e.target.value, 10))}
                    style={{
                      padding: '6px 24px 6px 12px',
                      fontSize: '13px',
                      borderRadius: '8px',
                      border: '1px solid var(--color-border)',
                      background: 'var(--bg-surface)',
                      color: 'var(--text-primary)',
                      fontWeight: 600,
                      cursor: 'pointer',
                      outline: 'none',
                      boxShadow: 'none',
                      transition: 'var(--transition-smooth)',
                      minHeight: '38px',
                    }}
                  >
                    <option value={0}>全部月份</option>
                    {MONTH_LABELS.map((label, idx) => (
                      <option key={idx + 1} value={idx + 1}>{label}</option>
                    ))}
                  </select>
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                  {Object.keys(summaryFiles).sort().map(year => (
                    <button 
                      key={year}
                      className={`btn ${activeYear === year ? 'btn-primary active' : 'btn-secondary'}`}
                      onClick={() => setActiveYear(year)}
                      style={{ minHeight: '36px', height: '36px', padding: '0 16px', fontSize: '13px' }}
                    >
                      {year} 年度 {activeYear === year ? '✓' : ''}
                    </button>
                  ))}
                  <button 
                    className={`btn ${activeYear === 'compare' ? 'btn-primary active' : 'btn-secondary'}`}
                    onClick={() => setActiveYear('compare')}
                    disabled={Object.keys(summaryFiles).length < 2}
                    style={{ minHeight: '36px', height: '36px', padding: '0 16px', fontSize: '13px' }}
                  >
                    📊 跨年度對比
                  </button>
                </div>
                
                {/* Year selector for cross-year comparison */}
                {activeYear === 'compare' && (
                  <div style={{ marginTop: '12px', paddingTop: '12px', borderTop: '1px solid var(--mck-border)', width: '100%' }}>
                    <span className="mck-label-sm">選擇對比年份：</span>
                    <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', marginTop: '4px' }}>
                      {Object.keys(summaryFiles).sort().map(year => {
                        const isSelected = compareYearSelection[year];
                        return (
                          <button
                            key={year}
                            className={`btn ${isSelected ? 'btn-primary' : 'btn-secondary'}`}
                            onClick={() => {
                              setCompareYearSelection(prev => ({
                                ...prev,
                                [year]: !prev[year]
                              }));
                            }}
                            style={{ minHeight: '28px', height: '28px', padding: '0 10px', fontSize: '12px' }}
                          >
                            {isSelected ? '✓' : '+'} {year}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {summaryData ? (
            <div className="mck-dashboard">
              {/* Sidebar Tabs */}
              <aside className="mck-sidebar">
                <div style={{ fontSize: '11px', fontWeight: '700', textTransform: 'uppercase', letterSpacing: '0.6px', padding: '0 8px 8px 8px', borderBottom: '1px solid var(--mck-border)', color: 'var(--mck-slate)' }}>
                  品檢報表工作表
                </div>
                {Object.keys(summaryData).map((sheetName) => (
                  <button
                    key={sheetName}
                    className={`mck-sheet-btn ${activeSheet === sheetName ? 'active' : ''}`}
                    onClick={() => setActiveSheet(sheetName)}
                  >
                    📝 {sheetName}
                  </button>
                ))}
              </aside>

              {/* Main Analysis Panels */}
              <div className="mck-main-content">
                {/* Checkbox pills */}
                <div className="mck-card">
                  <div className="mck-card-header">
                    <div>
                      <h2 className="mck-card-title">自訂篩選分析項目</h2>
                      <div className="mck-card-subtitle">點選下方標籤以新增或移除圖表分析的統計項目</div>
                    </div>
                    {(() => {
                      const rows = summaryData[activeSheet];
                      if (!rows || rows.length < 2) return null;
                      const headerRow = rows[1] || [];
                      const availableItems = [];
                      headerRow.forEach((h, idx) => {
                        const name = String(h || '').trim();
                        if (name && name !== '月份' && name !== '小計' && name !== 'NCA') {
                          availableItems.push({ name, idx });
                        }
                      });
                      
                      const isAllSelected = availableItems.length > 0 && selectedItems.length === availableItems.length;

                      return (
                        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                          <button 
                            className="btn btn-primary"
                            onClick={() => setSelectedItems(availableItems)}
                            disabled={isAllSelected}
                            style={{ 
                              minHeight: '32px', 
                              height: '32px', 
                              padding: '0 12px', 
                              fontSize: '12px',
                              opacity: isAllSelected ? 0.6 : 1,
                              cursor: isAllSelected ? 'not-allowed' : 'pointer'
                            }}
                          >
                            ✅ 一鍵全選
                          </button>
                          {selectedItems.length > 0 && (
                            <button 
                              className="btn btn-secondary"
                              onClick={() => setSelectedItems([])}
                              style={{ minHeight: '32px', height: '32px', padding: '0 12px', fontSize: '12px' }}
                            >
                              🗑 一鍵清空
                            </button>
                          )}
                        </div>
                      );
                    })()}
                  </div>
                  <div className="mck-pill-container">
                    {(() => {
                      const rows = summaryData[activeSheet];
                      if (!rows || rows.length < 2) return <p className="mck-empty-text">該工作表無有效數據欄位</p>;
                      const headerRow = rows[1] || [];
                      const items = [];
                      headerRow.forEach((h, idx) => {
                        const name = String(h || '').trim();
                        if (name && name !== '月份' && name !== '小計' && name !== 'NCA') {
                          items.push({ name, idx });
                        }
                      });
                      return items.map((item) => {
                        const active = selectedItems.some(i => i.idx === item.idx);
                        return (
                          <div 
                            key={item.idx} 
                            className={`mck-pill ${active ? 'active' : ''}`}
                            onClick={() => handleTogglePill(item)}
                          >
                            <span className="mck-pill-checkbox"></span>
                            {item.name}
                          </div>
                        );
                      });
                    })()}
                  </div>
                </div>

                {/* Plot Panel */}
                <div className="mck-card">
                  <div className="mck-card-header">
                    <div>
                      <h2 className="mck-card-title">{activeSheet} 數據趨勢</h2>
                      <div className="mck-card-subtitle">以月份為 X 軸，數量為 Y 軸呈現趨勢</div>
                    </div>
                  </div>
                  {selectedItems.length > 0 ? (
                    <div className="mck-chart-container">
                      <canvas ref={chartRef}></canvas>
                    </div>
                  ) : (
                    <div className="mck-empty-text" style={{ textAlign: 'center', padding: '64px 0' }}>
                      ⚠️ 請在上方選擇至少一個分析項目以繪製圖表。
                    </div>
                  )}
                </div>

                {/* Executive Insights Deck */}
                {selectedItems.length > 0 && (
                  <div className="mck-kpi-deck">
                    <div className="mck-kpi-card">
                      <div className="mck-kpi-label">篩選項目總筆數</div>
                      <div className="mck-kpi-val">{dashboardInsights.total.toLocaleString()}</div>
                      <div className="mck-kpi-desc">已選取子項累計</div>
                    </div>
                    <div className="mck-kpi-card">
                      <div className="mck-kpi-label">月均量</div>
                      <div className="mck-kpi-val">{dashboardInsights.avg.toLocaleString()}</div>
                      <div className="mck-kpi-desc">月平均品檢次數</div>
                    </div>
                    <div className="mck-kpi-card">
                      <div className="mck-kpi-label">高峰月份</div>
                      <div className="mck-kpi-val">{dashboardInsights.peakVal.toLocaleString()}</div>
                      <div className="mck-kpi-desc">最高峰出現在 {dashboardInsights.peakMonth}</div>
                    </div>
                  </div>
                )}

                {/* Data Table Card */}
                {selectedItems.length > 0 && dbRows.length > 0 && (
                  <div className="mck-card">
                    <div className="mck-card-header">
                      <h2 className="mck-card-title">數據明細報表</h2>
                    </div>
                    <div className="mck-table-wrapper">
                      <table className="mck-table">
                        <thead>
                          <tr>
                            {dbCols.map((col, idx) => (
                              <th key={idx} style={{ textAlign: idx === 0 ? 'left' : 'right' }}>{col}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {dbRows.map((row, rIdx) => {
                            const isTotal = row[0] === '小計';
                            return (
                              <tr key={rIdx} className={isTotal ? 'total-row' : ''}>
                                <td style={{ fontWeight: isTotal ? '700' : 'normal' }}>{row[0]}</td>
                                {row.slice(1).map((val, cIdx) => (
                                  <td key={cIdx} style={{ textAlign: 'right' }}>
                                    {val.toLocaleString()}
                                  </td>
                                ))}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

              </div>
            </div>
          ) : (
            <div className="panel-card" style={{ padding: '64px', textAlign: 'center' }}>
              <span style={{ fontSize: '48px' }}>📂</span>
              <h3 style={{ margin: '16px 0 8px 0', fontSize: '18px', fontWeight: '600' }}>尚未載入品檢統計報表</h3>
              <p className="mck-desc-text">
                請點擊上方按鈕，載入生成的品檢報表統計 .xlsx 檔案。載入後，系統將依麥肯錫風格為您呈現自訂動態製圖與數據分析。
              </p>
            </div>
          )}
        </div>
      )}

      {/* Tab 2: Original Extraction & Mappings Extractor */}
      {activeTab === 'extractor' && (
        <div className="main-grid">
          {/* Left Panel: Files Scanning controls */}
          <aside ref={leftPanelRef} className="mck-main-content">
            {/* Folder scanner card */}
            <section className="panel-card">
              <h2 className="card-title">📁 品管檔案夾掃描</h2>
              
              <div 
                className="upload-zone"
                onDragOver={handleDragOver}
                onDrop={handleDrop}
                role="region"
                aria-label="檔案上傳區域"
              >
                <div className="upload-icon">📥</div>
                <p style={{ fontWeight: 500 }}>拖曳檔案或資料夾至此</p>
                <p style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>支援批次讀取多個 QC 報表檔案</p>
                
                <div className="upload-btn-group">
                  <button 
                    className="btn btn-primary" 
                    style={{ flex: 1 }}
                    onClick={() => folderInputRef.current.click()}
                    aria-label="選取資料夾"
                  >
                    📂 選取資料夾
                  </button>
                </div>

                <input 
                  type="file" 
                  ref={folderInputRef} 
                  webkitdirectory="true" 
                  directory="true" 
                  multiple 
                  onChange={handleFolderChange} 
                  style={{ display: 'none' }} 
                />
              </div>

              {scannedRows.length > 0 && (
                <div style={{ marginTop: '16px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                    已加載: <strong>{folderName}</strong>
                  </div>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button 
                    className="btn btn-primary" 
                    style={{ flex: 1, minHeight: '36px' }}
                    onClick={() => exportToCSV(filteredRows, folderName || "品管報表")}
                    aria-label="匯出 CSV"
                  >
                    📄 匯出 CSV
                  </button>

                  <button 
                    className="btn btn-danger btn-icon" 
                    onClick={() => setScannedRows([])}
                    title="清除結果"
                    aria-label="清除結果"
                  >
                    🗑
                  </button>
                  </div>
                </div>
              )}
            </section>

            {uploadedFiles.length > 0 && (
              <section className="panel-card mck-border-gold">
                <h2 className="card-title" style={{ color: '#856404', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  🛠️ 年度統計報表一鍵生成器
                </h2>
                <p style={{ fontSize: '13px', color: '#664d03', marginBottom: '16px', lineHeight: '1.5' }}>
                  偵測到載入的資料夾包含 <strong>{uploadedFiles.length}</strong> 個品檢檔案。您可以在此選取解析引擎並一鍵執行瀏覽器端原生的 ETL 運算，輸出含有完整分流彙整版面的年度 Excel 統計檔。
                </p>
                
                {isProcessingETL ? (
                  <div>
                    <div style={{ fontSize: '13px', fontWeight: 600, color: 'var(--mck-navy)', marginBottom: '8px' }}>
                      ⚙️ 正在執行品管資料 ETL 轉換中... ({etlProgress?.current}/{etlProgress?.total})
                    </div>
                    <div style={{ fontSize: '12px', color: 'var(--mck-slate)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      處理檔案: {etlProgress?.filename}
                    </div>
                    <div style={{ width: '100%', height: '8px', background: '#E2E8F0', borderRadius: '4px', marginTop: '8px', overflow: 'hidden' }}>
                      <div style={{ width: `${(etlProgress?.current / etlProgress?.total) * 100}%`, height: '100%', background: 'var(--mck-navy)', transition: 'width 0.1s ease' }}></div>
                    </div>
                    <button 
                      className="btn btn-danger" 
                      onClick={() => {
                        setIsETLCancelled(true);
                        isETLCancelledRef.current = true;
                      }}
                      style={{ marginTop: '12px' }}
                    >
                      ❌ 取消處理
                    </button>
                  </div>
                ) : (
                  <div>
                    {/* 主流程操作區 */}
                    <div className="mck-section-divider">
                      <div className="mck-section-label">
                        🚀 主流程操作
                      </div>
                      <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <span style={{ fontSize: '13px', fontWeight: 600, color: '#664d03' }}>報表年度：</span>
                          <select 
                            className="mck-select-gold"
                            value={etlYear}
                            onChange={(e) => {
                              setEtlYear(parseInt(e.target.value, 10));
                              setCachedCounts(null);
                            }}
                            disabled={isScanning}
                            style={{ opacity: isScanning ? 0.6 : 1, cursor: isScanning ? 'not-allowed' : 'pointer' }}
                          >
                            {Array.from({ length: 31 }, (_, i) => 2010 + i).map(year => (
                              <option key={year} value={year}>{year} 年</option>
                            ))}
                          </select>
                        </div>

                         <button 
                           className="btn btn-primary"
                           onClick={() => handleRunBrowserETL(etlYear)}
                           disabled={isScanning}
                           style={{ opacity: isScanning ? 0.6 : 1, cursor: isScanning ? 'not-allowed' : 'pointer' }}
                         >
                           {isScanning ? '🔍 正在解析原始檔案中...' : `📊 輸出 ${etlYear} 品檢報表統計 (JSON)`}
                         </button>

                         <button 
                           className="btn btn-secondary"
                           onClick={() => handleExportExcel(etlYear)}
                           disabled={isScanning || !cachedCounts}
                           style={{ opacity: (isScanning || !cachedCounts) ? 0.6 : 1, cursor: (isScanning || !cachedCounts) ? 'not-allowed' : 'pointer' }}
                         >
                           📥 下載 Excel 報表
                         </button>
                         
                         <button 
                           className="btn btn-secondary"
                           onClick={() => handleExportIndividualReports(etlYear)}
                          disabled={isScanning}
                          style={{ opacity: isScanning ? 0.6 : 1, cursor: isScanning ? 'not-allowed' : 'pointer' }}
                        >
                          {isScanning ? '🔍 解析中...' : `📦 輸出 ${etlYear} 獨立報表 (全部 QC)`}
                        </button>
                      </div>
                    </div>

                    {/* 資料分析工具區 - 已整合到「品管檔案夾掃描」中 */}
                    <div className="mck-section-divider">
                      <div className="mck-section-label">
                        🔍 資料分析工具
                      </div>
                      <p style={{ fontSize: '13px', color: 'var(--text-secondary)', margin: '0' }}>
                        請先在上方點擊「📂 選取資料夾」按鈕，選取包含 QC 報表的資料夾後即可自動顯示資料夾結構與 QC 標籤查看結果。
                      </p>
                    </div>

                    {/* 資料管理工具區 - 已整合到「品管檔案夾掃描」中 */}
                    <div>
                      <div className="mck-section-label">
                        🗂 資料管理工具
                      </div>
                      <p style={{ fontSize: '13px', color: 'var(--text-secondary)', margin: '0' }}>
                        請先在上方選擇 QC 報表檔案後使用。
                      </p>
                    </div>
                  </div>
                )}
              </section>
            )}

            {/* Mappings Editor */}
            <section className="panel-card">
              <h2 className="card-title">⚙️ 表單編碼與名稱對照表</h2>
              
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '12px' }}>
                <button className="btn btn-secondary" style={{ fontSize: '12px', minHeight: '32px', padding: '4px 8px' }} onClick={handleExportMappings}>
                  📤 匯出 JSON
                </button>
                <label className="btn btn-secondary" style={{ fontSize: '12px', minHeight: '32px', padding: '4px 8px', cursor: 'pointer' }}>
                  📥 匯入 JSON
                  <input type="file" accept=".json" onChange={handleImportMappings} style={{ display: 'none' }} />
                </label>
                <button className="btn btn-danger" style={{ fontSize: '12px', minHeight: '32px', padding: '4px 8px' }} onClick={handleResetMappings}>
                  🔄 恢復預設
                </button>
              </div>

              <div className="mapping-list">
                {Object.keys(mappings).map((code) => (
                  <div key={code} className="mapping-item">
                    <span className="mapping-input mapping-code">{code}</span>
                    <input 
                      type="text" 
                      className="mapping-input mapping-name" 
                      value={mappings[code]} 
                      onChange={(e) => {
                        const updated = { ...mappings, [code]: e.target.value };
                        updateMappings(updated);
                      }} 
                    />
                    <button 
                      className="btn btn-danger" 
                      style={{ minHeight: '32px', width: '32px', padding: 0 }}
                      onClick={() => handleDeleteMapping(code)}
                    >
                      ✕
                    </button>
                  </div>
                ))}
              </div>

              <form onSubmit={handleAddMapping} style={{ borderTop: '1px solid var(--color-border)', paddingTop: '16px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <div style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-secondary)' }}>新增對照關係：</div>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <input 
                    type="text" 
                    className="mapping-input" 
                    placeholder="QC10006-R02" 
                    value={newCode} 
                    onChange={(e) => setNewCode(e.target.value)} 
                    style={{ width: '120px' }}
                  />
                  <input 
                    type="text" 
                    className="mapping-input" 
                    placeholder="對應表單名稱" 
                    value={newName} 
                    onChange={(e) => setNewName(e.target.value)} 
                    style={{ flex: 1 }}
                  />
                </div>
                <button type="submit" className="btn btn-primary" style={{ width: '100%', minHeight: '36px' }}>
                  ＋ 新增至對照表
                </button>
              </form>
            </section>
          </aside>

          {/* Right Panel: Result data table */}
          <section 
            className="panel-card" 
            style={{ 
              display: 'flex', 
              flexDirection: 'column', 
              flex: 1, 
              maxHeight: leftPanelHeight ? `${leftPanelHeight}px` : 'none',
              overflow: 'hidden' 
            }}
          >
            <div className="table-controls">
              <h2 className="card-title" style={{ margin: 0 }}>📊 工作表表單編碼提取結果</h2>
              
              <div className="filter-group">
                <input 
                  type="text" 
                  className="search-input" 
                  placeholder="搜尋結果..." 
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  aria-label="搜尋結果"
                />
                
                <select 
                  className="filter-select"
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                  aria-label="狀態篩選"
                >
                  <option value="all">所有狀態</option>
                  <option value="matched">成功識別</option>
                  <option value="unmatched">缺對照</option>
                  <option value="none">無編碼</option>
                  <option value="error">讀取錯誤</option>
                </select>
                
                {Object.values(columnFilters).some(arr => arr.length > 0) && (
                  <button 
                    type="button"
                    className="btn btn-secondary" 
                    style={{ 
                      fontSize: '12px', 
                      minHeight: '32px', 
                      padding: '4px 12px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '4px',
                      animation: 'popoverFadeIn 0.2s ease-out'
                    }}
                    onClick={() => {
                      setColumnFilters({
                        fileName: [],
                        filePath: [],
                        sheetName: [],
                        foundCode: [],
                        foundName: [],
                        status: []
                      });
                      setActiveFilterPopover(null);
                    }}
                    title="重設所有欄位篩選條件"
                    aria-label="重設所有欄位篩選條件"
                  >
                    🧹 重設篩選
                  </button>
                )}
              </div>
            </div>

            {scannedRows.some(row => row.etlStatus === '狀態異常') && (
              <div style={{ 
                padding: '12px 16px', 
                borderRadius: '8px', 
                marginBottom: '16px', 
                display: 'flex', 
                alignItems: 'center', 
                justifyContent: 'space-between', 
                gap: '12px',
                width: '100%',
                border: '1px solid #FEB2B2',
                backgroundColor: '#FFF5F5',
                color: 'var(--color-error)'
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ fontSize: '16px' }}>⚠️</span>
                  <span style={{ fontWeight: 600, fontSize: '13px' }}>
                    告警：偵測到有工作表在 ETL 校驗中被判定為「狀態異常」（例如：QC 表單資料不完整、缺月份資料夾或讀取錯誤），請點擊右方按鈕排查。
                  </span>
                </div>
                <button 
                  className="btn btn-danger" 
                  style={{ minHeight: '28px', height: '28px', padding: '0 10px', fontSize: '11px', whiteSpace: 'nowrap' }}
                  onClick={() => {
                    setColumnFilters(prev => ({
                      ...prev,
                      etlStatus: ['狀態異常']
                    }));
                  }}
                >
                  🔍 立即篩選異常
                </button>
              </div>
            )}

            {isScanning ? (
              <div className="empty-state" style={{ padding: '128px 0' }}>
                <div className="upload-icon" style={{ animation: 'spin 2s linear infinite' }}>🔄</div>
                <p style={{ fontWeight: 500 }}>正在解析 Excel 報表，請稍候... {scanProgress ? `(${scanProgress.current}/${scanProgress.total})` : ''}</p>
                {scanProgress && (
                  <div style={{ width: '280px', height: '6px', background: '#E2E8F0', borderRadius: '3px', marginTop: '12px', overflow: 'hidden', margin: '12px auto 0' }}>
                    <div style={{ width: `${(scanProgress.current / scanProgress.total) * 100}%`, height: '100%', background: 'var(--mck-navy)', transition: 'width 0.15s ease' }}></div>
                  </div>
                )}
                <p style={{ color: 'var(--text-secondary)', fontSize: '13px', marginTop: '8px' }}>依檔案數量及大小而定</p>
              </div>
            ) : filteredRows.length > 0 ? (
              <div className="table-wrapper">
                {activeFilterPopover && (
                  <div className="filter-backdrop" onClick={() => setActiveFilterPopover(null)} />
                )}
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>序號</th>
                      {renderFilterableHeader('檔案名稱', 'fileName')}
                      {renderFilterableHeader('檔案路徑', 'filePath')}
                      {renderFilterableHeader('工作表名稱', 'sheetName')}
                      {renderFilterableHeader('表單編碼', 'foundCode')}
                      {renderFilterableHeader('表單對照名稱', 'foundName')}
                      {renderFilterableHeader('狀態', 'status')}
                      {renderFilterableHeader('是否納入ETL計算', 'etlStatus')}
                      {renderFilterableHeader('原因說明', 'etlReason')}
                      {renderFilterableHeader('更新時間', 'etlTimestamp')}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.map((row, idx) => (
                      <tr key={idx}>
                        <td style={{ color: 'var(--text-secondary)' }}>{idx + 1}</td>
                        <td style={{ fontWeight: 500 }}>{row.fileName}</td>
                        <td className="filepath-cell">{row.filePath}</td>
                        <td>{row.sheetName}</td>
                        <td style={{ fontFamily: 'monospace', fontWeight: 600 }}>{row.foundCode}</td>
                        <td>{row.foundName}</td>
                        <td>
                          {row.status === 'matched' && (
                            <span className="status-badge status-matched">✓ 成功識別</span>
                          )}
                          {row.status === 'unmatched' && (
                            <span className="status-badge status-unmatched">⚠ 缺對照</span>
                          )}
                          {row.status === 'none' && (
                            <span className="status-badge status-none">無編碼</span>
                          )}
                          {row.status === 'error' && (
                            <span className="status-badge status-error">✗ 讀取錯誤</span>
                          )}
                        </td>
                        <td>
                          {row.etlStatus === '已納入' && (
                            <span className="status-badge status-matched">已納入</span>
                          )}
                          {row.etlStatus === '未納入' && (
                            <span className="status-badge status-none">未納入</span>
                          )}
                          {row.etlStatus === '狀態異常' && (
                            <span className="status-badge status-error">狀態異常</span>
                          )}
                        </td>
                        <td style={{ fontSize: '13px', color: row.etlStatus === '狀態異常' ? 'var(--alert-error, #EF4444)' : 'var(--text-secondary)' }}>{row.etlReason || '-'}</td>
                        <td style={{ color: 'var(--text-secondary)', fontSize: '13px' }}>{row.etlTimestamp || 'N/A'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="empty-state" style={{ flex: 1, display: 'flex', justifyContent: 'center' }}>
                <span className="empty-icon">📈</span>
                <h3>暫無統計資料</h3>
                <p>請由左側批次選取或拖入 Excel 檔案，即可在此處展開提取與編碼比對結果。</p>
              </div>
            )}
          </section>
        </div>
      )}

      {/* Footer */}
      <footer style={{
        marginTop: '24px',
        paddingTop: '16px',
        borderTop: '1px solid var(--mck-border)',
        textAlign: 'center',
        color: 'var(--mck-slate)',
        fontSize: '13px'
      }}>
        Developed by Wesley Chang @Mouldex, 2026.
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
          清除篩選
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

# 交接與重啟指南 (Handover Resume Guide)

## 當前專案狀態
- **專案名稱**：FileName_WorkSheet_Extract (Mouldex QMS 醫療器材品保檢驗數據提取與跨年度確效系統)
- **分支**：`main`
- **部署環境**：GitHub Pages（透過 GitHub Actions 自動確效並部署）
- **最新完成進度**：
  1. 高階醫材 QMS 工作台風格重構（去 AI 味、3-Stage 流暢工作引導）。
  2. 跨年度 QMS 數據與 ISO 13485 主任品質稽核員 Prompt 大模型一鍵匯出模組 (`src/utils/llmExport.js`)。
  3. 全專案 ESLint 零錯誤、Vite Production Build 零錯誤、軟體確效腳本 100% 通過。

---

## 核心架構摘要

### 1. 三段式品保工作台 (3-Stage Industrial MedTech Workbench)
- **STAGE 01 (原始檢驗表單掃描與批次校驗)**：
  - 批次讀取多層目錄，自動檢驗 7 大類 QC 編碼、Date Code 格式及空白樣板防呆守衛。
  - 支援異常偏差一鍵排查、多欄位篩選彈窗與 CSV 檢驗清單匯出。
- **STAGE 02 (QMS ETL 清洗與法規報表產出)**：
  - 雙引擎即時 ETL 運算、年度切換。
  - 輸出 QMS 年度統計 JSON、法規稽核 Excel 報表、7 大類獨立 QC 分冊。
  - 提供「立即同步至階段 03 儀表板」一鍵直達按鈕。
- **STAGE 03 (跨年度品質趨勢與大模型診斷)**：
  - 單年度/跨年度對比（2023、2024、2025、2026...）、月份篩選、組件堆疊柱狀圖。
  - 4 大數位 LED 檢驗 KPI 指標（全期檢驗總數、最高峰月份、單月最高量、月均負荷）。
  - 一鍵開啟大模型 (LLM) 匯出視窗。

### 2. 大模型 (LLM) 數據匯出引擎 (`src/utils/llmExport.js`)
- 彙整所有載入年份數據，生成涵蓋 YoY 成長趨勢、季節性波動、Setup vs 巡檢比例、關鍵組件 80/20 Pareto 風險分析及 ISO 13485 CAPA 建議的專業 Markdown 指令與機器可讀 JSON 矩陣。
- 支援彈窗即時預覽、一鍵複製至剪貼簿、下載 `.md` 指令檔與 `.json` 數據包。

### 3. 動態欄位解析 (Dynamic Columns)
- `getRawSubCategory` 自動動態解析所有出現的品檢子類別（如 Biometrix, Vivus, 射出A, 射出C, 射出D(組件) 等）。
- 新品項無須手動新增至代碼，自動展開為報表欄位。

### 4. 特殊製程與格式防禦
- **UUID 路徑防禦**：`isUUID` 過濾器支援標準 UUID、32 位十六進位、及長度 ≥ 24 的隨機字串。
- **民國年與點分隔符支援**：支援三位數民國年（如 `112/03/15`）及 `.` 分隔符。
- **QC10007-R03 特殊月份覆蓋與 O4 繞過**：針對 Tubing、射出、射出A、射出C、射出D(組件) 之檔名與路徑月份覆蓋邏輯。

---

## 關鍵檔案索引

| 檔案 | 職責 |
|---|---|
| `src/App.jsx` | 3-Stage 醫材工作台主介面、快取、圖表、UI 控制 |
| `src/utils/llmExport.js` | 跨年度 QMS 數據萃取與 ISO 13485 大模型 Prompt 生成引擎 |
| `src/utils/browserETL.js` | ETL 核心清洗引擎、動態欄位展開、日期與月份解析 |
| `src/utils/excelParser.js` | Excel 表單掃描、SheetJS 解析、JSON 匯入解析器 |
| `src/utils/db.js` | QC 表單編碼與名稱對照表 IndexedDB / LocalStorage 持久化 |
| `src/index.css` | 高階醫材 QMS 儀器工作台設計系統與色票規範 |
| `scratch/validate_qc_etl.cjs` | 自動化軟體確效測試腳本 (Validation Suite) |
| `DEV_LOG.md` | 開發日誌（完整記錄所有需求、RCA 根因分析與 CAPA 措施） |

---

## 開發與確效驗證流程

```bash
# 1. 本地開發伺服器
npm run dev

# 2. 程式碼規範檢查
npx eslint src/

# 3. 軟體確效測試 (Mandatory Software Validation)
node scratch/validate_qc_etl.cjs

# 4. 生產環境建構打包
npm run build

# 5. 推送部署（通過本地驗證後向使用者請求許可後執行）
git push origin main
```

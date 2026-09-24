# 品檢報表統計 ETL Pipeline

本專案為自動化品檢報表統計與分析平台。採用純前端瀏覽器端（Vite + React SPA）的架構，使用者可以直接匯入本地資料夾或選取檔案，即時在瀏覽器中執行 ETL (提取、轉換、載入) 管線，並直接下載統計 Excel 報表與進行 McKinsey 風格的互動式儀表板分析。

## 功能

- **多來源資料彙整**：自動掃描上傳之資料夾與檔案，辨識 7 大類 QC 檢驗記錄：
  - 原物料進料品檢 (QC10002-R02)
  - QIP 尺寸檢驗 (QC10004-R02)
  - 裝配對樣巡檢 (QC10006-R01)
  - 半成品品檢 (QC10006-R02)
  - 完成品品檢 (QC10007-R01)
  - 零組件入庫品檢 (QC10007-R03)
  - 出貨檢驗 (QC10008-R02)
- **智能分類統計**：按月份、品項/類別自動統計筆數。
- **民國年日期支援**：自動解析民國年格式（如 `112/03/15` → 2023 年 3 月）及點號分隔（如 `112.03.15`），涵蓋 `parseDateFromString` 與 `findDateInSheetFallback`。
- **半成品品檢雙重驗證**：要求工作表 QC 編碼為 `QC10006-R02` **且** 檔名/路徑含 `半成品品檢表`，避免誤判。
- **空白樣板守衛**：自動識別並跳過批號（儲存格 `G4`）為空的空白樣板檔案，防禦幻象數據（特定射出類別豁免）。
- **McKinsey 互動儀表板與 LLM 導出**：React 應用，使用者可任意勾選品項、以月份為 X 軸生成堆疊柱狀圖，支援跨年度對比模式與月份篩選；並提供一鍵導出 ISO 13485 醫療器材大模型 (LLM) 診斷 Prompt。
- **品檢對照提取與過濾**：瀏覽器端 ETL 列出每個工作表是否已被納入 ETL 統計之「原因說明」，並配備如同 Excel 的獨立欄位篩選彈窗。
- **獨立報表統計**：一次產出年度品檢報表統計 Excel 檔案。
- **JSON 匯入儀表板**：支援將 ETL 匯出之 JSON 格式報表（`QC_Annual_Summary_v1`）直接載入 McKinsey 儀表板進行分析，無需重新掃描 Excel 檔案。
- **ETL 結果快取**：初次轉換完成後自動快取結果資料，後續切換格式或輸出獨立報表時無需重新掃描，大幅提升重複輸出體驗。
- **GitHub Pages 自動部署**：推送至 `main` 分支後，透過 GitHub Actions 自動構建並部署至 GitHub Pages。
- **Inset Focus 內凹聚焦視覺語彙**：全系統採「外凸 (raised) 卡片承載內容、內凹 (inset) 表面承載輸入」的柔雙陰影設計（Neumorphism-Lite），搭配長春藍 (Periwinkle) 單一主色、柔彩膠囊標籤、標題漸層藍規線與內凹表頭；並嚴格遵循「最小字體不得小於 13px」與 MedTech High-Precision 階梯式排版規範。

## 使用方式

```bash
# 安裝依賴（僅首次）
npm install

# 啟動互動式儀表板與瀏覽器端 ETL 管線（開發模式）
npm run dev

# 代碼品質與語法檢查
npm run lint

# 部署生產環境（輸出至 dist/）
npm run build
```

## 前端 3 階段工作流程 (MedTech Workflow)

```
[階段 01：品檢編碼對照與檔案篩選]
  用戶選取或拖曳本地資料夾 / 多個 Excel 檔案 ➔ 即時解析 Column A 之 QC 編碼 ➔ 提供 Excel 等級多欄位篩選與異常排查
        │
        ▼
[階段 02：QMS ETL 數據清洗與法規報表產出]
  自動進行月份判定、去重與防呆過濾 ➔ 下載年度品檢報表統計 (.xlsx) ➔ 一鍵同步至分析儀表板
        │
        ▼
[階段 03：品質數據分析儀表板與 LLM 導出]
  多維度堆疊柱狀圖分析 ➔ 跨年度趨勢比對 ➔ 一鍵生成 ISO 13485 醫療器材大模型 (LLM) 診斷 Prompt 與全年度數據包
```

## 目錄結構

```
├── docs/                         # 開發相關文件與 wiki
│   ├── README.md                  # 文件索引（本 wiki 首頁：SSOT 分工表與維護規則）
│   ├── qc-path-analysis/         # 品檢映射路徑分析文檔
│   ├── today-requirements-*.md   # 歷史需求追溯日誌
│   ├── 狀態異常訊息.md           # ETL 狀態異常觸發條件說明
│   └── handover_resume_guide.md  # 交接與重啟指南
├── src/                          # React SPA 互動儀表板與前端 ETL
│   ├── App.jsx                   # 主介面（3 階段工作流控制、圖表與彈窗）
│   ├── index.css                 # 全域樣式 (Inset Focus 設計代幣：柔雙陰影 + 長春藍主色 + 最小 13px 排版階梯)
│   ├── main.jsx                  # React 入口
│   └── utils/
│       ├── browserETL.js         # ETL 核心（動態欄位 + 民國年 + 去重過濾）
│       ├── db.js                 # QC 表單編號對照表 (localStorage 管理)
│       ├── excelParser.js        # 瀏覽器端高效 Excel 解析引擎
│       └── llmExport.js          # ISO 13485 醫療器材大模型 Prompt 與數據包建構器
├── public/                       # 靜態資源
├── scratch/                      # 開發與確效測試腳本 (gitignored)
│   └── validate_qc_etl.cjs       # QC ETL 數據映射規則自動確效工具
├── .github/workflows/            # GitHub Actions 自動部署配置
│   └── deploy.yml                # 自動構建與 GitHub Pages 部署腳本
├── .agents/                      # 專案層 AI 代理規則（QIP 提取規則 + Ponytail 開發模式）
├── DEV_LOG.md                    # 開發日誌（含 RCA + CAPA 歷史記錄）
├── eslint.config.js              # ESLint 規範設定
├── package.json                  # 專案依賴與腳本
├── vite.config.js                # Vite 配置檔
└── .gitignore                    # Git 忽略規則
```

## 資訊架構與 SSOT 原則 (Single Source of Truth)

本專案的開發資訊依「單一事實來源 (SSOT) + MECE」原則分屬三個互斥且窮盡的源頭；其他文件只引用、不複述：

| 資訊域 | SSOT 檔案 | 內容範圍 (MECE 分工) |
|---|---|---|
| 專案總覽、功能規格、報表結構、ETL 統計規則 | `README.md`（本檔） | 使用者可見的功能與規格、目錄樹、QC10007-R03 子類別規則表 |
| 時序開發記錄（需求 / RCA / CAPA / 開發節點） | `DEV_LOG.md` | 依日期倒序之唯一開發史；所有變更節點以此為準 |
| 主題式技術文件（wiki） | `docs/`（入口 `docs/README.md`） | 狀態異常觸發條件、QC10002-R02 映射分析、歷史需求快照、交接指南 |

跨平台一致性約束（架構同構性與數據一致性）：

- **部署基底路徑**：唯一來源為 `vite.config.js` 的 `base: '/QC_Annual_Summary/'`，本機開發與 GitHub Pages 部署同構。
- **JSON 資料格式識別**：唯一來源為 `src/utils/browserETL.js` 匯出之 `JSON_FORMAT_ID`（`QC_Annual_Summary_v1`），匯出與匯入兩側共用同一常數，禁止再寫死字面值。
- **最小字級**：全系統 ≥ 13px，由 `src/index.css` 設計代幣層與排版階梯強制。

## 技術棧

- **Vite + React** — 前端 SPA 與開發伺服器
- **SheetJS (xlsx)** — 瀏覽器端 Excel 檔案讀寫
- **Chart.js** — 堆疊柱狀圖與 KPI 圖表呈現
- **Google Fonts** — Outfit 與 Noto Sans TC 高級字型
- **GitHub Actions** — 自動化構建與 GitHub Pages 部署

## 輸出 Excel 報表結構

> **注意**：由於採用動態欄位架構，下表為典型品項；實際欄位會依掃描到的品項自動展開。

| 工作表 | 典型欄位結構 |
|---|---|
| 原物料品檢(QC10002-R02) | 月份 + 原料 + 物料-* + 射出D + 小計 |
| QIP(QC10004-R02) | 押出/射出 Setup + 押出/射出 巡檢（4欄 + 4欄 並列） |
| 裝配對樣巡檢(QC10006-R01) | 月份 + 裝配巡檢 |
| 半成品品檢(QC10006-R02) | 月份 + 裝配C + (動態品項) + 小計 |
| 完成品品檢(QC10007-R01 R02) | 月份 + (動態品項) + 小計 |
| 零組件入庫品檢(QC10007-R03) | 月份 + Tubing + 射出 + 射出A + 射出C + 射出D(組件) + 裝配A + 裝配B + 裝配C |
| 出貨檢驗(QC10008-R02) | 月份 + (動態品項) + 小計 |

---

## 零組件入庫品檢 (QC10007-R03) 子類別特殊映射與統計規則

為了防止跨月日期導致統計紊亂，並提供準確的品檢數據，`QC10007-R03` 下的子類別採用了特定的月份提取與空白樣板過濾定義：

| 子分類 (subCat) | 月份判定依據 | O4 日期檢查 | 空白樣板檢測 (G4 批號) |
|---|---|---|---|
| **裝配A**、**裝配B**、**裝配C** | **檔案名稱後綴字母** (A-L 對應 1-12 月) | 🚫 捨棄且 Bypass | ⚠️ 啟用（空值過濾） |
| **射出D(組件)** | **檔案名稱後綴字母** (A-L 對應 1-12 月) | 🚫 捨棄且 Bypass | 🚫 豁免（空值不跳過） |
| **Tubing** | **資料夾名稱後綴月份** (如 `-02` 對應 2 月) | 🚫 捨棄且 Bypass | ⚠️ 啟用（空值過濾） |
| **射出(廠內)**、**射出A**、**射出C** | **資料夾名稱後綴月份** (如 `-03` 對應 3 月) | 🚫 捨棄且 Bypass | 🚫 豁免（空值不跳過） |
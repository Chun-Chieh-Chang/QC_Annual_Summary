# 文件索引 (docs — 專案 Wiki)

本目錄為本專案的主題式技術文件（wiki）。依 SSOT/MECE 原則：`README.md` 管功能規格、`DEV_LOG.md` 管時序開發史、本目錄管主題深潛；各文件互斥分工如下，其他文件僅引用、不複述。

| 文件 | 定位 | 資訊域 (MECE) |
|---|---|---|
| [狀態異常訊息.md](./狀態異常訊息.md) | 現行文件 (living) | ETL「狀態異常」之 `etlReason` 觸發條件與程式碼錨點（SSOT） |
| [qc-path-analysis/QC10002-R02-原物料品檢.md](./qc-path-analysis/QC10002-R02-原物料品檢.md) | 現行文件 (living) | QC10002-R02 四層映射鏈分析、風險點與改進建議（SSOT） |
| [handover_resume_guide.md](./handover_resume_guide.md) | 現行文件 (living) | 交接與重啟：專案狀態摘要、關鍵檔案索引、確效流程入口 |
| [today-requirements-2026-07-06.md](./today-requirements-2026-07-06.md) | 歷史快照 (historical) | 2026-07-06 當日需求記錄；狀態欄已以補註方式標示後續實作結果，不作為現行規格依據 |

維護規則：

1. 新增文件時必須在此表登錄（保持 MECE 窮盡）。
2. 現行事實（規格、行號錨點）只更新於對應之 SSOT 現行文件；`DEV_LOG.md` 僅記錄時序，不追改歷史行號。
3. 歷史快照文件不得改寫當日內容，僅可以「補註」方式更新狀態並註明日期。
# Direction Approved · STAGE 04 UI 方向決策

## 客戶篩選 UI（原 44 顆 pill）
- 展示三方向初稿：`design-demos/customer-filter.html`（A 按鈕+搜尋彈出層 / B 摘要膠囊分層 / C 客戶側欄清單）
- **用戶選定：方向 C** — 客戶側欄常駐清單：搜尋、勾選、筆數橫條（長度=訂單量）、底部即時覆蓋率（已選 X / N · 覆蓋 Y 筆 · %）
- 用戶原話：「客戶篩選我選方向C」

## 月份趨勢圖（原雙軸柱+線複合圖）
- 展示三方向初稿：`design-demos/trend-chart.html`（A 小倍數雙圖 / B 單一量綱+切換 / C 月份帶狀表）
- **用戶選定：方向 B** — 單一量綱 + 右上分頁切換（訂單數 / 數量萬pcs）；柱體帶資料標籤、Y 軸隱藏、無雙軸
- 用戶原話：「趨勢圖我選方向B」（截圖批註：雙軸複合呈現「很沒品味」）

## 實作記錄
- 兩案皆已實作進 `src/App.jsx` STAGE 04；decisions 見 DEV_LOG 當日條目。

# Project Evolution & Changelog

本文件紀錄 BOM to I need 工具的完整開發歷程。

## 🚀 v1.5：缺價主料警示標籤與單獨 Missing Price List 匯出 (Missing Price Warning & List Export)
**時間點：2026-05-30**

### 變更摘要
1. **單獨 Missing Price List 匯出**：
   - 點擊「一鍵打包下載 (ZIP)」時，系統會自動在壓縮包內新增一個名為 `Missing Price List.xlsx` 的檔案。
   - 彙整所有載入的 BOM 檔案中，缺失單價（單價為 0，且排除 `8B`/`8BB` 裝配階）的主料，並按 13 欄位標準格式匯出。
   - `A1` 儲存格記錄來源檔案名稱（不含副檔名）。
2. **卡片右上角缺主料警示**：
   - 若上傳的 BOM 檔案有任何主料缺失單價，會在檔案卡片右上方狀態標籤的下方，動態顯示 `⚠️ 缺主料價格 (N)` 的警示。
   - 警示會隨著使用者手動補登或「一鍵套用智慧單價」而即時遞減，直到所有缺價主料皆處理完畢後，警示會自動消失。

### ✅ 驗證步驟
- [x] 執行 `node verify.js` 驗證，轉換完全正常。
- [x] 網頁端補登及一鍵套用功能操作流暢，右上角警示數量同步更新且能在全部補齊時自動消失。

---

## 🚀 v1.4：被動元件分類精度優化與 VBA 程式更新 (Passive Categorization Optimization & VBA Sync)
**時間點：2026-05-30**

### 變更摘要
1. **電容分類 (CAP.) 規則優化**：
   - 將比對條件從寫死的 `2EC`, `2EG`, `2EL` 擴大為只要是被動電容 `2E` 開頭且排除 `2EV`（壓敏電阻）者。
   - 修正直插式電解電容如 `2ED012541SH8EB` (品名為 `CAP, AL`) 因前綴不符被誤判為「機構件/ Other」的缺陷。
2. **直插電阻/電容 (DIP Res./Cap.) 規則更新**：
   - 判斷清單中新增了直插電阻 `2R` 開頭的前綴。
   - 修正直插繞線電阻如 `2RD0008500GP` (品名為 `RES, WIREWOUND`) 被漏判為「機構件/ Other」的缺陷。
   - 邏輯保持互斥：含有 `SMD` 關鍵字的 `2R` 開頭貼片電阻會被 `!desc.includes('SMD')` 排除，正確流向 `SMD Res./Cap.`。
3. **VBA 代碼同步修改與註記**：
   - 針對 VBA 程序（`VBA.txt`）中的 `SUMPRODUCT` 篩選公式進行同步升級，確保導出 Excel 後的手動執行版公式與網頁版 Excel 導出公式維持 100% 相同，並加上詳細中文註解。
4. **物料分類邏輯文件**：
   - 於 `BOM_CLASSIFICATION.md` 獨立產出詳細的 13 大類物料分類規則與運作邏輯說明文件。

### ✅ 驗證步驟
- [x] 執行 `node verify.js` 驗證 `good.xls`、`good 2.xls` 及 `good 3.xls` 生產的 Excel 計算，3/3 完全通過。
- [x] 比對優化後的分類金額，原分類至 `機構件/ Other` 的繞線電阻（`2RD...`）與電解電容（`2ED...`）已被精準分流至 `DIP Res./Cap.` 與 `CAP.`，佔比更貼近實際。

---

## 🚀 v1.3：推薦料號顯示、一鍵套用與 UI 強化 (Recommended Part Display & UX Enhancement)
**時間點：2026-05-30**

### 變更摘要
1. **缺主料分析表新增「推薦料號」欄位**：
   - 在「智慧推薦單價」左側新增「推薦料號」欄，顯示來源 BOM 的完整料號。
   - 自動計算對比料號與目標料號的公共前綴，前綴相同頭 **≥ 6 碼**者以**藍色粗體**高亮顯示，便於快速辨識料號跟總中的差異。
2. **「⚡ 一鍵套用推薦單價」按鈕**：
   - 在「處理檔案清單」標題旁新增標志性橙色脈動按鈕，**僅在有實際可套用的推薦單價時才顯示**。
   - 點擊後一次批次將所有缺主料的智慧推薦單價填入，並自動觸發全底重算。
3. **處理檔案清單標題設計強化**：
   - 標題放大至 1.4rem、加上漸層色彩設計，搜對主題顏色。
4. **Drop Zone 說明文字更新**：
   - 新增說明文字：「將執行BOM表格式修正、價格計算、用料比例分析、缺料對比。」

### ✅ 驗證步驟
- [x] 執行 `node -e "new Function(script)"` 語法靜態驗證 SYNTAX OK。
- [x] 執行 `node verify.js`，3/3 通過。
- [x] 瀏覽器驗證：推薦料號欄位顯示正確，公共前綴藍色高亮正常。
- [x] 一鍵套用按鈕在有缺料時顯現，套用後自動消失。

---

## 🚀 v1.2：欄位位置對齊、智慧單價推薦與補登系統 (Alignment & Smart Pricing Recommendation)
**時間點：2026-05-30**

### 變更摘要
1. **Excel 元數據欄位對齊微調**：
   - 依據最新需求調整 BOM 頁面 metadata 排列位置，確保與原始格式及審查排版一致：
     - Row 4 (成品/半成品料號)：自 Column D 移回 **Column B (B4)**。
     - Row 7 (BOM Cost)：自 Column D 移回 **Column B (B7)**。
     - Row 8 (構成總筆數)：自 Column D 移回 **Column B (B8)**。
     - Row 9 (UOM)：自 Column D 移回 **Column B (B9)**。
     - Row 10 (品名規格)：自 Column D 移回 **Column B (B10)**。
     - Row 5 (BOM 版本) 與 Row 6 (Change NO) 保留在 **Column D (D5, D6)**。
2. **物料拆解分析工作表樣式調整**：
   - 將「物料拆解分析」工作表內所有單元格字型顏色改為預設（黑色），移除了原本全藍色的樣式。
3. **新增網頁端「缺主料價格分析」與「智慧推薦單價 (跨 BOM 掃描)」**：
   - 在網頁預覽明細中新增獨立的第四個子分頁「缺主料價格分析」，用以即時檢視單價為 0 的主要零件（非 PCBA 板或 8B/8BB 組裝件）。
   - 整合跨 BOM 掃描的推薦演算法：
     - **精確推薦**：在當前載入的其它 BOM 中，若有完全相同的料號且單價大於 0，則推薦該單價。
     - **相似推薦**：若存在前綴相同（前五碼或更多匹配）且單價大於 0 的料號（通常為版次遞增累進的料號），則推薦最近版本的相似單價。
   - 支援在網頁上點擊「套用」推薦單價，或手動「補登價格」；套用/補登後系統將即時重新計算整份 BOM、大分類及物料拆解的總額與佔比，並在下載導出時將修正後的單價直接寫入 Excel。


### 🐛 Bug 修復紀錄

#### Bug #1：`findMetaRow` 重複宣告導致整個 Script 無法執行
- **症狀**：網頁完全無反應——上傳按鈕點擊、拖曳檔案均失效，不產生任何錯誤提示。
- **根本原因**：`processBOMWorkbook` 函式內部有兩個 `const findMetaRow = ...` 宣告（分別位於函式頭部及 `// Rebuild metadata layout` 區段），ES6 嚴格規則中同一 scope 不允許重複 `const` 宣告，瀏覽器在解析階段就拋出 `ReferenceError: Identifier 'findMetaRow' has already been declared`，導致整個 `<script>` 區塊完全不執行。
- **為何難以察覺**：頁面外觀正常、Console 不一定明顯顯示（需主動開啟 DevTools 才看到），且 `node verify.js` 使用 Node.js 執行，不受瀏覽器嚴格 const 作用域影響，故驗證通過但瀏覽器仍失效。
- **修復**：刪除第二個重複的 `const findMetaRow` 宣告（原 line ~1317）。
- **防範**：今後凡是在函式內 hoist 或移動宣告時，必須同步刪除原位置的舊宣告。

#### Bug #2：三層巢狀 Template Literal 造成 JS 解析失敗
- **症狀**：與 Bug #1 相同——整個 script 無法執行。
- **根本原因**：`renderResults()` 中 `cardHtml` 大型 template literal 內，missing tab 那段使用了 `${missingItems.map(item => { ... return \`...${rec.score}...\`; }).join('')}`，造成三層反引號嵌套（外層 cardHtml、中層 ternary template、內層 map callback return template）。JavaScript 解析器對於「map callback 內又有 if/let/const/return + 第三層 template literal」的組合產生解析歧義，引發靜默錯誤。
- **修復**：將 missing rows 渲染邏輯抽出至獨立函式 `renderMissingRowsHtml()`，改用字串串接（`+`），完全消除巢狀 template literal。
- **防範**：凡是大型 template literal 內的 `${}` 中需要多行邏輯（if、let、const、巢狀 template），一律抽成獨立函式再呼叫。

### ✅ 驗證步驟
- [x] 執行 `node verify.js`，3/3 測試通過，Excel 三個分頁驗證正確。
- [x] 執行 `node -e "new Function(script)"` 語法靜態驗證，確認無重複宣告或解析錯誤。
- [x] 瀏覽器開啟 `http://localhost:8000/index.html`，上傳 BOM 檔案正常，分頁切換正常，缺主料分析顯示正確。
- [x] 跨 BOM 推薦：同時載入多個 BOM 後，缺主料分析頁顯示精確或相似推薦，點擊「套用」後即時重算總價。
- [x] 修改匯率後自動觸發 `recalculateAll()`，所有卡片數字同步更新。

---

## 🚀 v1.1：樣式、匯率與預覽優化版本 (Style & Preview Enhancement)
**時間點：2026-05-29**

### 變更摘要
1. **Excel 樣式與色彩對齊**：
   - 整合 `xlsx-js-style` 套件，導出時完美支援儲存格高亮。
   - BOM 頁面金額為 0 且非 8B/8BB PCBA 時儲存格填滿黃色 (`#FFFF00`)。
   - BOM 資料列依料號前綴標註字型顏色：`8B*` 標為灰色 (`#A9A9A9`)、`4W*` 標為藍色 (`#0000FF`)、`5CC*`/`3FH*`/`5OG*` 標為橘色 (`#FFA500`)。
   - 「物料拆解分析」頁面儲存格字型全面改為藍色 (`#0000FF`)。
   - BOM 頁面元數據區 E4、G4、L10 指標文字標為粗體與紅色 (`#FF0000`)。
   - 導出的試算表字型設定：中文文字使用「微軟正黑體」、英文/數字使用「Arial」。
2. **全域匯率限制與儲存功能**：
   - 全域美金匯率輸入欄位精確至小數點後兩位，並透過 `localStorage` 機制記憶最後一次調整的價格。
3. **網頁預覽分頁重構**：
   - 修復「大分類統計」與「物料拆解分析」分頁在 UI 切換時的隱藏顯示問題。
   - 網頁預覽統計表與分析表新增「百分比」欄位及「合計列 (Total Row)」。
   - 已整理的 BOM 檔案上傳時，網頁預覽會動態讀取明細並計算產出統計與拆解分析。
4. **細節微調與進版**：
   - 網頁解析成功標記文字調整為「已完成BOM解析」。
   - 打包下載的檔名維持與上傳檔名相同，取消 `_fixed` 後綴。
   - 專案版本號全面進版至 v1.1，並配置全新現代簡潔風的 favicon 圖示。

---

## 🚀 v1.0：初始版本 (Initial Release)
**時間點：2026-05-29**

### 變更摘要
1. **多 BOM 批次處理系統**：支援拖放上傳或點擊選取多個 BOM 檔案（CSV/XLS/XLSX）。
2. **檔案類型識別與篩選**：
   - 自動檢測檔案狀態：`processed`（已整理）、`fixed`（已修正偏移的原始 BOM）、`raw_normal`（格式正常的原始 BOM）、`error`（損毀或非 BOM 格式）。
   - 提供「略過已整理的 BOM 檔案」過濾開關，勾選後下載的 ZIP 包自動排除已整理檔案。
3. **格式與偏移自動修正**：
   - 自動將標題列定位至 Row 12（0-indexed Row 11）。
   - 自動修正移位之子料行，將移位至 Column 11-14 的資料正確移回 Column 9-12。
   - 重建 Rows 4-10 (0-indexed 3-9) 的元數據佈局，值統一移至 Column D 以防止溢出截斷。
4. **I Need 格式重構**：
   - 插入 Column L 「乘積」欄位，公式為 `=IFERROR(F13*K13, 0)`。
   - 移除原 BOM 工作表右側的統計摘要區，確保 BOM 明細乾淨。
5. **獨立統計與拆解分頁 (New feature)**：
   - 新增分頁「大分類統計」：整理 8 大類（1xxx 主動至 8xxx 磁性元件）的費用與佔比，配備 `SUMIFS` 函數與預計算值。
   - 新增分頁「物料拆解分析」：整理 13 大類（DIODE、IC、MOS、CAP 等）的拆解統計，配備與 VBA 對應的 `SUMPRODUCT` 跨表公式與預計算值。
6. **網頁即時互動預覽**：
   - 使用者可在 UI 上直接展開每個檔案的整理成果。
   - 提供三個子分頁切換：BOM 明細（附帶黃/灰/藍/橘高亮色彩規則）、大分類統計表、物料拆解分析表，以及詳細的偏移欄位「修正日誌」。
7. **本機化無網部署**：
   - 整合本機 SheetJS (`xlsx.full.min.js`) 與 JSZip (`jszip.min.js`)。

### 驗證步驟
- [x] 無網路下開啟 `index.html`，UI 顯示正常，元件讀取無誤。
- [x] 上傳 `good.xls`、`good 2.xls` 及 `good 3.xls`，系統標示為 `raw_normal`，並能預覽與導出。
- [x] 下載 ZIP 包，使用 Excel 開啟，分頁 `BOM`、`大分類統計`、`物料拆解分析` 結構正確，跨表公式求值完全正常。
- [x] 執行本地 Node.js 驗證腳本 `node verify.js`，全部測試通過。

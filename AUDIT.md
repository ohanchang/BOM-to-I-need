# Project Audit & Technical Debt

本文件追蹤專案中的程式碼品質、使用性問題以及未來的優化點。

---

## 評分標準
| 符號 | 難度 | 必要性 |
|------|------|--------|
| ⭐ | 簡單 | 可有可無 |
| ⭐⭐ | 中等 | 建議做 |
| ⭐⭐⭐ | 較高 | 重要 |
| ⭐⭐⭐⭐ | 極高 | 必須做 |

---

## 🚀 待處理項目 (Pending)

### 1. 📊 成本佔比圖表視覺化
- **描述**：網頁預覽中僅提供表格，若能提供大分類與物料拆解佔比的圓餅圖或柱狀圖，視覺效果會更佳。
- **建議**：拷貝 `BOM Compare` 的 Chart.js 並在預覽區整合圓餅圖。
- | 難度 | 必要性 |
  |------|--------|
  | ⭐⭐ | ⭐ |

---

## ✅ 已完成改進 (Resolved)

| # | 分類 | 項目 | 狀態 |
|---|------|------|------|
| 1 | 功能 | 批次上傳與 ZIP 封包匯出 | ✅ 完成 |
| 2 | 功能 | 原始/已整理類型自動偵測與略過設定 | ✅ 完成 |
| 3 | 功能 | 標題與移位子料欄位偏移修正 | ✅ 完成 |
| 4 | 功能 | 獨立「大分類統計」跨表 `SUMIFS` 工作表 | ✅ 完成 |
| 5 | 功能 | 獨立「物料拆解分析」跨表 `SUMPRODUCT` 工作表 | ✅ 完成 |
| 6 | UI | 網頁端即時資料格切換預覽與修正日誌 | ✅ 完成 |
| 7 | 本機化 | 離線化庫依賴整合 (SheetJS & JSZip) | ✅ 完成 |
| 8 | 樣式 | 導出 Excel 單元格色彩與字型格式套用 (`xlsx-js-style`) | ✅ 完成 |
| 9 | UI | 全域匯率小數點後兩位限制與 `localStorage` 記憶功能 | ✅ 完成 |
| 10 | UI | 網頁端大分類/物料拆解表格百分比欄位與合計列 | ✅ 完成 |
| 11 | 功能 | 已整理 BOM 檔案在網頁預覽時動態大分類與物料拆解計算 | ✅ 完成 |
| 12 | 功能 | Excel 導出元數據 D4/D7-D10 移回 B4/B7-B10 | ✅ 完成 |
| 13 | 功能 | 「物料拆解分析」工作表樣式顏色回復預設黑色 | ✅ 完成 |
| 14 | 功能 | 網頁預覽新增「缺主料價格分析」與「跨 BOM 智慧單價推薦」及即時補登 | ✅ 完成 |

---

## 💡 未來創意發想 (Roadmap)
- **智慧規格分析**：自動分析規格項目中的關鍵描述，並對异常的物料分類給予警示（如品名寫 Diode 但料號非 1D 起頭）。
- **報價有效期限追蹤**：對於報價日期超過 180 天的料件給予高亮提示。

---

## 🔍 Debug 經驗與排錯技巧

本章節記錄開發過程中遇到的重要 Bug，以及如何有效定位與排查。

### 案例一：頁面完全無反應（按鈕、拖曳均失效）

**現象描述**
- 頁面外觀、樣式、版面顯示完全正常
- 點擊上傳按鈕或拖曳檔案到 Drop Zone 毫無反應
- 沒有任何視覺錯誤提示

**錯誤診斷流程**

```
步驟 1：懷疑是 HTML 結構問題
→ 用 Node.js 確認 <script> tag 數量與閉合

步驟 2：懷疑是 JS 語法錯誤
→ node -e "new Function(script)"  ← 語法靜態驗證
→ 結果：有語法錯誤但 verify.js 沒報（因為 Node.js 與瀏覽器作用域規則不同！）

步驟 3：打開瀏覽器 DevTools (F12) → Console
→ 看到 "Identifier 'findMetaRow' has already been declared"

步驟 4：搜尋重複宣告
→ 在 processBOMWorkbook 內發現兩個 const findMetaRow = ...
→ 刪除重複宣告，問題解決
```

**根本原因**：同一函式 scope 中重複 `const` 宣告，瀏覽器在解析（parse）階段就拋出 ReferenceError，整個 `<script>` 區塊從未執行。

**教訓**
1. `node verify.js` ≠ 瀏覽器驗證。Node.js 執行模式與瀏覽器嚴格解析規則有差異
2. 頁面外觀正常 ≠ Script 正常。必須主動開 DevTools Console
3. 凡是移動或 hoist 宣告，**原位置的舊宣告必須同步刪除**

---

### 案例二：深層巢狀 Template Literal 靜默失敗

**現象描述**
- 症狀與案例一相同（完全無反應）
- 但 `node -e "new Function(script)"` 顯示 SYNTAX OK
- 仍然失效

**根本原因**：三層巢狀 template literal + map callback 內含多行邏輯

```javascript
// 危險寫法（三層巢狀）
cardHtml = `
  ${items.map(item => {
    const label = `相似 (${item.score} 碼)`;   // ← 第3層！
    return `<td>${label}</td>`;                 // ← 第3層！
  }).join('')}
`;
```

部分 JS 引擎在解析這種結構時會靜默失敗（不報錯但不執行）。

**修復方式**：抽出獨立函式，改用字串串接

```javascript
// 安全寫法：獨立函式 + 字串串接
function renderRowHtml(item) {
  const label = '相似 (' + item.score + ' 碼)';
  return '<td>' + label + '</td>';
}

// Template literal 中只呼叫函式
cardHtml = `${items.map(renderRowHtml).join('')}`;
```

**防範規則**：凡大型 template literal 的 `${...}` 中需要 `if`、`let`、`const`、`return` 或嵌套 template literal，**一律抽成獨立函式**。

---

### 快速 Debug 工具箱

| 問題類型 | 工具指令 | 說明 |
|---------|---------|------|
| 確認 JS 語法 | `node -e "const fs=require('fs');const html=fs.readFileSync('index.html','utf8');const s=html.match(/<script>([\s\S]*?)<\/script>/)[1];try{new Function(s);console.log('OK')}catch(e){console.log(e.message)}"` | 靜態語法驗證 |
| 確認 script tag 數量 | `node -e "const html=require('fs').readFileSync('index.html','utf8');console.log([...html.matchAll(/<script/g)].length,'script tags')"` | 排查多 tag 問題 |
| 尋找重複宣告 | 在腳本文字中搜尋 `const xxx` 並比對次數 | 定位 already declared |
| 確認函式是否載入 | 在 DevTools Console 輸入 `typeof functionName` | undefined = script 未執行 |
| 強制重整頁面 | `Ctrl+Shift+R`（Windows）或 `Cmd+Shift+R`（Mac） | 清除快取重載 |

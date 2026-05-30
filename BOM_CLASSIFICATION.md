# BOM 物料拆解分類邏輯文件 (BOM Classification Rules)

本文件詳細記錄 BOM 成本進階分析工具中將物料自動拆解為 **13 大類** 的分類邏輯。此邏輯同步實作於網頁前端 JavaScript 預覽與導出的 Excel `物料拆解分析` 工作表公式（`SUMPRODUCT`）中。

---

## 13 大類分類定義與規則

分類判斷依據主要依循兩個核心欄位：
1.  **料號 (Part No.)**：英文與數字首置前綴碼。
2.  **規格項目-名稱說明 (Description)**：品名中是否含有特定關鍵字（如 `SMD`、`FAN`、`WIRE` 等）。

以下為各分類的詳細匹配邏輯（依優先序排定，採互斥設計）：

| 類別名稱 (Category) | 匹配規則 (JS 判斷) | Excel SUMPRODUCT 公式結構 | 說明 |
| :--- | :--- | :--- | :--- |
| **1. DIODE** | 料號開頭為 `1D` | `LEFT(PartNo, 2)="1D"` | 主動元件：二極體 |
| **2. IC** | 料號開頭為 `1I`、`1S` 或 `1OT` | `(LEFT(PartNo,2)="1I")+(LEFT(PartNo,2)="1S")+(LEFT(PartNo,3)="1OT")>0` | 主動元件：積體電路與光電耦合器 |
| **3. MOS/ TR.** | 料號開合為 `1T` 或 `1M` | `(LEFT(PartNo,2)="1T")+(LEFT(PartNo,2)="1M")>0` | 主動元件：電晶體、MOSFET |
| **4. CAP.** | 料號開頭為 `2E` 且排除 `2EV`，且品名不包含 `SMD` | `(LEFT(PartNo,2)="2E")*(LEFT(PartNo,3)<>"2EV")*ISERR(SEARCH("SMD",Desc))` | 被動元件：直插式電解電容與固態電容 (如 `2EC`, `2ED`, `2EG`, `2EL` 等) |
| **5. DIP Res./Cap.** | 料號開頭為 `2CX`, `2CY`, `2EV`, `2FN`, `2FS`, `2OP`, `2VR`, `3RD`, `3UC` 或 `2R` 之一，且品名不包含 `SMD` | `(LEFT(PartNo,3)="2CX")+(...)+(LEFT(PartNo,2)="2R")>0*ISERR(SEARCH("SMD",Desc))` | 直插被動元件：直插電阻（如繞線電阻 `2RD`）、安全電容、壓敏電阻、熱敏電阻等 |
| **6. SMD Res./Cap.** | 料號開頭為 `2` 且品名包含 `SMD` | `(LEFT(PartNo,1)="2")*ISNUMBER(SEARCH("SMD",Desc))` | 貼片被動元件：貼片電阻、貼片電容 |
| **7. Magnetic** | 料號開頭為 `7` 或 `8`（排除 `8B`） | `((LEFT(PartNo,1)="7")+(LEFT(PartNo,1)="8")>0)*(LEFT(PartNo,2)<>"8B")` | 磁性元件：變壓器、電感、磁芯等 |
| **8. PCB** | 料號開頭為 `3B` 或 `8B` | `(LEFT(PartNo,2)="3B")+(LEFT(PartNo,2)="8B")>0` | 印刷電路板與組裝基板 (PCBA) |
| **9. FAN** | 品名包含 `FAN` 且不包含 `FAN GUARD`，且料號開頭不是 `1` | `ISNUMBER(SEARCH("FAN",Desc))*ISERR(SEARCH("FAN GUARD",Desc))*(LEFT(PartNo,1)<>"1")` | 散熱風扇（排除風扇鐵網與主動控制 IC） |
| **10. FAN GUARD** | 品名包含 `FAN GUARD` | `ISNUMBER(SEARCH("FAN GUARD",Desc))` | 風扇防護鐵網 |
| **11. CASE** | 品名包含 `CASE` 或 `CHASSIS` | `ISNUMBER(SEARCH("CASE",Desc))+ISNUMBER(SEARCH("CHASSIS",Desc))>0` | 五金外殼與底盤 |
| **12. O/P WIRE** | 料號開頭為 `4W` 或品名包含 `WIRE`（排除品名含 `GUARD` 項目） | `((LEFT(PartNo,2)="4W")+ISNUMBER(SEARCH("WIRE",Desc))>0)*ISERR(SEARCH("GUARD",Desc))` | 輸出線材與配線 |
| **13. 機構件/ Other** | 所有不符合上述 1-12 類規則的項目 | `BOM總額 - SUM(其他12類金額)` | 包含開關、插座、螺絲、散熱片、絕緣片、膠水、標籤貼紙等 |

---

## 精確分類邏輯代碼對照 (JavaScript)

在系統中，BOM 每一列的分類判斷代碼如下：

```javascript
let cat = '機構件/ Other';

if (partNo.startsWith('1D')) {
  cat = 'DIODE';
} else if (partNo.startsWith('1I') || partNo.startsWith('1S') || partNo.startsWith('1OT')) {
  cat = 'IC';
} else if (partNo.startsWith('1T') || partNo.startsWith('1M')) {
  cat = 'MOS/ TR.';
} else if (partNo.startsWith('2E') && !partNo.startsWith('2EV') && !desc.includes('SMD')) {
  cat = 'CAP.';
} else if (['2CX', '2CY', '2EV', '2FN', '2FS', '2OP', '2VR', '3RD', '3UC', '2R'].some(p => partNo.startsWith(p)) && !desc.includes('SMD')) {
  cat = 'DIP Res./Cap.';
} else if (partNo.startsWith('2') && desc.includes('SMD')) {
  cat = 'SMD Res./Cap.';
} else if ((partNo.startsWith('7') || partNo.startsWith('8')) && !partNo.startsWith('8B')) {
  cat = 'Magnetic';
} else if (partNo.startsWith('3B') || partNo.startsWith('8B')) {
  cat = 'PCB';
} else if (desc.includes('FAN') && !desc.includes('FAN GUARD') && !partNo.startsWith('1')) {
  cat = 'FAN';
} else if (desc.includes('FAN GUARD')) {
  cat = 'FAN GUARD';
} else if (desc.includes('CASE') || desc.includes('CHASSIS')) {
  cat = 'CASE';
} else if ((partNo.startsWith('4W') || desc.includes('WIRE')) && !desc.includes('GUARD')) {
  cat = 'O/P WIRE';
}
```

---

## 2026/05/30 分類邏輯優化更新紀錄

### 1. CAP. (電容類別) 擴大兼容
*   **調整前**：僅精確比對 `2EC`、`2EG`、`2EL`。
*   **調整後**：改為 `2E` 開頭（排除壓敏電阻 `2EV`）。
*   **原因說明**：解決了如 `good 3.xls` 中直插電解電容 `2ED012541SH8EB` (品名為 `CAP,AL`) 被錯判為 `機構件/ Other` 的問題。

### 2. DIP Res./Cap. (直插電阻) 補齊前綴
*   **調整前**：前綴比對名單未包含 `2R`。
*   **調整後**：在比對名單中加入 `2R`。
*   **原因說明**：直插式繞線電阻或水泥電阻（如 `good.xls` 中之 `2RD0008500GP`，品名為 `RES,WIREWOUND`）以 `2RD` 開頭，之前由於不含 `SMD` 關鍵字，也落入無前綴匹配的 `機構件/ Other`，現已正確分類。

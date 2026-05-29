# ACHB RFQ 製程估價系統

ACHB 內部使用的報價估算工具，純前端單頁 HTML，不需要伺服器或安裝任何軟體。

## 使用方式

直接用瀏覽器開啟 `rfq-estimate-V3.5.html` 即可使用，無需網路連線。

## 功能

- **單件模式**：單一零件多製程途層估價，支援廠內加工、廠外委外、外購物料
- **組合件模式**：多工件組合估價，每個工件可設定每組用量（pcs/set）
- 加工費率依 2026-05-27 最新機台成本自動計算
- 多數量級距同時報價（可自訂 qty tiers）
- 每個級距支援個別毛利設定
- 匯出 Excel（SheetJS）、JSON 儲存與載入
- 匯率輸入，自動換算 NT$/USD

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `rfq-estimate-V3.5.html` | 最新版本（目前使用版本） |
| `RFQ_estimate_CHANGELOG.html` | 版本更新紀錄 |
| `RFQ_estimate_SOP.html` | 使用說明 SOP |

## 版本

目前版本：**V3.5**　　詳細更新紀錄請見 `RFQ_estimate_CHANGELOG.html`

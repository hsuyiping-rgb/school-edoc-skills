# build_report.js 的資料格式

寫一個 JSON，然後 `node scripts/build_report.js <資料.json> <輸出.docx>`。

版面、配色、字型、表格框線、分頁都由腳本處理。你只要負責內容。

## 行內標記

任何內文字串都可以用：

| 寫法 | 效果 |
|---|---|
| `**文字**` | 粗體深色，用來標出承辦人要動手改的關鍵字 |
| `` `採購法 §26 III` `` | 等寬深綠，法條專用 |

其他一律是內文。不要在字串裡放 `\n`——需要換行就拆成陣列的多個元素。

## 骨架

```json
{
  "meta": {
    "title": "115年度自然科學教室改善及教具採購（範例）\n招標文件審查報告",
    "subtitle": "新北市○○區○○國民小學．依政府採購法及其子法所為之書面審查",
    "fileNo": "115/0000/1",
    "caseNo": "11500",
    "reviewDate": "115.03.03",
    "creator": "○○國小總務處"
  },
  "facts": [
    ["預算金額", "NT$880,000"],
    ["金額級距", "逾公告金額1/10、未達公告金額"],
    ["招標方式", "公開取得企劃書（§49）"]
  ],
  "verdict": {
    "heading": "審查結論：不宜依現行文件上網，應先補正",
    "paragraphs": ["招標方式與級距均相符⋯但有四項⋯"],
    "tally": [["high", 4], ["fix", 5], ["advice", 4], ["ok", 7]]
  },
  "sections": [ ... ],
  "closing": {
    "heading": "建議處理順序",
    "items": ["先做**分項預算概估**⋯", "重寫**需求說明書與估價單**⋯"],
    "fill": "EDEFEC"
  },
  "footer": "本審查係就所附招標文件之書面內容為之⋯"
}
```

`meta.title` 用 `\n` 分行（這是唯一容許 `\n` 的地方）。`facts` 每列三格，不足會自動補空。`tally` 的鍵只能是 `high` / `fix` / `advice` / `ok`。

## sections 的元素

每個 section 可以帶下列任意組合，依序輸出：

```json
{
  "pageBreak": true,
  "heading": "壹、重大爭議　（上網前必須處理）",
  "note": "下列各點涉及法規適用錯誤或限制競爭之虞⋯",
  "paragraphs": ["一般段落，支援行內標記"],
  "findings": [ ... ],
  "table": { ... },
  "listTable": { ... },
  "numbered": ["編號段落"],
  "bullets": ["項目符號段落"],
  "afterParagraphs": ["表格之後的收尾段落"],
  "callout": { ... },
  "spacerAfter": 60
}
```

輸出順序是固定的（paragraphs → findings → table → listTable → numbered → bullets → afterParagraphs → callout）。要把話說在表格**後面**，用 `afterParagraphs`，不要為此另開一個沒有標題的 section——那會讓分頁很難看。

### findings（審查發現）

```json
{
  "id": "A1",
  "severity": "high",
  "title": "採購性質認定為「財物」可能有誤，連帶影響廠商資格與契約範本",
  "blocks": [
    { "label": "事實", "text": "估價單 23 項中，第 1、2、3、4、23 項為⋯" },
    { "label": "法據", "text": "`採購法 §7 I`「工程」指⋯；`§7 IV` 兼有性質者⋯" },
    { "label": "連動", "bullets": ["第一點⋯", "第二點⋯"] },
    { "label": "建議", "text": "上網前先做出**分項預算概估表**⋯" }
  ]
}
```

`severity` 是 `high` / `fix` / `advice` / `ok`。`blocks` 的 `label` 自訂——常用「事實／法據/依據／問題／連動／試算/清單／另涉／建議」。label 含「建議」「處理」「作法」者會自動用主色標示，因為那是承辦人真正要看的部分。

一個 block 可以只有 `text`、只有 `bullets`，或兩者都有（`text` 當引言）。

### table（一般表格，如時程檢核）

```json
{
  "header": ["日期", "原訂程序", "檢核", "說明"],
  "widths": [1560, 2100, 1100, 4570],
  "mono": [0],
  "rows": [
    ["3/5（四）", "刊登公告", "起算日", "依`招標期限標準 §11`，公告當日算入"]
  ]
}
```

`widths` 要加總為 9330。`mono` 是要用等寬字的欄位索引（日期、金額）。`keyCol` 可指定某欄為淺底標題欄。

### listTable（左窄右寬、右欄條列）

規格補正表與共通條款用這個。

```json
{
  "header": ["項次／品名與數量", "應補列之規格內容"],
  "widths": [1900, 7430],
  "monoKey": true,
  "rows": [
    {
      "key": "3\n防撞牆墊\n200 片",
      "bullets": ["必須載明單片尺寸⋯", "構造分層與厚度⋯"]
    },
    { "key": "同等品", "text": "本表所列規格為最低要求⋯" }
  ]
}
```

`key` 用 `\n` 分行，第一行會突顯。`monoKey: true` 時第一行用等寬主色（適合項次編號）；共通條款那種純文字鍵就不要開。`bullets` 與 `text` 擇一。

已符合清單也用這個：`{"key": "§49", "text": "採公開取得企劃書辦理⋯"}` 搭配 `monoKey: true`。

### callout（強調方塊）

```json
{
  "heading": "建議處理順序",
  "paragraphs": ["段落"],
  "items": ["編號項目"],
  "tally": [["high", 4]],
  "fill": "EDEFEC",
  "topColor": "A32A21",
  "keepTogether": true
}
```

`topColor` 會畫一條粗色頂線（結論方塊用紅色）。`keepTogether: true` 讓方塊不被切在頁面中間——收尾的處理順序方塊建議開，並在該 section 加 `pageBreak`。

## 兩份文件的慣用結構

**審查報告**

```
meta + facts + verdict
壹、重大爭議        findings（severity: high）
貳、應補正事項      findings（severity: fix）
參、建議事項        findings（severity: advice）
肆、時程合規檢核    table
伍、已符合規定之處  listTable（monoKey）
closing（建議處理順序）+ footer
```

**規格補正建議表**

```
meta
壹、問題確認        paragraphs
貳、重點品項規格表  listTable（monoKey: false，左欄是「處理器」「記憶體」這類項目名）
                   ＋ afterParagraphs 放替代作法之類的補充
參、其餘品項規格補正 listTable（monoKey，左欄是「3\n防撞牆墊\n200 片」）
肆、共通條款        listTable（monoKey: false）
伍、須先行確認事項  numbered
footer
```

兩份文件都有現成範例可以照抄：`assets/example-review.json`、`assets/example-spec.json`。

## 產出後

一定要跑 `python scripts/render_check.py <輸出.docx>` 並讀輸出的 PNG。常見要修的：

- **`widths` 加總不等於 9330** → 欄寬跑掉
- **收尾方塊被拆成兩頁** → `keepTogether: true`，必要時在該 section 加 `pageBreak`
- **`pageBreak` 造成整頁空白** → 前一節只溢出一兩行時，硬分頁會浪費一整頁。拿掉 `pageBreak` 讓它自然流動，通常比較好看

分頁是看了 render 才知道的事，不要憑空加 `pageBreak`。

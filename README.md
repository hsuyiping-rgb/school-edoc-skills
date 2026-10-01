# 學校公文核閱與採購審查技能（Claude Code Skills）

給新北市立學校校長、主任使用的三個 Claude Code 技能，協助核閱二代公文系統的待核公文，並在遇到採購簽時自動做法規審查。Agent 只當幕僚：讀文、擬意見、整理風險，**決行與送出一律由校長本人確認**。

| 技能 | 用途 | 觸發說法 |
|---|---|---|
| `ntpc-edoc-review` | 核閱新北市二代公文系統待核公文：讀文與附件、依《新北市政府文書處理要點》分流、擬批示意見、經校長逐件同意後送出 | 「核閱公文」「批公文」「待核公文」 |
| `tender-review` | 上網公告前審查招標文件（投標須知、契約、需求說明書、估價單、監辦通知單），產出 Word 審查報告 | 「審查這個招標案」「這份文件可以上網了嗎」 |
| `eng-procurement-review` | 工程採購各階段（招標、決標、變更設計、驗收、結算）的簽核判讀，給「可簽／緩簽／退回」與批示文字 | 「這個工程案我可以簽嗎」「驗收可以簽了嗎」 |

`ntpc-edoc-review` 核閱到採購簽時，會自動調用另外兩個技能。

## 安裝

把三個資料夾複製到 `~/.claude/skills/` 下即可（Windows 為 `%USERPROFILE%\.claude\skills\`）。

需要的環境：

- Claude Code，以及 Claude in Chrome 擴充功能（`ntpc-edoc-review` 用它操作公文系統）
- Python 3：`pypdf`、`PyMuPDF`（fitz）、`python-docx`、`pandas`、`openpyxl`；Windows 讀 `.doc` 另需 `pywin32`
- Node.js：`docx` 套件（`tender-review` 產 Word 報告用）
- 讀取舊版 `.doc`：Windows 需有 Microsoft Word，其他平台需 LibreOffice

## 使用前要改的地方

- `ntpc-edoc-review/SKILL.md`「參考」一節：改成你存放《新北市政府文書處理要點》條文 PDF 的路徑。
- 查名冊時的過濾關鍵字：`read_attachments.py --grep <本校校名>`。
- 審查報告的存放位置：`SKILL.md` 第 4 步的 `<工作區>/採購案/`。
- 腳本內的 `DOC_NO = '0000000000'` 是佔位值，執行時由 Agent 替換成實際文號。

## 護欄

- 登入、驗證碼、自然人憑證 PIN 一律由使用者本人輸入，Agent 不代填。
- 「決行」「退承辦人」按下即送出、無二次確認，Agent 必須逐件取得同意才送。
- 公文內容與個資不寫入任何檔案；下載的附件在核閱完成後移到資源回收筒。
- 重大案件（霸凌、性平、人事、公告金額以上採購等）只整理案情與選項，不代為決定。

## 限制與免責

- 系統操作方法依 2026 年 9 月實測的新北市二代公文系統（SPEED Super Desk）版面撰寫，系統改版後腳本可能失效。
- 法規內容依撰寫時的版本整理，採購規範與契約範本會修正；`eng-procurement-review/references/pending-changes.md` 列有需追蹤的版本。引用條文前請以全國法規資料庫及新北市政府公告原文為準。
- 審查結果只就書面文件判讀，不含現場勘查與預算合理性評估，不構成法律意見。
- `tender-review/assets/` 的範例報告已去識別化，校名與案號為虛構。

## 授權

MIT License，詳見 [LICENSE](LICENSE)。

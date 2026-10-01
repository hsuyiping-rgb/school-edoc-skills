#!/usr/bin/env python3
"""把一整批招標文件轉成 UTF-8 純文字。

支援 .docx / .doc / .pdf / .xlsx / .xls / .ods / .csv / .txt。

.doc 需要 Word（Windows）或 LibreOffice。公文 PDF 常是 Big5 CID 編碼，
PyMuPDF 能正確取出，但 stdout 若為 cp950 會再次亂碼——所以一律寫檔，不印內容。

  python extract_docs.py <資料夾或檔案> [-o 輸出資料夾]
"""
import argparse
import os
import subprocess
import sys
import tempfile

EXTS = {".docx", ".doc", ".pdf", ".xlsx", ".xls", ".ods", ".csv", ".txt"}


def w(path, text):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def from_docx(path):
    import docx
    d = docx.Document(path)
    parts = [p.text for p in d.paragraphs]
    for i, tbl in enumerate(d.tables):
        parts.append(f"\n--- 表格 {i + 1} ---")
        for row in tbl.rows:
            parts.append(" | ".join(c.text.replace("\n", " ").strip() for c in row.cells))
    return "\n".join(parts)


def from_pdf(path):
    import fitz
    doc = fitz.open(path)
    out = []
    for i, page in enumerate(doc):
        out.append(f"=== 第 {i + 1} 頁 ===")
        out.append(page.get_text())
    return "\n".join(out)


def from_sheet(path):
    """試算表 -> 文字。優先 openpyxl（不需 pandas），其次 csv，最後才試 pandas。
    舊版 .xls（BIFF/OLE2）由 extract() 先呼叫 sheet_to_xlsx 轉檔後再進本函式。"""
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        import csv as _csv
        rows = None
        for enc in ("utf-8-sig", "cp950", "utf-16"):
            try:
                with open(path, newline="", encoding=enc) as f:
                    rows = list(_csv.reader(f))
                break
            except Exception:
                rows = None
        if not rows:
            return None
        out = []
        for row in rows:
            cells = [str(c).strip() for c in row]
            if any(cells):
                out.append(" | ".join(cells))
        return "\n".join(out)

    if ext == ".xlsx":
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            out = []
            for name in wb.sheetnames:
                ws = wb[name]
                out.append(f"=== 工作表：{name} ===")
                for row in ws.iter_rows(values_only=True):
                    cells = ["" if c is None else str(c).strip() for c in row]
                    if any(cells):
                        out.append(" | ".join(cells).rstrip(" |"))
            wb.close()
            return "\n".join(out)
        except Exception as e:
            print(f"  openpyxl 讀取失敗（{e}），改試 pandas", file=sys.stderr)

    try:
        import pandas as pd
    except ImportError:
        return None
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None)
    except Exception:
        return None
    out = []
    for name, df in sheets.items():
        out.append(f"=== 工作表：{name} ===")
        for _, row in df.iterrows():
            cells = ["" if str(v) == "nan" else str(v).strip() for v in row.tolist()]
            if any(cells):
                out.append(" | ".join(cells))
    return "\n".join(out)


def sheet_to_xlsx(path, tmpdir):
    """舊版 .xls（BIFF/OLE2）或 .ods -> .xlsx。openpyxl 只讀 xlsx，
    pandas 讀 .xls 需 xlrd 而 xlrd 2.x 已移除該支援，故用 Excel COM，其次 LibreOffice。"""
    out = os.path.join(tmpdir, os.path.splitext(os.path.basename(path))[0] + ".conv.xlsx")
    if sys.platform == "win32":
        try:
            import win32com.client as wc
            app = wc.Dispatch("Excel.Application")
            try:
                app.Visible = False
                app.DisplayAlerts = False
            except Exception:
                pass
            try:
                wb = app.Workbooks.Open(os.path.abspath(path), 0, True)
                wb.SaveAs(out, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
                wb.Close(False)
                if os.path.exists(out):
                    return out
            finally:
                try:
                    app.Quit()
                except Exception:
                    pass
        except Exception as e:
            print(f"  Excel COM 失敗（{e}），改試 LibreOffice", file=sys.stderr)
    for exe in ("soffice", "libreoffice"):
        try:
            subprocess.run(
                [exe, "--headless", "--convert-to", "xlsx", "--outdir", tmpdir, os.path.abspath(path)],
                check=True, capture_output=True, timeout=180,
            )
            cand = os.path.join(tmpdir, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
            if os.path.exists(cand):
                return cand
        except Exception:
            continue
    return None


def doc_to_docx(path, tmpdir):
    """.doc -> .docx，優先用 Word COM，其次 LibreOffice。回傳新路徑或 None。"""
    if sys.platform == "win32":
        try:
            import win32com.client as wc
            app = wc.Dispatch("Word.Application")
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            try:
                d = app.Documents.Open(os.path.abspath(path), False, True)
                out = os.path.join(tmpdir, os.path.basename(path) + ".txt")
                # 16 = wdFormatDocumentDefault(.docx)，7 = wdFormatEncodedText
                text = d.Content.Text
                d.Close(False)
                return ("text", text)
            finally:
                try:
                    app.Quit()
                except Exception:
                    pass
        except Exception as e:
            print(f"  Word COM 失敗（{e}），改試 LibreOffice", file=sys.stderr)
    for exe in ("soffice", "libreoffice"):
        try:
            subprocess.run(
                [exe, "--headless", "--convert-to", "docx", "--outdir", tmpdir, os.path.abspath(path)],
                check=True, capture_output=True, timeout=180,
            )
            cand = os.path.join(tmpdir, os.path.splitext(os.path.basename(path))[0] + ".docx")
            if os.path.exists(cand):
                return ("docx", cand)
        except Exception:
            continue
    return None


def extract(path, tmpdir):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return from_docx(path)
    if ext == ".pdf":
        return from_pdf(path)
    if ext in (".xlsx", ".xls", ".ods", ".csv"):
        got = from_sheet(path)
        if got and got.strip():
            return got
        # 舊版 .xls / .ods 讀不動時，先轉成 .xlsx 再讀一次
        if ext in (".xls", ".ods"):
            conv = sheet_to_xlsx(path, tmpdir)
            if conv:
                return from_sheet(conv)
        return None
    if ext == ".txt":
        for enc in ("utf-8", "cp950", "utf-16"):
            try:
                with open(path, encoding=enc) as f:
                    return f.read()
            except Exception:
                continue
        return None
    if ext == ".doc":
        got = doc_to_docx(path, tmpdir)
        if not got:
            return None
        kind, val = got
        return val if kind == "text" else from_docx(val)
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", help="招標文件資料夾或單一檔案")
    ap.add_argument("-o", "--out", default=None, help="輸出資料夾（預設 <src>/_text）")
    args = ap.parse_args()

    src = os.path.abspath(args.src)
    if os.path.isdir(src):
        files = [os.path.join(src, f) for f in sorted(os.listdir(src))
                 if os.path.splitext(f)[1].lower() in EXTS
                 and not f.startswith("~$")]  # Word 開檔時的鎖定暫存檔
        default_out = os.path.join(src, "_text")
    else:
        files = [src]
        default_out = os.path.join(os.path.dirname(src), "_text")

    out_dir = os.path.abspath(args.out or default_out)
    os.makedirs(out_dir, exist_ok=True)

    if not files:
        print("找不到可處理的檔案", file=sys.stderr)
        return 1

    ok = fail = 0
    with tempfile.TemporaryDirectory() as tmp:
        for f in files:
            name = os.path.basename(f)
            try:
                text = extract(f, tmp)
            except Exception as e:
                text = None
                print(f"  錯誤 {name}: {e}", file=sys.stderr)
            if text is None:
                print(f"  略過 {name}（無法解析）")
                fail += 1
                continue
            dest = os.path.join(out_dir, name + ".txt")
            w(dest, text)
            print(f"  {name}  ->  {len(text)} 字")
            ok += 1

    print(f"\n完成：{ok} 個成功、{fail} 個略過")
    print(f"輸出於 {out_dir}")
    print("\n提醒：用 Read 逐份讀過，特別是估價單的表格內容。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

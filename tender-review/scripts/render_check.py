#!/usr/bin/env python3
"""把 .docx 轉成 PDF 並輸出每頁 PNG，供目視檢查排版。

  python render_check.py <檔案.docx> [-o 輸出資料夾] [--dpi 100]

Windows 走 Word COM，其他平台走 LibreOffice。之後用 Read 開那些 PNG，
確認表格沒破版、方塊沒被切在頁面中間、中文字型正確。
"""
import argparse
import os
import subprocess
import sys


def to_pdf(docx_path, out_dir):
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf = os.path.join(out_dir, base + ".pdf")
    if sys.platform == "win32":
        try:
            import win32com.client as wc
            app = wc.Dispatch("Word.Application")
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            try:
                d = app.Documents.Open(os.path.abspath(docx_path), False, True)
                d.SaveAs2(os.path.abspath(pdf), 17)  # 17 = wdFormatPDF
                pages = d.ComputeStatistics(2)       # 2 = wdStatisticPages
                d.Close(False)
                return pdf, pages
            finally:
                try:
                    app.Quit()
                except Exception:
                    pass
        except Exception as e:
            print(f"Word COM 失敗（{e}），改試 LibreOffice", file=sys.stderr)
    for exe in ("soffice", "libreoffice"):
        try:
            subprocess.run(
                [exe, "--headless", "--convert-to", "pdf", "--outdir", out_dir, os.path.abspath(docx_path)],
                check=True, capture_output=True, timeout=240,
            )
            if os.path.exists(pdf):
                return pdf, None
        except Exception:
            continue
    return None, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--dpi", type=int, default=100)
    args = ap.parse_args()

    out_dir = os.path.abspath(args.out or os.path.join(os.path.dirname(os.path.abspath(args.docx)), "_render"))
    os.makedirs(out_dir, exist_ok=True)

    pdf, pages = to_pdf(args.docx, out_dir)
    if not pdf:
        print("轉 PDF 失敗：需要 Word（Windows）或 LibreOffice。", file=sys.stderr)
        return 1

    try:
        import fitz
    except ImportError:
        print(f"已產生 {pdf}，但缺 PyMuPDF 無法出圖（pip install pymupdf）")
        return 0

    doc = fitz.open(pdf)
    base = os.path.splitext(os.path.basename(args.docx))[0]
    made = []
    for i, page in enumerate(doc):
        png = os.path.join(out_dir, f"{base}-p{i + 1:02d}.png")
        page.get_pixmap(dpi=args.dpi).save(png)
        made.append(png)
    print(f"共 {len(doc)} 頁，圖檔於 {out_dir}")
    for m in made:
        print("  " + m)
    print("\n用 Read 開這些 PNG 檢查：表格破版、方塊被切在頁面中間、字型是否正確。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""讀取公文附件文字，供核閱比對。

用法：
  python read_attachments.py 檔案1 [檔案2 ...]          # 印出各檔全文（每檔上限 MAXCHARS）
  python read_attachments.py --since 分鐘數 [--grep 關鍵字]  # 讀「下載」資料夾最近 N 分鐘內的新檔
    --grep <本校校名>  只印含關鍵字的行（查名單時用，避免把整份個資名冊印出來）

支援 pdf、odt/ods/odp、docx/xlsx、txt/csv；圖片與掃描檔會標示「無文字層」，需改用截圖判讀。
"""
import os, re, sys, time, zipfile, html

MAXCHARS = 6000
DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")


def _xml_text(zf, name):
    raw = zf.read(name).decode("utf-8", "ignore")
    raw = re.sub(r"</(text:p|text:h|table:table-row|w:p|row)>", "\n", raw)
    raw = re.sub(r"<(table:table-cell|w:tab|c)[^>]*>", "\t", raw)
    return html.unescape(re.sub(r"<[^>]+>", "", raw))


def extract(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        import pypdf
        r = pypdf.PdfReader(path)
        t = "\n".join((p.extract_text() or "") for p in r.pages)
        return t if t.strip() else f"[PDF {len(r.pages)} 頁，無文字層：請改用截圖判讀]"
    if ext in (".odt", ".ods", ".odp"):
        with zipfile.ZipFile(path) as z:
            return _xml_text(z, "content.xml")
    if ext == ".docx":
        with zipfile.ZipFile(path) as z:
            return _xml_text(z, "word/document.xml")
    if ext == ".xlsx":
        with zipfile.ZipFile(path) as z:
            names = ["xl/sharedStrings.xml"] + [n for n in z.namelist() if n.startswith("xl/worksheets/sheet")]
            return "\n".join(_xml_text(z, n) for n in names if n in z.namelist())
    if ext in (".txt", ".csv"):
        return open(path, encoding="utf-8", errors="ignore").read()
    return f"[{ext} 無法擷取文字（圖片或其他格式），請改用截圖判讀]"


def main(argv):
    grep = None
    if "--grep" in argv:
        i = argv.index("--grep"); grep = argv[i + 1]; del argv[i:i + 2]
    if argv and argv[0] == "--since":
        cutoff = time.time() - float(argv[1]) * 60
        files = sorted((os.path.join(DOWNLOADS, f) for f in os.listdir(DOWNLOADS)),
                       key=os.path.getmtime)
        files = [f for f in files if os.path.isfile(f) and os.path.getmtime(f) >= cutoff
                 and not f.endswith((".crdownload", ".tmp"))]
    else:
        files = argv
    for f in files:
        print(f"===== {os.path.basename(f)} =====")
        try:
            t = extract(f)
        except Exception as e:
            t = f"[讀取失敗：{e}]"
        if grep:
            hits = [l for l in t.splitlines() if grep in l]
            print("\n".join(hits) if hits else f"（無「{grep}」）")
        else:
            print(t[:MAXCHARS] + ("\n…（截斷）" if len(t) > MAXCHARS else ""))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1:])

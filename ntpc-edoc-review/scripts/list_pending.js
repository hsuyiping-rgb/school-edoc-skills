// 在待辦理區（或待簽收區）清單畫面執行：列出每件公文。
// 欄位：序號 | 類別(收/創) | 文號 | 主旨 | 限辦日期 | 承辦單位 | 承辦人 | 上一流程送出時間
// 待簽收區的欄位順序不同，若結果怪異，先 screenshot 對照表頭再調整索引。
const out = [];
function walk(w) {
  try {
    w.document.querySelectorAll('tr').forEach(tr => {
      const c = [...tr.cells].map(td => td.innerText.trim());
      if (c.length > 6 && /^\d+$/.test(c[0]))
        out.push([c[0], c[2], c[6], c[7].slice(0, 80), c[8], c[9], c[11], c[12]].join(' | '));
    });
    for (let i = 0; i < w.frames.length; i++) walk(w.frames[i]);
  } catch (e) {}
}
walk(window);
out.join('\n')

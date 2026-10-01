// 在清單畫面執行：把各 frame 的 window.open 改成同分頁跳轉，再點指定文號。
// 公文會在同一分頁開啟 SUPERDESK/Editor（否則會開成群組外的彈出視窗，Agent 看不到）。
// 使用前把 DOC_NO 換成文號。回傳 true 代表已點到；之後等約 6 秒再 get_page_text。
const DOC_NO = '0000000000';
let found = false;
function walk(w) {
  try {
    w.open = function (u) { top.location.href = new URL(u, w.location.href).href; return top; };
    const a = [...w.document.querySelectorAll('a')].find(x => x.textContent.trim() === DOC_NO);
    if (a && !found) { found = true; a.click(); }
    for (let i = 0; i < w.frames.length; i++) walk(w.frames[i]);
  } catch (e) {}
}
walk(window);
found

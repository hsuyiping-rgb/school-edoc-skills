// 在 OdMainFrame.aspx 載入後執行：點左側選單「待辦理區」（預設第一個＝本人待辦理區）。
// 改 AREA 可切到「待簽收區」。回傳 true 代表已點到。
const AREA = '待辦理區';
let done = false;
function walk(w) {
  try {
    if (!done && w.location.pathname.includes('OdLeftMenu')) {
      const s = [...w.document.querySelectorAll('span')].find(x => x.textContent.trim() === AREA);
      if (s) { s.click(); done = true; }
    }
    for (let i = 0; i < w.frames.length; i++) walk(w.frames[i]);
  } catch (e) {}
}
walk(window);
done

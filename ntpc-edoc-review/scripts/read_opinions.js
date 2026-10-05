// 在公文 Editor 畫面執行：依序切換簽核意見區的「承辦／會辦／核判」頁籤，讀出全部意見。
// get_page_text 只讀得到目前顯示的頁籤（預設「承辦」），會辦人的意見要切到「會辦」才看得到。
// 回傳三段文字；某頁籤沒有意見時顯示「（無）」。
const out = [];
const area = () => {
  const t = document.body.innerText;
  const i = t.indexOf('承辦會辦核判');
  const j = t.indexOf('簽核流程', i);
  return i < 0 ? '' : t.slice(i + 6, j).trim();
};
for (const name of ['承辦', '會辦', '核判']) {
  const chip = [...document.querySelectorAll('.comment-filter .tag')].find(e => e.textContent.trim() === name);
  if (!chip) { out.push('【' + name + '】（找不到頁籤）'); continue; }
  chip.click();
  await new Promise(r => setTimeout(r, 800));
  out.push('【' + name + '】' + (area() || '（無）'));
}
out.join('\n')

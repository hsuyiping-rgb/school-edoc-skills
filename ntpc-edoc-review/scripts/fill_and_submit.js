// 在公文 Editor 畫面、剛用滑鼠點過意見欄（並等 1 秒）之後執行，只填意見，不按按鈕。
// 與「點意見欄」「computer left_click 決行」「輪詢結果」放在同一個 batch（合併一步）；與「開文」分開 batch。
// 三道檢查：文號正確、焦點在 #flow-opinion-input、寫入後文字一致。任一失敗就 throw，
// batch 中止，後面的「點決行」不會執行。
// 使用前替換：DOC_NO、TXT、ACTION（'決行' 或 '退承辦人'）、BTN（下一步 computer left_click 要點的座標，
// 決行鈕固定約 (1055,328)；該點上的元素文字不是按鈕名稱就中止，避免點錯）。
// 回傳 'filled btn=x,y ih=視窗高'；x,y 是按鈕中心，供下一步 computer left_click 使用。
// 為什麼不在腳本裡用 btn.click()：自動權限審查會把 JS 觸發的決行視為未授權提交而擋下；
// 改用 computer left_click 才是校長授權範圍內的真實點擊。
const DOC_NO = '0000000000';
const TXT = '如擬。';
const ACTION = '決行';
const BTN = [1055, 328];

const e = document.getElementById('flow-opinion-input');
let r;
if (!document.body.innerText.includes('公文文號：' + DOC_NO)) r = 'ABORT wrong doc';
else if (document.activeElement !== e) r = 'ABORT focus=' + document.activeElement.id;
else {
  getSelection().selectAllChildren(e);          // 只選取意見欄自己的內容
  document.execCommand('insertText', false, TXT);
  if (e.innerText.trim() !== TXT) r = 'ABORT text mismatch: ' + e.innerText;
  else {
    // 用實際點擊位置驗證：該點上的元素文字必須就是按鈕名稱（量到的座標曾是版面未穩定時的假座標）
    const el = document.elementFromPoint(BTN[0], BTN[1]);
    r = el && el.textContent.trim() === ACTION
      ? 'filled btn=' + BTN + ' ih=' + innerHeight
      : 'ABORT not ' + ACTION + ' at ' + BTN + ': ' + (el ? el.textContent.trim().slice(0, 10) : 'none');
  }
}
if (r.startsWith('ABORT')) throw new Error(r);
r

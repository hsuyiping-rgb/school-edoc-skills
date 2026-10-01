// 在公文 Editor 畫面、剛用滑鼠點過意見欄之後執行（與點擊放在同一個 batch，但與「開文」分開 batch）。
// 三道檢查：文號正確、焦點在 #flow-opinion-input、寫入後文字一致。全數通過才按按鈕。
// 使用前替換：DOC_NO、TXT、ACTION（'決行' 或 '退承辦人'）。
// 回傳 'submitted' 表示已送出；'ABORT ...' 表示什麼都沒送，重點意見欄後再跑一次。
const DOC_NO = '0000000000';
const TXT = '如擬。';
const ACTION = '決行';

const e = document.getElementById('flow-opinion-input');
let r;
if (!document.body.innerText.includes('公文文號：' + DOC_NO)) r = 'ABORT wrong doc';
else if (document.activeElement !== e) r = 'ABORT focus=' + document.activeElement.id;
else {
  getSelection().selectAllChildren(e);          // 只選取意見欄自己的內容
  document.execCommand('insertText', false, TXT);
  if (e.innerText.trim() !== TXT) r = 'ABORT text mismatch: ' + e.innerText;
  else {
    const btn = ACTION === '決行'
      ? [...document.querySelectorAll('a.k-button')].find(b => b.textContent.trim() === '決行')
      : [...document.querySelectorAll('a')].find(b => b.textContent.trim() === '退承辦人');
    if (!btn) r = 'ABORT button not found: ' + ACTION;
    else { btn.click(); r = 'submitted'; }
  }
}
r

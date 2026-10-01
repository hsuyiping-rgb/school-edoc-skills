// 在公文 Editor 畫面執行：把左側文件抽屜裡每個附件的下載鈕各點一次。
// 檔案會存到 Windows「下載」資料夾，再用 read_attachments.py 讀取。
// 回傳附件數；0 代表這件公文沒有系統內附件（可能改放在「附件下載區」，見操作手冊）。
// 只點下載鈕、每個一次：點附件名稱本身也會觸發下載，會造成重複檔案。
// SKIP：已下載過的前幾個（重跑時用），預設 0。
const SKIP = 0;
const btns = [...document.querySelectorAll('.doc-list button.icon-download')];
btns.slice(SKIP).forEach((b, i) => setTimeout(() => b.click(), i * 1500));
const no = (document.body.innerText.match(/公文文號：\d+/) || [''])[0];
no + ' 附件數=' + btns.length + (SKIP ? '（略過前 ' + SKIP + ' 個）' : '')

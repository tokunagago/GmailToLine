const lineToken = "XXXXXXXXXXXXXXXX";

/*
 * GmailからLINEへ転送する
 */
function main() {
 newMessage = fetchContactMail()
 if(newMessage.length > 0){
   for(let i = newMessage.length-1; i >= 0; i--){
     sendLine(newMessage[i])
   }
 }
}

/*
 * メールを取得する
 * @return メールの配列
 */
function fetchContactMail() { 
 const mailAdressesStr = getMailAdressesStr();
 // 検索条件文字列を生成（未読メール かつ 検索対象のメール）
 const strTerms = 'is:unread (' + mailAdressesStr + ')';
 
 // メールを取得
 const myThreads = GmailApp.search(strTerms);
 const myMssages = GmailApp.getMessagesForThreads(myThreads);
 const valMssages = [];
 for(let i = 0; i < myMssages.length;i++){
   // 受信日時
   valMssages[i] = " " + (myMssages[i].slice(-1)[0].getDate().getMonth()+1) + "/"+ myMssages[i].slice(-1)[0].getDate().getDate()
   + " " + myMssages[i].slice(-1)[0].getDate().getHours() + ":" + myMssages[i].slice(-1)[0].getDate().getMinutes();
   const from = "\n[送信元]" + myMssages[i].slice(-1)[0].getFrom();
   const title = "\n\n[件名]" + myMssages[i].slice(-1)[0].getSubject();
   valMssages[i] = valMssages[i] + from + title;
 }
 return valMssages;
}

/*
 * 検索対象のメールアドレスを取得する
 * @return 検索対象のメールアドレス
 */
function getMailAdressesStr() {
  const mySheet = SpreadsheetApp.getActiveSheet();
  const lastRow = mySheet.getLastRow();
  const ROW_2 = 2;
  const COLUMN_1 = 1;
  // 1行目は項目名のため、2行1列目から 全行数-1行分 を取得する
  const adress = mySheet.getRange(ROW_2,COLUMN_1,lastRow-1).getValues();
  
  let mailAdressesStr = "";
  for (let i = 0; i <= adress.length-1; i++) {
    mailAdressesStr += adress[i];
    if (i != adress.length-1) {
      mailAdressesStr += " OR "
    }
  }
  return mailAdressesStr;
}

/*
 * LINEへ送る
 * @param massage 新着メール
 */
function sendLine(massage){
 const payload = {'message' : massage};
 const options ={
   "method" : "post",
   "payload" : payload,
   "headers" : {"Authorization" : "Bearer "+ lineToken}  
 };
 UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}
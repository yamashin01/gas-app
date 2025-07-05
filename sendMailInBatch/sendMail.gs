/**
 * メーリングの一括送信
 * 1. メール内容シートの、「タイトル」「本文」を読み取ってメールの件名と本文にセットする
 * 2. メール文を作成する
 * 3. 送付先シートの、名前とメールアドレスを取得する
 * 4. 送付対象か確認する
 * 5. メール送付する
 * 6. 送信先をクリアする
 */
function sendMailInBatch() {
  // メール送信の可否のダイアログ表示
  let response = ui.alert("送信実行", `メール送信を実行しますか？`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {
    ui.alert("メール送信を中止しました。");
    return;
  }

  // 送付先シートから送付先の情報を取得
  const destInfoArray = addressSheet.getRange(4, 2, addressSheet.getLastRow() - 2, 3).getValues();

  // メール内容シートから、メールタイトルと本文、添付ファイルを取得
  const title = mailSheet.getRange(2, 2).getValue();
  const content = mailSheet.getRange(3, 2).getValue();
  const fileId = extractFileIdFromUrl(mailSheet.getRange(4, 2).getValue());
  const attachFile = fileId ? DriveApp.getFileById(fileId) : "";

  const options = attachFile ? {
                                "cc" : "hello-coding-info@future-tech-association.org",
                                //  "bcc" : "bbb@auto-worker.com",
                                "name" : "Hello Coding School講師",
                                "attachments": attachFile
                                }:{
                                "cc" : "hello-coding-info@future-tech-association.org",
                                //  "bcc" : "bbb@auto-worker.com",
                                "name" : "Hello Coding School講師",
                                };

  // 添付ファイル確認のダイアログ表示
  if (fileId) {
    response = ui.alert("添付ファイルの確認", "添付ファイルが添付されていますが良いでしょうか？", ui.ButtonSet.YES_NO);
    if (response == ui.Button.NO) {
      ui.alert("メール送信を中止しました。");
      return;
    }
  }


  // メール文面の確認ダイアログ表示
  response = ui.alert("文面の確認", `内容は下記で問題ないでしょうか？\n\nタイトル：${title}\n\n本文：${content}`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {
    ui.alert("メール送信を中止しました。");
    return;
  }

  // 送付先情報を1行ずつループ
  let count = 0;
  let errCount = 0;
  for (let destInfoList of destInfoArray) {
    // 送付対象でない場合、スキップ
    if (!destInfoList[2]) continue;

    // メールアドレスが空欄の場合はスキップ
    if (!destInfoList[1]) continue;

    // 送付先の名前とアドレスをname, addressに入力
    const name = destInfoList[0];
    const address = destInfoList[1];

    // 送信するメール文面の作成
    const body = name ? `${name}様\n\n${content}` : content;
    console.log(body);

    try {
      // メール送付
      GmailApp.sendEmail(address, title, body, options);
    } catch(err) {
      console.error(`${address}へのメール送信に失敗しました。\n${err.message}`);
      errCount++;
    }

    // 送付した数をカウント
    count++;
  }
  
  // メール送信完了のダイアログ表示
  ui.alert(`${count - errCount}件のメール送信を完了しました。`);
  if (errCount > 0) ui.alert(`${errCount}件のメール送信に失敗しました。`);

  // 送付先シートの名前とメールアドレス、出欠欄をクリアし、送信対象をfalseにする
  addressSheet.getRange(4, 2, addressSheet.getLastRow(), 2).clearContent();
  addressSheet.getRange(4, 4, addressSheet.getLastRow(), 1).setValue(false);
  addressSheet.getRange(4, 6, addressSheet.getLastRow(), 1).clearContent();

}

/**
 * URLからファイルIDを抽出する
 */
const extractFileIdFromUrl = (url) => {
  // 開始文字列と終了文字列を取得（下記フォーマットを想定）
  // https://drive.google.com/file/d/<ファイルID>/view?usp=drive_link
  const startString = 'https://drive.google.com/file/d/';
  const endString = '/view?usp=drive_link';

  // URLからデータを取得
  const startIndex = url.indexOf(startString) + startString.length;
  const endIndex = url.indexOf(endString, startIndex);
  const fileId = url.substring(startIndex, endIndex);  
  
  return fileId;
}
/**
 * ツールバーに一括送信ボタンを追加する
 */
function onOpen() {
  // メニューボタン名
  const menu = ui.createMenu("メール送信");
  
  // メニュー内の実行ボタン
  menu.addItem("一括送信", "sendMailInBatch");
  menu.addItem("送付先のクリア", "clearAddressList");
  
  // スプレッドシートに反映
  menu.addToUi();
}
/**
 * スプレッドシート情報
 */
const ss = SpreadsheetApp.getActiveSpreadsheet();
const addressSheet = ss.getSheetByName("送付先");
const mailSheet    = ss.getSheetByName("メール内容");

const ui = SpreadsheetApp.getUi();

/**
 * 設定情報
 */
const CONFIG = {
  CC_ADDRESS: "",
  SENDER_NAME: "",
  SEND_DELAY: 1000, // ミリ秒
  MAX_DAILY_EMAILS: 100, // 送信可能なメールの件数、無料版は100件、有料版は1,500件
  MAX_FILE_SIZE: 25 * 1024 * 1024 // 25MB
};

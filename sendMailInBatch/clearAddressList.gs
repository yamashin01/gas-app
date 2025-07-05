/**
 * 送付先シートに記載のアドレスをクリアする
 */
const clearAddressList = () => {
  try {
    // 送付先シートの名前とメールアドレスをクリア
    const addressSheet = ss.getSheetByName("送付先");
    addressSheet.getRange(4, 2,addressSheet.getLastRow(), 2).clearContent();

    // 送信対象をすべてfalseにする
    addressSheet.getRange(4, 4, addressSheet.getLastRow(), 1).setValue(false);

    // 送付先クリア完了のダイアログ表示
    ui.alert(`送付先をクリアしました。`);

  } catch(e) {
    console.error(e.message);
    ui.alert(`送付先のクリアに失敗しました。`);
  }
  
}

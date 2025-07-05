/**
 * 送付先シートの内容をクリアする
 */
function clearAddressSheet() {
  try {
    const lastRow = addressSheet.getLastRow();
    if (lastRow >= 4) {
      // 名前とメールアドレスをクリア
      addressSheet.getRange(4, 2, lastRow - 3, 2).clearContent();
      // 送信対象をfalseにする
      addressSheet.getRange(4, 4, lastRow - 3, 1).setValue(false);
      // メモ欄をクリア
      addressSheet.getRange(4, 5, lastRow - 3, 1).clearContent();
    }
    ui.alert(`送付先をクリアしました。`);

  } catch (error) {
    console.error('シートクリアエラー:', error);
    ui.alert('警告', 'データのクリアに失敗しました。', ui.ButtonSet.OK);
  }
}


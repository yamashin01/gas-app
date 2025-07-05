/**
 * メーリングの一括送信（メイン関数）
 * 1. 入力データの検証
 * 2. 送付先の取得
 * 3. 送信確認（件数表示）
 * 4. メール内容の取得
 * 5. メール一括送信
 * 6. 結果表示
 * 7. 送信先をクリアする
 */
function sendMailInBatch() {
  try {
    // 入力データの検証
    if (!validateInputData()) return;
    
    // 送付先の取得（送信確認前に件数を知るため）
    const destinations = getValidDestinations();
    if (destinations.length === 0) {
      ui.alert('エラー', '送信対象が見つかりません。', ui.ButtonSet.OK);
      return;
    }
    
    // 送信確認（件数表示）
    const message = `${destinations.length}件のメール送信を実行しますか？`;
    const response = ui.alert("送信実行", message, ui.ButtonSet.YES_NO);
    if (response == ui.Button.NO) {
      ui.alert("メール送信を中止しました。");
      return;
    }

    // メール内容の取得
    const mailData = getMailContent();
    if (!mailData) return;
    
    // 送信件数の確認
    if (!checkSendingLimit(destinations.length)) return;
    
    // メール一括送信
    const result = sendEmails(destinations, mailData);
    
    // 結果表示
    showResult(result);
    
    // 送信先をクリア
    clearAddressSheet();
    
  } catch (error) {
    console.error('メール送信処理でエラーが発生しました:', error);
    ui.alert('エラー', `処理中にエラーが発生しました: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 入力データの検証
 */
function validateInputData() {
  try {
    // メール内容の検証
    const title = mailSheet.getRange(2, 2).getValue();
    const content = mailSheet.getRange(3, 2).getValue();
    
    if (!title || title.toString().trim() === '') {
      ui.alert('エラー', 'メールタイトルが空です。', ui.ButtonSet.OK);
      return false;
    }
    
    if (!content || content.toString().trim() === '') {
      ui.alert('エラー', 'メール本文が空です。', ui.ButtonSet.OK);
      return false;
    }
    
    // 送付先データの存在確認
    const lastRow = addressSheet.getLastRow();
    if (lastRow < 4) {
      ui.alert('エラー', '送付先データがありません。', ui.ButtonSet.OK);
      return false;
    }
    
    return true;
  } catch (error) {
    console.error('入力データ検証エラー:', error);
    ui.alert('エラー', 'データの読み取りに失敗しました。シート構成を確認してください。', ui.ButtonSet.OK);
    return false;
  }
}

/**
 * メール内容の取得
 */
function getMailContent() {
  try {
    const title = mailSheet.getRange(2, 2).getValue().toString().trim();
    const content = mailSheet.getRange(3, 2).getValue().toString().trim();
    const fileUrl = mailSheet.getRange(4, 2).getValue();
    
    let attachFile = null;
    if (fileUrl) {
      const fileId = extractFileIdFromUrl(fileUrl);
      if (fileId) {
        try {
          attachFile = DriveApp.getFileById(fileId);
          
          // ファイルサイズの確認
          const fileSize = attachFile.getSize();
          if (fileSize > CONFIG.MAX_FILE_SIZE) {
            ui.alert('エラー', `添付ファイルが25MBを超えています。(${Math.round(fileSize / 1024 / 1024)}MB)`, ui.ButtonSet.OK);
            return null;
          }
          
          // 添付ファイル確認のダイアログ表示
          const response = ui.alert("添付ファイルの確認", 
            `添付ファイル「${attachFile.getName()}」が添付されていますが良いでしょうか？`, 
            ui.ButtonSet.YES_NO);
          if (response == ui.Button.NO) {
            ui.alert("メール送信を中止しました。");
            return null;
          }
        } catch (fileError) {
          console.error('添付ファイル取得エラー:', fileError);
          ui.alert('エラー', '添付ファイルの取得に失敗しました。ファイルIDまたは権限を確認してください。', ui.ButtonSet.OK);
          return null;
        }
      }
    }
    
    // メール文面の確認ダイアログ表示
    const confirmMessage = `内容は下記で問題ないでしょうか？\n\nタイトル：${title}\n\n本文：${content}${attachFile ? `\n\n添付ファイル：${attachFile.getName()}` : ''}`;
    const response = ui.alert("文面の確認", confirmMessage, ui.ButtonSet.YES_NO);
    if (response == ui.Button.NO) {
      ui.alert("メール送信を中止しました。");
      return null;
    }
    
    return { title, content, attachFile };
  } catch (error) {
    console.error('メール内容取得エラー:', error);
    ui.alert('エラー', 'メール内容の取得に失敗しました。', ui.ButtonSet.OK);
    return null;
  }
}

/**
 * 有効な送付先の取得
 */
function getValidDestinations() {
  try {
    const lastRow = addressSheet.getLastRow();
    const destInfoArray = addressSheet.getRange(4, 2, lastRow - 3, 3).getValues();
    
    const validDestinations = [];
    const invalidEmails = [];
    
    // 各行をチェック
    for (let i = 0; i < destInfoArray.length; i++) {
      const row = destInfoArray[i];
      const name = row[0];
      const email = row[1];
      const shouldSend = row[2];
      
      // 送付対象でない場合はスキップ
      if (!shouldSend) continue;
      
      // メールアドレスが空の場合はスキップ
      if (!email || email.toString().trim() === '') continue;
      
      const emailStr = email.toString().trim();
      
      // メールアドレス形式チェック
      if (!isValidEmail(emailStr)) {
        invalidEmails.push({
          row: i + 4, // 実際のシート行番号
          name: name ? name.toString().trim() : '名前なし',
          email: emailStr
        });
        continue;
      }
      
      // 有効な送付先として追加
      validDestinations.push({
        name: name ? name.toString().trim() : '',
        email: emailStr,
        shouldSend: shouldSend
      });
    }
    
    // 無効なメールアドレスがある場合、ユーザーに通知
    if (invalidEmails.length > 0) {
      let errorMessage = `以下の無効なメールアドレスが見つかりました:\n\n`;
      invalidEmails.forEach(item => {
        errorMessage += `• 行${item.row}: ${item.name} - ${item.email}\n`;
      });
      errorMessage += `\nこれらのアドレスは送信対象から除外されます。続行しますか？`;
      
      const response = ui.alert('無効なメールアドレス', errorMessage, ui.ButtonSet.YES_NO);
      if (response == ui.Button.NO) {
        ui.alert("メール送信を中止しました。");
        return [];
      }
    }
    
    return validDestinations;
  } catch (error) {
    console.error('送付先取得エラー:', error);
    ui.alert('エラー', '送付先データの取得に失敗しました。', ui.ButtonSet.OK);
    return [];
  }
}

/**
 * メールアドレスの形式チェック
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * 送信件数制限の確認
 */
function checkSendingLimit(count) {
  if (count > CONFIG.MAX_DAILY_EMAILS) {
    const proceed = ui.alert('確認', 
      `${count}件の送信を行います。Gmailの1日制限（${CONFIG.MAX_DAILY_EMAILS}件）を超える可能性があります。続行しますか？`, 
      ui.ButtonSet.YES_NO);
    if (proceed == ui.Button.NO) {
      ui.alert("メール送信を中止しました。");
      return false;
    }
  }
  return true;
}

/**
 * メール一括送信
 */
function sendEmails(destinations, mailData) {
  const { title, content, attachFile } = mailData;
  
  // メールオプションの設定
  const options = {};
  
  // CCアドレスの設定
  if (CONFIG.CC_ADDRESSES && CONFIG.CC_ADDRESSES.length > 0) {
    const validCcAddresses = CONFIG.CC_ADDRESSES.filter(email => 
      email && email.trim() !== '' && isValidEmail(email.trim())
    );
    if (validCcAddresses.length > 0) {
      options.cc = validCcAddresses.join(',');
    }
  }
  
  // 送信者名の設定
  if (CONFIG.SENDER_NAME && CONFIG.SENDER_NAME.trim() !== '') {
    options.name = CONFIG.SENDER_NAME.trim();
  }
  
  // 添付ファイルの設定
  if (attachFile) {
    options.attachments = [attachFile];
  }
  
  let successCount = 0;
  let errorCount = 0;
  const errors = [];
  
  console.log(`メール送信開始: ${destinations.length}件`);
  
  for (let i = 0; i < destinations.length; i++) {
    const destination = destinations[i];
    const { name, email } = destination;
    
    // 進捗表示
    console.log(`進捗: ${i + 1}/${destinations.length} - ${email}`);
    
    // 送信するメール文面の作成
    const body = name ? `${name}様\n\n${content}` : content;
    
    try {
      // メール送付
      GmailApp.sendEmail(email, title, body, options);
      successCount++;
      
      // API制限回避のため待機
      if (i < destinations.length - 1) {
        Utilities.sleep(CONFIG.SEND_DELAY);
      }
      
    } catch (error) {
      console.error(`${email}へのメール送信に失敗しました: ${error.message}`);
      errorCount++;
      errors.push({ email, error: error.message });
      
      // 重要なエラーの場合は処理を中断
      if (error.message.includes('Daily email quota exceeded') || 
          error.message.includes('quota exceeded')) {
        ui.alert('エラー', '1日のメール送信上限に達しました。処理を中断します。', ui.ButtonSet.OK);
        break;
      }
      
      // 認証エラーの場合も中断
      if (error.message.includes('Authentication') || 
          error.message.includes('Permission')) {
        ui.alert('エラー', '認証エラーが発生しました。処理を中断します。', ui.ButtonSet.OK);
        break;
      }
    }
  }
  
  return { successCount, errorCount, errors, total: destinations.length };
}

/**
 * 送信結果を表示する
 */
function showResult(result) {
  const { successCount, errorCount, errors, total } = result;
  
  // 成功メッセージ
  if (successCount > 0) {
    ui.alert('送信完了', `${successCount}件のメール送信を完了しました。`, ui.ButtonSet.OK);
  }
  
  // エラーメッセージ
  if (errorCount > 0) {
    let errorMessage = `${errorCount}件のメール送信に失敗しました。\n\n`;
    errors.forEach(error => {
      errorMessage += `• ${error.email}: ${error.error}\n`;
    });
    ui.alert('送信エラー', errorMessage, ui.ButtonSet.OK);
  }
  
  // ログ出力
  console.log(`送信結果 - 成功: ${successCount}, 失敗: ${errorCount}, 総数: ${total}`);
}

/**
 * URLからファイルIDを抽出する
 */
function extractFileIdFromUrl(url) {
  if (!url || typeof url !== 'string') return null;
  
  // 複数のGoogleドライブURL形式に対応
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9-_]+)\//, // 一般的な形式
    /id=([a-zA-Z0-9-_]+)/, // 古い形式
    /\/d\/([a-zA-Z0-9-_]+)/, // 短縮形式
    /drive\.google\.com\/.*\/([a-zA-Z0-9-_]+)/ // その他の形式
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  console.warn('ファイルIDの抽出に失敗しました:', url);
  return null;
}
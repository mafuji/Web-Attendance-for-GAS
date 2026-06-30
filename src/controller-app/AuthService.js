/**
 * 認証
 */

// パスワードを検証し、正しければメイン画面のHTMLを返す
function verifyPassword(inputPassword) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    return { success: false, message: 'Configシートが見つかりません。' };
  }
  
  // 変更：2列目（B列）の2行目（インデックスは2行目=2, 2列目=2）の値（password）を取得
  const correctPassword = sheet.getRange(2, 2).getValue().toString();
  
  if (inputPassword === correctPassword) {
    // 認証成功: メインページのHTMLを生成して文字列として返す
    const sessionId = Utilities.getUuid();
    const template = getHtmlTemplate('index');
    template.sessionId = sessionId;
    
    return {
      success: true,
      html: template.evaluate().getContent()
    };
  } else {
    // 認証失敗
    return { success: false, message: 'パスワードが正しくありません。' };
  }
}
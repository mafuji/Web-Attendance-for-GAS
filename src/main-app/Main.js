/**
 * メインロジック
 */
// バージョン情報
const CURRENT_VERSION = "v1.0.0"; 

function getAppVersion() { return CURRENT_VERSION; } // バージョン取得
function getAppUrl() { return ScriptApp.getService().getUrl(); } // URL取得

// エントリーポイント
function doGet(e) {
  const page = e.parameter.p; // URLパラメータ "p" を取得
  const email = Session.getActiveUser().getEmail(); // 💡ユーザーのメアドを一度取得

  // --- [高速化] ログインユーザーの情報をUserシートから1回でまとめて取得 ---
  let isRegisteredUser = false;
  let isUserAdmin = false;

  if (email) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('User');
    const lastRow = userSheet.getLastRow();
    
    if (lastRow >= 2) {
      // A列(Email)とD列(isAdmin)の範囲だけを1回だけ取得
      const userData = userSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      const currentUserRow = userData.find(row => row[0] === email);
      
      if (currentUserRow) {
        isRegisteredUser = true;         // 行が存在すれば登録済み
        isUserAdmin = currentUserRow[3] === true; // 4列目がTRUEなら管理者
      }
    }
  }

  // 1. 管理者ページの出し分け
  if (page === 'summary') {
    // 💡 事前チェック済みの変数を使う（isAdmin()の再実行を防止）
    if (!isUserAdmin) {
      return HtmlService.createHtmlOutput('アクセス権限がありません。')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return getHtmlTemplate('summary')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 2. アクティブユーザーの登録認証
  // 💡 事前チェック済みの変数を使う（isRegistered()の再実行を防止）
  if (!isRegisteredUser){
    return getHtmlTemplate("unknown-user").evaluate();
  }

  // セッションIDを生成
  const sessionId = generateSessionId();

  // 設定を取得
  const config = getConfig();

  // 入退室入力画面テンプレート作成
  const template = getHtmlTemplate("index");
  template.sessionId = sessionId;
  template.requirePassword = config.requirePassword;
  template.useIpControl = config.useIpControl;
  template.allowedIps = config.allowedIps;

  // 入退室入力画面を表示
  return template
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1') 
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

// htmlから別のファイルを呼び出すための関数（スクリプトレットに記載）
function include(fileName){
  return getHtmlTemplate(fileName).evaluate().getContent();
}

/**
 * ============================================================================
 * Merged.gs用（全htmlファイルを連想配列化して1ファイルにまとめる）
 * ============================================================================
 */
function getHtmlTemplate(fileName){
  // 開発時はファイルから読込
  try {
      return HtmlService.createTemplateFromFile(fileName);
  } catch (e) {  
    const templates = {}

    // GitHub Actionによりファイルからhtml文字列を自動生成してtemplatesに格納

    return HtmlService.createTemplate(templates[fileName]);
  }
}

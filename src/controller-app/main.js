/**
 * メインロジック
 */

// バージョン情報
const CURRENT_VERSION = "v1.0.0"; 
function getAppVersion() { return CURRENT_VERSION; }

function doGet() {
  // 最初にログイン画面を表示
  return getHtmlTemplate('login')
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
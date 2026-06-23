function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('管理メニュー');
  
  menu
    .addItem('パスワード自動更新を無効にする', 'deleteTrigger')
    .addItem('パスワード自動更新を有効にする', 'createTrigger');

  menu.addToUi();
}
function isTriggerExists(functionName) {
  // 現在のプロジェクトに設定されているすべてのトリガーを取得
  const triggers = ScriptApp.getProjectTriggers();
  
  // ループで1つずつ関数名をチェック
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      return true; // 見つかったらその時点で true を返して終了
    }
  }
  return false; // ループが終わっても見つからなければ false
}

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

// パスワードを検証し、正しければメイン画面のHTMLを返す
function verifyPassword(inputPassword) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    return { success: false, message: 'Configシートが見つかりません。' };
  }
  
  // 3列目（C列）の2行目（インデックスは2行目=2, 3列目=3）の値を取得
  const correctPassword = sheet.getRange(2, 3).getValue().toString();
  
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

// セッションIDの更新（Merge処理）
function mergeControllerData(sessionId, ipAddress) {
  // 本体Appシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config"); 
  const appSsId = configSheet.getRange("B2").getValue();
  const appSs = SpreadsheetApp.openById(appSsId);
  let sheet = appSs.getSheetByName('Controller');

  // Controllerデータ取得
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const expiredAt = new Date(now.getTime() + (5 * 60 * 1000)); // 5分後
  
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sessionId) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > -1) {
    // UPDATE: IPと期限を更新
    sheet.getRange(foundRow, 2, 1, 2).setValues([[ipAddress, expiredAt]]);
  } else {
    // INSERT: 新規追加
    sheet.appendRow([sessionId, ipAddress, expiredAt]);
  }
  
  // ついでに最新のQR用データを返してあげると効率的です（任意）
  //return getLatestQrData(); // QRにする場合こっち
  return getLatestPassword(); 
}

// 最新のパスワードを返す（JSから呼び出す）
function getLatestPassword() {
  // 設定を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const appSsId = configSheet.getRange("B2").getValue();

  // 本体appの設定シートからパスワード取得
  const appSs = SpreadsheetApp.openById(appSsId);
  const appConfigSheet = appSs.getSheetByName("Config");
  const currentPwd = appConfigSheet.getRange("A2").getValue(); 
  
  return currentPwd;
}

//================================================================================
// トリガー用
//================================================================================

// 1分おきのトリガーを設定する（一度だけ実行する）
function createTrigger() {
  // クリア
  deleteTrigger();

  // 新規作成
  ScriptApp.newTrigger('triggerHub')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function deleteTrigger() {
  // クリア
  const triggers = ScriptApp.getProjectTriggers();

  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }    
}

function triggerHub() {
  const now = new Date();
  const min = now.getMinutes();
  const hour = now.getHours();
  const date = now.getDate();

  // パスワード書き換え（毎分）
  try {
    rotatePassword(); 
  } catch(e) {
    console.error("エラー:rotatePassword:", e);
  }

  // タスク追加例：毎日夜の23:00に実行
  // if (hour === 23 && min === 0) {
  //   try {
  //     autoBackUpLogSheet(); // ライブラリ側で関数を増やす
  //   } catch(e) {
  //     console.error(e);
  //   }
  // }
}

// 本体appのパスワードを書き換える
function rotatePassword() {
  // 設定を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config"); 
  const appSsId = configSheet.getRange("B2").getValue();

  // ランダムなパスワード生成 (4桁の数字)
  const newPwd = Math.floor(1000 + Math.random() * 9000).toString();
  
  // 本体appの設定シートに書き込み
  const appSs = SpreadsheetApp.openById(appSsId);
  const appConfigSheet = appSs.getSheetByName("Config");

  // 誤書込み防止策（列名で識別）
  const pwdColumn = appConfigSheet.getRange("A1").getValue();
  const previousPwdColumn = appConfigSheet.getRange("B1").getValue();

  if (pwdColumn !== "password" || previousPwdColumn !== "previous_password") {
  console.log(`${pwdColumn}, ${previousPwdColumn}`);
  console.log("書き込み対象のシートIDがWeb入退室シートではない可能性があるので、処理を中断します。");
  return;
 }

  const previousPwd = appConfigSheet.getRange("A2").getValue();
  appConfigSheet.getRange("B2").setValue(previousPwd); // 現在のパスワードを過去パスワード欄に移動
  appConfigSheet.getRange("A2").setValue(newPwd);
}

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
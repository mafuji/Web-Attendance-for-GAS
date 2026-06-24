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
// カスタムメニュー・トリガー用
//================================================================================
/**
 * ============================================================================
 * 【約束事1】アプリ固有のカスタムメニュー構成を返す
 * ============================================================================
 * ?? シンプルトリガー制限を回避するため、内部での ScriptApp 呼び出しを撤廃！
 */
function getAppMenuConfig() {
  return [
    // 表示を固定のシンプルな文言にします
    { type: "item", name: "?? パスワード自動更新のON / OFFを切り替える", functionName: "main_toggleTrigger" }
  ];
}

/**
 * ============================================================================
 * 【約束事2】「アプリを公開する」ボタンを押したときに自動作成するトリガー
 * ============================================================================
 */
function getAppTriggerConfig() {
  return [
    { 
      // ?? 実行したい関数名
      functionName: "rotatePassword", 
      // ?? GASの本物のメソッド名と引数をそのまま配列で指定！
      methods: [
        { name: "everyMinutes", args: [1] }
      ]
    }
  ];
}

/**
 * ============================================================================
 * アプリ固有のロジック・トリガー制御部
 * ============================================================================
 */

/**
 * トリガーのON/OFFを切り替えるカスタムメニュー関数
 * ?? ユーザーがボタンを「クリックした」後は、すべての権限が使える（シンプルトリガーではない）ため、
 * ここで ScriptApp を使うのは100%安全です。
 */
function main_toggleTrigger() {
  const ui = SpreadsheetApp.getUi();
  
  if (isAppTriggerRunning()) {
    // 稼働中なら止める
    deleteAppTriggerOnly();
    ui.alert('定期処理の停止', '?? パスワードの自動更新を停止しました。', ui.ButtonSet.OK);
  } else {
    // 停止中なら動かす
    createAppTriggerOnly();
    ui.alert('定期処理の開始', '?? パスワードの自動更新を開始しました。\n今後1分おきに自動実行されます。', ui.ButtonSet.OK);
  }
}

/**
 * 現在、rotatePassword トリガーが稼働中かどうかを判定する
 */
function isAppTriggerRunning() {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(t => t.getHandlerFunction() === 'rotatePassword');
}

/**
 * rotatePassword トリガーを直接新規作成する
 */
function createAppTriggerOnly() {
  deleteAppTriggerOnly();

  ScriptApp.newTrigger('rotatePassword')
    .timeBased()
    .everyMinutes(1)
    .create();
}

/**
 * rotatePasswordだけをピンポイントで削除する
 */
function deleteAppTriggerOnly() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rotatePassword') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
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
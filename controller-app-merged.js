// --- File: main.js ---
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
    const templates = {
  "base.css": "<style>\n  /* 画面全体を固定してスクロールバーを出さない */\n  html, body {\n    margin: 0;\n    padding: 0;\n    width: 100%;\n    height: 100%;\n    overflow: hidden;\n    background-color: #f8f9fa; /* 目に優しい極薄のグレー */\n    font-family: 'Helvetica Neue', Arial, sans-serif;\n  }\n\n  /* タブレット画面いっぱいに要素を縦並び・中央配置 */\n  .container {\n    display: flex;\n    flex-direction: column;\n    justify-content: space-between; /* 上・中・下にきれいに分配 */\n    align-items: center;\n    width: 100%;\n    height: 100%;\n    box-sizing: border-box;\n    padding: 4vh 0; /* 上下の余白 */\n  }\n\n  /* タイトル（上部） */\n  h2 {\n    font-size: 4vw; /* 画面幅に応じたサイズ */\n    margin: 0;\n    color: #444;\n    letter-spacing: 4px;\n  }\n\n  /* 💡 超巨大4桁数字（中央） */\n  .passcode-display {\n    font-size: 24vw; /* 画面横幅の約1/4サイズ＝画面いっぱいに広がる */\n    font-weight: 900; /* 圧倒的太文字 */\n    line-height: 1;\n    letter-spacing: 2vw; /* 数字同士の間隔 */\n    text-indent: 2vw;    /* letter-spacingによる右寄りを中央に戻す */\n    color: #111;\n    font-family: monospace; /* 数字の幅を統一してガタつきを防止 */\n  }\n\n  /* ステータス・時計（下部） */\n  .status {\n    font-size: 2.5vw;\n    color: #666;\n    margin: 0;\n  }\n</style>",
  "index": "<!DOCTYPE html>\n<html>\n  <head>\n    <base target=\"_top\">\n    <?!= include(\"base.css\") ?>\n  </head>\n  <body>\n    <div class=\"container\">\n      <h2>Web入退室</h2>\n      <div id=\"passcode\" class=\"passcode-display\">----</div>     \n      <div class=\"status\" id=\"status\">同期中...</div>\n    </div>\n  \n    <script>\n      const sessionId = \"<?= sessionId ?>\";\n    </script>\n    \n    <?!= include(\"index.js\") ?>\n  </body>\n</html>",
  "index.js": "<script>\n  // IPアドレスを取得して、GASのMerge処理を直接呼び出す関数\n  async function updatePasscode() {\n    document.getElementById(\"status\").innerText = \"同期中...\";\n    \n    try {\n      const response = await fetch('https://api.ipify.org?format=json');\n      const data = await response.json();\n      const ipAddress = data.ip;\n\n      google.script.run\n        .withSuccessHandler((receivedPasscode) => {\n          // 取得した4桁のパスワードを画面に表示\n          document.getElementById(\"passcode\").innerText = receivedPasscode;\n          \n          const now = new Date();\n          document.getElementById(\"status\").innerText = \"最終更新: \" + now.toLocaleTimeString();\n        })\n        .withFailureHandler((err) => {\n          console.error(\"GASエラー:\", err);\n          document.getElementById(\"status\").innerText = \"同期失敗\";\n        })\n        .mergeControllerData(sessionId, ipAddress);\n    } catch (err) {\n      console.error(\"通信エラー:\", err);\n      document.getElementById(\"status\").innerText = \"接続エラー\";\n    }\n  }\n\n  // 初回実行と5秒ごとの定期実行\n  updatePasscode();\n  setInterval(updatePasscode, 5000);\n</script>",
  "login.css": "<style>\n  body {\n    font-family: sans-serif;\n    display: flex;\n    justify-content: center;\n    align-items: center;\n    height: 100vh;\n    margin: 0;\n    background-color: #f5f5f5;\n  }\n  .login-container {\n    background: white;\n    padding: 30px;\n    border-radius: 8px;\n    box-shadow: 0 4px 6px rgba(0,0,0,0.1);\n    width: 100%;\n    max-width: 320px;\n    text-align: center;\n  }\n  h3 {\n    margin-bottom: 20px;\n    color: #333;\n  }\n  input[type=\"password\"] {\n    width: 100%;\n    padding: 10px;\n    margin-bottom: 15px;\n    border: 1px solid #ddd;\n    border-radius: 4px;\n    box-sizing: border-box;\n  }\n  button {\n    width: 100%;\n    padding: 10px;\n    background-color: #4CAF50;\n    color: white;\n    border: none;\n    border-radius: 4px;\n    cursor: pointer;\n    font-size: 16px;\n  }\n  button:disabled {\n    background-color: #cccccc;\n  }\n  .error {\n    color: red;\n    font-size: 14px;\n    margin-top: 10px;\n    min-height: 20px;\n  }\n</style>",
  "login": "<!DOCTYPE html>\n<html>\n  <head>\n    <base target=\"_top\">\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n    <?!= include(\"login.css\") ?>\n  </head>\n  <body>\n\n    <div class=\"login-container\" id=\"login-box\">\n      <h3>パスワード認証</h3>\n      <input type=\"password\" id=\"password\" placeholder=\"パスワードを入力\">\n      <button id=\"submit-btn\" onclick=\"submitPassword()\">ログイン</button>\n      <div class=\"error\" id=\"error-msg\"></div>\n    </div>\n\n    <?!= include(\"login.js\") ?>\n  </body>\n</html>",
  "login.js": "<script>\n  function submitPassword() {\n    const password = document.getElementById('password').value;\n    const errorMsg = document.getElementById('error-msg');\n    const btn = document.getElementById('submit-btn');\n    \n    if (!password) {\n      errorMsg.textContent = 'パスワードを入力してください。';\n      return;\n    }\n    \n    btn.disabled = true;\n    errorMsg.textContent = '認証中...';\n    \n    google.script.run\n      .withSuccessHandler(function(response) {\n        if (response.success) {\n          document.open();\n          document.write(response.html);\n          document.close();\n        } else {\n          errorMsg.textContent = response.message;\n          btn.disabled = false;\n        }\n      })\n      .withFailureHandler(function(err) {\n        errorMsg.textContent = '通信エラーが発生しました。';\n        btn.disabled = false;\n      })\n      .verifyPassword(password);\n  }\n\n  // Enterキーでの送信制御\n  document.getElementById('password').addEventListener('keypress', function(e) {\n    if (e.key === 'Enter') submitPassword();\n  });\n</script>"
}

    // GitHub Actionによりファイルからhtml文字列を自動生成してtemplatesに格納

    return HtmlService.createTemplate(templates[fileName]);
   }
}
// --- File: main.js ---
function doGet(e) {
  // クエリパラメータでサイト出し分け
  const page = e.parameter.p; // URLパラメータ "p" を取得

  if (page === 'summary') {
    // 管理者チェック
    if (!isAdmin()) {
      return HtmlService.createHtmlOutput('アクセス権限がありません。')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return getHtmlTemplate('summary')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // アクティブユーザーのメアド認証
  // 未登録ユーザー or メアド取得失敗⇒再ログインページ
  if (!isRegistered()){
    return getHtmlTemplate("unknown_user").evaluate();
  }

  // セッションIDを生成
  const sessionId = generateSessionId();

  // 設定
  const config = getConfig();

  // 入退室入力画面テンプレート作成（各変数をHTMLテンプレートにセッションIDを直接埋め込む）
  const template = getHtmlTemplate("index");
  template.sessionId = sessionId;
  template.requirePassword = config.requirePassword;
  template.useIpControl = config.useIpControl;
  template.allowedIps = config.allowedIps;

  // 入退室入力画面を表示
  return template
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1') // プログラム側でも追加可能
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 枠外への干渉を許可
}

// htmlから別のファイルを呼び出すための関数（スクリプトレットに記載）
function include(fileName){
  return getHtmlTemplate(fileName).evaluate().getContent();
}

// URL取得
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// 設定オブジェクトを返す
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シート
  const configSheet = ss.getSheetByName('Config');
  const useIpControl = configSheet.getRange('C2').getValue();
  const requirePassword = configSheet.getRange('D2').getValue();

  // 許可IPリスト
  const controllerSheet = ss.getSheetByName('Controller');
  const data = controllerSheet.getDataRange().getValues();

  return {
    requirePassword: requirePassword,
    useIpControl: useIpControl,
    allowedIps: data.slice(1).map(row => row[1]) // IP列(B列)のみ返す
  }; 
}

// パスワード認証
function authorized(pwdInput) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const correctPwd = configSheet.getRange("A2").getValue();

  return String(pwdInput).trim() === String(correctPwd).trim();
}

// 登録済みユーザーチェック
function isRegistered() {
  const email = Session.getActiveUser().getEmail();  

  if (!email) {
    return false;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('User');
  const data = userSheet.getDataRange().getValues();
  
  // emailが一致するかどうか
  return data.slice(1).some(row => row[0] === email);
}

// ユーザーが管理者かどうかシートを見て判定する
function isAdmin() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('User');
  const data = userSheet.getDataRange().getValues();
  
  // emailが一致し、かつ4列目(index 3)がTRUEの行を探す
  return data.slice(1).some(row => row[0] === email && row[3] === true);
}

// セッションID生成
function generateSessionId() {
  // 現在時刻とアクティブユーザーの取得
  const timestamp = new Date();
  const userEmail = Session.getActiveUser().getEmail();  

  // Usersシートでアクティブユーザーの行を特定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('User');
  
  // シートのデータを取得（A列:Email, E列:session_id, F列:session_expires_at）
  const lastRow = userSheet.getLastRow();
  if (lastRow < 2) {
    console.warn("ユーザーデータが存在しません。");
    return;
  }
  
  const data = userSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // A列(Email)のみ取得
  let targetRow = -1;

  // ログイン中のユーザーのメールアドレスと一致する行を検索
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === userEmail) {
      targetRow = i + 2; // インデックスは0から、行番号は2から始まるため
      break;
    }
  }

  if (targetRow === -1) {
    console.error("ユーザーがUserシートに見つかりませんでした: " + userEmail);
    return;
  }

  // GUID生成
  const sessionId = Utilities.getUuid();

  // 有効期限の計算 (現在時刻 + 10分)
  const expiresAt = new Date(timestamp.getTime() + 10 * 60 * 1000);

  // スプレッドシートへの書き込み
  // E列(5列目)にsession_id、F列(6列目)にsession_expires_at
  userSheet.getRange(targetRow, 5).setValue(sessionId);
  userSheet.getRange(targetRow, 6).setValue(expiresAt);

  console.log(`User: ${userEmail} のセッションを更新しました。ID: ${sessionId}`);

  // 💡 GAS側の書き込みをここで「強制同期・確定」させる（超重要）
  SpreadsheetApp.flush(); 

  return sessionId; // 💡 生成したIDを返す  
}

// セッションの有効性チェック
// @param {string} sessionId - クライアントから送られてきたセッションID
// @return {boolean} セッションが有効であればtrue、そうでなければfalse
function isSessionValid(sessionId) {
  if (!sessionId) return false;

  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = new Date(); // 現在時刻（サーバー時間）

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('User');
  
  const lastRow = userSheet.getLastRow();
  if (lastRow < 2) return false;

  // A列(Email), E列(session_id), F列(session_expires_at)のデータを取得
  // getRange(行, 列, 行数, 列数)
  const data = userSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // ユーザーの行を特定して検証
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const emailInSheet = row[0];      // A列
    const sessionIdInSheet = row[4];  // E列
    const expiresAt = row[5];         // F列

    if (emailInSheet === userEmail) {
      // 1. セッションIDが一致するか
      // 2. 有効期限が現在時刻より先か
      if (sessionIdInSheet === sessionId && expiresAt instanceof Date && expiresAt > timestamp) {
        console.log("セッション認証成功: " + userEmail);
        return true;
      } else {
        console.warn("セッション無効または期限切れ: " + userEmail);
        return false;
      }
    }
  }

  console.warn("ユーザーが見つかりません: " + userEmail);
  return false;
}

// 打刻レコード追加
function insertRecord(status, sessionId) {
  // セッションの有効性チェック
  if (!isSessionValid(sessionId)) {
    return null;
  }

  // ログシート取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AttendanceLog");

  // 現在時刻とアクティブユーザーの取得
  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail();

  // 登録
  sheet.appendRow([timestamp, user, status]);

  // 打刻した時刻を返す
  const formattedTime = Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd HH:mm:ss");
  return formattedTime;
}

// パスコードを検証し、正しければ打刻処理を行う（一括処理版）
function verifyPasswordAndInsert(status, sessionId, inputPassword) {
  // 1. セッションの有効性チェック（insertRecord内でもやりますが、ミスマッチ防止でここでも通します）
  if (!isSessionValid(sessionId)) {
    return null;
  }

  // 2. 💡 パスコードはActiveSpreadsheet内のConfigシートを参照
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  
  if (!configSheet) {
    return { success: false, message: 'システムエラー:「Config」シートが見つかりません。' };
  }
  
  // ConfigシートのA2セルからパスコードを取得
  const correctPassword = configSheet.getRange("A2").getValue().toString(); 
  const previousPassword = configSheet.getRange("B2").getValue().toString(); 
  console.log(`cerrect:${correctPassword}, previous:${previousPassword}, input:${inputPassword}`);

  // 3. パスコード判定
  if (inputPassword !== correctPassword && inputPassword !== previousPassword) {
    return { 
      success: false, 
      message: 'パスコードが一致しません。最新のパスコードを入力してください。' 
    };
  }

  // 4. 💡 パスコードが一致していたら、元の insertRecord を呼び出す
  try {
    const formattedTime = insertRecord(status, sessionId);
    
    // insertRecord側でセッション切れ等によりnullが返ってきた場合
    if (formattedTime === null) {
      return null;
    }

    // HTML（JS）側が期待しているオブジェクト形式で返す
    return {
      success: true,
      timestamp: formattedTime
    };

  } catch(e) {
    console.error("打刻エラー: ", e);
    return { 
      success: false, 
      message: '打刻処理中にエラーが発生しました: ' + e.message 
    };
  }
}

// アクティブユーザーのステータスとセッションIDを取得
function getStatusOfActiveUser() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. セッションIDの取得 (Userシート) ---
  const userSheet = ss.getSheetByName('User');
  const userData = userSheet.getDataRange().getValues();
  let sessionId = null;

  // A列:Email, E列:session_id
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === email) {
      sessionId = userData[i][4] || null; // E列
      break;
    }
  }

  // --- 2. 打刻ステータスの取得 (AttendanceLogシート) ---
  const logSheet = ss.getSheetByName("AttendanceLog");
  const lastRow = logSheet.getLastRow();
  let lastAction = null;

  if (lastRow >= 2) {
    const logValues = logSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    // 逆順（新しい順）にスキャン
    for (let i = logValues.length - 1; i >= 0; i--) {
      const row = logValues[i];
      if (row[1] === email) {
        lastAction = {
          timestamp: Utilities.formatDate(row[0], "JST", "yyyy/MM/dd HH:mm:ss"),
          status: row[2] // C列: ステータス
        };
        break;
      }
    }
  }

  // --- 3. 結果をまとめて返す ---
  return {
    email: email,
    sessionId: sessionId, // セッションIDのみ
    lastAction: lastAction
  };
}

// 期間指定付きのデータ取得 + 現在の稼働状況
// @param {string} startStr "yyyy-mm-dd"
// @param {string} endStr "yyyy-mm-dd"
function getSummaryData(startStr, endStr) {
  if (!isAdmin()) { 
    throw new Error("権限がありません");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('AttendanceLog');
  const userSheet = ss.getSheetByName('User');

  const startDate = new Date(startStr);
  startDate.setHours(0,0,0,0);
  const endDate = new Date(endStr);
  endDate.setHours(23,59,59,999);

  const userRows = userSheet.getDataRange().getValues();
  const logRows = logSheet.getDataRange().getValues();

  // --- 1. 全ユーザーのベースマップ作成 ---
  const userMap = {};
  let totalUserCount = 0; // 全アクティブユーザー数

  userRows.slice(1).forEach(row => {
    const [email, name, isInactive] = row;
    if (!isInactive && email) {
      totalUserCount++;
      userMap[email] = { 
        name: name, 
        email: email, 
        logs: [], 
        inCount: 0, 
        rate: 0,
        isCurrentlyIn: false // 現在入室フラグ
      };
    }
  });

  // --- 2. ログの解析（全期間の最新状態と、指定期間の集計） ---
  // ログを時系列順に処理するため、シートが古い順であることを前提とします
  logRows.slice(1).forEach(row => {
    const [timestamp, email, status] = row;
    if (!timestamp || !email || !status || !userMap[email]) return;

    const logDate = new Date(timestamp);

    // A. 「現在の入室状態」を判定 (全ログを走査して最新の状態に更新し続ける)
    if (status === 'IN') {
      userMap[email].isCurrentlyIn = true;
    } else if (status === 'OUT') {
      userMap[email].isCurrentlyIn = false;
    }

    // B. 「指定期間内」の集計
    if (logDate >= startDate && logDate <= endDate) {
      const dateStr = Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd");
      const timeStr = Utilities.formatDate(timestamp, "JST", "HH:mm");

      if (status === 'IN') {
        userMap[email].inCount++;
        userMap[email].logs.push({ date: dateStr, in: timeStr, out: '', duration: '', rawIn: timestamp });
      } else if (status === 'OUT') {
        const userLogs = userMap[email].logs;
        const lastLog = userLogs[userLogs.length - 1];
        if (lastLog && lastLog.in && !lastLog.out) {
          lastLog.out = timeStr;
          const diffMs = timestamp - lastLog.rawIn;
          const hours = Math.floor(diffMs / (1000 * 60 * 60));
          const mins = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
          lastLog.duration = `${hours}h ${mins}m`;
        }
      }
    }
  });

  // --- 3. 統計の算出 ---
  const diffDays = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
  const targetDays = diffDays; 

  let totalRateSum = 0;
  let currentlyInCount = 0; // 現在入室中の合計人数

  const userList = Object.values(userMap).map(user => {
    // 1. 現在入室中ならカウント
    if (user.isCurrentlyIn) currentlyInCount++;

    // 2. 指定期間にログがあるユーザーのみリスト化の対象とする（既存仕様維持）
    if (user.logs.length > 0) {
      user.rate = targetDays > 0 ? Math.round((user.inCount / targetDays) * 100) : 0;
      totalRateSum += user.rate;
      user.logs.forEach(l => delete l.rawIn);
      user.logs.reverse();
      return user;
    }
    return null;
  }).filter(u => u !== null);

  return {
    stats: {
      averageRate: userList.length > 0 ? Math.round(totalRateSum / userList.length) : 0,
      currentlyIn: currentlyInCount,    // 現在の入室人数
      totalActiveUsers: totalUserCount, // 分母となる全ユーザー数
      activeInPeriod: userList.length   // 期間内に一度でも打刻した人数
    },
    users: userList
  };
}

//================================================================================
// トリガー用
//================================================================================

// トリガーを設定する
function createTrigger() {
  // クリア
  const triggers = ScriptApp.getProjectTriggers();
  
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }  
  // 新規作成
  ScriptApp.newTrigger('triggerHub')
    .timeBased()
    .everyMinutes(1)
    .create();
}

// 複数の処理を1つのトリガーにまとめる
function triggerHub() {
  const now = new Date();
  const min = now.getMinutes();
  const hour = now.getHours();
  const date = now.getDate();

  // Controllerシートから期限切れのセッション情報を削除する（毎分）
  try {
    deleteExpiredSessions(); 
  } catch(e) {
    console.error("エラー:deleteExpiredSessions:", e);
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

// Controllerシートから期限切れのセッション情報を削除する
function deleteExpiredSessions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Controller');

  const lastRow = sheet.getLastRow();
  // 💡 データがそもそも存在しない（ヘッダー行以下がない）場合は即終了
  if (lastRow < 2) return;

  const lastColumn = sheet.getLastColumn();
  const values = sheet.getDataRange().getValues();
  
  const header = values[0]; // 1行目（ヘッダー）を退避
  const now = new Date().getTime();
  const DATE_COL_INDEX = 2; // C列 (0始まりで2)

  // 💡 生き残るデータ（期限内データ）だけを格納する配列
  const keepRows = [header]; 

  // 2行目（インデックス1）以降をチェック
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rawCell = row[DATE_COL_INDEX];

    // 日付が不正、または空欄の行は安全のため残す（スキップ）
    if (!rawCell || isNaN(Date.parse(rawCell))) {
      keepRows.push(row);
      continue;
    }

    const expiredAt = new Date(rawCell).getTime();

    // 💡 期限内のデータだけを配列にキープする（未来の時刻 ＞ 現在時刻）
    if (expiredAt >= now) {
      keepRows.push(row);
    }
  }

  // 💡 データに変化（削除対象）があった場合のみシートを更新
  if (keepRows.length < values.length) {
    // 1. 一旦シートのデータ部分をすべてクリア
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
    
    // 2. 残ったデータがあれば、2行目以降に一括で書き込み
    if (keepRows.length > 1) {
      const dataToSet = keepRows.slice(1); // ヘッダーを除いた残りのデータ
      sheet.getRange(2, 1, dataToSet.length, lastColumn).setValues(dataToSet);
    }
    
    console.log(`期限切れセッションをクリアしました。残データ: ${keepRows.length - 1}件`);
  }
}

function getHtmlTemplate(fileName){
  // 開発時はファイルから読込
  try {
      return HtmlService.createTemplateFromFile(fileName);
  } catch (e) {  
    const templates = {
  "base.css": "<style>\n  /* 全体に適用 */\n  body {\n    font-family: 'Noto Sans JP', sans-serif;\n  }\n\n  /* 共通のボタン設定 */\n  button {\n    padding: 12px 32px;\n    font-size: 16px;\n    font-weight: bold;\n    border: none;\n    border-radius: 8px; /* 角丸でモダンに */\n    cursor: pointer;\n    transition: all 0.3s ease; /* 動きをなめらかに */\n    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); /* 軽い影で浮かせる */\n    outline: none;\n  }\n\n  /* ボタンを押した瞬間の沈み込み */\n  button:active {\n    transform: translateY(0);\n    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);\n  }\n\n  /* ローディング全体の配置 */\n  .loading-container {\n    display: flex;\n    flex-direction: column;\n    align-items: center;\n    justify-content: center;\n    height: 100vh; /* 画面いっぱいに表示 */\n    color: #666;\n  }\n\n  /* くるくる回るアニメーション */\n  .spinner {\n    width: 40px;\n    height: 40px;\n    border: 4px solid #f3f3f3;\n    border-top: 4px solid #3498db; /* 青色 */\n    border-radius: 50%;\n    animation: spin 1s linear infinite;\n    margin-bottom: 10px;\n  }\n  \n  @keyframes spin {\n    0% { transform: rotate(0deg); }\n    100% { transform: rotate(360deg); }\n  }\n\n  /* --- ヘッダーレイアウト --- */\n  .app-header {\n    display: flex;\n    justify-content: space-between;\n    align-items: center;\n    padding: 0 20px;\n    background-color: #2c3e50;\n    color: white;\n    height: 60px;\n    position: sticky;\n    top: 0;\n    z-index: 1000;\n    box-shadow: 0 2px 8px rgba(0,0,0,0.15);\n  }\n\n  .header-logo {\n    font-weight: bold;\n    font-size: 1.1rem;\n    letter-spacing: 1px;\n  }\n\n  .header-right {\n    position: relative;\n  }\n\n  /* --- ハンバーガーボタン --- */\n  .menu-trigger {\n    background: none;\n    border: none;\n    width: 20px; \n    height: 16px; /* 偶数にして計算を安定させる */\n    position: relative;\n    cursor: pointer;\n    display: block;\n    padding: 0;\n    overflow: visible; /* 描画が切れるのを防ぐ */\n  }\n\n  .menu-trigger span {\n    display: block;\n    position: absolute;\n    width: 100%;\n    height: 2px;\n    background: #fff;\n    transition: all .3s ease;\n    border-radius: 2px;\n    left: 0;\n  }\n\n  /* 各線の位置（キリの良い数字で配置） */\n  .menu-trigger span:nth-child(1) { top: 0; }\n  .menu-trigger span:nth-child(2) { top: 7px; } \n  .menu-trigger span:nth-child(3) { top: 14px; } /* topからの絶対値 */\n\n  /* --- アニメーション時の挙動（中央 7px に集める） --- */\n  .menu-trigger.is-active span:nth-child(1) { \n    transform: translateY(7px) rotate(-45deg); \n  }\n  .menu-trigger.is-active span:nth-child(2) { \n    opacity: 0; \n    transform: translateX(-5px); /* 左に逃がすことで消え方を綺麗に */\n  }\n  .menu-trigger.is-active span:nth-child(3) { \n    transform: translateY(-7px) rotate(45deg); \n  }\n\n  /* --- ドロップダウンメニュー --- */\n  .dropdown-menu {\n    position: absolute;\n    top: 50px;\n    right: -10px;\n    background: white;\n    color: #333;\n    box-shadow: 0 10px 30px rgba(0,0,0,0.2);\n    border-radius: 12px;\n    width: 240px;\n    border: 1px solid #eee;\n    overflow: hidden;\n  }\n\n  .dropdown-menu ul {\n    list-style: none;\n    margin: 0;\n    padding: 0;\n  }\n\n  .dropdown-menu li a {\n    display: block;\n    padding: 18px 20px;\n    text-decoration: none;\n    color: #333;\n    font-size: 1rem;\n    transition: background 0.2s;\n  }\n\n  .dropdown-menu li a:hover {\n    background: #f0f4f8;\n    color: #007bff;\n  }\n\n  /* --- メニュー最下部のアカウントエリア --- */\n  .menu-footer {\n    padding: 15px 20px;\n    background-color: #f8f9fa;\n    border-top: 2px solid #eee; /* ここでボーダー分け */\n  }\n\n  .user-label {\n    font-size: 0.65rem;\n    color: #999;\n    font-weight: bold;\n    margin-bottom: 4px;\n  }\n\n  .user-email {\n    font-size: 0.8rem;\n    color: #666;\n    word-break: break-all;\n    line-height: 1.4;\n  }\n\n  /* --- アニメーション --- */\n  .fade-enter-active, .fade-leave-active { transition: opacity .2s, transform .2s; }\n  .fade-enter-from, .fade-leave-to { opacity: 0; transform: translateY(-10px); }\n\n  /* --- コンテンツ調整 --- */\n  .content-container { padding: 20px; }\n  .status-card { margin-top: 30px; text-align: center; color: #555; }\n</style>",
  "head_common": "<base target=\"_top\">\n<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n<link href=\"https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap\" rel=\"stylesheet\">\n<meta charset=\"UTF-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">",
  "index.css": "<style>\n  // --- メイン --- \n  // 全体のレイアウト調整 \n  #app {\n    display: flex;\n    gap: 20px; // ボタン同士の隙間 \n    justify-content: center;\n    margin-top: 50px;\n  }\n\n  // 入室ボタン (青系) \n  .button-in {\n    font-size: 32px;\n    background-color: #4A90E2;\n    color: white;\n  }\n\n  .button-in:hover {\n    background-color: #357ABD;\n    transform: translateY(-2px); // 少し浮き上がる \n    box-shadow: 0 6px 12px rgba(74, 144, 226, 0.3);\n  }\n\n  // 退室ボタン (緑系) \n  .button-out {\n    font-size: 32px;\n    background-color: #2ecc71; // 鮮やかなエメラルドグリーン \n    color: #ffffff;            // 文字は白で見やすく \n    border: none;              // 枠線を消すとモダンになります \n    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); // 軽い影で立体感を \n    transition: all 0.3s ease; // 変化を滑らかに \n  }\n\n  .button-out:hover {\n    background-color: #27ae60; // ホバー時は少し濃い緑に \n    color: #ffffff;\n    transform: translateY(-2px); // 少し浮き上がる演出 \n    box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15); // 影も強く \n  }\n\n  // 直近の打刻 \n  .last-action {\n    margin-top: 20px;\n    font-size: 0.9rem;\n    color: #666;\n    text-align: center;\n    border-top: 1px solid #eee;\n    padding-top: 10px;\n  } \n\n  // セッション切れ表示\n  .expired-container {\n    display: flex;\n    justify-content: center;\n    align-items: center;\n    min-height: 80vh;\n    padding: 20px;\n  }\n\n  .expired-card {\n    background: white;\n    padding: 40px 30px;\n    border-radius: 12px;\n    box-shadow: 0 4px 20px rgba(0,0,0,0.08);\n    text-align: center;\n    max-width: 400px;\n    width: 100%;\n  }\n\n  .expired-icon {\n    font-size: 50px;\n    margin-bottom: 15px;\n  }\n\n  .expired-card h2 {\n    color: #333;\n    margin-bottom: 15px;\n    font-size: 1.5rem;\n  }\n\n  .expired-card p {\n    color: #666;\n    line-height: 1.6;\n    margin-bottom: 25px;\n    font-size: 0.95rem;\n  }  \n</style>",
  "index": "<!DOCTYPE html>\n<html lang=\"ja\">\n  <head>\n    <?!= include(\"head_common\") ?>\n    <?!= include(\"base.css\") ?>\n    <?!= include(\"index.css\") ?>\n    <title>Web入退室</title>\n    <script src=\"https://cdn.jsdelivr.net/npm/sweetalert2@11\" defer></script>\n  </head>\n  <body>\n    <script>\n      // GAS側からレンダリング時に直接注入されるので、絶対に空にならない\n      const SERVER_SESSION_ID =\"<?= sessionId ?>\";\n      const requirePassword = <?!= JSON.stringify(requirePassword) ?>;\n    </script>\n\n    <!-- IPチェック -->\n    <script>\n      // 1. GAS側からIPチェックを行うかのフラグを受け取る\n      const useIpControl = <?!= JSON.stringify(useIpControl) ?>;\n      \n      // サーバーから渡された許可IPリストをJS変数として展開\n      const allowedIps = <?!= JSON.stringify(allowedIps) ?>;\n\n      // 非同期でIP判定し、許可外なら画面を書き換える\n      (async function checkIp() {\n        // 2. フラグが false (チェックしない) なら、ここで処理を終了してスキップ\n        if (!useIpControl) {\n          console.log(\"IPチェックはスキップされました。\");\n          return; \n        }\n\n        try {\n          const res = await fetch('https://api.ipify.org?format=json');\n          const data = await res.json();\n          \n          if (!allowedIps.includes(data.ip)) {\n            // 許可されていないIPなら画面を白紙にしてメッセージを表示\n            document.body.innerHTML = \"<h1>アクセス拒否</h1><p>許可されていないIPアドレスからのアクセスです。</p>\";\n          }\n        } catch(e) {\n          document.body.innerHTML = \"<h1>通信エラー</h1><p>ネットワークを確認してください。</p>\";\n        }\n      })();      \n    </script>\n\n    <div id=\"main\">\n      <!-- 読み込み中表示 -->\n      <div v-if=\"!isReady\" class=\"loading-container\">\n        <div class=\"spinner\"></div>\n        <p>読み込み中...</p>\n      </div>\n      \n      <!-- セッション切れ -->\n      <div v-else-if=\"isSessionExpired\" class=\"expired-container\">\n        <div class=\"expired-card\">\n          <div class=\"expired-icon\">⚠️</div>\n          <h2>セッションの有効期限切れ</h2>\n          <p>もう一度打刻用端末からQRコードを読み取ってアクセスしてください。</p>\n        </div>\n      </div>\n\n      <!-- ヘッダー -->\n      <header class=\"app-header\" v-if=\"isReady && !isSessionExpired\">\n        <div class=\"header-logo\">Web入退室</div>\n\n        <div class=\"header-right\">\n          <button class=\"menu-trigger\" @click=\"isMenuOpen = !isMenuOpen\" :class=\"{ 'is-active': isMenuOpen }\">\n            <span></span>\n            <span></span>\n            <span></span>\n          </button>\n\n          <transition name=\"fade\">\n            <nav class=\"dropdown-menu\" v-show=\"isMenuOpen\">\n              <ul>\n                <li v-if=\"isAdmin\"><a href=\"#\" @click.prevent=\"openDashboard\">ダッシュボード</a></li>\n                <li><a href=\"#\" @click.prevent=\"logout\">ログアウト</a></li>\n                <li class=\"menu-footer\">\n                  <span class=\"user-label\">ログイン中：</span>\n                  <span class=\"user-email\">{{ email }}</span>\n                </li>\n              </ul>\n            </nav>\n          </transition>\n        </div>\n      </header>\n\n      <!-- アプリ画面 -->\n      <div id=\"app\" v-if=\"isReady && !isSessionExpired\">\n        <button @click=\"sendIn\" v-show=\"!isIn\" class=\"button-in\">入室</button>\n        <button @click=\"sendOut\" v-show=\"isIn\" class=\"button-out\">退室</button>\n      </div>\n      <div class=\"last-action\" v-if=\"lastAction.time && !isSessionExpired\">\n          前回の打刻：{{ lastAction.time }} （{{ lastAction.status }}）<br/>\n      </div>\n    </div>\n    <script src=\"https://cdn.jsdelivr.net/npm/vue@3.1.5\"></script>\n    <?!= include(\"index.js\") ?>\n  </body>\n</html>",
  "index.js": "<script>\n  Vue.createApp({\n    data() {\n      return {\n        isReady: false,\n        isIn: false,\n        isAdmin: false,\n        isMenuOpen: false, \n        email: \"\",\n        sessionId: SERVER_SESSION_ID,\n        requirePassword: requirePassword,\n        isSessionExpired: false,\n        lastAction: { time: \"\", status: \"\" }\n      }\n    },\n    async mounted() {\n      this.fetchUserData();\n    },\n    methods: {\n      // ユーザー情報の取得（直呼び出し）\n      fetchUserData() {\n        google.script.run\n          .withFailureHandler(() => { this.isReady = true; })\n          .withSuccessHandler(res => {\n            this.email = res.email;\n            if (res.lastAction) {\n              this.lastAction = {\n                time: res.lastAction.timestamp,\n                status: res.lastAction.status === \"IN\" ? \"入室\" : \"退室\"\n              };\n              this.isIn = (res.lastAction.status === \"IN\");\n            }\n            this.checkAdmin();\n          })\n          .getStatusOfActiveUser();\n      },\n\n      // 管理者判定（直呼び出し）\n      checkAdmin() {\n        google.script.run\n          .withSuccessHandler(res => {\n            this.isAdmin = res;\n            this.isReady = true; // すべてのデータが揃ってから表示\n          })\n          .isAdmin();\n      },\n\n      // パスワード認証を挟んだ打刻処理\n      async send(status) {\n        let password = null;\n\n        // 1. 入室（IN）のときだけパスワード入力を求める\n        if (status === \"IN\" && requirePassword === true) {\n          const { value: typedPassword } = await Swal.fire({\n            title: 'パスワード入力',\n            text: '',\n            input: 'password',\n            inputAttributes: {\n              maxlength: '4',\n              inputmode: 'numeric',\n              pattern: '[0-9]*',\n              style: 'text-align: center; font-size: 24px; letter-spacing: 4px;'\n            },\n            showCancelButton: true,\n            cancelButtonText: 'キャンセル',\n            confirmButtonText: '次へ',\n            inputValidator: (value) => {\n              if (!value || value.length !== 4 || isNaN(value)) {\n                return '4桁の数字を入力してください';\n              }\n            }\n          });\n\n          // キャンセルされた場合は処理終了\n          if (!typedPassword) return;\n          password = typedPassword;\n        }\n\n        // 2. 処理中ローディング表示\n        Swal.fire({\n          title: '処理中...',\n          allowOutsideClick: false,\n          allowEscapeKey: false,\n          showConfirmButton: false,\n          didOpen: () => {\n            Swal.showLoading();\n          }\n        });\n\n        // 3. パスワード検証 ＆ 打刻リクエスト（直呼び出しのチェーン構成）\n        const runner = google.script.run\n          .withFailureHandler(e => {\n            Swal.fire(\"エラー\", \"通信に失敗しました: \" + e, \"error\");\n          })\n          .withSuccessHandler(res => {\n            // セッション切れのハンドリング\n            if (res === null) {\n              this.isSessionExpired = true;\n              Swal.close();\n              return;\n            }\n            \n            let displayTime = \"\";\n\n            if (status === \"IN\" && requirePassword === true) {\n              // verifyPasswordAndInsert の結果オブジェクトを処理\n              if (res.success === false) {\n                Swal.fire(\"認証失敗\", res.message || \"パスワードが正しくありません。\", \"error\");\n                return;\n              }\n              displayTime = res.timestamp;\n            } else {\n              // insertRecord の結果（直に日時文字列が返る）を処理\n              displayTime = res;\n            }\n\n            // 打刻成功時の画面反映\n            this.isIn = (status === \"IN\");\n            this.lastAction = {\n              time: displayTime,\n              status: status === \"IN\" ? \"入室\" : \"退室\"\n            };\n\n            // 完了メッセージに差し替え\n            Swal.fire({\n              title: '完了',\n              text: status === \"IN\" ? \"入室を記録しました\" : \"退室を記録しました\",\n              icon: 'success',\n              timer: 2000,\n              showConfirmButton: false\n            });\n          });\n\n        // 💡 状況に応じてGAS側の呼び出す関数と引数を出し分ける\n        if (status === \"IN\" && requirePassword === true) {\n          // verifyPasswordAndInsert(status, sessionId, inputPassword) のシグネチャに合わせる\n          runner.verifyPasswordAndInsert(status, this.sessionId, password);\n        } else {\n          // insertRecord(status, sessionId) のシグネチャに合わせる\n          runner.insertRecord(status, this.sessionId);\n        }\n      },\n\n      sendIn() { this.send(\"IN\"); },\n      sendOut() { this.send(\"OUT\"); },\n\n      // メニューアクション（ダッシュボード移動）\n      openDashboard() {\n        this.isMenuOpen = false;\n        google.script.run\n          .withSuccessHandler(url => {\n            window.top.location.href = url + \"?p=summary\";\n          })\n          .getAppUrl();\n      },\n\n      // ログアウト処理\n      logout() {\n        this.isMenuOpen = false;\n        Swal.fire({\n          title: 'ログアウトしますか？',\n          text: \"Googleアカウントのログアウトページへ移動します\",\n          icon: 'question',\n          showCancelButton: true,\n          confirmButtonText: 'ログアウト',\n          cancelButtonText: 'キャンセル'\n        }).then((result) => {\n          if (result.isConfirmed) {\n            google.script.run\n              .withSuccessHandler(function(url) {\n                var logoutUrl = 'https://accounts.google.com/Logout?continue=' + \n                                encodeURIComponent('https://accounts.google.com/ServiceLogin?continue=' + url);\n                window.top.location.href = logoutUrl;\n              })\n              .getAppUrl();\n          }\n        });\n      }\n    }\n  }).mount(\"#main\");\n</script>",
  "summary.css": "<style>\n/* 全体の土台 */\n  .dashboard-body {\n    background-color: #f8fafc; /* slate-50 */\n    min-height: 100vh;\n    font-family: sans-serif;\n    color: #0f172a; /* slate-900 */\n    margin: 0;\n  }\n  \n  .container-main {\n    max-width: 64rem; /* max-w-5xl */\n    margin: 0 auto;\n    padding: 2.5rem 1rem;\n  }\n\n  /* 共通の統計カード */\n  .stats-card {\n    background: #ffffff;\n    padding: 1.5rem;\n    border-radius: 1rem;\n    border: 1px solid #e2e8f0;\n    box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);\n  }\n  .header-section {\n    margin-bottom: 2.5rem;\n    display: flex;\n    flex-direction: column;\n    justify-content: space-between;\n    gap: 1rem;\n  }\n\n  /* PCサイズでは横並び */\n  @media (min-width: 768px) {\n    .header-section { flex-direction: row; align-items: flex-end; }\n  }\n\n  .header-title {\n    font-size: 1.875rem;\n    font-weight: 800;\n    color: #1e293b;\n    margin: 0;\n  }\n\n  .date-filter {\n    display: flex;\n    align-items: center;\n    gap: 0.5rem;\n    background: #ffffff;\n    padding: 0.5rem;\n    border-radius: 0.75rem;\n    border: 1px solid #e2e8f0;\n  }\n\n  .date-input {\n    background: transparent;\n    border: none;\n    font-size: 0.875rem;\n    outline: none;\n    text-align: center;\n  }\n\n  .stats-grid {\n    display: grid;\n    grid-template-columns: 1fr;\n    gap: 1rem;\n    margin-bottom: 2rem;\n  }\n\n  @media (min-width: 768px) {\n    .stats-grid { grid-template-columns: repeat(2, 1fr); }\n  }\n\n  .stats-label {\n    font-size: 0.875rem;\n    font-weight: 500;\n    color: #64748b;\n    text-transform: uppercase;\n    margin: 0;\n  }\n\n  .stats-value {\n    font-size: 1.875rem;\n    font-weight: 700;\n    margin-top: 0.25rem;\n    margin-bottom: 0;\n    display: flex;\n    align-items: baseline;\n    gap: 0.25rem;    \n  }\n\n  .text-blue-600 { color: #2563eb; }\n\n  /* アコーディオン制御 */\n  details > summary { list-style: none; outline: none; }\n  details > summary::-webkit-details-marker { display: none; }\n  \n  details[open] .group-open-rotate { \n    transform: rotate(180deg); \n  }\n\n  .user-list-container {\n    display: flex;\n    flex-direction: column;\n    gap: 0.75rem;\n  }  \n\n  /* ステータスバッジ */\n  .status-badge {\n    padding: 0.2rem 0.6rem;\n    border-radius: 9999px;\n    font-size: 0.75rem;\n    font-weight: 700;\n    line-height: 1;\n    display: inline-flex;\n    align-items: center;\n  }\n\n  /* 在室中（緑） */\n  .badge-in {\n    background-color: #ecfdf5; /* emerald-50 */\n    color: #059669; /* emerald-600 */\n    border: 1px solid #10b981; /* emerald-500 */\n  }\n\n  /* 不在（グレー） */\n  .badge-out {\n    background-color: #f1f5f9; /* slate-100 */\n    color: #64748b; /* slate-500 */\n    border: 1px solid #cbd5e1; /* slate-300 */\n  }  \n</style>",
  "summary": "<!DOCTYPE html>\n<html lang=\"ja\">\n  <head>\n    <?!= include(\"head_common\") ?>    \n    <?!= include(\"base.css\") ?>\n    <?!= include(\"summary.css\") ?>\n    <title>Web入退室ダッシュボード</title>\n    <script src=\"https://cdn.tailwindcss.com\" defer></script>\n  </head>\n  <body class=\"dashboard-body\">\n    <div class=\"container-main\">\n      <header class=\"header-section\">\n        <h1 class=\"header-title\">Web入退室ダッシュボード</h1>\n        \n        <div class=\"date-filter\">\n          <input type=\"date\" class=\"date-input\" value=\"2024-03-01\">\n          <span class=\"date-separator\">〜</span>\n          <input type=\"date\" class=\"date-input\" value=\"2024-03-31\">\n        </div>\n      </header>\n\n      <section class=\"stats-grid\">\n        <div class=\"stats-card\">\n          <p class=\"stats-label\">全体の平均出席率</p>\n          <p id=\"stat-avg-rate\" class=\"stats-value text-blue-600\">--%</p>\n        </div>\n        \n        <div class=\"stats-card\">\n          <p class=\"stats-label\">現在の在室状況</p>\n          <p id=\"stat-current-in\" class=\"stats-value\">\n            -- <span style=\"font-size: 1.125rem; color: #64748b; font-weight: 500;\">/ -- 名</span>\n          </p>\n        </div>\n      </section>\n\n      <div id=\"user-list\" class=\"user-list-container\"></div>\n    </div>\n\n    <?!= include(\"summary.js\") ?>\n  </body>\n</html>",
  "summary.js": "<script>\n  document.addEventListener('DOMContentLoaded', () => {\n    setDefaultDates(); // 1. デフォルト日付をセット\n    fetchData();       // 2. データ取得\n  });\n\n  // 当月初(1日)〜今日をセットする\n  function setDefaultDates() {\n    const now = new Date();\n    const currentYear = now.getFullYear();\n    const currentMonth = now.getMonth() + 1; // 1-12月\n\n    const formattedMonth = String(currentMonth).padStart(2, '0'); // 月を2桁（01, 02...）に整形\n    const startDate = `${currentYear}-${formattedMonth}-01`; // 当月の月初（01日）を作成\n    const endDate = now.toISOString().split('T')[0]; // 今日を yyyy-mm-dd形式で取得\n\n    document.querySelectorAll('.date-input')[0].value = startDate;\n    document.querySelectorAll('.date-input')[1].value = endDate;\n  }\n\n  // サーバーからデータを取得（期間を指定・直呼び出しに修正）\n  function fetchData() {\n    const listContainer = document.getElementById('user-list');\n    const start = document.querySelectorAll('.date-input')[0].value;\n    const end = document.querySelectorAll('.date-input')[1].value;\n\n    listContainer.innerHTML = '<p class=\"text-center py-10 text-slate-400\">期間内のデータを集計中...</p>';\n\n    // GAS側の引数の定義に沿って、start と end を直接渡す\n    google.script.run\n      .withSuccessHandler(renderSummary)\n      .withFailureHandler(err => {\n        listContainer.innerHTML = `<p class=\"text-red-500\">エラー: ${err.message}</p>`;\n      })\n      .getSummaryData(start, end);\n  }\n\n  // 日付が変わったら自動で再読み込みする場合\n  document.querySelectorAll('.date-input').forEach(el => {\n    el.addEventListener('change', fetchData);\n  });\n\n  // 取得したデータを画面に描画\n  function renderSummary(data) {\n    const listContainer = document.getElementById('user-list');\n\n    // 在室者⇒不在者の並びにソート\n    data.users.sort((a, b) => b.isCurrentlyIn - a.isCurrentlyIn);\n\n    // 統計カードの更新\n    document.getElementById('stat-avg-rate').textContent = `${data.stats.averageRate}%`;\n    \n    const currentInEl = document.getElementById('stat-current-in');\n    if (currentInEl) {\n      currentInEl.innerHTML = `\n        ${data.stats.currentlyIn} <span style=\"font-size: 1.125rem; color: #64748b; font-weight: 500;\">/ ${data.stats.totalActiveUsers} 名</span>\n      `;\n    }\n\n    // データがない場合のガード\n    if (!data.users || data.users.length === 0) {\n      listContainer.innerHTML = '<p class=\"text-center py-10 text-slate-400\">期間内のデータがありません。</p>';\n      return;\n    }\n\n    // リストの描画\n    listContainer.innerHTML = data.users.map(user => {\n      // 在室・不在のバッジを定義\n      const statusBadge = user.isCurrentlyIn \n        ? `<span class=\"status-badge badge-in\">在室</span>` \n        : `<span class=\"status-badge badge-out\">不在</span>`;\n\n      // HTMLを返す（テンプレートリテラル）\n      return `\n        <details class=\"group bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm transition-all duration-300 open:ring-2 open:ring-blue-500/20\">\n          <summary class=\"flex items-center justify-between p-5 cursor-pointer hover:bg-slate-50 transition-colors gap-4\">\n            <div class=\"min-w-0 flex items-center gap-3\"> \n              <div class=\"min-w-0\">\n                <div class=\"flex items-center gap-2\">\n                  <h3 class=\"font-bold text-slate-800 text-lg truncate\">${user.name}</h3>\n                  ${statusBadge}\n                </div>\n                <p class=\"text-sm text-slate-400 truncate\">${user.email}</p>\n              </div>\n            </div>\n            \n            <div class=\"flex items-center gap-8 flex-shrink-0\">\n              <div class=\"text-right\">\n                <p class=\"text-[10px] font-bold text-slate-400 uppercase tracking-widest\">出席率</p>\n                <p class=\"text-xl font-black ${user.rate < 80 ? 'text-orange-500' : 'text-emerald-500'}\">\n                  ${user.rate}<span class=\"text-sm ml-0.5\">%</span>\n                </p>\n              </div>\n              <div class=\"group-open-rotate transition-transform duration-300 text-slate-300\">\n                <svg xmlns=\"http://www.w3.org/2000/svg\" class=\"h-6 w-6\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"currentColor\">\n                  <path stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" />\n                </svg>\n              </div>\n            </div>\n          </summary>\n          <div class=\"px-5 pb-5 overflow-x-auto\">\n            <div class=\"bg-slate-50 rounded-xl p-4 border border-slate-100\">\n              <table class=\"w-full text-left text-sm\">\n                <thead>\n                  <tr class=\"text-slate-400 font-medium border-b border-slate-200\">\n                    <th class=\"py-3 px-2\">日付</th>\n                    <th class=\"py-3\">入室</th>\n                    <th class=\"py-3\">退室</th>\n                    <th class=\"py-3 text-right\">在室</th>\n                  </tr>\n                </thead>\n                <tbody class=\"text-slate-600 divide-y divide-slate-200\">\n                  ${user.logs.map(log => `\n                    <tr class=\"hover:bg-slate-100 transition-colors\">\n                      <td class=\"py-3 px-2 font-medium text-slate-700\">${log.date}</td>\n                      <td class=\"py-3\">${log.in}</td>\n                      <td class=\"py-3\">${log.out || '--:--'}</td>\n                      <td class=\"py-3 text-right font-mono font-medium\">${log.duration || '---'}</td>\n                    </tr>\n                  `).join('')}\n                </tbody>\n              </table>\n            </div>\n          </div>\n        </details>\n      `;\n    }).join('');\n  }\n</script>",
  "unknown_user": "<!DOCTYPE html>\n<html>\n  <div style=\"font-family: sans-serif; padding: 20px;\">\n    <p style=\"color: red; font-weight: bold;\">認証失敗: 未登録のユーザーか、Google Workspaceアカウントでログインしていません。</p>\n    <p>以下のボタンを押してユーザー登録済みのGoogle Workspaceアカウントを選択し直してください。</p>\n    <p>※ユーザー登録が済んでいない場合は管理者にお問い合わせ下さい。</p>\n    <a href=\"https://accounts.google.com/Logout?continue=https://accounts.google.com/ServiceLogin?continue=${ScriptApp.getService().getUrl()}\"\n      target=\"_top\" \n      style=\"display: inline-block; padding: 10px 20px; background: #4285f4; color: white; text-decoration: none; border-radius: 4px;\">\n      アカウントを切り替える\n    </a>\n  </div>\n</html>\n"
}

    // GitHub Actionによりファイルからhtml文字列を自動生成してtemplatesに格納

    return HtmlService.createTemplate(templates[fileName]);
   }
}
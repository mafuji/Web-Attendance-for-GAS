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

  // 　 GAS側の書き込みをここで「強制同期・確定」させる（超重要）
  SpreadsheetApp.flush(); 

  return sessionId; // 　 生成したIDを返す  
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

  // 2. 　 パスコードはActiveSpreadsheet内のConfigシートを参照
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

  // 4. 　 パスコードが一致していたら、元の insertRecord を呼び出す
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
// カスタムメニュー・トリガー用
//================================================================================

/**
 * ============================================================================
 * 【約束事1】アプリ固有のカスタムメニュー構成を返す
 * ============================================================================
 */
function getAppMenuConfig() {
  // ?? 現時点ではメニューなし（将来追加したくなったらこの配列の中にオブジェクトを増やす）
  return [
    // 例: { type: "item", name: "? 新機能の実行", functionName: "app_newFeature" }
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
      functionName: "deleteExpiredSessions", 
      // ?? GASの本物のメソッド名と引数をそのまま配列で指定！
      methods: [
        { name: "everyMinutes", args: [1] }
      ]
    }
  ];
}

// Controllerシートから期限切れのセッション情報を削除する
function deleteExpiredSessions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Controller');

  const lastRow = sheet.getLastRow();
  // 　 データがそもそも存在しない（ヘッダー行以下がない）場合は即終了
  if (lastRow < 2) return;

  const lastColumn = sheet.getLastColumn();
  const values = sheet.getDataRange().getValues();
  
  const header = values[0]; // 1行目（ヘッダー）を退避
  const now = new Date().getTime();
  const DATE_COL_INDEX = 2; // C列 (0始まりで2)

  // 　 生き残るデータ（期限内データ）だけを格納する配列
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

    // 　 期限内のデータだけを配列にキープする（未来の時刻 ＞ 現在時刻）
    if (expiredAt >= now) {
      keepRows.push(row);
    }
  }

  // 　 データに変化（削除対象）があった場合のみシートを更新
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
    const templates = {}

    // GitHub Actionによりファイルからhtml文字列を自動生成してtemplatesに格納

    return HtmlService.createTemplate(templates[fileName]);
   }
}
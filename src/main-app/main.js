// バージョン情報
const CURRENT_VERSION = "v1.0.0"; 
function getAppVersion() { return CURRENT_VERSION; }

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
    return getHtmlTemplate("unknown_user").evaluate();
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

// URL取得
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// 設定オブジェクトを返す（CIDR範囲指定対応版）
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シート
  const configSheet = ss.getSheetByName('Config');
  const useIpControl = configSheet.getRange('C2').getValue();
  const requirePassword = configSheet.getRange('D2').getValue();

  // AllowedIpシートから許可IPリストを取得
  const allowedIpSheet = ss.getSheetByName('AllowedIp');
  if (!allowedIpSheet) {
    return { requirePassword: requirePassword, useIpControl: useIpControl, allowedIps: [] };
  }
  
  const data = allowedIpSheet.getDataRange().getValues();
  // ヘッダーを除いたA列の文字列リスト（空行は除外）
  const ipRules = data.slice(1).map(row => String(row[0]).trim()).filter(Boolean);

  return {
    requirePassword: requirePassword,
    useIpControl: useIpControl,
    allowedIps: ipRules // ここに "192.168.1.0/24" などがそのまま入る
  }; 
}

// パスワード認証
function authorized(pwdInput) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const correctPwd = configSheet.getRange("A2").getValue();

  return String(pwdInput).trim() === String(correctPwd).trim();
}

// 登録済みユーザーチェック（高速版）
function isRegistered() {
  const email = Session.getActiveUser().getEmail();  

  if (!email) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('User');
  const lastRow = userSheet.getLastRow();
  
  if (lastRow < 2) return false;

  // 💡 A列（Email）のデータだけをピンポイントで一括取得（2行目から、1列分だけ）
  const emailValues = userSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  // 1次元配列にフラット化して検索
  return emailValues.map(row => row[0]).includes(email);
}

// ユーザーが管理者かどうかシートを見て判定する（高速版）
function isAdmin() {
  const email = Session.getActiveUser().getEmail();
  
  if (!email) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('User');
  const lastRow = userSheet.getLastRow();
  
  if (lastRow < 2) return false;
  
  // 💡 A列（Email）からD列（管理者フラグ）までの4列だけを取得範囲にする
  const userData = userSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  // emailが一致し、かつ4列目(index 3)がTRUEの行を探す
  return userData.some(row => row[0] === email && row[3] === true);
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

  //   GAS側の書き込みをここで「強制同期・確定」させる（超重要）
  SpreadsheetApp.flush(); 

  return sessionId; //   生成したIDを返す  
}

// セッションの有効性チェック
function isSessionValid(sessionId) {
  if (!sessionId) return false;

  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = new Date(); // 現在時刻（サーバー時間）

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('User');
  
  const lastRow = userSheet.getLastRow();
  if (lastRow < 2) return false;

  // A列(Email), E列(session_id), F列(session_expires_at)のデータを取得
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
  // 💡 1. 通常呼び出し時はここでセッションチェック
  if (!isSessionValid(sessionId)) {
    return null;
  }

  // 💡 2. 実際の書き込み処理は、下の共通関数（_executeInsert）に丸投げする
  return _executeInsert(status);
}

// パスコードを検証し、正しければ打刻処理を行う（一括処理版）
function verifyPasswordAndInsert(status, sessionId, inputPassword) {
  // 💡 1. ここでセッションチェック（1回目）
  if (!isSessionValid(sessionId)) {
    return null;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  
  if (!configSheet) {
    return { success: false, message: 'システムエラー:「Config」シートが見つかりません。' };
  }
  
  const correctPassword = configSheet.getRange("A2").getValue().toString(); 
  const previousPassword = configSheet.getRange("B2").getValue().toString(); 

  if (inputPassword !== correctPassword && inputPassword !== previousPassword) {
    return { 
      success: false, 
      message: 'パスコードが一致しません。最新のパスコードを入力してください。' 
    };
  }

  // 💡 2. パスワードが合っていれば、二度目のセッションチェックをパスして直接書き込む
  try {
    const formattedTime = _executeInsert(status); // 👈 共通関数を直接呼ぶ（2回目のチェックをスキップ！）
    
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

// 💡 共通の書き込みロジック（シート操作のみを行う軽い関数）
function _executeInsert(status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AttendanceLog");

  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail();

  // ログに登録
  sheet.appendRow([timestamp, user, status]);

  // 打刻した時刻を返す
  return Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd HH:mm:ss");
}

// アクティブユーザーのステータスとセッションIDを取得
function getStatusOfActiveUser() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Userシートのメモリ上での一発検索 ---
  const userSheet = ss.getSheetByName('User');
  const userData = userSheet.getDataRange().getValues(); // 全データを2次元配列として一括取得
  
  // A列(Email)が一致する行を高速検索
  const userRow = userData.find(row => row[0] === email);
  
  let userName = "ゲストユーザー";
  let sessionId = null;
  
  if (userRow) {
    userName = userRow[1] || "名前未設定"; // B列: Name
    sessionId = userRow[4] || null;        // E列: session_id
  }

  // --- 2. AttendanceLogシートの高速スキャン ---
  const logSheet = ss.getSheetByName("AttendanceLog");
  const lastRow = logSheet.getLastRow();
  let lastAction = null;

  if (lastRow >= 2) {
    const logValues = logSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const latestLog = logValues.reverse().find(row => row[1] === email);
    
    if (latestLog) {
      lastAction = {
        timestamp: Utilities.formatDate(latestLog[0], "JST", "yyyy/MM/dd HH:mm:ss"), // A列: 日時
        status: latestLog[2] // C列: ステータス (IN / OUT)
      };
    }
  }

  // --- 3. 結果をまとめて返す ---
  return {
    email: email,
    name: userName,
    sessionId: sessionId, 
    lastAction: lastAction
  };
}

// 期間指定付きのデータ取得 + 現在の稼働状況（マスタ全表示版）
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
  let totalUserCount = 0;

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
        isCurrentlyIn: false
      };
    }
  });

  // --- 2. 現在の入室状態（isCurrentlyIn）を配列の逆順から一発判定 ---
  const checkedUsers = new Set();
  for (let i = logRows.length - 1; i >= 1; i--) {
    const [_, email, status] = logRows[i];
    if (!email || !userMap[email] || checkedUsers.has(email)) continue;
    
    if (status === 'IN') userMap[email].isCurrentlyIn = true;
    checkedUsers.add(email);
    if (checkedUsers.size === totalUserCount) break; 
  }

  // --- 3. 指定期間内のログだけをピンポイントで解析 ---
  logRows.slice(1).forEach(row => {
    const [timestamp, email, status] = row;
    if (!timestamp || !email || !status || !userMap[email]) return;

    const logTimeMs = timestamp instanceof Date ? timestamp.getTime() : new Date(timestamp).getTime();
    if (logTimeMs < startDate.getTime() || logTimeMs > endDate.getTime()) return;

    const dateStr = Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(timestamp, "JST", "HH:mm");

    if (status === 'IN') {
      userMap[email].inCount++;
      userMap[email].logs.push({ date: dateStr, in: timeStr, out: '', duration: '', rawIn: logTimeMs });
    } else if (status === 'OUT') {
      const userLogs = userMap[email].logs;
      const lastLog = userLogs[userLogs.length - 1];
      if (lastLog && lastLog.in && !lastLog.out) {
        lastLog.out = timeStr;
        const diffMs = logTimeMs - lastLog.rawIn;
        const hours = Math.floor(diffMs / (1000 * 60 * 60));
        const mins = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
        lastLog.duration = `${hours}h ${mins}m`;
      }
    }
  });

  // --- 4. 統計の算出 ---
  const diffDays = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
  const targetDays = diffDays; 

  let totalRateSum = 0;
  let currentlyInCount = 0;

  const userList = Object.values(userMap).map(user => {
    if (user.isCurrentlyIn) currentlyInCount++;

    user.rate = targetDays > 0 ? Math.round((user.inCount / targetDays) * 100) : 0;
    totalRateSum += user.rate;

    if (user.logs.length > 0) {
      user.logs.forEach(l => delete l.rawIn);
      user.logs.reverse(); // 新しいログ順にする
    }
    
    return user; 
  });

  return {
    stats: {
      averageRate: userList.length > 0 ? Math.round(totalRateSum / userList.length) : 0,
      currentlyIn: currentlyInCount,
      totalActiveUsers: totalUserCount,
      activeInPeriod: userList.length 
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
  return [];
}

/**
 * ============================================================================
 * 【約束事2】「アプリを公開する」ボタンを押したときに自動作成するトリガー
 * ============================================================================
 * 変更：期限切れセッション削除の処理が不要になったため、設定配列を空に変更しました。
 */
function getAppTriggerConfig() {
  return [];
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
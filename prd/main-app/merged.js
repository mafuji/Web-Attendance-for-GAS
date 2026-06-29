// --- File: main.js ---
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

// 設定オブジェクトを返す（AllowedIp + AllowedIpFromAsn 統合対応版）
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シート
  const configSheet = ss.getSheetByName('Config');
  const useIpControl = configSheet.getRange('C2').getValue();
  const requirePassword = configSheet.getRange('D2').getValue();

  let ipRules = [];

  // ① AllowedIpシートから許可IPリストを取得
  const allowedIpSheet = ss.getSheetByName('AllowedIp');
  if (allowedIpSheet) {
    const data = allowedIpSheet.getDataRange().getValues();
    const rules = data.slice(1).map(row => String(row[0]).trim()).filter(Boolean);
    ipRules = ipRules.concat(rules);
  }

  // ② AllowedIpFromAsnシートから許可IPリストを取得（★ここを追加）
  const allowedIpFromAsnSheet = ss.getSheetByName('AllowedIpFromAsn');
  if (allowedIpFromAsnSheet) {
    const dataFromAsn = allowedIpFromAsnSheet.getDataRange().getValues();
    const rulesFromAsn = dataFromAsn.slice(1).map(row => String(row[0]).trim()).filter(Boolean);
    ipRules = ipRules.concat(rulesFromAsn);
  }

  return {
    requirePassword: requirePassword,
    useIpControl: useIpControl,
    allowedIps: ipRules // 両方のシートのルールがマージされてここに入る
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

// 💡 修正：打刻レコード追加（引数に userIp を追加）
function insertRecord(status, sessionId, userIp) {
  // 1. セッションチェック
  if (!isSessionValid(sessionId)) return null;

  // 2. ⚡サーバー側でIPアドレスの検証
  if (!_checkIpOnServer(userIp)) {
    throw new Error("許可されていないネットワーク（研究室外）からの打刻はできません。");
  }

  return _executeInsert(status, "", userIp);
}

// パスワードを検証し、正しければ打刻
function verifyPasswordAndInsert(status, sessionId, inputPassword, userIp) {
  // 1. セッションチェック
  if (!isSessionValid(sessionId)) return null;

  // 2. ⚡サーバー側でIPアドレスの検証
  if (!_checkIpOnServer(userIp)) {
    return { success: false, message: '許可されていないネットワーク（研究室外）からの打刻はできません。' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) return { success: false, message: 'システムエラー:「Config」シートが見つかりません。' };
  
  const correctPassword = configSheet.getRange("A2").getValue().toString(); 
  const previousPassword = configSheet.getRange("B2").getValue().toString(); 

  if (inputPassword !== correctPassword && inputPassword !== previousPassword) {
    return { success: false, message: 'パスコードが一致しません。最新のパスコードを入力してください。' };
  }

  try {
    const formattedTime = _executeInsert(status, inputPassword, userIp);
    return { success: true, timestamp: formattedTime };
  } catch(e) {
    console.error("打刻エラー: ", e);
    return { success: false, message: '打刻処理中にエラーが発生しました: ' + e.message };
  }
}

// レコード追加処理
function _executeInsert(status, inputPassword, userIp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AttendanceLog");

  const timestamp = new Date();
  const userEmail = Session.getActiveUser().getEmail();
  
  // ユーザー名をUserシートから取得（名前をログに残すため）
  const userSheet = ss.getSheetByName('User');
  const userData = userSheet.getDataRange().getValues();
  const userRow = userData.find(row => row[0] === userEmail);
  const userName = userRow ? (userRow[1] || "名前未設定") : "ゲストユーザー";

  // 値の正規化（nullやundefinedを弾く）
  const pwdLog = inputPassword ? "'" + String(inputPassword) : "";
  const ipLog = userIp ? String(userIp) : "";
  const noteLog = ""; // G列: note (初期値は空)

  // ⚡【新仕様】に完全準拠して1行追加
  // A: user_email, B: user_name, C: time_stamp, D: status, E: input_password, F: ip_address, G: note
  sheet.appendRow([
    userEmail,
    userName,
    timestamp,
    status,
    pwdLog,
    ipLog,
    noteLog
  ]);

  return Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd HH:mm:ss");
}

// 💡 新規追加：サーバー側のCIDR対応IPチェックロジック
function _checkIpOnServer(userIp) {
  const config = getConfig();
  if (!config.useIpControl) return true; // 設定でIPチェックがオフならスルー
  if (!userIp) return false; // IPが送られてきていなければ拒否

  const allowedRules = config.allowedIps;
  
  return allowedRules.some(rule => {
    if (rule === userIp) return true;

    if (rule.includes('/')) {
      const parts = rule.split('/');
      const range = parts[0];
      const bits = parseInt(parts[1], 10);
      
      const ipNum = _ipToLong(userIp);
      const rangeNum = _ipToLong(range);
      
      if (ipNum === null || rangeNum === null) return false;

      const mask = bits === 0 ? 0 : (~0 << (32 - bits));
      return (ipNum & mask) === (rangeNum & mask);
    }
    return false;
  });
}

function _ipToLong(ip) {
  const parts = ip.split('.');
  if (parts.length !== 4) return null;
  return parts.reduce((acc, part) => (acc << 8) + parseInt(part, 10), 0) >>> 0;
}

// アクティブユーザーのステータスとセッションIDを取得
function getStatusOfActiveUser() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Userシートのメモリ上での一発検索 ---
  const userSheet = ss.getSheetByName('User');
  const userData = userSheet.getDataRange().getValues();
  const userRow = userData.find(row => row[0] === email);
  
  let userName = "ゲストユーザー";
  let sessionId = null;
  
  if (userRow) {
    userName = userRow[1] || "名前未設定";
    sessionId = userRow[4] || null;
  }

  // --- 2. AttendanceLogシートの高速スキャン ---
  const logSheet = ss.getSheetByName("AttendanceLog");
  const lastRow = logSheet.getLastRow();
  let lastAction = null;

  if (lastRow >= 2) {
    // ⚡ 取得範囲をA列(1)〜D列(4)までに拡張
    const logValues = logSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    // A列（row[0]）にemailが入っているので、そこを基準に検索
    const latestLog = logValues.reverse().find(row => row[0] === email);
    
    if (latestLog) {
      lastAction = {
        timestamp: Utilities.formatDate(latestLog[2], "JST", "yyyy/MM/dd HH:mm:ss"), // ⚡ C列(index 2)が日時
        status: latestLog[3] // ⚡ D列(index 3)がステータス (IN / OUT)
      };
    }
  }

  return {
    email: email,
    name: userName,
    sessionId: sessionId, 
    lastAction: lastAction
  };
}

// 期間指定付きのデータ取得 + 現在の稼働状況（マスタ全表示版）
// 💡 大改造：集計ロジックの列割り当て変更
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

  // --- 2. 現在の入室状態（isCurrentlyIn）の一発判定 ---
  const checkedUsers = new Set();
  for (let i = logRows.length - 1; i >= 1; i--) {
    // ⚡ 新列仕様：A列[0]=email, D列[3]=status
    const email = logRows[i][0];
    const status = logRows[i][3];
    
    if (!email || !userMap[email] || checkedUsers.has(email)) continue;
    
    if (status === 'IN') userMap[email].isCurrentlyIn = true;
    checkedUsers.add(email);
    if (checkedUsers.size === totalUserCount) break; 
  }

  // --- 3. 指定期間内のログだけをピンポイントで解析 ---
  logRows.slice(1).forEach(row => {
    // ⚡ 新列仕様から必要な要素だけをマッピング
    // A: email [0], C: timestamp [2], D: status [3]
    const email = row[0];
    const timestamp = row[2];
    const status = row[3];
    
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
      delete user.logs.rawIn;
      user.logs.forEach(l => delete l.rawIn);
      user.logs.reverse(); 
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
/**
 * 【約束事1】アプリ固有のカスタムメニュー構成を返す
 */
function getAppMenuConfig() {
  return [
    { type: "item", name: "🟢 ASN自動更新（1日1回）をONにする", functionName: "main_createAsnTriggerOnly" },
    { type: "item", name: "🛑 ASN自動更新（1日1回）をOFFにする", functionName: "main_deleteAsnTriggerOnly" },
    { type: "item", name: "⚡ ASNからCIDRを今すぐ手動更新", functionName: "main_refreshCidrFromAsnNow" },
    { type: "separator" }, // 境界線
    { type: "item", name: "🗑️ 未登録ユーザーのログ一括削除", functionName: "main_deleteUnknownUserLogs" } 
  ];
}
/**
 * カスタムメニュー「未登録ユーザーのログ一括削除」の実装
*/
function main_deleteUnknownUserLogs() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. 実行前の確認アラート
  const response = ui.alert(
    '確認', 
    'Userシートに登録されていないユーザーの打刻ログをすべて削除します。よろしいですか？', 
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  // 2. 実処理の関数を呼び出し、削除件数を受け取る
  const deletedCount = deleteUnknownUserLogsOnly();

  // 3. 実行後の結果アラート
  if (deletedCount === 0) {
    ui.alert('結果', '未登録ユーザーのログは見つかりませんでした。データはすべて綺麗です！', ui.ButtonSet.OK);
  } else {
    ui.alert('完了', `未登録ユーザーのログを ${deletedCount} 件削除しました。`, ui.ButtonSet.OK);
  }
}

/**
 * ★新規追加：メニューからASN自動更新をONにする
 */
function main_createAsnTriggerOnly() {
  createAsnTriggerOnly();

  const ui = SpreadsheetApp.getUi();
  ui.alert('定期処理の開始', '🟢 ASNの自動更新を開始しました。\n今後、毎日深夜1時〜2時の間に最新のCIDRへ自動洗い替えされます。', ui.ButtonSet.OK);
}

/**
 * ★新規追加：メニューからASN自動更新をOFFにする
 */
function main_deleteAsnTriggerOnly() {
  deleteAsnTriggerOnly();

  const ui = SpreadsheetApp.getUi();
  ui.alert('定期処理の停止', '🛑 ASNの自動更新を停止しました。', ui.ButtonSet.OK);
}
/**
 * カスタムメニューから手動でASN同期を実行するためのラッパー
 */
function main_refreshCidrFromAsnNow() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認', 'AllowedAsnシートに記載されたASN情報を元に、CIDRリストを今すぐ最新に洗い替えします。よろしいですか？', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    refreshCidrFromAsn(); // ⑥の実処理を呼び出し
    ui.alert('完了', '🟢 ASNからCIDRリストの同期・洗い替えが完了しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', '同期中にエラーが発生しました: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * ============================================================================
 * 【約束事2】「アプリを公開する」ボタンを押したときに自動作成するトリガー
 * ============================================================================
 */
function getAppTriggerConfig() {
  return [
    {
      // 1日1回深夜にASN情報を同期するトリガー定義
      functionName: "refreshCidrFromAsn",
      methods: [
        { name: "everyDays", args: [1] },
        { name: "atHour", args: [1] } // 深夜1時〜2時の間に実行
      ]
    }    
  ];
}

/**
 * ============================================================================
 * 処理ロジック
 * ============================================================================
 */
/**
 * ★新規追加：refreshCidrFromAsn トリガーを直接新規作成する
 */
function createAsnTriggerOnly() {
  deleteAsnTriggerOnly(); // 重複防止

  ScriptApp.newTrigger('refreshCidrFromAsn')
    .timeBased()
    .everyDays(1)
    .atHour(1) // 深夜1時〜2時の間に実行
    .create();
}
/**
 * ★新規追加：refreshCidrFromAsnだけをピンポイントで削除する
 */
function deleteAsnTriggerOnly() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshCidrFromAsn') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
/**
 * 💡 ASNからCIDRを取得して AllowedIpFromAsn を洗い替えする実処理（HackerTarget版）
 */
function refreshCidrFromAsn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. AllowedAsn シートから取得対象のASN一覧を取り出す
  const asnSheet = ss.getSheetByName('AllowedAsn');
  if (!asnSheet) {
    console.warn("AllowedAsnシートが見つかりません。処理をスキップします。");
    return;
  }
  const asnData = asnSheet.getDataRange().getValues();
  // ヘッダーを除き、"AS1234" や "1234" から数字のみを抽出して、"AS12345" の形式に統一
  const asnList = asnData.slice(1).map(row => {
    const match = String(row[0]).match(/\d+/);
    return match ? `AS${match[0]}` : null;
  }).filter(Boolean);

  if (asnList.length === 0) {
    console.warn("AllowedAsnシートに対象のASNが記載されていません。");
    return;
  }

  // 2. 各ASNに対して HackerTarget API を叩いてCIDRを一括取得
  let allPrefixes = [];
  asnList.forEach(asn => {
    // HackerTargetのURLに整形 (例: https://api.hackertarget.com/aslookup/?q=AS2500)
    const url = `https://api.hackertarget.com/aslookup/?q=${asn}`;
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      console.log(response);
      if (response.getResponseCode() === 200) {
        const text = response.getContentText();
        
        // HackerTargetは改行区切りのプレーンテキストを返すので行ごとに分割
        const lines = text.split('\n');
        
        lines.forEach(line => {
          const trimmedLine = line.trim();
          // 空行や、エラーメッセージ（"error"等から始まる行）を除外
          if (trimmedLine && !trimmedLine.toLowerCase().startsWith('error')) {
            // 最初に見つかるASN情報行（ASN,NAME）は無視し、CIDR表記（スラッシュを含む行）だけを抽出
            if (trimmedLine.includes('/')) {
              allPrefixes.push(trimmedLine);
            }
          }
        });
      } else {
        console.warn(`ASN: ${asn} の情報取得に失敗しました。ステータスコード: ${response.getResponseCode()}`);
      }
    } catch (e) {
      console.error(`ASN: ${asn} の通信エラー: ` + e.toString());
    }
  });

  // 重複を除去
  allPrefixes = [...new Set(allPrefixes)].filter(Boolean);

  // 3. AllowedIpFromAsn シートのデータをクリアして新リストで洗い替え
  let targetSheet = ss.getSheetByName('AllowedIpFromAsn');
  if (!targetSheet) {
    targetSheet = ss.insertSheet('AllowedIpFromAsn');
  }
  
  targetSheet.clearContents(); // ヘッダー含め全クリア
  targetSheet.getRange(1, 1).setValue('allowed_cidr_from_asn'); // 新たにヘッダー書き込み

  if (allPrefixes.length > 0) {
    // 2次元配列に変換して一括書き込み
    const writeData = allPrefixes.map(prefix => [prefix]);
    targetSheet.getRange(2, 1, writeData.length, 1).setValues(writeData);
  }

  console.log(`AllowedIpFromAsn を更新(HackerTarget)しました。取得ASN件数: ${asnList.length}, 総CIDR数: ${allPrefixes.length}`);
  SpreadsheetApp.flush();
}
/**
 * Userシートに存在しないユーザーの打刻データをAttendanceLogシートから全て削除する（実処理）
 * @return {number} 削除したレコード件数
 */
function deleteUnknownUserLogsOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Userシートから登録済みEmailの一覧を一括取得してSetに格納
  const userSheet = ss.getSheetByName('User');
  const userLastRow = userSheet.getLastRow();
  if (userLastRow < 2) return 0; // ユーザーがいない場合は処理不可

  const registeredEmails = new Set(
    userSheet.getRange(2, 1, userLastRow - 1, 1).getValues().map(row => String(row[0]).trim())
  );

  // 2. AttendanceLogシートの全データを一括取得
  const logSheet = ss.getSheetByName('AttendanceLog');
  const logLastRow = logSheet.getLastRow();
  if (logLastRow < 2) return 0; // ログがない場合は処理なし
  
  const logMaxColumns = logSheet.getLastColumn();
  const logRange = logSheet.getRange(2, 1, logLastRow - 1, logMaxColumns);
  const logValues = logRange.getValues();

  // 3. メモリ上で「登録済みユーザーのログだけ」を残す
  const filteredLogs = logValues.filter(row => {
    const email = row[0] ? String(row[0]).trim() : "";
    return registeredEmails.has(email);
  });

  // 削除対象の件数を計算
  const deletedCount = logValues.length - filteredLogs.length;
  if (deletedCount === 0) return 0;

  // 4. シートをクリアして書き戻し
  logRange.clearContent();
  if (filteredLogs.length > 0) {
    logSheet.getRange(2, 1, filteredLogs.length, logMaxColumns).setValues(filteredLogs);
  }

  // GAS側の書き込みを強制確定
  SpreadsheetApp.flush();

  return deletedCount;
}

/**
 * ============================================================================
 * merged.gs用
 * ============================================================================
 */
function getHtmlTemplate(fileName){
  // 開発時はファイルから読込
  try {
      return HtmlService.createTemplateFromFile(fileName);
  } catch (e) {  
    const templates = {
  "base.css": "<style>\n  /* 全体に適用 */\n  body {\n    font-family: 'Noto Sans JP', sans-serif;\n  }\n\n  /* 共通のボタン設定 */\n  button {\n    padding: 12px 32px;\n    font-size: 16px;\n    font-weight: bold;\n    border: none;\n    border-radius: 8px; /* 角丸でモダンに */\n    cursor: pointer;\n    transition: all 0.3s ease; /* 動きをなめらかに */\n    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); /* 軽い影で浮かせる */\n    outline: none;\n  }\n\n  /* ボタンを押した瞬間の沈み込み */\n  button:active {\n    transform: translateY(0);\n    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);\n  }\n\n  /* ローディング全体の配置 */\n  .loading-container {\n    display: flex;\n    flex-direction: column;\n    align-items: center;\n    justify-content: center;\n    height: 100vh; /* 画面いっぱいに表示 */\n    color: #666;\n  }\n\n  /* くるくる回るアニメーション */\n  .spinner {\n    width: 40px;\n    height: 40px;\n    border: 4px solid #f3f3f3;\n    border-top: 4px solid #3498db; /* 青色 */\n    border-radius: 50%;\n    animation: spin 1s linear infinite;\n    margin-bottom: 10px;\n  }\n  \n  @keyframes spin {\n    0% { transform: rotate(0deg); }\n    100% { transform: rotate(360deg); }\n  }\n\n  /* --- ヘッダーレイアウト --- */\n  .app-header {\n    display: flex;\n    justify-content: space-between;\n    align-items: center;\n    padding: 0 20px;\n    background-color: #2c3e50;\n    color: white;\n    height: 60px;\n    position: sticky;\n    top: 0;\n    z-index: 1000;\n    box-shadow: 0 2px 8px rgba(0,0,0,0.15);\n  }\n\n  .header-logo {\n    font-weight: bold;\n    font-size: 1.1rem;\n    letter-spacing: 1px;\n  }\n\n  .header-right {\n    position: relative;\n  }\n\n  /* --- ハンバーガーボタン --- */\n  .menu-trigger {\n    background: none;\n    border: none;\n    width: 20px; \n    height: 16px; /* 偶数にして計算を安定させる */\n    position: relative;\n    cursor: pointer;\n    display: block;\n    padding: 0;\n    overflow: visible; /* 描画が切れるのを防ぐ */\n  }\n\n  .menu-trigger span {\n    display: block;\n    position: absolute;\n    width: 100%;\n    height: 2px;\n    background: #fff;\n    transition: all .3s ease;\n    border-radius: 2px;\n    left: 0;\n  }\n\n  /* 各線の位置（キリの良い数字で配置） */\n  .menu-trigger span:nth-child(1) { top: 0; }\n  .menu-trigger span:nth-child(2) { top: 7px; } \n  .menu-trigger span:nth-child(3) { top: 14px; } /* topからの絶対値 */\n\n  /* --- アニメーション時の挙動（中央 7px に集める） --- */\n  .menu-trigger.is-active span:nth-child(1) { \n    transform: translateY(7px) rotate(-45deg); \n  }\n  .menu-trigger.is-active span:nth-child(2) { \n    opacity: 0; \n    transform: translateX(-5px); /* 左に逃がすことで消え方を綺麗に */\n  }\n  .menu-trigger.is-active span:nth-child(3) { \n    transform: translateY(-7px) rotate(45deg); \n  }\n\n  /* --- ドロップダウンメニュー --- */\n  .dropdown-menu {\n    position: absolute;\n    top: 50px;\n    right: -10px;\n    background: white;\n    color: #333;\n    box-shadow: 0 10px 30px rgba(0,0,0,0.2);\n    border-radius: 12px;\n    width: 240px;\n    border: 1px solid #eee;\n    overflow: hidden;\n  }\n\n  .dropdown-menu ul {\n    list-style: none;\n    margin: 0;\n    padding: 0;\n  }\n\n  .dropdown-menu li a {\n    display: block;\n    padding: 18px 20px;\n    text-decoration: none;\n    color: #333;\n    font-size: 1rem;\n    transition: background 0.2s;\n  }\n\n  .dropdown-menu li a:hover {\n    background: #f0f4f8;\n    color: #007bff;\n  }\n\n  /* --- メニュー最下部のアカウントエリア --- */\n  .menu-footer {\n    padding: 15px 20px;\n    background-color: #f8f9fa;\n    border-top: 2px solid #eee; /* ここでボーダー分け */\n  }\n\n  .user-label {\n    font-size: 0.65rem;\n    color: #999;\n    font-weight: bold;\n    margin-bottom: 4px;\n  }\n\n  .user-email {\n    font-size: 0.8rem;\n    color: #666;\n    word-break: break-all;\n    line-height: 1.4;\n  }\n\n  /* --- アニメーション --- */\n  .fade-enter-active, .fade-leave-active { transition: opacity .2s, transform .2s; }\n  .fade-enter-from, .fade-leave-to { opacity: 0; transform: translateY(-10px); }\n\n  /* --- コンテンツ調整 --- */\n  .content-container { padding: 20px; }\n  .status-card { margin-top: 30px; text-align: center; color: #555; }\n</style>",
  "head_common": "<base target=\"_top\">\n<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n<link href=\"https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap\" rel=\"stylesheet\">\n<meta charset=\"UTF-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">",
  "index.css": "<style>\n  /* --- メイン --- */\n  /* 全体のレイアウト調整 */\n  #app {\n    display: flex;\n    flex-direction: column; /* ユーザー情報とボタンを「縦並び」にする */\n    align-items: center;    /* 中央寄せ */\n    margin-top: 50px;\n    gap: 25px;              /* ユーザー情報とボタンの間の隙間 */\n  }\n\n  /* 入室ボタン (青系) */\n  .button-in {\n    font-size: 32px;\n    background-color: #4A90E2;\n    color: white;\n  }\n\n  .button-in:hover {\n    background-color: #357ABD;\n    transform: translateY(-2px); /* 少し浮き上がる */\n    box-shadow: 0 6px 12px rgba(74, 144, 226, 0.3);\n  }\n\n  /* 退室ボタン (緑系) */\n  .button-out {\n    font-size: 32px;\n    background-color: #2ecc71; /* 鮮やかなエメラルドグリーン */\n    color: #ffffff;            /* 文字は白で見やすく */\n    border: none;              /* 枠線を消すとモダンになります */\n    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); /* 軽い影で立体感を */\n    transition: all 0.3s ease; /* 変化を滑らかに */\n  }\n\n  .button-out:hover {\n    background-color: #27ae60; /* ホバー時は少し濃い緑に */\n    color: #ffffff;\n    transform: translateY(-2px); /* 少し浮き上がる演出 */\n    box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15); /* 影も強く */\n  }\n\n  /* 直近の打刻 */\n  .last-action {\n    margin-top: 20px;\n    font-size: 0.9rem;\n    color: #666;\n    text-align: center;\n    border-top: 1px solid #eee;\n    padding-top: 10px;\n  } \n\n  /* セッション切れ表示 */\n  .expired-container {\n    display: flex;\n    justify-content: center;\n    align-items: center;\n    min-height: 80vh;\n    padding: 20px;\n  }\n\n  .expired-card {\n    background: white;\n    padding: 40px 30px;\n    border-radius: 12px;\n    box-shadow: 0 4px 20px rgba(0,0,0,0.08);\n    text-align: center;\n    max-width: 400px;\n    width: 100%;\n  }\n\n  .expired-icon {\n    font-size: 50px;\n    margin-bottom: 15px;\n  }\n\n  .expired-card h2 {\n    color: #333;\n    margin-bottom: 15px;\n    font-size: 1.5rem;\n  }\n\n  .expired-card p {\n    color: #666;\n    line-height: 1.6;\n    margin-bottom: 25px;\n    font-size: 0.95rem;\n  }  \n\n  /* ボタン配置エリア（元々の#appの横並びをここに移植） */\n  .button-area {\n    display: flex;\n    gap: 20px;              /* ボタン同士の隙間 */\n    justify-content: center;\n  }\n\n  /* --- かっこいいユーザー情報ボックス --- */\n  .user-info-box {\n    display: flex;\n    flex-direction: column; /* ラベルとメアドを縦並びに */\n    align-items: center;\n    background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);\n    padding: 16px 40px;\n    border-radius: 12px;\n    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05); /* 既存のカードに合わせた軽い影 */\n    border: 1px solid #e9ecef;\n    max-width: 90%;\n    transition: all 0.3s ease;\n  }\n\n  /* ホバー時に少し浮き上がる（全体のボタンの仕様と統一） */\n  .user-info-box:hover {\n    transform: translateY(-1px);\n    box-shadow: 0 6px 20px rgba(0, 0, 0, 0.08);\n  }\n\n  /* 小さめのグレーラベル */\n  .user-info-label {\n    font-size: 0.75rem;\n    color: #999;\n    font-weight: bold;\n    letter-spacing: 1px;\n    margin-bottom: 4px;\n    text-transform: uppercase;\n  }\n\n  /* 大きくてかっこいいメールアドレス */\n  .user-info-email {\n    font-size: 1.35rem;     /* ボタン（32px）に負けない大きめのサイズ */\n    color: #2c3e50;         /* ヘッダーの色（#2c3e50）と統一してスマートに */\n    font-weight: 700;\n    letter-spacing: 0.5px;\n    word-break: break-all;  /* 長いメアドでも画面外に崩れないようにガード */\n  }\n</style>",
  "index": "<!DOCTYPE html>\n<html lang=\"ja\">\n  <head>\n    <?!= include(\"head_common\") ?>\n    <?!= include(\"base.css\") ?>\n    <?!= include(\"index.css\") ?>\n    <title>Web入退室</title>\n    <script src=\"https://cdn.jsdelivr.net/npm/sweetalert2@11\" defer></script>\n  </head>\n  <body>\n    <script>\n      // GAS側からレンダリング時に直接注入されるので、絶対に空にならない\n      const SERVER_SESSION_ID =\"<?= sessionId ?>\";\n      const requirePassword = <?!= JSON.stringify(requirePassword) ?>;\n    </script>\n\n    <div id=\"main\">\n      <div v-if=\"!isReady\" class=\"loading-container\">\n        <div class=\"spinner\"></div>\n        <p>読み込み中...</p>\n      </div>\n      \n      <div v-else-if=\"isSessionExpired\" class=\"expired-container\">\n        <div class=\"expired-card\">\n          <div class=\"expired-icon\">⚠️</div>\n          <h2>セッションの有効期限切れ</h2>\n          <p>ブラウザをリロードしてください。</p>\n        </div>\n      </div>\n\n      <header class=\"app-header\" v-if=\"isReady && !isSessionExpired\">\n        <div class=\"header-logo\">Web入退室</div>\n\n        <div class=\"header-right\">\n          <button class=\"menu-trigger\" @click=\"isMenuOpen = !isMenuOpen\" :class=\"{ 'is-active': isMenuOpen }\">\n            <span></span>\n            <span></span>\n            <span></span>\n          </button>\n\n          <transition name=\"fade\">\n            <nav class=\"dropdown-menu\" v-show=\"isMenuOpen\">\n              <ul>\n                <li v-if=\"isAdmin\"><a href=\"#\" @click.prevent=\"openDashboard\">ダッシュボード</a></li>\n                <li><a href=\"#\" @click.prevent=\"logout\">ログアウト</a></li>\n                <li class=\"menu-footer\">\n                  <span class=\"user-label\">ログイン中：</span>\n                  <span class=\"user-email\">{{ email }}</span>\n                </li>                \n              </ul>\n            </nav>\n          </transition>\n        </div>\n      </header>\n\n      <div id=\"app\" v-if=\"isReady && !isSessionExpired\">\n        <div class=\"user-info-box\">\n          <span class=\"user-info-label\">ユーザー：</span>\n          <span class=\"user-info-email\">{{ name }} さん</span>\n        </div>\n\n        <div class=\"button-area\">\n          <button @click=\"sendIn\" v-show=\"!isIn\" class=\"button-in\">入室</button>\n          <button @click=\"sendOut\" v-show=\"isIn\" class=\"button-out\">退室</button>\n        </div>\n        <div class=\"last-action\" v-if=\"lastAction.time && !isSessionExpired\">\n            前回の打刻：{{ lastAction.time }} （{{ lastAction.status }}）<br/>\n        </div>        \n      </div>\n    </div>\n    <script src=\"https://cdn.jsdelivr.net/npm/vue@3.1.5\"></script>\n    <?!= include(\"index.js\") ?>\n  </body>\n</html>",
  "index.js": "<script>\n  Vue.createApp({\n    data() {\n      return {\n        isReady: false,\n        isIn: false,\n        isAdmin: false,\n        isMenuOpen: false, \n        email: \"\", // 💡 残す\n        name: \"\",  // 💡 追加\n        sessionId: SERVER_SESSION_ID,\n        requirePassword: requirePassword,\n        isSessionExpired: false,\n        lastAction: { time: \"\", status: \"\" }\n      }\n    },\n    async mounted() {\n      // 💡 2つの通信を同時にスタートさせる（並列呼び出し）\n      this.fetchUserData();\n      this.checkAdmin();\n    },\n    methods: {\n      // ユーザー情報の取得\n      fetchUserData() {\n        google.script.run\n          .withFailureHandler(() => { this.checkReady(); })\n          .withSuccessHandler(res => {\n            this.email = res.email;\n            this.name = res.name;\n            if (res.lastAction) {\n              this.lastAction = {\n                time: res.lastAction.timestamp,\n                status: res.lastAction.status === \"IN\" ? \"入室\" : \"退室\"\n              };\n              this.isIn = (res.lastAction.status === \"IN\");\n            }\n            this.checkReady(); // 💡 終わったら準備完了チェックへ\n          })\n          .getStatusOfActiveUser();\n      },\n\n      // 管理者判定\n      checkAdmin() {\n        google.script.run\n          .withFailureHandler(() => { this.checkReady(); })\n          .withSuccessHandler(res => {\n            this.isAdmin = res;\n            this.checkReady(); // 💡 終わったら準備完了チェックへ\n          })\n          .isAdmin();\n      },\n\n      // 💡 新設：両方のデータが揃ったら画面を表示する判定関数\n      checkReady() {\n        // email（UserDataの結果）と isAdmin の結果がどちらも取得できたら表示フラグを立てる\n        if ((this.email || this.isReady) && (this.isAdmin !== undefined)) {\n          this.isReady = true;\n        }\n      },\n\n      // パスワード認証とIP送信を挟んだ打刻処理\n      async send(status) {\n        let password = null;\n\n        Swal.fire({\n          title: '処理中...',\n          allowOutsideClick: false,\n          allowEscapeKey: false,\n          showConfirmButton: false,\n          didOpen: () => { Swal.showLoading(); }\n        });\n\n        // 💡 ⚡ボタンが押された瞬間に、超高速で自身のIPを取得（画面起動時ではなく今やる）\n        let userIp = \"\";\n        try {\n          const res = await fetch('https://api.ipify.org?format=json');\n          const data = await res.json();\n          userIp = data.ip;\n        } catch(e) {\n          Swal.fire(\"エラー\", \"ネットワーク環境が確認できませんでした。研究室のWi-Fiに接続されているか確認してください。\", \"error\");\n          return;\n        }\n\n        if (status === \"IN\" && requirePassword === true) {\n          Swal.close(); // パスワード入力のために一度ローディングを閉じる\n          const { value: typedPassword } = await Swal.fire({\n            title: 'パスワード入力',\n            input: 'password',\n            inputAttributes: {\n              maxlength: '4',\n              inputmode: 'numeric',\n              pattern: '[0-9]*',\n              style: 'text-align: center; font-size: 24px; letter-spacing: 4px;'\n            },\n            showCancelButton: true,\n            cancelButtonText: 'キャンセル',\n            confirmButtonText: '次へ',\n            inputValidator: (value) => {\n              if (!value || value.length !== 4 || isNaN(value)) {\n                return '4桁の数字を入力してください';\n              }\n            }\n          });\n\n          if (!typedPassword) return;\n          password = typedPassword;\n\n          // 再度ローディングを表示\n          Swal.fire({ title: '処理中...', allowOutsideClick: false, showConfirmButton: false, didOpen: () => { Swal.showLoading(); } });\n        }\n\n        const runner = google.script.run\n          .withFailureHandler(e => {\n            Swal.fire(\"エラー\", \"通信に失敗しました: \" + e, \"error\");\n          })\n          .withSuccessHandler(res => {\n            if (res === null) {\n              this.isSessionExpired = true;\n              Swal.close();\n              return;\n            }\n            \n            let displayTime = \"\";\n\n            if (status === \"IN\" && requirePassword === true) {\n              if (res.success === false) {\n                // 💡 GAS側から「IPエラー」か「パスワードエラー」のメッセージが返ってくる\n                Swal.fire(\"打刻失敗\", res.message, \"error\");\n                return;\n              }\n              displayTime = res.timestamp;\n            } else {\n              displayTime = res;\n            }\n\n            this.isIn = (status === \"IN\");\n            this.lastAction = {\n              time: displayTime,\n              status: status === \"IN\" ? \"入室\" : \"退室\"\n            };\n\n            Swal.fire({\n              title: '完了',\n              text: status === \"IN\" ? \"入室を記録しました\" : \"退室を記録しました\",\n              icon: 'success',\n              timer: 2000,\n              showConfirmButton: false\n            });\n          });\n\n        // 💡 修正：末尾に取得した userIp を引数として追加してGASにまとめてブチ込む（通信は1回のみ）\n        if (status === \"IN\" && requirePassword === true) {\n          runner.verifyPasswordAndInsert(status, this.sessionId, password, userIp);\n        } else {\n          runner.insertRecord(status, this.sessionId, userIp);\n        }\n      },\n\n      sendIn() { this.send(\"IN\"); },\n      sendOut() { this.send(\"OUT\"); },\n\n      openDashboard() {\n        this.isMenuOpen = false;\n        google.script.run\n          .withSuccessHandler(url => {\n            window.top.location.href = url + \"?p=summary\";\n          })\n          .getAppUrl();\n      },\n\n      logout() {\n        this.isMenuOpen = false;\n        Swal.fire({\n          title: 'ログアウトしますか？',\n          text: \"Googleアカウントのログアウトページへ移動します\",\n          icon: 'question',\n          showCancelButton: true,\n          confirmButtonText: 'ログアウト',\n          cancelButtonText: 'キャンセル'\n        }).then((result) => {\n          if (result.isConfirmed) {\n            google.script.run\n              .withSuccessHandler(function(url) {\n                var logoutUrl = 'https://accounts.google.com/Logout?continue=' + \n                                encodeURIComponent('https://accounts.google.com/ServiceLogin?continue=' + url);\n                window.top.location.href = logoutUrl;\n              })\n              .getAppUrl();\n          }\n        });\n      }\n    }\n  }).mount(\"#main\");\n</script>",
  "summary.css": "<style>\n/* 全体の土台 */\n  .dashboard-body {\n    background-color: #f8fafc; /* slate-50 */\n    min-height: 100vh;\n    font-family: sans-serif;\n    color: #0f172a; /* slate-900 */\n    margin: 0;\n  }\n  \n  .container-main {\n    max-width: 64rem; /* max-w-5xl */\n    margin: 0 auto;\n    padding: 2.5rem 1rem;\n  }\n\n  /* 共通の統計カード */\n  .stats-card {\n    background: #ffffff;\n    padding: 1.5rem;\n    border-radius: 1rem;\n    border: 1px solid #e2e8f0;\n    box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);\n  }\n  .header-section {\n    margin-bottom: 2.5rem;\n    display: flex;\n    flex-direction: column;\n    justify-content: space-between;\n    gap: 1rem;\n  }\n\n  /* PCサイズでは横並び */\n  @media (min-width: 768px) {\n    .header-section { flex-direction: row; align-items: flex-end; }\n  }\n\n  .header-title {\n    font-size: 1.875rem;\n    font-weight: 800;\n    color: #1e293b;\n    margin: 0;\n  }\n\n  .date-filter {\n    display: flex;\n    align-items: center;\n    gap: 0.5rem;\n    background: #ffffff;\n    padding: 0.5rem;\n    border-radius: 0.75rem;\n    border: 1px solid #e2e8f0;\n  }\n\n  .date-input {\n    background: transparent;\n    border: none;\n    font-size: 0.875rem;\n    outline: none;\n    text-align: center;\n  }\n\n  .stats-grid {\n    display: grid;\n    grid-template-columns: 1fr;\n    gap: 1rem;\n    margin-bottom: 2rem;\n  }\n\n  @media (min-width: 768px) {\n    .stats-grid { grid-template-columns: repeat(2, 1fr); }\n  }\n\n  .stats-label {\n    font-size: 0.875rem;\n    font-weight: 500;\n    color: #64748b;\n    text-transform: uppercase;\n    margin: 0;\n  }\n\n  .stats-value {\n    font-size: 1.875rem;\n    font-weight: 700;\n    margin-top: 0.25rem;\n    margin-bottom: 0;\n    display: flex;\n    align-items: baseline;\n    gap: 0.25rem;    \n  }\n\n  .text-blue-600 { color: #2563eb; }\n\n  /* アコーディオン制御 */\n  details > summary { list-style: none; outline: none; }\n  details > summary::-webkit-details-marker { display: none; }\n  \n  details[open] .group-open-rotate { \n    transform: rotate(180deg); \n  }\n\n  .user-list-container {\n    display: flex;\n    flex-direction: column;\n    gap: 0.75rem;\n  }  \n\n  /* ステータスバッジ */\n  .status-badge {\n    padding: 0.2rem 0.6rem;\n    border-radius: 9999px;\n    font-size: 0.75rem;\n    font-weight: 700;\n    line-height: 1;\n    display: inline-flex;\n    align-items: center;\n  }\n\n  /* 在室中（緑） */\n  .badge-in {\n    background-color: #ecfdf5; /* emerald-50 */\n    color: #059669; /* emerald-600 */\n    border: 1px solid #10b981; /* emerald-500 */\n  }\n\n  /* 不在（グレー） */\n  .badge-out {\n    background-color: #f1f5f9; /* slate-100 */\n    color: #64748b; /* slate-500 */\n    border: 1px solid #cbd5e1; /* slate-300 */\n  }  \n</style>",
  "summary": "<!DOCTYPE html>\n<html lang=\"ja\">\n  <head>\n    <?!= include(\"head_common\") ?>    \n    <?!= include(\"base.css\") ?>\n    <?!= include(\"summary.css\") ?>\n    <title>Web入退室ダッシュボード</title>\n    <script src=\"https://cdn.tailwindcss.com\" defer></script>\n  </head>\n  <body class=\"dashboard-body\">\n    <div class=\"container-main\">\n      <header class=\"header-section\">\n        <h1 class=\"header-title\">Web入退室ダッシュボード</h1>\n        \n        <div class=\"date-filter flex items-center gap-2\"> <input type=\"date\" class=\"date-input\" value=\"2024-03-01\">\n          <span class=\"date-separator\">〜</span>\n          <input type=\"date\" class=\"date-input\" value=\"2024-03-31\">\n          <button id=\"search-btn\" class=\"ml-2 px-4 py-2 bg-blue-600 text-white font-bold rounded-xl text-sm hover:bg-blue-700 transition-colors shadow-sm\">\n            集計表示\n          </button>\n        </div>\n      </header>\n\n      <section class=\"stats-grid\">\n        <div class=\"stats-card\">\n          <p class=\"stats-label\">全体の平均出席率</p>\n          <p id=\"stat-avg-rate\" class=\"stats-value text-blue-600\">--%</p>\n        </div>\n        \n        <div class=\"stats-card\">\n          <p class=\"stats-label\">現在の在室状況</p>\n          <p id=\"stat-current-in\" class=\"stats-value\">\n            -- <span style=\"font-size: 1.125rem; color: #64748b; font-weight: 500;\">/ -- 名</span>\n          </p>\n        </div>\n      </section>\n\n      <div id=\"user-list\" class=\"user-list-container\"></div>\n    </div>\n\n    <?!= include(\"summary.js\") ?>\n    <?!= include(\"update_alert\") ?>\n  </body>\n</html>",
  "summary.js": "<script>\n  document.addEventListener('DOMContentLoaded', () => {\n    setDefaultDates(); // 1. デフォルト日付をセット\n    fetchData();       // 2. データ取得\n    document.getElementById('search-btn').addEventListener('click', fetchData);\n  });\n\n  // 当月初(1日)〜今日をセットする\n  function setDefaultDates() {\n    const now = new Date();\n    const currentYear = now.getFullYear();\n    const currentMonth = now.getMonth() + 1; // 1-12月\n\n    const formattedMonth = String(currentMonth).padStart(2, '0'); // 月を2桁（01, 02...）に整形\n    const startDate = `${currentYear}-${formattedMonth}-01`; // 当月の月初（01日）を作成\n    const endDate = now.toISOString().split('T')[0]; // 今日を yyyy-mm-dd形式で取得\n\n    document.querySelectorAll('.date-input')[0].value = startDate;\n    document.querySelectorAll('.date-input')[1].value = endDate;\n  }\n\n  // サーバーからデータを取得（期間を指定・直呼び出しに修正）\n  function fetchData() {\n    const listContainer = document.getElementById('user-list');\n    const start = document.querySelectorAll('.date-input')[0].value;\n    const end = document.querySelectorAll('.date-input')[1].value;\n\n    listContainer.innerHTML = '<p class=\"text-center py-10 text-slate-400\">期間内のデータを集計中...</p>';\n\n    // GAS側の引数の定義に沿って、start と end を直接渡す\n    google.script.run\n      .withSuccessHandler(renderSummary)\n      .withFailureHandler(err => {\n        listContainer.innerHTML = `<p class=\"text-red-500\">エラー: ${err.message}</p>`;\n      })\n      .getSummaryData(start, end);\n  }\n\n  // 取得したデータを画面に描画\n  function renderSummary(data) {\n    const listContainer = document.getElementById('user-list');\n\n    // 在室者⇒不在者の並びにソート\n    data.users.sort((a, b) => b.isCurrentlyIn - a.isCurrentlyIn);\n\n    // 統計カードの更新\n    document.getElementById('stat-avg-rate').textContent = `${data.stats.averageRate}%`;\n    \n    const currentInEl = document.getElementById('stat-current-in');\n    if (currentInEl) {\n      currentInEl.innerHTML = `\n        ${data.stats.currentlyIn} <span style=\"font-size: 1.125rem; color: #64748b; font-weight: 500;\">/ ${data.stats.totalActiveUsers} 名</span>\n      `;\n    }\n\n    // データがない場合のガード\n    if (!data.users || data.users.length === 0) {\n      listContainer.innerHTML = '<p class=\"text-center py-10 text-slate-400\">期間内のデータがありません。</p>';\n      return;\n    }\n\n    // リストの描画\n    listContainer.innerHTML = data.users.map(user => {\n      // 在室・不在のバッジを定義\n      const statusBadge = user.isCurrentlyIn \n        ? `<span class=\"status-badge badge-in\">在室</span>` \n        : `<span class=\"status-badge badge-out\">不在</span>`;\n\n      // HTMLを返す（テンプレートリテラル）\n      return `\n        <details class=\"group bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm transition-all duration-300 open:ring-2 open:ring-blue-500/20\">\n          <summary class=\"flex items-center justify-between p-5 cursor-pointer hover:bg-slate-50 transition-colors gap-4\">\n            <div class=\"min-w-0 flex items-center gap-3\"> \n              <div class=\"min-w-0\">\n                <div class=\"flex items-center gap-2\">\n                  <h3 class=\"font-bold text-slate-800 text-lg truncate\">${user.name}</h3>\n                  ${statusBadge}\n                </div>\n                <p class=\"text-sm text-slate-400 truncate\">${user.email}</p>\n              </div>\n            </div>\n            \n            <div class=\"flex items-center gap-8 flex-shrink-0\">\n              <div class=\"text-right\">\n                <p class=\"text-[10px] font-bold text-slate-400 uppercase tracking-widest\">出席率</p>\n                <p class=\"text-xl font-black ${user.rate < 80 ? 'text-orange-500' : 'text-emerald-500'}\">\n                  ${user.rate}<span class=\"text-sm ml-0.5\">%</span>\n                </p>\n              </div>\n              <div class=\"group-open-rotate transition-transform duration-300 text-slate-300\">\n                <svg xmlns=\"http://www.w3.org/2000/svg\" class=\"h-6 w-6\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"currentColor\">\n                  <path stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" />\n                </svg>\n              </div>\n            </div>\n          </summary>\n          <div class=\"px-5 pb-5 overflow-x-auto\">\n            <div class=\"bg-slate-50 rounded-xl p-4 border border-slate-100\">\n              <table class=\"w-full text-left text-sm\">\n                <thead>\n                  <tr class=\"text-slate-400 font-medium border-b border-slate-200\">\n                    <th class=\"py-3 px-2\">日付</th>\n                    <th class=\"py-3\">入室</th>\n                    <th class=\"py-3\">退室</th>\n                    <th class=\"py-3 text-right\">在室</th>\n                  </tr>\n                </thead>\n                <tbody class=\"text-slate-600 divide-y divide-slate-200\">\n                  ${user.logs.map(log => `\n                    <tr class=\"hover:bg-slate-100 transition-colors\">\n                      <td class=\"py-3 px-2 font-medium text-slate-700\">${log.date}</td>\n                      <td class=\"py-3\">${log.in}</td>\n                      <td class=\"py-3\">${log.out || '--:--'}</td>\n                      <td class=\"py-3 text-right font-mono font-medium\">${log.duration || '---'}</td>\n                    </tr>\n                  `).join('')}\n                </tbody>\n              </table>\n            </div>\n          </div>\n        </details>\n      `;\n    }).join('');\n  }\n</script>",
  "unknown_user": "<!DOCTYPE html>\n<html>\n  <div style=\"font-family: sans-serif; padding: 20px;\">\n    <p style=\"color: red; font-weight: bold;\">認証失敗: 未登録のユーザーか、Google Workspaceアカウントでログインしていません。</p>\n    <p>以下のボタンを押してユーザー登録済みのGoogle Workspaceアカウントを選択し直してください。</p>\n    <p>※ユーザー登録が済んでいない場合は管理者にお問い合わせ下さい。</p>\n    <a href=\"https://accounts.google.com/Logout?continue=https://accounts.google.com/ServiceLogin?continue=${ScriptApp.getService().getUrl()}\"\n      target=\"_top\" \n      style=\"display: inline-block; padding: 10px 20px; background: #4285f4; color: white; text-decoration: none; border-radius: 4px;\">\n      アカウントを切り替える\n    </a>\n  </div>\n</html>\n",
  "update_alert": "<div id=\"update-toast\" style=\"\n  display: none; \n  position: fixed; \n  bottom: 20px; \n  right: 20px; \n  background-color: #1e293b; \n  color: #ffffff; \n  padding: 16px 20px; \n  border-radius: 12px; \n  box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3); \n  z-index: 99999; \n  max-width: 300px; /* 案内文が少し長くなるため、幅を少し広げました */\n  font-family: sans-serif;\n  font-size: 13px;\n  line-height: 1.5;\n  border: 1px solid #334155;\n  animation: slideInToast 0.3s ease-out;\n\">\n  <div style=\"font-weight: bold; margin-bottom: 6px; color: #fbbf24; display: flex; align-items: center; gap: 6px;\">\n    📢 システムアップデートのお知らせ\n  </div>\n  <div id=\"update-toast-msg\" style=\"color: #cbd5e1; margin-bottom: 12px;\">\n    </div>\n  <div style=\"text-align: right;\">\n    <button onclick=\"document.getElementById('update-toast').style.display='none';\" style=\"\n      background-color: #475569; \n      color: #ffffff; \n      border: none; \n      padding: 6px 14px; \n      border-radius: 6px; \n      cursor: pointer;\n      font-weight: bold;\n      font-size: 12px;\n    \">\n      閉じる\n    </button>\n  </div>\n</div>\n\n<style>\n@keyframes slideInToast {\n  from { transform: translateY(100px); opacity: 0; }\n  to { transform: translateY(0); opacity: 1; }\n}\n</style>\n\n<script>\n  window.addEventListener('load', function() {\n    // アップデートの確認を行うコア関数\n    function checkSystemUpdate() {\n      google.script.run\n        .withSuccessHandler(function(result) {\n          if (result && result.hasUpdate) {\n            const toast = document.getElementById('update-toast');\n            const msg = document.getElementById('update-toast-msg');\n            \n            // トーストにメッセージを注入して表示\n            msg.innerText = `新しいバージョン (${result.latestVersion}) が利用可能です。\\n\\n管理者は、スプレッドシートの「アプリメニュー」から更新を実行するか、マニュアルに従って最新のGASコードを取得してデプロイをバージョンアップしてください。`;\n            \n            toast.style.display = 'block';\n          }\n        })\n        .withFailureHandler(function(error) {\n          console.error(\"アップデート確認エラー:\", error);\n        })\n        .checkUpdateForHtml();\n    }\n\n    // 1. 初回起動時：起動から「1.5秒後（1500ミリ秒）」に最初のチェックを走らせる\n    setTimeout(checkSystemUpdate, 1500); \n\n    // 2. 定期実行：その後は「24時間（86,400,000ミリ秒）」ごとに自動チェック\n    setInterval(checkSystemUpdate, 86400000); \n  });\n</script>"
}

    // GitHub Actionによりファイルからhtml文字列を自動生成してtemplatesに格納

    return HtmlService.createTemplate(templates[fileName]);
  }
}
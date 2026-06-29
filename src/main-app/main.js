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
    const templates = {}

    // GitHub Actionによりファイルからhtml文字列を自動生成してtemplatesに格納

    return HtmlService.createTemplate(templates[fileName]);
  }
}
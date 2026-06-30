/**
 * 打刻・集計関係
 */

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
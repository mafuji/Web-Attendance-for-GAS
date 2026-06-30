/**
 * ユーザー関係
 */

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

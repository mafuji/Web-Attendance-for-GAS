/**
 * アプリメンテナンス
 */

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
/**
 * パスワード管理
 */

// 最新のパスワードを返す（JSから呼び出す）
function getLatestPassword() {
  // 設定を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  // 変更：1列目（A列）の2行目から本体AppのID（app_sheet_id）を取得
  const appSsId = configSheet.getRange("A2").getValue();

  // 本体appの設定シートからパスワード取得
  const appSs = SpreadsheetApp.openById(appSsId);
  const appConfigSheet = appSs.getSheetByName("Config");
  const currentPwd = appConfigSheet.getRange("A2").getValue(); 
  
  return currentPwd;
}

// 本体appのパスワードを書き換える
function rotatePassword() {
  // 設定を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config"); 
  // 変更：1列目（A列）の2行目から本体AppのID（app_sheet_id）を取得
  const appSsId = configSheet.getRange("A2").getValue();

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

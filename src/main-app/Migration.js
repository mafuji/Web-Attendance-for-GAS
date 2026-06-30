/**
 * シート構造変更などのコード以外の更新スクリプト（Updater.gsから呼び出す）
 * @param {string} oldVersion - アップデート前のバージョン (例: "v1.0.0")
 * @param {string} newVersion - アップデート後のバージョン (例: "v2.0.0")
 */

// 実装例

function runMigration(oldVersion, newVersion) {
/*
  if (oldVersion === newVersion) {
    Logger.log(`バージョンに変更がないため、マイグレーションをスキップします (${oldVersion})`);
    return;
  }
  
  Logger.log(`[Migration] 移行処理を開始します: ${oldVersion} -> ${newVersion}`);
  
  // --- 引数のバージョンを元に、ピンポイントで移行ロジックを実行 ---
  if (oldVersion === 'v1.0.0' && newVersion === 'v2.0.0') {
    _migrate_v1_to_v2();
  }
  
  if (oldVersion === 'v2.0.0' && newVersion === 'v3.0.0') {
    // _migrate_v2_to_v3(); // 将来の拡張用
  }

  // 念のため、適用完了したバージョンをスプレッドシートのスクリプトプロパティにも刻印しておく
  // （手動実行されたときなどの整合性チェック用）
  PropertiesService.getScriptProperties().setProperty('INSTALLED_VERSION', newVersion);
  Logger.log(`[Migration] 移行処理が正常に終了しました。`);
*/
}

/*
// v1.0.0 -> v2.0.0 へのシート構造変更
function _migrate_v1_to_v2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // ここに具体的なシート追加や列挿入のロジックを書く
  SpreadsheetApp.flush();
}
*/
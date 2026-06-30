/**
 * アプリ固有メニュー
 */
/**
 * ============================================================================
 * 【約束事1】アプリ固有のカスタムメニュー構成を返す
 * ============================================================================
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
 * メニューからASN自動更新をONにする
 */
function main_createAsnTriggerOnly() {
  createAsnTriggerOnly();

  const ui = SpreadsheetApp.getUi();
  ui.alert('定期処理の開始', '🟢 ASNの自動更新を開始しました。\n今後、毎日深夜1時〜2時の間に最新のCIDRへ自動洗い替えされます。', ui.ButtonSet.OK);
}

/**
 * メニューからASN自動更新をOFFにする
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
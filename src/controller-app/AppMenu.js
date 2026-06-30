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
    { 
      type: "item",
      name: "🟢 パスワード自動更新をONにする", 
      functionName: "main_createRotatePwdTriggerOnly" 
    },
    {
      type: "item", 
      name: "🛑 パスワード自動更新をOFFにする", 
      functionName: "main_deleteRotatePwdTriggerOnly" 
    }
  ];
}

function main_createRotatePwdTriggerOnly() {
  createRotatePwdTriggerOnly();

  const ui = SpreadsheetApp.getUi()
  ui.alert('定期処理の開始', '🟢 パスワードの自動更新を開始しました。\n今後1分おきに自動実行されます。', ui.ButtonSet.OK);
}

function main_deleteRotatePwdTriggerOnly() {
  deleteRotatePwdTriggerOnly();

  const ui = SpreadsheetApp.getUi()
  ui.alert('定期処理の停止', '🛑 パスワードの自動更新を停止しました。', ui.ButtonSet.OK);
}
/**
 * アプリ固有トリガー
 */

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
 * トリガー操作
 * ============================================================================
 */
/**
 * refreshCidrFromAsn トリガーを直接新規作成する
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
 * refreshCidrFromAsnだけをピンポイントで削除する
 */
function deleteAsnTriggerOnly() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshCidrFromAsn') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
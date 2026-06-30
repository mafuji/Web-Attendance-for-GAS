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
      // 実行したい関数名
      functionName: "rotatePassword", 
      // GASの本物のメソッド名と引数をそのまま配列で指定！
      methods: [
        { name: "everyMinutes", args: [1] }
      ]
    }
  ];
}

/**
 * rotatePasswordトリガーを作成する
 */
function createRotatePwdTriggerOnly() {
  deleteRotatePwdTriggerOnly();

  ScriptApp.newTrigger('rotatePassword')
    .timeBased()
    .everyMinutes(1)
    .create();
}

/**
 * rotatePasswordだけをピンポイントで削除する
 */
function deleteRotatePwdTriggerOnly() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rotatePassword') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
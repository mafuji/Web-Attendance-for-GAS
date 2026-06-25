// ==========================================
// 共通設定（GitHub情報）
// ==========================================
const GITHUB_USER = "mafuji";
const GITHUB_REPO = "Web-Attendance-for-GAS";
const CURRENT_VERSION = "v1.0.0"; // アップデート後にここも自動で書き換わります

// ==========================================
// アプリごとの固有設定
// ==========================================
const TARGET_APP_DIR = "controller-app"; // アプリごとに書き換えて複製する
const TARGET_FILE = "merged.js";

/**
 * 共通トースト通知ヘルパー
 */
function showToast(message, title = '⚙️ システム') {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  } catch(e) {
    Logger.log(`[Toast] ${title}: ${message}`);
  }
}

/**
 * アップデートチェック＆実行のメイン関数
 */
function checkAndExecuteUpdate() {
  const url = getGitHubApiUrl();
  const token = PropertiesService.getScriptProperties().getProperty("GH_TOKEN");
  
  const response = UrlFetchApp.fetch(url, { 
    "muteHttpExceptions": true,
    "headers": token ? { "Authorization": "token " + token } : {}
  });
  
  if (response.getResponseCode() !== 200) {
    Logger.log(`GitHubからの情報取得に失敗しました。ステータス: ${response.getResponseCode()}`);
    Logger.log(`詳細: ${response.getContentText()}`);
    return;
  }
  
  const resData = JSON.parse(response.getContentText());
  let latestRelease;
  
  if (Array.isArray(resData)) {
    if (resData.length === 0) return;
    latestRelease = resData[0]; // 開発モード：最新のプレリリース
  } else {
    latestRelease = resData; // 本番モード：最新の正式リリース
  }
  
  const latestVersion = latestRelease.tag_name;
  Logger.log(`現在のバージョン: ${CURRENT_VERSION} / 最新バージョン: ${latestVersion}`);
  
  if (latestVersion !== CURRENT_VERSION) {
    Logger.log("新バージョンを検知しました。アップデート処理を開始します...");
    
    const rawCode = fetchFromGitHub(latestVersion, TARGET_FILE);
    if (!rawCode) return;

    const rawManifest = fetchFromGitHub(latestVersion, "appsscript.json");
    if (!rawManifest) return;
    
    const success = updateProjectFiles(rawCode, rawManifest, latestVersion);
    
    if (success) {
      Logger.log(`🎉 バージョン ${latestVersion} へのアップデートが正常に完了しました！`);
    }
  } else {
    Logger.log("すでに最新の状態です。");
  }
}

/**
 * 環境に応じたGitHub APIのURLを取得
 */
function getGitHubApiUrl() {
  const envMode = PropertiesService.getScriptProperties().getProperty("ENV_MODE");
  
  if (envMode === "development") {
    Logger.log("--- [開発モード] プレリリースを含む最新を取得します ---");
    return `https://api.github.com/repos/${GITHUB_USER}/${GITHUB_REPO}/releases`;
  } else {
    Logger.log("--- [本番モード] 正式リリースのみを取得します ---");
    return `https://api.github.com/repos/${GITHUB_USER}/${GITHUB_REPO}/releases/latest`;
  }
}

/**
 * 環境に応じて、GitHubまたはjsDelivrから各種アセットファイルを安全にダウンロードする
 */
function fetchFromGitHub(version, fileName) {
  const envMode = PropertiesService.getScriptProperties().getProperty("ENV_MODE");
  let downloadUrl = "";
  let options = { "muteHttpExceptions": true };

  const remotePath = `prd/${TARGET_APP_DIR}/${fileName}`;

  if (envMode === "development") {
    downloadUrl = `https://raw.githubusercontent.com/${GITHUB_USER}/${GITHUB_REPO}/${version}/${remotePath}`;
    const token = PropertiesService.getScriptProperties().getProperty("GH_TOKEN");
    if (token) {
      options["headers"] = { "Authorization": "token " + token };
    }
  } else {
    downloadUrl = `https://cdn.jsdelivr.net/gh/${GITHUB_USER}/${GITHUB_REPO}@${version}/${remotePath}`;
  }

  Logger.log(`[DL開始] 接続先: ${downloadUrl}`);
  const response = UrlFetchApp.fetch(downloadUrl, options);
  
  if (response.getResponseCode() !== 200) {
    Logger.log(`[Error] ${fileName} のダウンロードに失敗しました。ステータス: ${response.getResponseCode()}`);
    return null;
  }
  return response.getContentText();
}

/**
 * Apps Script API を叩いて、プロジェクト内のファイルを一括書き換え
 */
function updateProjectFiles(newCode, newManifest, latestVersion) {
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  
  const getUrl = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
  const getResponse = UrlFetchApp.fetch(getUrl, {
    method: "get",
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });
  
  if (getResponse.getResponseCode() !== 200) {
    Logger.log("現在のプロジェクト構成の取得に失敗しました。");
    return false;
  }
  
  const projectContent = JSON.parse(getResponse.getContentText());
  let mergedFileExists = false;
  
  projectContent.files = projectContent.files.map(file => {
    if (file.name === "merged") {
      file.source = newCode;
      mergedFileExists = true;
    }
    
    if (file.name === "updater") {
      file.source = file.source.replace(
        /const CURRENT_VERSION = ".*?";/,
        `const CURRENT_VERSION = "${latestVersion}";`
      );
    }
    
    if (file.name === "appsscript") {
      try {
        JSON.parse(newManifest); 
        file.source = newManifest;
        Logger.log("--- [マニフェスト同期] GitHub上のappsscript.jsonをそのまま適用しました ---");
      } catch (e) {
        Logger.log("⚠️ GitHubから取得したマニフェストファイルが不正なJSON形式のため、同期をスキップしました: " + e.toString());
      }
    }
    return file;
  });
  
  if (!mergedFileExists) {
    projectContent.files.push({
      name: "merged",
      type: "SERVER_JS",
      source: newCode
    });
  }
  
  const putUrl = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
  const putResponse = UrlFetchApp.fetch(putUrl, {
    method: "put",
    headers: { "Authorization": "Bearer " + token },
    contentType: "application/json",
    payload: JSON.stringify(projectContent),
    muteHttpExceptions: true
  });
  
  if (putResponse.getResponseCode() !== 200) {
    Logger.log("Apps Script APIへのコード書き込みに失敗しました: " + putResponse.getContentText());
    return false;
  }
  
  Logger.log("--- コードおよびマニフェストの書き換えが成功しました。続いて自動デプロイを開始します ---");

  // デプロイとトリガーの作成・更新へ進む
  createDeployAndTrigger();
  
  return true;
}

/**
 * デプロイとトリガー作成（および既存デプロイの更新）
 */
function createDeployAndTrigger(){
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  const ui = SpreadsheetApp.getUi(); // スコープバグ修正
  
  try {
    showToast('最新のコードでWebアプリを公開中...', '⚙️ セットアップ');
    
    // 1. 新しい「版（バージョン）」を必ず作成する
    const versionUrl = `https://script.googleapis.com/v1/projects/${scriptId}/versions`;
    const createVersion = UrlFetchApp.fetch(versionUrl, {
      method: "post",
      headers: { "Authorization": "Bearer " + token },
      contentType: "application/json",
      payload: JSON.stringify({ "description": `Deploy via Update System (${CURRENT_VERSION})` }),
      muteHttpExceptions: true
    });
    
    let vNum = 1;
    if (createVersion.getResponseCode() === 200) {
      vNum = JSON.parse(createVersion.getContentText()).versionNumber;
      Logger.log(`🟢 新しい版 (版 ${vNum}) を作成しました。`);
    } else {
      throw new Error("新しい版の作成に失敗しました: " + createVersion.getContentText());
    }
    
    // 既存の本物デプロイがあるかチェック
    const deployUrl = `https://script.googleapis.com/v1/projects/${scriptId}/deployments`;
    const getDeploy = UrlFetchApp.fetch(deployUrl, { method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
    
    let existingDeployId = null;
    if (getDeploy.getResponseCode() === 200) {
      const deployData = JSON.parse(getDeploy.getContentText());
      if (deployData.deployments) {
        const realDeploy = deployData.deployments.find(d => 
          d.deploymentConfig && 
          d.deploymentConfig.manifestFileName === "appsscript" && 
          d.updateTime !== "1970-01-01T00:00:00Z"
        );
        if (realDeploy) existingDeployId = realDeploy.deploymentId;
      }
    }
    
    let deployedData = null;
    
    if (existingDeployId) {
      // 【仕様改善】既存デプロイがある場合は、新しい版(vNum)で既存デプロイを更新(PUT)する
      const updateDeployUrl = `${deployUrl}/${existingDeployId}`;
      const updateResponse = UrlFetchApp.fetch(updateDeployUrl, {
        method: "put",
        headers: { "Authorization": "Bearer " + token },
        contentType: "application/json",
        payload: JSON.stringify({
          "deploymentConfig": {
            "versionNumber": vNum,
            "manifestFileName": "appsscript",
            "description": `Updated to version ${vNum}`
          }
        }),
        muteHttpExceptions: true
      });
      
      if (updateResponse.getResponseCode() === 200) {
        deployedData = JSON.parse(updateResponse.getContentText());
        Logger.log(`🟢 既存のデプロイ [${existingDeployId}] を 新しい版 ${vNum} に更新しました。`);
      } else {
        throw new Error("デプロイの更新に失敗しました: " + updateResponse.getContentText());
      }
    } else {
      // 無ければ新規デプロイ(POST)
      const deployResponse = UrlFetchApp.fetch(deployUrl, {
        method: "post",
        headers: { "Authorization": "Bearer " + token },
        contentType: "application/json",
        payload: JSON.stringify({
          "versionNumber": vNum,
          "manifestFileName": "appsscript",
          "description": "Initial production deploy"
        }),
        muteHttpExceptions: true
      });
      
      if (deployResponse.getResponseCode() === 200) {
        deployedData = JSON.parse(deployResponse.getContentText());
        Logger.log("🟢 初代デプロイを樹立しました。");
      } else {
        throw new Error("デプロイの作成に失敗しました: " + deployResponse.getContentText());
      }
    }

    // ----------------------------------------------------
    // タスク2: 全トリガーの一括設置
    // ----------------------------------------------------
    showToast('自動更新およびアプリのトリガーを設置中...', '⚙️ セットアップ');
    const allTriggers = ScriptApp.getProjectTriggers();

    // ① updater自身の「毎日深夜3時」の自動更新トリガー
    const hasUpdateTrigger = allTriggers.some(t => t.getHandlerFunction() === 'checkAndExecuteUpdate');
    if (!hasUpdateTrigger) {
      ScriptApp.newTrigger('checkAndExecuteUpdate')
        .timeBased()
        .everyDays(1)
        .atHour(3)
        .create();
      Logger.log("🟢 自動更新トリガー（毎日深夜3時）を設置しました。");
    }

    // ② merged.gs側からアプリ固有のトリガー設定を動的作成
    if (typeof getAppTriggerConfig === 'function') {
      const appTriggers = getAppTriggerConfig();
      
      if (Array.isArray(appTriggers)) {
        appTriggers.forEach(config => {
          const alreadyExists = allTriggers.some(t => t.getHandlerFunction() === config.functionName);
          
          if (!alreadyExists && config.functionName && Array.isArray(config.methods)) {
            let builder = ScriptApp.newTrigger(config.functionName).timeBased();
            
            config.methods.forEach(method => {
              if (typeof builder[method.name] === 'function') {
                // メソッドチェーンの戻り値を常に安全に代入
                builder = builder[method.name].apply(builder, method.args || []);
              }
            });
            
            builder.create();
            Logger.log(`🟢 アプリトリガー [${config.functionName}] を設置しました。`);
          }
        });
      }
    }

    // ----------------------------------------------------
    // 結果発表
    // ----------------------------------------------------
    let successMessage = '🎉 アプリの公開とすべてのトリガー設定が完了しました！\n\n';
    if (deployedData && deployedData.entryPoints && deployedData.entryPoints.length > 0) {
      successMessage += `🔗 生成された公開URL:\n${deployedData.entryPoints[0].webApp.url}`;
    } else if (existingDeployId) {
      // 更新時はURLがすでにわかっている、または既存デプロイ一覧から再取得可能
      successMessage += `🔗 公開URL（既存）が最新コードに更新されました。\nシステム管理メニューの「現在の公開状況を確認する」からURLを取得できます。`;
    } else {
      successMessage += '（※すでに公開済みのWebアプリURLが維持されています）';
    }
    
    ui.alert('セットアップ完了', successMessage, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('❌ エラー発生', '処理中にエラーが発生しました:\n' + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * スプレッドシートが開かれたときに自動実行される関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('アプリメニュー');

  try {
    if (typeof getAppMenuConfig === 'function') {
      const appMenus = getAppMenuConfig();
      if (Array.isArray(appMenus)) {
        appMenus.forEach(item => {
          if (item.type === "item" && item.name && item.functionName) {
            menu.addItem(item.name, item.functionName);
          } else if (item.type === "separator") {
            menu.addSeparator();
          }
        });
      }
    }
  } catch (e) {
    Logger.log("⚠️ アプリ固有メニューの取得に失敗しました: " + e.toString());
  }

  menu.addSeparator();

  menu.addSubMenu(
    ui.createMenu('🛠️ システム管理')
      .addItem('🚀 アプリを公開する', 'menu_setupInitialDeploymentAndTriggers')
      .addItem('📄 現在の公開状況を確認する', 'menu_checkCurrentDeploymentStatus')
      .addSeparator()
      .addItem('🔄 アプリを更新する', 'menu_forceExecuteUpdate')
      .addItem('▶️ アプリの自動更新を開始する', 'menu_startAutoUpdateTrigger')
      .addItem('⏸️ アプリの自動更新を一時停止する', 'menu_stopAutoUpdateTrigger')
      .addSeparator()
      .addItem('🛑 アプリの公開を停止する', 'menu_terminateSystem')
  );

  menu.addToUi();
}

/**
 * 1. 「🚀 アプリを公開する」の実装
 */
function menu_setupInitialDeploymentAndTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'アプリの公開（初回セットアップ）',
    'Webアプリを公開し、自動更新を含むすべての定期実行トリガーを設置します。よろしいですか？\n（※初回実行時のみ、Googleによるアクセス権限の承認が必要です）',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  createDeployAndTrigger();
}

/**
 * 2. 「🔄 修正パッチを適用する」の実装
 */
function menu_forceExecuteUpdate() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '修正パッチの適用（手動アップデート）',
    'GitHubから最新のプログラムを今すぐダウンロードし、現在のURLを維持したままアプリを更新します。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  showToast('最新パッチを適用中（GitHub通信＆デプロイを実行）...', '⚙️ パッチ適用');

  try {
    checkAndExecuteUpdate();
    ui.alert('アップデート完了', '🎉 最新の修正パッチの適用に成功しました！\n\nWebアプリのURLは変わっていません。画面をリロードしてご確認ください。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ パッチ適用失敗', 'エラーが発生しました:\n' + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 3. 「🛑 アプリの公開を停止する」の実装
 */
function menu_terminateSystem() {
  const ui = SpreadsheetApp.getUi();
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  
  const res1 = ui.alert(
    '⚠️ 警告：アプリの公開停止',
    '現在公開中のWebアプリURLを無効化し、自動更新を含むすべての定期実行トリガーを削除します。\n本当に実行してもよろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (res1 !== ui.Button.YES) return;

  const res2 = ui.alert(
    '最終確認',
    'この処理を実行すると、利用中のユーザーはアプリへアクセスできなくなります。本当の後戻りはできませんが、よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (res2 !== ui.Button.YES) return;

  showToast('すべての定期実行トリガーを削除中...', '🛑 公開停止処理');
  
  try {
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`🗑️ トリガーを削除しました: ${trigger.getHandlerFunction()}`);
    });

    showToast('公開中のWebアプリURLを無効化中...', '🛑 公開停止処理');
    const deployUrl = `https://script.googleapis.com/v1/projects/${scriptId}/deployments`;
    
    const getDeployResponse = UrlFetchApp.fetch(deployUrl, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    if (getDeployResponse.getResponseCode() === 200) {
      const deployData = JSON.parse(getDeployResponse.getContentText());
      
      if (deployData.deployments && deployData.deployments.length > 0) {
        const trueDeployments = deployData.deployments.filter(d => d.updateTime !== "1970-01-01T00:00:00Z");
        
        if (trueDeployments.length > 0) {
          trueDeployments.forEach(d => {
            const deleteUrl = `${deployUrl}/${d.deploymentId}`;
            const deleteResponse = UrlFetchApp.fetch(deleteUrl, {
              method: "delete",
              headers: { "Authorization": "Bearer " + token },
              muteHttpExceptions: true
            });
            
            if (deleteResponse.getResponseCode() === 200) {
              Logger.log(`🗑️ 既存のデプロイを削除しました: ${d.deploymentId}`);
            } else {
              Logger.log(`❌ デプロイ [${d.deploymentId}] の削除に失敗: ` + deleteResponse.getContentText());
            }
          });
        }
      }
    } else {
      throw new Error("デプロイ一覧の取得に失敗しました: " + getDeployResponse.getContentText());
    }

    ui.alert(
      '公開停止完了',
      '🛑 アプリの公開停止および後片付けがすべて完了しました！\n\n・WebアプリURLは完全に無効化されました。\n・すべての定期トリガーが消去されました。',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ 処理失敗', 'エラーが発生しました:\n' + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 4. 「📄 現在の公開状況を確認する」の実装
 */
function menu_checkCurrentDeploymentStatus() {
  const ui = SpreadsheetApp.getUi();
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  
  showToast('現在のデプロイおよびトリガーの状態を調査中...', '🔎 状況確認');

  try {
    const deployUrl = `https://script.googleapis.com/v1/projects/${scriptId}/deployments`;
    const getDeployResponse = UrlFetchApp.fetch(deployUrl, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    let deployInfoText = "🔴 公開されていません（アクティブなWebアプリはありません）";
    let webAppUrl = "";

    if (getDeployResponse.getResponseCode() === 200) {
      const deployData = JSON.parse(getDeployResponse.getContentText());
      
      if (deployData.deployments && deployData.deployments.length > 0) {
        const trueDeploy = deployData.deployments.find(d => 
          d.deploymentConfig && 
          d.deploymentConfig.manifestFileName === "appsscript" && 
          d.updateTime !== "1970-01-01T00:00:00Z"
        );
        
        if (trueDeploy) {
          const config = trueDeploy.deploymentConfig;
          const updateTime = new Date(trueDeploy.updateTime).toLocaleString('ja-JP');
          
          deployInfoText = `🟢 公開中（アクティブ）\n` +
                           `・デプロイID: ${trueDeploy.deploymentId}\n` +
                           `・適用中の版: 版 ${config.versionNumber}\n` +
                           `・最終更新日時: ${updateTime}\n` +
                           `・メモ: ${config.description || "なし"}`;
          
          if (trueDeploy.entryPoints && trueDeploy.entryPoints.length > 0) {
            webAppUrl = trueDeploy.entryPoints[0].webApp.url;
          }
        }
      }
    } else {
      throw new Error("デプロイ情報の取得に失敗しました。");
    }

    const allTriggers = ScriptApp.getProjectTriggers();
    let triggerInfoText = "";
    
    const hasUpdateTrigger = allTriggers.some(t => t.getHandlerFunction() === 'checkAndExecuteUpdate');
    triggerInfoText += hasUpdateTrigger ? "・🔄 自動更新システム: 🟢 稼働中（毎日深夜）\n" : "・🔄 自動更新システム: 🛑 停止中\n";
    
    if (typeof getAppTriggerConfig === 'function') {
      const appTriggers = getAppTriggerConfig();
      if (Array.isArray(appTriggers)) {
        appTriggers.forEach(config => {
          const isRunning = allTriggers.some(t => t.getHandlerFunction() === config.functionName);
          triggerInfoText += `・📅 アプリ機能 [${config.functionName}]: ${isRunning ? "🟢 稼働中" : "🛑 停止中"}\n`;
        });
      }
    }

    let message = `【Webアプリの公開ステータス】\n${deployInfoText}\n\n` +
                  `【定期実行トリガーの稼働状況】\n${triggerInfoText}`;
                  
    if (webAppUrl) {
      message += `\n【公開URL】（※コピーして利用してください）\n${webAppUrl}`;
    }

    ui.alert('📊 現在の公開・稼働状況', message, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('❌ 状況確認失敗', 'エラーが発生しました:\n' + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * システム管理：毎晩の自動更新トリガーを設置（開始）する
 */
function menu_startAutoUpdateTrigger() {
  const ui = SpreadsheetApp.getUi();
  const allTriggers = ScriptApp.getProjectTriggers();
  
  // 重複チェック
  const hasUpdateTrigger = allTriggers.some(t => t.getHandlerFunction() === 'checkAndExecuteUpdate');
  
  if (hasUpdateTrigger) {
    ui.alert('確認', '🔄 毎晩の自動更新トリガーは、すでに「稼働中」です。', ui.ButtonSet.OK);
    return;
  }
  
  // トリガー新規作成
  ScriptApp.newTrigger('checkAndExecuteUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(3) // 深夜3時〜4時の間に発火
    .create();
    
  showToast('自動更新トリガーを設置しました。', '▶️ 稼働開始');
  ui.alert('稼働開始', '🟢 毎晩深夜3時の自動更新システムを有効化しました。', ui.ButtonSet.OK);
}

/**
 * システム管理：毎晩の自動更新トリガーを削除（一時停止）する
 */
function menu_stopAutoUpdateTrigger() {
  const ui = SpreadsheetApp.getUi();
  const allTriggers = ScriptApp.getProjectTriggers();
  let deleted = false;
  
  // checkAndExecuteUpdate のトリガーだけを狙い撃ちして削除
  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkAndExecuteUpdate') {
      ScriptApp.deleteTrigger(trigger);
      deleted = true;
    }
  });
  
  if (deleted) {
    showToast('自動更新トリガーを削除しました。', '⏸️ 一時停止');
    ui.alert('一時停止完了', '⏸️ 毎晩の自動更新を一時停止しました。\n（※ユーザーのアプリ利用や、手動でのアップデート機能はそのまま使えます）', ui.ButtonSet.OK);
  } else {
    ui.alert('確認', '🛑 自動更新トリガーはすでに停止しているか、設置されていません。', ui.ButtonSet.OK);
  }
}

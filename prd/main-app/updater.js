// ==========================================
// 共通設定（GitHub情報）
// ==========================================
const GITHUB_USER = "mafuji";
const GITHUB_REPO = "Web-Attendance-for-GAS";
const CURRENT_VERSION = "v1.0.0"; // アップデート後にここも自動で書き換わります

// ==========================================
// アプリごとの固有設定
// ==========================================
const TARGET_APP_DIR = "main-app"; // アプリごとに書き換えて複製する
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
 * バージョンチェックを行い、結果をUIに通知する（手動確認用）
 */
function menu_checkVersionAndNotify() {
  const ui = SpreadsheetApp.getUi();
  showToast('GitHubから最新のバージョン情報を取得中...', '🔎 バージョン確認');

  // GitHubのversion.jsonから最新情報を取得
  const url = `https://raw.githubusercontent.com/${GITHUB_USER}/${GITHUB_REPO}/main/src/shared/version.json`;
  const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
  
  if (response.getResponseCode() !== 200) {
    ui.alert('❌ 確認失敗', `バージョン情報の取得に失敗しました。\nステータス: ${response.getResponseCode()}`, ui.ButtonSet.OK);
    return;
  }
  
  const resData = JSON.parse(response.getContentText());
  const latestVersion = resData.version;
  const updateNotes = resData.description || "新機能の追加およびバグ修正"; // version.jsonに更新内容があれば表示

  if (latestVersion !== CURRENT_VERSION) {
    // 新バージョンがある場合のUI通知
    const alertMessage = `📢 新しいバージョンが利用可能です！\n\n` +
                         `現在のバージョン: ${CURRENT_VERSION}\n` +
                         `最新のバージョン: ${latestVersion}\n\n` +
                         `【更新内容】\n${updateNotes}\n\n` +
                         `今すぐアプリを最新版にアップデートしますか？\n` +
                         `（※現在のWebアプリURLを維持したまま更新されます）`;
    
    const choice = ui.alert('🔄 アップデート通知', alertMessage, ui.ButtonSet.YES_NO);
    
    if (choice === ui.Button.YES) {
      executeDirectUpdate(latestVersion);
    }
  } else {
    // すでに最新の場合のUI通知
    ui.alert('✅ 最新状態です', `ご利用中のバージョンは最新です。\n\n現在のバージョン: ${CURRENT_VERSION}`, ui.ButtonSet.OK);
  }
}

/**
 * ダイアログから直接呼び出されるアップデート実行関数
 */
function executeDirectUpdate(latestVersion) {
  const ui = SpreadsheetApp.getUi();
  showToast('最新プログラムをダウンロード中...', '⚙️ アプリ更新');

  try {
    // コードとマニフェストを取得
    const rawCode = fetchFromGitHub(latestVersion, TARGET_FILE);
    if (!rawCode) throw new Error(`${TARGET_FILE} の取得に失敗しました。`);

    const rawManifest = fetchFromGitHub(latestVersion, "appsscript.json");
    if (!rawManifest) throw new Error("appsscript.json の取得に失敗しました。");
    
    const success = updateProjectFiles(rawCode, rawManifest, latestVersion);
    
    if (success) {
      ui.alert('🎉 アップデート完了', `バージョン ${latestVersion} への更新が正常に完了しました！\n\nWebアプリのURLは変わっていません。`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ 更新失敗', 'アップデート処理中にエラーが発生しました:\n' + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 指定されたバージョンのアセットファイルをGitHubのRawから直接安全にダウンロードする
 */
function fetchFromGitHub(version, fileName) {
  const remotePath = `prd/${TARGET_APP_DIR}/${fileName}`;
  const downloadUrl = `https://raw.githubusercontent.com/${GITHUB_USER}/${GITHUB_REPO}/${version}/${remotePath}`;

  Logger.log(`[DL開始] 接続先: ${downloadUrl}`);
  const response = UrlFetchApp.fetch(downloadUrl, { "muteHttpExceptions": true });
  
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
    throw new Error("現在のプロジェクト構成の取得に失敗しました。");
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
    throw new Error("Apps Script APIへのコード書き込みに失敗しました: " + putResponse.getContentText());
  }
  
  Logger.log("--- コードおよびマニフェストの書き換えが成功しました。続いて自動デプロイを開始します ---");

  // デプロイとアプリ用トリガーの作成・更新へ進む
  createDeployAndTrigger();
  
  return true;
}

/**
 * デプロイとトリガー作成（メインコントロール関数）
 */
function createDeployAndTrigger() {
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  const ui = SpreadsheetApp.getUi(); 
  
  try {
    // 1. デプロイ作業の実行
    showToast('最新のコードでWebアプリを公開中...', '⚙️ セットアップ');
    const deployedData = executeDeployment(scriptId, token);

    // 2. アプリ固有トリガーの作成作業を実行
    showToast('アプリ固有のトリガーを設置中...', '⚙️ セットアップ');
    setupProjectTriggers();

    // 結果発表
    let successMessage = '🎉 アプリの公開とトリガー設定が完了しました！\n\n';
    if (deployedData && deployedData.entryPoints && deployedData.entryPoints.length > 0) {
      successMessage += `🔗 生成された公開URL:\n${deployedData.entryPoints[0].webApp.url}`;
    } else {
      successMessage += `🔗 公開URLが最新コードに更新されました。\nシステム管理メニューの「現在の公開状況を確認する」からURLを取得できます。`;
    }
    
    ui.alert('セットアップ完了', successMessage, ui.ButtonSet.OK);

  } catch (err) {
    ui.alert('❌ エラー発生', '処理中にエラーが発生しました:\n' + err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 【機能1】デプロイ作業（新しい版の作成、およびデプロイの作成・更新）
 */
function executeDeployment(scriptId, token) {
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
      if (!deployedData.deploymentId) deployedData.deploymentId = existingDeployId;
      Logger.log(`🟢 既存のデプロイ [${existingDeployId}] を 新しい版 ${vNum} に更新しました。`);
    } else {
      throw new Error("デプロイの更新に失敗しました: " + updateResponse.getContentText());
    }
  } else {
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

  return deployedData;
}

/**
 * 【機能2】トリガー作成作業（アプリ固有トリガーのみ）
 */
function setupProjectTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();

  // merged.gs側からアプリ固有のトリガー設定を動的作成
  if (typeof getAppTriggerConfig === 'function') {
    const appTriggers = getAppTriggerConfig();
    
    if (Array.isArray(appTriggers)) {
      appTriggers.forEach(config => {
        const alreadyExists = allTriggers.some(t => t.getHandlerFunction() === config.functionName);
        
        if (!alreadyExists && config.functionName && Array.isArray(config.methods)) {
          let builder = ScriptApp.newTrigger(config.functionName).timeBased();
          
          config.methods.forEach(method => {
            if (typeof builder[method.name] === 'function') {
              builder = builder[method.name].apply(builder, method.args || []);
            }
          });
          
          builder.create();
          Logger.log(`🟢 アプリトリガー [${config.functionName}] を設置しました。`);
        }
      });
    }
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
      .addItem('🚀 アプリを公開する（初回）', 'menu_setupInitialDeploymentAndTriggers')
      .addItem('📄 現在の公開状況を確認する', 'menu_checkCurrentDeploymentStatus')
      .addSeparator()
      .addItem('🔎 バージョン情報を確認する', 'menu_checkVersionAndNotify')
      .addSeparator()
      .addItem('🛑 アプリの公開を停止する', 'menu_terminateSystem')
  );

  menu.addToUi();
}

/**
 * 1. 「🚀 アプリを公開する（初回）」の実装
 */
function menu_setupInitialDeploymentAndTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'アプリの公開（初回セットアップ）',
    'Webアプリを公開し、アプリに必要な定期実行トリガーを設置します。よろしいですか？\n（※初回実行時のみ、Googleによるアクセス権限の承認が必要です）',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  createDeployAndTrigger();
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
    '現在公開中のWebアプリURLを無効化し、すべての定期実行トリガーを削除します。\n本当に実行してもよろしいですか？',
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
 * HTML（フロントエンド）から呼び出され、アップデートの有無を判定する
 * @return {Object} 判定結果とバージョン情報
 */
function checkUpdateForHtml() {
  const url = `https://raw.githubusercontent.com/${GITHUB_USER}/${GITHUB_REPO}/main/src/shared/version.json`;
  
  try {
    const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    if (response.getResponseCode() !== 200) {
      return { hasUpdate: false, error: "バージョン取得失敗" };
    }
    
    const resData = JSON.parse(response.getContentText());
    const latestVersion = resData.version;
    const description = resData.description || "";
    
    // 現在のバージョンとGitHubの最新バージョンを比較
    return {
      hasUpdate: latestVersion !== CURRENT_VERSION,
      currentVersion: CURRENT_VERSION,
      latestVersion: latestVersion,
      description: description
    };
  } catch(e) {
    return { hasUpdate: false, error: e.toString() };
  }
}
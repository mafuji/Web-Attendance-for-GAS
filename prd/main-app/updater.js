// ==========================================
// 共通設定（GitHub情報）
// ==========================================
const GITHUB_USER = "mafuji";
const GITHUB_REPO = "Web-Attendance-for-GAS";
const CURRENT_VERSION = "v1.0.0-beta.3"; // アップデート後にここも自動で書き換わります

// ==========================================
// アプリごとの固有設定（※アプリに応じて書き換える）
// ==========================================
// 💡 prdフォルダ内の階層構造に合わせて、ディレクトリ名も含めて指定します
const TARGET_APP_DIR = "main-app"; // controller-app の場合はここを書き換える
const TARGET_FILE = "merged.js";

/**
 * アップデートチェック＆実行のメイン関数
 */
function checkAndExecuteUpdate() {
  const url = getGitHubApiUrl();
  const token = PropertiesService.getScriptProperties().getProperty("GH_TOKEN");
  
  // GitHub APIを叩く際は、トークンがあればヘッダーに乗せてIP制限（レートリミット）を回避
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
    
    // ----------------------------------------------------
    // 【改修点1】新しい統合ソースコード(.js)をダウンロード
    // ----------------------------------------------------
    const rawCode = fetchFromGitHub(latestVersion, TARGET_FILE);
    if (!rawCode) return;

    // ----------------------------------------------------
    // 【改修点2】新しい設定マニフェスト(.json)をダウンロード
    // ----------------------------------------------------
    const rawManifest = fetchFromGitHub(latestVersion, "appsscript.json");
    if (!rawManifest) return;
    
    // 3. 自身のコードと設定ファイルを新しいもので上書き
    const success = updateProjectFiles(rawCode, rawManifest, latestVersion);
    
    if (success) {
      Logger.log(`🎉 バージョン ${latestVersion} へのアップデートが正常に完了しました！`);
      Logger.log("※画面をリロードするか、一度スクリプトエディタを閉じて開き直してください。");
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
 * 【改修点3】環境に応じて、GitHubまたはjsDelivrから各種アセットファイル（JS/JSON）を安全にダウンロードする
 */
function fetchFromGitHub(version, fileName) {
  const envMode = PropertiesService.getScriptProperties().getProperty("ENV_MODE");
  let downloadUrl = "";
  let options = { "muteHttpExceptions": true };

  // 💡 prd/アプリ名/ファイル名 の形に新パスを組み立てる
  const remotePath = `prd/${TARGET_APP_DIR}/${fileName}`;

  if (envMode === "development") {
    // 【開発モード】キャッシュのないGitHub rawから直接取得（トークンを使用）
    downloadUrl = `https://raw.githubusercontent.com/${GITHUB_USER}/${GITHUB_REPO}/${version}/${remotePath}`;
    
    const token = PropertiesService.getScriptProperties().getProperty("GH_TOKEN");
    if (token) {
      options["headers"] = { "Authorization": "token " + token };
    }
  } else {
    // 【本番モード】回数制限なしの無敵のCDN（jsDelivr）から高速取得
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
 * 【改修点4】Apps Script API を叩いて、プロジェクト内のファイルを一括書き換え ＋ スマート自動デプロイ
 */
function updateProjectFiles(newCode, newManifest, latestVersion) {
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  const envMode = PropertiesService.getScriptProperties().getProperty("ENV_MODE");
  
  // ====================================================
  // 1. 現在のプロジェクト構成（ファイル一覧）を取得
  // ====================================================
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
  
  // ====================================================
  // 2. 各ファイルの書き換え処理（JSおよびappsscript.jsonの反映）
  // ====================================================
  projectContent.files = projectContent.files.map(file => {
    // ① merged.js の上書き
    if (file.name === "merged") {
      file.source = newCode;
      mergedFileExists = true;
    }
    
    // ② updater.js の CURRENT_VERSION の書き換え
    if (file.name === "updater") {
      file.source = file.source.replace(
        /const CURRENT_VERSION = ".*?";/,
        `const CURRENT_VERSION = "${latestVersion}";`
      );
    }
    
    // ③ appsscript.json (マニフェスト) のGitHub同期＆公開範囲調整
    if (file.name === "appsscript") {
      try {
        // 💡 GitHub側から引っ張ってきた最新の設定ファイル（newManifest）をベースに適用
        let manifestObj = JSON.parse(newManifest);
        
        if (manifestObj.webapp) {
          manifestObj.webapp.executeAs = "USER_DEPLOYING";
          manifestObj.webapp.access = "ANYONE";
        }
        file.source = JSON.stringify(manifestObj, null, 2);
      } catch (e) {
        Logger.log("マニフェストファイルの自動書き換え中にエラーが発生しました: " + e.toString());
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
  
  // ====================================================
  // 3. 変更した構成をAPIで一括プッシュ（保存）
  // ====================================================
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

  // ====================================================
  // 4〜7. デプロイ作成処理
  // ====================================================
  // (※元のコードと完全に同一のため、ここの中身の記述は省略します)
  
  return true;
}

/**
 * ============================================================================
 * 基盤システム（枠組み・インターフェース制御部）
 * ============================================================================
 */

/**
 * スプレッドシートが開かれたときに自動実行される関数（基盤側で一元管理）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 💡 まずはベースとなる「アプリメニュー」の看板を作成
  const menu = ui.createMenu('アプリメニュー');

  // ----------------------------------------------------
  // 1. 【メイン処理】merged.gs 側からメニューデータを吸い上げて合流
  // ----------------------------------------------------
  try {
    // merged.gs 側に、約束の関数「getAppMenuConfig」が存在するか確認
    if (typeof getAppMenuConfig === 'function') {
      const appMenus = getAppMenuConfig();
      
      // 配列で届いたメニューデータを安全に1つずつマッピング
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
    // 💡 万が一 merged 側が壊れていて吸い出しに失敗しても、ログに留めて処理を続行！
    Logger.log("⚠️ アプリ固有メニューの取得に失敗しました: " + e.toString());
  }

  // アプリメニューとシステム管理の境界線
  menu.addSeparator();

  // ----------------------------------------------------
  // 2. 【固定処理】updater 自身が提供する必須のシステム管理メニュー
  // ----------------------------------------------------
  menu.addSubMenu(
    ui.createMenu('🛠️ システム管理')
      .addItem('🚀 アプリを公開する', 'menu_setupInitialDeploymentAndTriggers')
      .addItem('📄 現在の公開状況を確認する', 'menu_checkCurrentDeploymentStatus') // 💡 追加！
      .addSeparator()
      .addItem('🔄 修正パッチを適用する', 'menu_forceExecuteUpdate')
      .addSeparator()
      .addItem('🛑 アプリの公開を停止する', 'menu_terminateSystem')
  );

  // 画面に一括反映
  menu.addToUi();
}

/**
 * ============================================================================
 * メニューボタンに対応するプレースホルダー関数（中身は後ほど実装）
 * ============================================================================
 */

/**
 * 1. 「🚀 アプリを公開する」の実装
 */
function menu_setupInitialDeploymentAndTriggers() {
  const ui = SpreadsheetApp.getUi();
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  
  const response = ui.alert(
    'アプリの公開（初回セットアップ）',
    'Webアプリを公開し、自動更新を含むすべての定期実行トリガーを設置します。よろしいですか？\n（※初回実行時のみ、Googleによるアクセス権限の承認が必要です）',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // 画面右下にローディング通知を表示
  const toast = (msg) => SpreadsheetApp.getActiveSpreadsheet().toast(msg, '⚙️ セットアップ');
  
  try {
    // ----------------------------------------------------
    // タスク1: 最新の「版」を作成して、初代デプロイを樹立する
    // ----------------------------------------------------
    toast('最新のコードでWebアプリを公開中...');
    
    // 既存の本物デプロイがあるかチェック（2回押しによる二重デプロイ防止）
    const deployUrl = `https://script.googleapis.com/v1/projects/${scriptId}/deployments`;
    const getDeploy = UrlFetchApp.fetch(deployUrl, { method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
    
    let hasRealDeploy = false;
    if (getDeploy.getResponseCode() === 200) {
      const deployData = JSON.parse(getDeploy.getContentText());
      if (deployData.deployments) {
        hasRealDeploy = deployData.deployments.some(d => 
          d.deploymentConfig && 
          d.deploymentConfig.manifestFileName === "appsscript" && 
          d.updateTime !== "1970-01-01T00:00:00Z"
        );
      }
    }
    
    let deployedData = null;
    
    if (hasRealDeploy) {
      Logger.log("すでに公開済みのデプロイが存在するため、新規作成はスキップします。");
    } else {
      // 履歴コミットにあたる「新バージョン(版)」を作成
      const versionUrl = `https://script.googleapis.com/v1/projects/${scriptId}/versions`;
      const createVersion = UrlFetchApp.fetch(versionUrl, {
        method: "post",
        headers: { "Authorization": "Bearer " + token },
        contentType: "application/json",
        payload: JSON.stringify({ "description": "Initial deploy via menu button" }),
        muteHttpExceptions: true
      });
      
      let vNum = 1;
      if (createVersion.getResponseCode() === 200) {
        vNum = JSON.parse(createVersion.getContentText()).versionNumber;
      }
      
      // 【仕様3】無ければ新規デプロイ(POST)
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
    // タスク2: 全トリガーの一括設置（重複チェック付き）
    // ----------------------------------------------------
    toast('自動更新およびアプリのトリガーを設置中...');
    const allTriggers = ScriptApp.getProjectTriggers();

    // ① updater自身の「毎日深夜3時」の自動更新トリガー
    const hasUpdateTrigger = allTriggers.some(t => t.getHandlerFunction() === 'checkAndExecuteUpdate');
    if (!hasUpdateTrigger) {
      ScriptApp.newTrigger('checkAndExecuteUpdate')
        .timeBased()
        .everyDays(1)
        .atHour(3) // 💡 深夜3時〜4時の間に実行
        .create();
      Logger.log("🟢 自動更新トリガー（毎日深夜3時）を設置しました。");
    }

    // ② merged.gs側からアプリ固有のトリガー設定を吸い上げて動的作成
    if (typeof getAppTriggerConfig === 'function') {
      const appTriggers = getAppTriggerConfig();
      
      if (Array.isArray(appTriggers)) {
        appTriggers.forEach(config => {
          // 重複チェック
          const alreadyExists = allTriggers.some(t => t.getHandlerFunction() === config.functionName);
          
          if (!alreadyExists && config.functionName && Array.isArray(config.methods)) {
            
            // 1. まずベースとなるタイムベースドトリガーの原型を作る
            let builder = ScriptApp.newTrigger(config.functionName).timeBased();
            
            // 2. merged側から指定されたメソッド（everyMinutes等）を動的に数珠つなぎで実行する
            config.methods.forEach(method => {
              if (typeof builder[method.name] === 'function') {
                // 💡 例： builder["everyMinutes"].apply(builder, [1]) と同義になり、GASのメソッドが動的に発火します
                builder = builder[method.name].apply(builder, method.args || []);
              }
            });
            
            // 3. 最後にトリガーを確定（create）
            builder.create();
            Logger.log(`🟢 アプリトリガー [${config.functionName}] を動的メソッドチェーンにより設置しました。`);
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
    } else {
      successMessage += '（※すでに公開済みのWebアプリURLが維持されています）';
    }
    
    ui.alert('セットアップ完了', successMessage, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('❌ エラー発生', '処理中にエラーが発生しました:\n' + e.toString(), ui.ButtonSet.OK);
  }
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
  
  // 画面右下に通知
  const toast = (msg) => SpreadsheetApp.getActiveSpreadsheet().toast(msg, '⚙️ パッチ適用');
  toast('最新パッチを適用中（GitHub通信＆デプロイを実行）...');

  try {
    // 💡 すでに完成している自動更新のコア関数をそのまま1行呼び出すだけ！
    // これにより、全く同じ通信・置換・デプロイロジックが安全に走ります。
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
  
  // 誤操作防止のため、2回確認を入れます
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

  const toast = (msg) => SpreadsheetApp.getActiveSpreadsheet().toast(msg, '🛑 公開停止処理');
  
  try {
    // ----------------------------------------------------
    // STEP 1. すべてのトリガーを完全に削除
    // ----------------------------------------------------
    toast('すべての定期実行トリガーを削除中...');
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`🗑️ トリガーを削除しました: ${trigger.getHandlerFunction()}`);
    });
    Logger.log("🟢 すべてのトリガーの削除が完了しました。");

    // ----------------------------------------------------
    // STEP 2. 本物のデプロイをすべてループ処理でアーカイブ（削除）
    // ----------------------------------------------------
    toast('公開中のWebアプリURLを無効化中...');
    const deployUrl = `https://script.googleapis.com/v1/projects/${scriptId}/deployments`;
    
    const getDeployResponse = UrlFetchApp.fetch(deployUrl, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    if (getDeployResponse.getResponseCode() === 200) {
      const deployData = JSON.parse(getDeployResponse.getContentText());
      
      if (deployData.deployments && deployData.deployments.length > 0) {
        // 1970年のテストデプロイ以外（＝本物のデプロイ）を抽出
        const trueDeployments = deployData.deployments.filter(d => d.updateTime !== "1970-01-01T00:00:00Z");
        
        if (trueDeployments.length > 0) {
          trueDeployments.forEach(d => {
            const deleteUrl = `${deployUrl}/${d.deploymentId}`;
            
            // 💡 該当のデプロイIDに対して DELETE リクエストを送信して消し去る！
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
        } else {
          Logger.log("削除対象となるアクティブな本物デプロイはありませんでした。");
        }
      }
    } else {
      throw new Error("デプロイ一覧の取得に失敗しました: " + getDeployResponse.getContentText());
    }

    // ----------------------------------------------------
    // 結果発表
    // ----------------------------------------------------
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
  
  const toast = (msg) => SpreadsheetApp.getActiveSpreadsheet().toast(msg, '🔎 状況確認');
  toast('現在のデプロイおよびトリガーの状態を調査中...');

  try {
    // ----------------------------------------------------
    // 1. デプロイ状況の取得
    // ----------------------------------------------------
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
        // 💡 1970年のテストデプロイ以外（本物）を一本釣り
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

    // ----------------------------------------------------
    // 2. トリガー稼働状況の取得
    // ----------------------------------------------------
    const allTriggers = ScriptApp.getProjectTriggers();
    let triggerInfoText = "";
    
    const hasUpdateTrigger = allTriggers.some(t => t.getHandlerFunction() === 'checkAndExecuteUpdate');
    triggerInfoText += hasUpdateTrigger ? "・🔄 自動更新システム: 🟢 稼働中（毎日深夜）\n" : "・🔄 自動更新システム: 🛑 停止中\n";
    
    // merged側からトリガー設定が取得できれば、それらも稼働しているかチェック
    if (typeof getAppTriggerConfig === 'function') {
      const appTriggers = getAppTriggerConfig();
      if (Array.isArray(appTriggers)) {
        appTriggers.forEach(config => {
          const isRunning = allTriggers.some(t => t.getHandlerFunction() === config.functionName);
          triggerInfoText += `・📅 アプリ機能 [${config.functionName}]: ${isRunning ? "🟢 稼働中" : "🛑 停止中"}\n`;
        });
      }
    }

    // ----------------------------------------------------
    // 3. 結果をダイアログで美しく表示
    // ----------------------------------------------------
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
/**
 * 設定（打刻パスワード・許可IP）
 */

// 設定オブジェクトを返す（AllowedIp + AllowedIpFromAsn 統合対応版）
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シート
  const configSheet = ss.getSheetByName('Config');
  const useIpControl = configSheet.getRange('C2').getValue();
  const requirePassword = configSheet.getRange('D2').getValue();

  let ipRules = [];

  // ① AllowedIpシートから許可IPリストを取得
  const allowedIpSheet = ss.getSheetByName('AllowedIp');
  if (allowedIpSheet) {
    const data = allowedIpSheet.getDataRange().getValues();
    const rules = data.slice(1).map(row => String(row[0]).trim()).filter(Boolean);
    ipRules = ipRules.concat(rules);
  }

  // ② AllowedIpFromAsnシートから許可IPリストを取得（★ここを追加）
  const allowedIpFromAsnSheet = ss.getSheetByName('AllowedIpFromAsn');
  if (allowedIpFromAsnSheet) {
    const dataFromAsn = allowedIpFromAsnSheet.getDataRange().getValues();
    const rulesFromAsn = dataFromAsn.slice(1).map(row => String(row[0]).trim()).filter(Boolean);
    ipRules = ipRules.concat(rulesFromAsn);
  }

  return {
    requirePassword: requirePassword,
    useIpControl: useIpControl,
    allowedIps: ipRules // 両方のシートのルールがマージされてここに入る
  }; 
}

// パスワード認証
function authorized(pwdInput) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const correctPwd = configSheet.getRange("A2").getValue();

  return String(pwdInput).trim() === String(correctPwd).trim();
}

// サーバー側のCIDR対応IPチェックロジック
function _checkIpOnServer(userIp) {
  const config = getConfig();
  if (!config.useIpControl) return true; // 設定でIPチェックがオフならスルー
  if (!userIp) return false; // IPが送られてきていなければ拒否

  const allowedRules = config.allowedIps;
  
  return allowedRules.some(rule => {
    if (rule === userIp) return true;

    if (rule.includes('/')) {
      const parts = rule.split('/');
      const range = parts[0];
      const bits = parseInt(parts[1], 10);
      
      const ipNum = _ipToLong(userIp);
      const rangeNum = _ipToLong(range);
      
      if (ipNum === null || rangeNum === null) return false;

      const mask = bits === 0 ? 0 : (~0 << (32 - bits));
      return (ipNum & mask) === (rangeNum & mask);
    }
    return false;
  });
}

function _ipToLong(ip) {
  const parts = ip.split('.');
  if (parts.length !== 4) return null;
  return parts.reduce((acc, part) => (acc << 8) + parseInt(part, 10), 0) >>> 0;
}

/**
 * 💡 ASNからCIDRを取得して AllowedIpFromAsn を洗い替えする実処理（IPinfo.io版）
 */
function refreshCidrFromAsn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. AllowedAsn シートから取得対象のASN一覧を取り出す
  const asnSheet = ss.getSheetByName('AllowedAsn');
  if (!asnSheet) {
    console.warn("AllowedAsnシートが見つかりません。処理をスキップします。");
    return;
  }
  const asnData = asnSheet.getDataRange().getValues();
  // ヘッダーを除き、"AS1234" や "1234" から数字のみを抽出して、"AS12345" の形式に統一
  const asnList = asnData.slice(1).map(row => {
    const match = String(row[0]).match(/\d+/);
    return match ? `AS${match[0]}` : null;
  }).filter(Boolean);

  if (asnList.length === 0) {
    console.warn("AllowedAsnシートに対象のASNが記載されていません。");
    return;
  }

  // 2. 各ASNに対して IPinfo API を叩いてCIDRを一括取得
  let allPrefixes = [];
  
  asnList.forEach(asn => {
    // IPinfo.io のエンドポイント（例: https://ipinfo.io/AS15169）
    const url = `https://ipinfo.io/${asn}`;
    const maxRetries = 3; // 共有IPガチャ対策のリトライ回数
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const responseCode = response.getResponseCode();
        
        if (responseCode === 200) {
          const json = JSON.parse(response.getContentText());
          
          // IPinfoのレスポンスからIPv4のprefixes（CIDR）配列を取得
          if (json.prefixes && Array.isArray(json.prefixes)) {
            json.prefixes.forEach(item => {
              // オブジェクトの中に netblock (CIDR文字列) が入っているため抽出
              if (item.netblock && item.netblock.includes('/')) {
                allPrefixes.push(item.netblock.trim());
              }
            });
            console.log(`[Success] ${asn} から ${json.prefixes.length} 件のCIDRを取得しました。`);
          } else {
            console.warn(`[Warning] ${asn} のデータ構造に prefixes が見つかりませんでした。`);
          }
          break; // 成功したためリトライループを抜ける
          
        } else if (responseCode === 429) {
          // もしGoogleのIP被りでレート制限を喰らった場合、数秒待って再試行
          console.warn(`[Rate Limit] ${asn} で制限(429)を検知。${attempt}回目のリトライをします...`);
          if (attempt < maxRetries) {
            Utilities.sleep(4000); // 4秒待機してGoogleの別IPガチャを狙う
          }
        } else {
          console.warn(`ASN: ${asn} の情報取得に失敗しました。ステータスコード: ${responseCode}`);
          break; 
        }
      } catch (e) {
        console.error(`ASN: ${asn} (試行 ${attempt}/${maxRetries}) の通信エラー: ` + e.toString());
        if (attempt < maxRetries) Utilities.sleep(4000);
      }
    }
    
    // APIへの連続アクセスによる負荷を軽減するためのわずかなウェイト（0.5秒）
    Utilities.sleep(500); 
  });

  // 重複を除去
  allPrefixes = [...new Set(allPrefixes)].filter(Boolean);

  // 3. AllowedIpFromAsn シートのデータをクリアして新リストで洗い替え
  let targetSheet = ss.getSheetByName('AllowedIpFromAsn');
  if (!targetSheet) {
    targetSheet = ss.insertSheet('AllowedIpFromAsn');
  }
  
  targetSheet.clearContents(); // ヘッダー含め全クリア
  targetSheet.getRange(1, 1).setValue('allowed_cidr_from_asn'); // 新たにヘッダー書き込み

  if (allPrefixes.length > 0) {
    // 2次元配列に変換して一括書き込み
    const writeData = allPrefixes.map(prefix => [prefix]);
    targetSheet.getRange(2, 1, writeData.length, 1).setValues(writeData);
  }

  console.log(`AllowedIpFromAsn を更新(IPinfo)しました。取得ASN件数: ${asnList.length}, 総CIDR数: ${allPrefixes.length}`);
  SpreadsheetApp.flush();
}
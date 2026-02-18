/**
 * Slack Bot投稿のリアクション取得スクリプト
 *
 * 指定チャンネル内の特定ワードを含むBot投稿を検索し、
 * メッセージURL・リアクション情報をGoogleスプレッドシートに出力する。
 *
 * === セットアップ手順 ===
 * 1. Slack Appを作成し、以下のBot Token Scopesを付与:
 *    - channels:history (パブリックチャンネルの履歴読み取り)
 *    - reactions:read (リアクション情報の読み取り)
 *    - channels:read (チャンネル情報の読み取り)
 *    ※ プライベートチャンネルの場合は groups:history, groups:read も追加
 * 2. Slack AppをワークスペースにインストールしてBot User OAuth Tokenを取得
 * 3. 下記の定数をプロジェクトに合わせて書き換える
 */

// ==================== 設定 ====================
const CONFIG = {
  SLACK_BOT_TOKEN: 'xoxb-xxxx-xxxx-xxxx',  // Slack Bot User OAuth Token
  CHANNEL_ID: 'C0XXXXXXXXX',                // 対象チャンネルID
  SEARCH_WORD: 'キーワード',                  // 検索ワード
  WORKSPACE_DOMAIN: 'your-workspace',        // ワークスペースのドメイン（xxxxx.slack.com の xxxxx 部分）
  SHEET_NAME: 'SlackReactions',              // 出力先シート名
};

/**
 * メイン関数 - スプレッドシートのメニューまたは手動で実行
 */
function fetchSlackBotReactions() {
  const messages = getChannelMessages_(CONFIG.CHANNEL_ID);
  const filtered = filterBotMessages_(messages, CONFIG.SEARCH_WORD);
  writeToSheet_(filtered);
}

/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Slack連携')
    .addItem('Bot投稿のリアクションを取得', 'fetchSlackBotReactions')
    .addToUi();
}

// ==================== Slack API ====================

/**
 * 指定チャンネルの全メッセージを取得（ページネーション対応）
 * @param {string} channelId - チャンネルID
 * @return {Object[]} メッセージの配列
 */
function getChannelMessages_(channelId) {
  const allMessages = [];
  let cursor = '';

  do {
    const url = 'https://slack.com/api/conversations.history'
      + '?channel=' + channelId
      + '&limit=200'
      + (cursor ? '&cursor=' + cursor : '');

    const response = callSlackApi_(url);

    if (!response.ok) {
      throw new Error('conversations.history failed: ' + response.error);
    }

    allMessages.push(...response.messages);
    cursor = (response.response_metadata && response.response_metadata.next_cursor) || '';
  } while (cursor);

  return allMessages;
}

/**
 * Bot投稿かつ検索ワードを含むメッセージをフィルタリング
 * @param {Object[]} messages - メッセージ配列
 * @param {string} searchWord - 検索ワード
 * @return {Object[]} フィルタ済みメッセージ情報の配列
 */
function filterBotMessages_(messages, searchWord) {
  const results = [];

  for (const msg of messages) {
    // Bot投稿の判定: bot_id が存在する or subtype が 'bot_message'
    const isBot = msg.bot_id || msg.subtype === 'bot_message';
    if (!isBot) continue;

    // 検索ワードを含むかチェック
    const text = msg.text || '';
    if (text.indexOf(searchWord) === -1) continue;

    // メッセージのパーマリンクを生成
    const permalink = buildPermalink_(CONFIG.WORKSPACE_DOMAIN, CONFIG.CHANNEL_ID, msg.ts);

    // リアクション情報を整形
    const reactions = formatReactions_(msg.reactions || []);

    results.push({
      timestamp: msg.ts,
      date: formatTimestamp_(msg.ts),
      text: text,
      permalink: permalink,
      botName: msg.username || msg.bot_profile?.name || 'Unknown Bot',
      reactions: reactions,
      reactionCount: getTotalReactionCount_(msg.reactions || []),
    });
  }

  // 日時の古い順にソート
  results.sort(function(a, b) {
    return parseFloat(a.timestamp) - parseFloat(b.timestamp);
  });

  return results;
}

/**
 * Slack API を呼び出す共通関数
 * @param {string} url - API エンドポイントURL
 * @return {Object} APIレスポンス
 */
function callSlackApi_(url) {
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.SLACK_BOT_TOKEN,
      'Content-Type': 'application/json; charset=utf-8',
    },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

/**
 * Slackメッセージのパーマリンクを生成
 * @param {string} domain - ワークスペースドメイン
 * @param {string} channelId - チャンネルID
 * @param {string} ts - メッセージのタイムスタンプ
 * @return {string} パーマリンクURL
 */
function buildPermalink_(domain, channelId, ts) {
  // ts "1234567890.123456" -> "p1234567890123456"
  const tsForUrl = 'p' + ts.replace('.', '');
  return 'https://' + domain + '.slack.com/archives/' + channelId + '/' + tsForUrl;
}

/**
 * Unixタイムスタンプを日本時間の日時文字列に変換
 * @param {string} ts - Slackタイムスタンプ
 * @return {string} "YYYY/MM/DD HH:mm:ss" 形式
 */
function formatTimestamp_(ts) {
  const date = new Date(parseFloat(ts) * 1000);
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

/**
 * リアクション配列を "emoji(count)" 形式の文字列に変換
 * @param {Object[]} reactions - リアクション配列
 * @return {string} 整形済み文字列
 */
function formatReactions_(reactions) {
  if (!reactions || reactions.length === 0) return '';

  return reactions.map(function(r) {
    return ':' + r.name + ': (' + r.count + ')';
  }).join(', ');
}

/**
 * リアクションの合計数を取得
 * @param {Object[]} reactions - リアクション配列
 * @return {number} 合計リアクション数
 */
function getTotalReactionCount_(reactions) {
  return reactions.reduce(function(sum, r) {
    return sum + r.count;
  }, 0);
}

// ==================== スプレッドシート書き込み ====================

/**
 * 結果をスプレッドシートに書き込む
 * @param {Object[]} data - 書き込むデータ配列
 */
function writeToSheet_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // シートが無ければ作成
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  // 既存データをクリア
  sheet.clearContents();

  // ヘッダー行
  const headers = ['投稿日時', 'Bot名', '投稿テキスト', 'メッセージURL', 'リアクション一覧', 'リアクション合計数'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーの書式設定
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');

  if (data.length === 0) {
    Logger.log('条件に一致するBot投稿が見つかりませんでした。');
    SpreadsheetApp.getUi().alert('条件に一致するBot投稿が見つかりませんでした。');
    return;
  }

  // データ行を作成
  const rows = data.map(function(item) {
    return [
      item.date,
      item.botName,
      item.text,
      item.permalink,
      item.reactions,
      item.reactionCount,
    ];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // 列幅を自動調整
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // フィルターを設定
  var range = sheet.getRange(1, 1, rows.length + 1, headers.length);
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  range.createFilter();

  Logger.log(rows.length + '件のBot投稿を出力しました。');
  SpreadsheetApp.getUi().alert(rows.length + '件のBot投稿をシートに出力しました。');
}

function sendGmailToSlack() {
  const slackWebhookUrl = ""; // ← あなたのSlack Webhook URL
  const threads = GmailApp.search('from:*@example.ac.jp newer_than:1d');  // ← ドメインで絞り込み

  for (let thread of threads) {
    const messages = thread.getMessages();
    for (let message of messages) {
      if (!message.isUnread()) continue;

      const subject = message.getSubject();
      const from = message.getFrom();
      const originalBody = message.getPlainBody();

      // ▼ 追加：過去のやり取り（引用部分）をカットして新規メッセージのみにする
      const cleanBody = extractNewMessage(originalBody);

      const payload = {
        text: `📬 *新着メール通知*\n*件名*: ${subject}\n*送信者*: ${from}\n*本文*:\n${cleanBody}`
      };

      UrlFetchApp.fetch(slackWebhookUrl, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
      });

      message.markRead(); // 通知済みとしてマーク（任意）
    }
  }
}

/**
 * メール本文から過去のやり取り（引用部分）をカットする関数
 */
function extractNewMessage(body) {
  const lines = body.split('\n');
  const result = [];
  
  // 一般的なメールソフトの「引用開始」を表すパターンを定義
  const quotePatterns = [
    /^20\d{2}[年/.-]\d{1,2}[月/.-]\d{1,2}.*<.*@.*>/, // Gmail等 (例: 2024年3月25日 10:00 Name <email@...>)
    /^On\s.*wrote:/i,                                // 英語Gmail等 (例: On Mon, Mar 25... wrote:)
    /^-+Original Message-+/i,                        // Outlook等の英語表記
    /^-+\s*元のメッセージ\s*-+/i,                      // Outlook等の日本語表記
    /^_{10,}/,                                       // 長いアンダーバーの区切り線
    /^>/                                             // 引用符「>」で始まる行
  ];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // もし引用開始のパターンが見つかったら、それ以降の行はすべて無視（カット）する
    if (quotePatterns.some(pattern => pattern.test(line))) {
      break; 
    }
    
    result.push(lines[i]);
  }

  // 余分な空白や改行を削って返す
  return result.join('\n').trim();
}
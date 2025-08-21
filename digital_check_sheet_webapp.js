// 【重要】事前に以下の設定を行ってください
// 1. このスクリプトが紐づいているスプレッドシートのIDを以下に設定
const SPREADSHEET_ID = '1U73wJb1vuip1NYav4EtwwDbwOV05tO8tRb6IYMZc6VQ'; 
// 2. GASのプロパティサービスでAPIキーを設定してください
// スクリプトエディタ → プロジェクトの設定 → スクリプトプロパティ → GEMINI_API_KEY を追加
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

// スプレッドシートのシート名
const SHEET_NAME = '回答ログ';

/**
 * WebサイトからのPOSTリクエストを処理するメイン関数
 */
function doPost(e) {
  try {
    Logger.log('doPost function started');
    const requestData = JSON.parse(e.postData.contents);
    const answers = requestData.answers;
    Logger.log('Received answers: ' + JSON.stringify(answers));

    // 1. 回答をスプレッドシートに記録
    Logger.log('Step 1: Logging answers to sheet');
    logAnswersToSheet(answers);
    
    // 2. スコアを計算
    Logger.log('Step 2: Calculating scores');
    const scores = calculateScores(answers);
    Logger.log('Scores calculated: ' + JSON.stringify(scores));
    
    // 3. Gemini APIで診断コメントを生成
    Logger.log('Step 3: Generating comment');
    const geminiComment = generateCommentWithGemini(scores);
    Logger.log('Comment generated successfully');
    
    // 4. PDFレポートを生成
    Logger.log('Step 4: Creating PDF report');
    const pdfUrl = createPdfReport(scores, geminiComment);
    Logger.log('PDF created: ' + pdfUrl);
    
    // 5. 成功レスポンスを返す
    const response = { status: 'success', pdfUrl: pdfUrl };
    Logger.log('Returning success response');
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    // エラーの詳細をJSON形式でフロントエンドに返す
    const errorResponse = {
      status: 'error',
      message: error.toString(),
      details: {
        message: error.message,
        fileName: error.fileName,
        lineNumber: error.lineNumber,
        stack: error.stack,
      }
    };
    return ContentService.createTextOutput(JSON.stringify(errorResponse)).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 回答データをスプレッドシートに記録する
 */
function logAnswersToSheet(answers) {
  try {
    Logger.log("Attempting to log answers to sheet...");
    Logger.log("Spreadsheet ID: " + SPREADSHEET_ID);
    Logger.log("Answers: " + JSON.stringify(answers));
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Spreadsheet opened successfully");
    
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log("Sheet not found, creating new sheet: " + SHEET_NAME);
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = ['タイムスタンプ', ...Array.from({ length: 20 }, (_, i) => `Q${i + 1}`)];
      sheet.appendRow(headers);
      Logger.log("Headers added: " + JSON.stringify(headers));
    }
    
    const rowData = [new Date(), ...answers];
    sheet.appendRow(rowData);
    Logger.log("Row appended successfully: " + JSON.stringify(rowData));
    
  } catch(e) {
    Logger.log("Failed to log answers to sheet: " + e.toString());
    Logger.log("Error details: " + JSON.stringify(e));
    throw e; // エラーを再スローして上位で処理
  }
}

/**
 * 回答からスコアを計算する
 */
function calculateScores(answers) {
  const categories = {
    'コミュニケーション・情報共有': { score: 0, count: 5 },
    '業務プロセス・効率化': { score: 0, count: 5 },
    'セキュリティ・情報管理': { score: 0, count: 5 },
    '経営・データ活用': { score: 0, count: 5 },
  };

  answers.forEach((answer, index) => {
    const score = (answer === 'yes') ? 1 : 0; // 「はい」=問題あり=1点、「いいえ」=問題なし=0点
    if (index < 5) categories['コミュニケーション・情報共有'].score += score;
    else if (index < 10) categories['業務プロセス・効率化'].score += score;
    else if (index < 15) categories['セキュリティ・情報管理'].score += score;
    else categories['経営・データ活用'].score += score;
  });

  const totalScore = Object.values(categories).reduce((sum, cat) => sum + cat.score, 0);
  return { categories, totalScore };
}

/**
 * Gemini APIを使って診断コメントを生成する
 */
function generateCommentWithGemini(scores) {
  if (GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY' || !GEMINI_API_KEY) {
    throw new Error("Gemini API Error: API key is not set in the script.");
  }

  const { categories, totalScore } = scores;
  
  let resultType = '';
  if (totalScore >= 15) resultType = '【赤信号】今すぐ改革必須！DX待ったなしタイプ';
  else if (totalScore >= 10) resultType = '【黄信号】課題が山積！アナログ業務見直しタイプ';
  else if (totalScore >= 5) resultType = '【青信号】あと一歩！デジタル化優等生タイプ';
  else resultType = '【素晴らしい！】DX推進リーダータイプ';

  const prompt = `
あなたは優秀な中小企業向けのITコンサルタントです。以下の診断結果データに基づき、具体的で示唆に富むアドバイスを生成してください。

【重要】絶対に守ってください：
- 個人への呼びかけ（「◯◯社長」「○○様」「社長様」「〜様」など）は一切使用しないでください
- 挨拶や感謝の言葉は一切不要です
- いきなり診断結果の解説から始めてください
- 一般的なアドバイスとして記述してください

# 診断結果データ
- 総合診断タイプ: ${resultType}
- 総合スコア: ${totalScore} / 20点
- カテゴリ別スコア:
  - コミュニケーション・情報共有: ${categories['コミュニケーション・情報共有'].score} / 5点
  - 業務プロセス・効率化: ${categories['業務プロセス・効率化'].score} / 5点
  - セキュリティ・情報管理: ${categories['セキュリティ・情報管理'].score} / 5点
  - 経営・データ活用: ${categories['経営・データ活用'].score} / 5点

# 指示
1. まず、総合診断タイプと総合スコアについて、その意味合いを解説してください。
2. 次に、カテゴリ別スコアの中で、最も点数が低いカテゴリを「最大の課題」として特定し、そのカテゴリでどのような問題が起きている可能性が高いかを、具体例を交えて鋭く指摘してください。
3. 最後に、その最大の課題を解決するための「最初の一歩」として、具体的で実行可能なアクションを2〜3個提案してください。
4. 全体を通して、専門用語は避け、中小企業の経営者に寄り添うような、丁寧かつ力強いトーンで記述してください。
5. 出力はMarkdown形式で、見出しや箇条書きを効果的に使用してください。文字数は400〜600字程度にまとめてください。
6. 冒頭の挨拶は一切不要です。いきなり診断結果の解説から始めてください。`;

  const url = `https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true // エラー時もレスポンスを取得するため
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const result = JSON.parse(responseBody);
    if (result.candidates && result.candidates.length > 0) {
      return result.candidates[0].content.parts[0].text;
    } else {
      throw new Error("Gemini API Error: Invalid response structure. " + responseBody);
    }
  } else {
    Logger.log(`Gemini API Error: Failed with status code ${responseCode}. Response: ${responseBody}`);
    // Gemini APIが利用できない場合は、デフォルトのコメントを返す
    return generateDefaultComment(scores);
  }
}

/**
 * Gemini APIが利用できない場合のデフォルトコメントを生成する
 */
function generateDefaultComment(scores) {
  const { categories, totalScore } = scores;
  
  let resultType = '';
  if (totalScore >= 15) resultType = '【赤信号】今すぐ改革必須！DX待ったなしタイプ';
  else if (totalScore >= 10) resultType = '【黄信号】課題が山積！アナログ業務見直しタイプ';
  else if (totalScore >= 5) resultType = '【青信号】あと一歩！デジタル化優等生タイプ';
  else resultType = '【素晴らしい！】DX推進リーダータイプ';

  // 最も点数が高いカテゴリ（問題が多い）を特定
  let highestCategory = '';
  let highestScore = 0;
  Object.keys(categories).forEach(key => {
    if (categories[key].score > highestScore) {
      highestScore = categories[key].score;
      highestCategory = key;
    }
  });

  return `診断結果の解説

総合診断タイプについて
あなたの会社は「${resultType}」です。総合スコアは${totalScore}/20点で、${totalScore >= 10 ? '改善の余地が大きい' : totalScore >= 5 ? '部分的に改善が必要' : '良好な状態'}です。

最大の課題
最も改善が必要な分野は「${highestCategory}」です。この分野では以下のような問題が起きている可能性があります：

* 業務効率の低下
* 情報共有の不備
* セキュリティリスク
* データ活用の不足

最初の一歩として
1. 現状の業務フローを見直す
2. 必要なツールの導入を検討する
3. 社員への研修を実施する

詳細な改善提案については、専門家にご相談ください。`;
}

/**
 * 診断結果を元にPDFレポートを作成し、URLを返す
 */
function createPdfReport(scores, geminiComment) {
    const { categories, totalScore } = scores;
    const docName = `デジタル化推進度チェックシート_診断レポート_${new Date().getTime()}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    const h2Style = { [DocumentApp.Attribute.FONT_SIZE]: 14, [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#0B57D0' };
    const normalStyle = { [DocumentApp.Attribute.FONT_SIZE]: 11, [DocumentApp.Attribute.BOLD]: false };

    body.appendParagraph('会社のIT健康診断！\nデジタル化推進度チェックシート').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('パーソナル診断レポート').setHeading(DocumentApp.ParagraphHeading.SUBTITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendHorizontalRule();

    let resultType = '';
    if (totalScore >= 15) resultType = '【赤信号】今すぐ改革必須！DX待ったなしタイプ';
    else if (totalScore >= 10) resultType = '【黄信号】課題が山積！アナログ業務見直しタイプ';
    else if (totalScore >= 5) resultType = '【青信号】あと一歩！デジタル化優等生タイプ';
    else resultType = '【素晴らしい！】DX推進リーダータイプ';

    body.appendParagraph('総合診断結果').setAttributes(h2Style);
    body.appendParagraph(`貴社の総合スコアは ${totalScore} / 20 点です。`).setAttributes(normalStyle);
    body.appendParagraph(`診断タイプ： ${resultType}`).setAttributes(normalStyle);
    body.appendParagraph('\n');

    body.appendParagraph('カテゴリ別スコア').setAttributes(h2Style);
    
    // 表形式でスコアを表示（レーダーチャートの代わり）
    const table = body.appendTable();
    
    // ヘッダー行
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell('カテゴリ').setAttributes({[DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 12});
    headerRow.appendTableCell('スコア').setAttributes({[DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 12});
    headerRow.appendTableCell('評価').setAttributes({[DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 12});
    headerRow.appendTableCell('詳細').setAttributes({[DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 12});
    
    // データ行
    Object.keys(categories).forEach(key => {
        const row = table.appendTableRow();
        row.appendTableCell(key).setAttributes({[DocumentApp.Attribute.FONT_SIZE]: 11});
        row.appendTableCell(`${categories[key].score} / 5`).setAttributes({[DocumentApp.Attribute.FONT_SIZE]: 11});
        
        // 評価と詳細
        let evaluation, details;
        if (categories[key].score >= 4) {
            evaluation = '要改善';
            details = '緊急の対応が必要です';
        } else if (categories[key].score >= 2) {
            evaluation = '注意';
            details = '改善の余地があります';
        } else {
            evaluation = '良好';
            details = '良好な状態です';
        }
        
        row.appendTableCell(evaluation).setAttributes({[DocumentApp.Attribute.FONT_SIZE]: 11});
        row.appendTableCell(details).setAttributes({[DocumentApp.Attribute.FONT_SIZE]: 11});
    });
    
    // 表の後に空行を追加
    body.appendParagraph('\n');
    
    // 総合評価の追加
    body.appendParagraph('総合評価').setAttributes(h2Style);
    const totalEvaluation = totalScore >= 15 ? '緊急対応が必要' : 
                           totalScore >= 10 ? '改善が必要' : 
                           totalScore >= 5 ? '部分的改善' : '良好';
    body.appendParagraph(`総合スコア ${totalScore}/20点: ${totalEvaluation}`).setAttributes(normalStyle);
    body.appendParagraph('\n');
    
    body.appendParagraph('ITコンサルタントによるAI分析コメント').setAttributes(h2Style);
    
    // Markdown記法を適切なGoogle Docs形式に変換
    geminiComment.split('\n').forEach(line => {
        const trimmedLine = line.trim();
        
        if (trimmedLine.startsWith('## ')) {
            // 見出し2
            body.appendParagraph(trimmedLine.substring(3)).setHeading(DocumentApp.ParagraphHeading.HEADING2);
        } else if (trimmedLine.startsWith('### ')) {
            // 見出し3
            body.appendParagraph(trimmedLine.substring(4)).setHeading(DocumentApp.ParagraphHeading.HEADING3);
        } else if (trimmedLine.startsWith('# ')) {
            // 見出し1
            body.appendParagraph(trimmedLine.substring(2)).setHeading(DocumentApp.ParagraphHeading.HEADING1);
        } else if (trimmedLine.startsWith('* ') || trimmedLine.startsWith('- ')) {
            // 箇条書き - Markdown強調記法を処理
            let listText = trimmedLine.substring(2);
            if (listText.includes('**')) {
                // 箇条書き内の太字処理
                const regex = /\*\*(.*?)\*\*/g;
                listText = listText.replace(regex, '$1');
                const listItem = body.appendListItem(listText).setGlyphType(DocumentApp.GlyphType.BULLET);
                // 太字部分を後から適用
                const textElements = listItem.getText().split(regex);
                if (textElements.length > 1) {
                    // 太字部分を検出して適用
                    const regex2 = /\*\*(.*?)\*\*/g;
                    let match;
                    while ((match = regex2.exec(trimmedLine.substring(2))) !== null) {
                        const boldText = match[1];
                        const textElement = listItem.findText(boldText);
                        if (textElement) {
                            textElement.setBold(true);
                        }
                    }
                }
            } else {
                body.appendListItem(listText).setGlyphType(DocumentApp.GlyphType.BULLET);
            }
        } else if (trimmedLine.match(/^\d+\.\s/)) {
            // 番号付きリスト - Markdown強調記法を処理
            let listText = trimmedLine.replace(/^\d+\.\s/, '');
            if (listText.includes('**')) {
                // 番号付きリスト内の太字処理
                const regex = /\*\*(.*?)\*\*/g;
                listText = listText.replace(regex, '$1');
                const listItem = body.appendListItem(listText).setGlyphType(DocumentApp.GlyphType.NUMBER);
                // 太字部分を後から適用
                const regex2 = /\*\*(.*?)\*\*/g;
                let match;
                while ((match = regex2.exec(trimmedLine.replace(/^\d+\.\s/, ''))) !== null) {
                    const boldText = match[1];
                    const textElement = listItem.findText(boldText);
                    if (textElement) {
                        textElement.setBold(true);
                    }
                }
            } else {
                body.appendListItem(listText).setGlyphType(DocumentApp.GlyphType.NUMBER);
            }
        } else if (trimmedLine === '') {
            // 空行
            body.appendParagraph('');
        } else {
            // 通常のテキスト - Markdownの強調記法を処理
            let processedLine = trimmedLine;
            
            // **太字** を太字に変換
            if (processedLine.includes('**')) {
                // 正規表現で**で囲まれた部分を検出して置換
                const regex = /\*\*(.*?)\*\*/g;
                let match;
                let lastIndex = 0;
                let paragraph = body.appendParagraph('');
                
                while ((match = regex.exec(processedLine)) !== null) {
                    // マッチ前の通常テキストを追加
                    if (match.index > lastIndex) {
                        const normalText = processedLine.substring(lastIndex, match.index);
                        if (normalText.trim() !== '') {
                            paragraph.appendText(normalText);
                        }
                    }
                    
                    // 太字部分を追加
                    const boldText = match[1];
                    const textElement = paragraph.appendText(boldText);
                    textElement.setBold(true);
                    
                    lastIndex = match.index + match[0].length;
                }
                
                // 残りの通常テキストを追加
                if (lastIndex < processedLine.length) {
                    const remainingText = processedLine.substring(lastIndex);
                    if (remainingText.trim() !== '') {
                        paragraph.appendText(remainingText);
                    }
                }
            } else {
                // 強調記法がない場合は通常のテキストとして追加
                body.appendParagraph(processedLine).setAttributes(normalStyle);
            }
        }
    });

    body.appendPageBreak();
    body.appendParagraph('次のステップのご案内').setAttributes(h2Style);
    body.appendParagraph('診断で明らかになった課題を解決するため、専門家があなたの会社に合わせた最適な解決策をご提案します。\n\n【Step1：情報収集から始めたい方へ】\n「明日からできる！情報セキュリティ対策 最初の10のステップ」の資料をご用意しています。ご希望の場合はお問い合わせください。\n\n【Step2：具体的に相談したい方へ】\n「IT課題の壁打ち 30分無料オンライン相談会」を毎月3社様限定で実施中です。以下の連絡先までお気軽にご連絡ください。\n連絡先: xxx-xxxx-xxxx / email: info@example.com');
    
    doc.saveAndClose();
    const pdfFile = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
    const pdfInDrive = DriveApp.createFile(pdfFile).setName(docName + '.pdf');
    pdfInDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return pdfInDrive.getUrl();
}

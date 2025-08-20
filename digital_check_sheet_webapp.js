// 【重要】事前に以下の設定を行ってください
// 1. このスクリプトが紐づいているスプレッドシートのIDを以下に設定
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'; 
// 2. Google AI Studio (https://aistudio.google.com/) で取得したAPIキーを設定
const GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY';

// スプレッドシートのシート名
const SHEET_NAME = '回答ログ';

/**
 * WebサイトからのPOSTリクエストを処理するメイン関数
 */
function doPost(e) {
  try {
    // POSTされたデータを取得
    const requestData = JSON.parse(e.postData.contents);
    const answers = requestData.answers; // ["yes", "no", "yes", ...]

    // 1. 回答をスプレッドシートに記録
    logAnswersToSheet(answers);

    // 2. スコアを計算
    const scores = calculateScores(answers);

    // 3. Gemini APIで診断コメントを生成
    const geminiComment = generateCommentWithGemini(scores);

    // 4. PDFレポートを生成
    const pdfUrl = createPdfReport(scores, geminiComment);
    
    // 5. 成功レスポンスを返す
    const response = {
      status: 'success',
      pdfUrl: pdfUrl
    };
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
    // エラーレスポンスを返す
    const response = {
      status: 'error',
      message: error.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 回答データをスプレッドシートに記録する
 */
function logAnswersToSheet(answers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['タイムスタンプ'];
    for (let i = 1; i <= 20; i++) {
      headers.push(`Q${i}`);
    }
    sheet.appendRow(headers);
  }
  const timestamp = new Date();
  sheet.appendRow([timestamp, ...answers]);
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
    const score = (answer === 'no') ? 1 : 0;
    if (index < 5) {
      categories['コミュニケーション・情報共有'].score += score;
    } else if (index < 10) {
      categories['業務プロセス・効率化'].score += score;
    } else if (index < 15) {
      categories['セキュリティ・情報管理'].score += score;
    } else {
      categories['経営・データ活用'].score += score;
    }
  });

  const totalScore = Object.values(categories).reduce((sum, cat) => sum + cat.score, 0);
  return { categories, totalScore };
}

/**
 * Gemini APIを使って診断コメントを生成する
 */
function generateCommentWithGemini(scores) {
  const { categories, totalScore } = scores;
  
  let resultType = '';
  if (totalScore >= 15) resultType = '【赤信号】今すぐ改革必須！DX待ったなしタイプ';
  else if (totalScore >= 10) resultType = '【黄信号】課題が山積！アナログ業務見直しタイプ';
  else if (totalScore >= 5) resultType = '【青信号】あと一歩！デジタル化優等生タイプ';
  else resultType = '【素晴らしい！】DX推進リーダータイプ';

  const prompt = `
あなたは優秀な中小企業向けのITコンサルタントです。以下の診断結果データに基づき、経営者に対してパーソナライズされた、具体的で示唆に富むアドバイスを生成してください。

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
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-preview-0514:generateContent?key=${GEMINI_API_KEY}`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{
        parts: [{ text: prompt }]
      }]
    })
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    return result.candidates[0].content.parts[0].text;
  } catch (e) {
    Logger.log("Gemini API Error: " + e.toString());
    return "AIによる分析中にエラーが発生しました。診断結果のスコアをご確認ください。";
  }
}


/**
 * 診断結果を元にPDFレポートを作成し、URLを返す
 */
function createPdfReport(scores, geminiComment) {
    const { categories, totalScore } = scores;

    // 1. Googleドキュメントを一時的に作成
    const docName = `デジタル化推進度チェックシート_診断レポート_${new Date().getTime()}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    // スタイルの設定
    const titleStyle = { [DocumentApp.Attribute.FONT_SIZE]: 18, [DocumentApp.Attribute.BOLD]: true };
    const h2Style = { [DocumentApp.Attribute.FONT_SIZE]: 14, [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#0B57D0' };
    const normalStyle = { [DocumentApp.Attribute.FONT_SIZE]: 11, [DocumentApp.Attribute.BOLD]: false };

    // ヘッダー
    body.appendParagraph('会社のIT健康診断！\nデジタル化推進度チェックシート').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('パーソナル診断レポート').setHeading(DocumentApp.ParagraphHeading.SUBTITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendHorizontalRule();

    // 診断結果
    let resultType = '';
    if (totalScore >= 15) resultType = '【赤信号】今すぐ改革必須！DX待ったなしタイプ';
    else if (totalScore >= 10) resultType = '【黄信号】課題が山積！アナログ業務見直しタイプ';
    else if (totalScore >= 5) resultType = '【青信号】あと一歩！デジタル化優等生タイプ';
    else resultType = '【素晴らしい！】DX推進リーダータイプ';

    body.appendParagraph('総合診断結果').setAttributes(h2Style);
    body.appendParagraph(`貴社の総合スコアは ${totalScore} / 20 点です。`).setAttributes(normalStyle);
    body.appendParagraph(`診断タイプ： ${resultType}`).setAttributes(normalStyle);
    body.appendParagraph('\n');

    // レーダーチャート画像
    const chartData = Object.keys(categories).map(key => categories[key].score).join(',');
    const chartLabels = Object.keys(categories).join('|');
    const chartUrl = `https://image-charts.com/chart?cht=r&chd=t:${chartData},5&chds=0,5&chs=400x400&chxt=x&chxl=0:|${chartLabels}&chco=3092DE&chls=2&chm=B,3092DE,0,0,0`;
    
    try {
        const chartBlob = UrlFetchApp.fetch(chartUrl).getBlob();
        body.appendParagraph('カテゴリ別スコア').setAttributes(h2Style);
        body.appendImage(chartBlob).setWidth(350).setHeight(350);
        body.getParagraphs().slice(-1)[0].setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } catch(e) {
        Logger.log("Chart API Error: " + e.toString());
        body.appendParagraph('チャートの生成に失敗しました。').setAttributes(normalStyle);
    }
    
    // GeminiによるAI分析コメント
    body.appendParagraph('ITコンサルタントによるAI分析コメント').setAttributes(h2Style);
    // Markdownを簡易的にパースして整形
    const lines = geminiComment.split('\n');
    lines.forEach(line => {
        if (line.startsWith('# ')) {
            body.appendParagraph(line.substring(2)).setHeading(DocumentApp.ParagraphHeading.HEADING2);
        } else if (line.startsWith('## ')) {
            body.appendParagraph(line.substring(3)).setHeading(DocumentApp.ParagraphHeading.HEADING3);
        } else if (line.startsWith('* ') || line.startsWith('- ')) {
            const listItem = body.appendListItem(line.substring(2));
            listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
        } else {
            body.appendParagraph(line).setAttributes(normalStyle);
        }
    });

    // 次のステップ (CTA)
    body.appendPageBreak();
    body.appendParagraph('次のステップのご案内').setAttributes(h2Style);
    body.appendParagraph('診断で明らかになった課題を解決するため、専門家があなたの会社に合わせた最適な解決策をご提案します。').setAttributes(normalStyle);
    body.appendParagraph('\n');
    body.appendParagraph('【Step1：まずは情報収集から始めたい方へ】').setAttributes({[DocumentApp.Attribute.BOLD]: true});
    body.appendParagraph('「明日からできる！情報セキュリティ対策 最初の10のステップ」の資料をご用意しています。ご希望の場合はお問い合わせください。');
    body.appendParagraph('\n');
    body.appendParagraph('【Step2：具体的に相談したい方へ】').setAttributes({[DocumentApp.Attribute.BOLD]: true});
    body.appendParagraph('「IT課題の壁打ち 30分無料オンライン相談会」を毎月3社様限定で実施中です。以下の連絡先までお気軽にご連絡ください。');
    body.appendParagraph('連絡先: xxx-xxxx-xxxx / email: info@example.com').setAttributes({[DocumentApp.Attribute.BOLD]: true});

    doc.saveAndClose();

    // 2. PDFに変換してURLを取得
    const pdfBlob = doc.getAs('application/pdf');
    const pdfFile = DriveApp.createFile(pdfBlob).setName(docName + '.pdf');
    
    // 誰でも閲覧できるように権限を設定
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 3. 一時的に作成したドキュメントを削除
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return pdfFile.getUrl();
}

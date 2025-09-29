// 【重要】事前に以下の設定を行ってください
// 1. このスクリプトが紐づいているスプレッドシートのIDを以下に設定
const SPREADSHEET_ID = '1U73wJb1vuip1NYav4EtwwDbwOV05tO8tRb6IYMZc6VQ'; 
// 2. GASのプロパティサービスでAPIキーを設定してください
// スクリプトエディタ → プロジェクトの設定 → スクリプトプロパティ → GEMINI_API_KEY を追加
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

// スプレッドシートのシート名
const SHEET_NAME = '回答ログ';

/**
 * GETリクエストを処理する関数（ウェブアプリの動作確認用）
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'デジタル化推進度チェックシート API が正常に動作しています',
    version: '1.0.0'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * WebサイトからのPOSTリクエストを処理するメイン関数
 */
function doPost(e) {
  try {
    Logger.log('doPost function started');
    
    // GETリクエストの場合はdoGetの結果を返す
    if (!e || !e.postData || !e.postData.contents) {
      Logger.log('GET request detected, redirecting to doGet');
      return doGet(e);
    }
    
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
  if (!GEMINI_API_KEY) {
    throw new Error("Gemini API Error: API key is not set in the script. Please set GEMINI_API_KEY in Script Properties.");
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

【スコアの意味】：
- 点数が高い（4-5点）= 問題が多い = 改善が必要
- 点数が低い（0-1点）= 問題が少ない = 良好な状態

【カテゴリの重要度順位】（同点時の優先度）：
1. セキュリティ・情報管理（最重要）
2. 経営・データ活用
3. 業務プロセス・効率化
4. コミュニケーション・情報共有

# 診断結果データ
- 総合診断タイプ: ${resultType}
- 総合スコア: ${totalScore} / 20点
- カテゴリ別スコア:
  - コミュニケーション・情報共有: ${categories['コミュニケーション・情報共有'].score} / 5点
  - 業務プロセス・効率化: ${categories['業務プロセス・効率化'].score} / 5点
  - セキュリティ・情報管理: ${categories['セキュリティ・情報管理'].score} / 5点
  - 経営・データ活用: ${categories['経営・データ活用'].score} / 5点

# 指示
1. まず、総合診断タイプと総合スコアについて、その意味合いを詳しく解説してください。

2. 次に、すべてのカテゴリについて、回答順序で詳細に分析してください：
   - 各カテゴリの現状評価（点数が高いほど問題が多い）
   - そのカテゴリで起きている可能性が高い具体的な問題点を3〜5個
   - 各問題点について、なぜそれが問題なのかの理由
   - 改善のための具体的で実践的なアドバイスを2〜3個
   - 各アドバイスについて、どのような効果が期待できるかの説明

3. カテゴリの並び順は以下の通りにしてください（回答順序）：
   - コミュニケーション・情報共有
   - 業務プロセス・効率化
   - セキュリティ・情報管理
   - 経営・データ活用

4. 各カテゴリは以下の形式で出力してください：
   ## 【カテゴリ名】スコア：X/5点
   ### 現状評価
   （点数に応じた詳細な評価コメント）
   
   ### 主な問題点とその影響
   * **問題点1**: 具体的な問題の説明と、なぜそれが問題なのかの理由
   * **問題点2**: 具体的な問題の説明と、なぜそれが問題なのかの理由
   * **問題点3**: 具体的な問題の説明と、なぜそれが問題なのかの理由
   
   ### 改善のための具体的アドバイス
   * **アドバイス1**: 具体的な改善方法と期待される効果
   * **アドバイス2**: 具体的な改善方法と期待される効果
   * **アドバイス3**: 具体的な改善方法と期待される効果

5. 全体を通して、専門用語は避け、中小企業の経営者に寄り添うような、丁寧かつ力強いトーンで記述してください。

6. 出力はMarkdown形式で、見出しや箇条書きを効果的に使用してください。文字数は1500〜2000字程度にまとめてください。

7. 冒頭の挨拶は一切不要です。いきなり診断結果の解説から始めてください。

8. 各問題点やアドバイスは、中小企業が実際に取り組める具体的な内容にしてください。

9. 可能な限り、数値や期間、具体的なツール名や手法名を含めて、実践的な内容にしてください。`;

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

  // カテゴリを回答順序でソート（コミュニケーション・情報共有 → 業務プロセス・効率化 → セキュリティ・情報管理 → 経営・データ活用）
  const categoryOrder = [
    'コミュニケーション・情報共有',
    '業務プロセス・効率化', 
    'セキュリティ・情報管理',
    '経営・データ活用'
  ];
  const sortedCategories = categoryOrder;

  let comment = `診断結果の詳細分析

## 総合診断結果について

あなたの会社は「${resultType}」に該当します。総合スコアは${totalScore}/20点で、${getOverallEvaluation(totalScore)}です。

この診断では、デジタル化推進における4つの重要な領域を評価しています。各領域のスコアが高いほど、その領域で改善が必要な課題が多いことを示しています。

`;

  // 各カテゴリについて詳細な分析コメントを生成
  sortedCategories.forEach(category => {
    const score = categories[category].score;
    const analysis = getCategoryAnalysis(category, score);
    
    comment += `## 【${category}】スコア：${score}/5点
### 現状評価
${analysis.evaluation}

### 主な問題点とその影響
`;
    
    analysis.problems.forEach(problem => {
      comment += `* **${problem.title}**: ${problem.description}\n`;
    });
    
    comment += `
### 改善のための具体的アドバイス
`;
    
    analysis.advice.forEach(adv => {
      comment += `* **${adv.title}**: ${adv.description}\n`;
    });
    
    comment += '\n';
  });

  comment += `## 次のステップについて

診断で明らかになった課題を解決するため、段階的な改善計画を立てることをお勧めします。特にスコアの高い領域から優先的に取り組むことで、効果的なデジタル化推進が期待できます。

詳細な改善提案については、専門家にご相談ください。`;
  
  return comment;
}

/**
 * 総合スコアに基づく評価を取得
 */
function getOverallEvaluation(totalScore) {
  if (totalScore >= 15) {
    return '緊急の対応が必要な状態です。デジタル化の抜本的見直しが求められます。';
  } else if (totalScore >= 10) {
    return '改善が必要な状態です。段階的なデジタル化推進が重要です。';
  } else if (totalScore >= 5) {
    return '部分的に改善が必要な状態です。さらなる効率化が期待できます。';
  } else {
    return '良好な状態です。継続的な改善により更なる向上が期待できます。';
  }
}

/**
 * カテゴリ別の詳細分析を取得
 */
function getCategoryAnalysis(category, score) {
  const analysisData = {
    'セキュリティ・情報管理': {
      5: {
        evaluation: '緊急の対応が必要な状態です。情報セキュリティリスクが非常に高い状況にあります。',
        problems: [
          { title: '情報漏洩リスクの高さ', description: '適切なセキュリティ対策が講じられておらず、顧客情報や機密情報の漏洩リスクが極めて高い状態です。これにより、法的責任や信用失墜のリスクが発生します。' },
          { title: 'サイバー攻撃への脆弱性', description: '基本的なセキュリティ対策が不十分で、マルウェア感染や不正アクセスの被害を受ける可能性が高い状態です。' },
          { title: 'データバックアップの不備', description: '重要なデータのバックアップが適切に行われておらず、システム障害時にデータ損失のリスクがあります。' },
          { title: 'アクセス権限管理の不備', description: '誰がどの情報にアクセスできるかの管理が不適切で、内部不正のリスクが高い状態です。' },
          { title: 'セキュリティ教育の不足', description: '従業員のセキュリティ意識が低く、人的要因によるセキュリティインシデントのリスクが高い状態です。' }
        ],
        advice: [
          { title: 'セキュリティ専門家への相談', description: '情報セキュリティの専門家に相談し、包括的なセキュリティ対策の策定と実装を進めることをお勧めします。初期費用はかかりますが、インシデント発生時の損失を大幅に軽減できます。' },
          { title: '段階的なセキュリティ強化', description: 'まずは基本的な対策（ウイルス対策ソフト、ファイアウォール、パスワード管理）から始め、徐々に高度な対策を導入する段階的アプローチを採用してください。' },
          { title: '従業員教育の実施', description: '定期的なセキュリティ教育を実施し、フィッシングメールの見分け方や安全なパスワードの設定方法などを教育してください。' }
        ]
      },
      4: {
        evaluation: '改善が必要な状態です。セキュリティリスクが高く、早急な対策が求められます。',
        problems: [
          { title: 'セキュリティ対策の不備', description: '基本的なセキュリティ対策が不十分で、情報漏洩やサイバー攻撃のリスクが高い状態です。' },
          { title: 'データ管理の課題', description: '重要なデータの管理方法が不適切で、データ損失や不正アクセスのリスクがあります。' },
          { title: 'アクセス制御の不備', description: '情報へのアクセス権限の管理が不十分で、内部不正のリスクが存在します。' }
        ],
        advice: [
          { title: 'セキュリティ対策の強化', description: 'ウイルス対策ソフトの導入、ファイアウォールの設定、定期的なセキュリティ更新の実施など、基本的なセキュリティ対策を強化してください。' },
          { title: 'データバックアップの確立', description: '重要なデータの定期的なバックアップを確立し、復旧手順を明確化してください。' }
        ]
      },
      3: {
        evaluation: '部分的に改善が必要な状態です。基本的なセキュリティ対策はある程度整っていますが、さらなる強化が推奨されます。',
        problems: [
          { title: 'セキュリティ対策の不統一', description: '一部のセキュリティ対策は実施されているものの、全社的に統一された対策が不足している状況です。' },
          { title: '継続的な監視の不足', description: 'セキュリティ状況の継続的な監視や評価が不十分で、新たな脅威への対応が遅れる可能性があります。' }
        ],
        advice: [
          { title: 'セキュリティポリシーの策定', description: '全社的なセキュリティポリシーを策定し、従業員への周知徹底を図ってください。' },
          { title: '定期的なセキュリティ監査', description: '定期的なセキュリティ監査を実施し、対策の効果を評価・改善してください。' }
        ]
      },
      2: {
        evaluation: '良好な状態ですが、さらなる改善の余地があります。',
        problems: [
          { title: '高度なセキュリティ対策の不足', description: '基本的なセキュリティ対策は整っているものの、より高度な対策の導入により、さらなる安全性の向上が期待できます。' }
        ],
        advice: [
          { title: '高度なセキュリティ対策の検討', description: '多要素認証、暗号化、侵入検知システムなど、より高度なセキュリティ対策の導入を検討してください。' }
        ]
      },
      1: {
        evaluation: '非常に良好な状態です。基本的なセキュリティ対策が適切に実施されています。',
        problems: [
          { title: '継続的な改善の必要性', description: '現状は良好ですが、セキュリティ脅威は日々進化しているため、継続的な改善が必要です。' }
        ],
        advice: [
          { title: '継続的なセキュリティ強化', description: '最新のセキュリティ脅威に対応するため、定期的な対策の見直しと強化を継続してください。' }
        ]
      },
      0: {
        evaluation: '優秀な状態です。セキュリティ対策が適切に実施されています。',
        problems: [
          { title: '維持・向上の継続', description: '現状のセキュリティレベルを維持し、さらなる向上を図る必要があります。' }
        ],
        advice: [
          { title: 'ベストプラクティスの共有', description: '他の部門や他社との情報交換により、さらなる改善の機会を探してください。' }
        ]
      }
    },
    '経営・データ活用': {
      5: {
        evaluation: '緊急の改善が必要な状態です。データ活用による経営判断の基盤が不十分です。',
        problems: [
          { title: 'データドリブン経営の欠如', description: '経営判断に必要なデータが適切に収集・分析されておらず、感覚的な経営判断に依存している状態です。これにより、機会損失や非効率な意思決定が発生します。' },
          { title: 'KPI管理の不備', description: '重要な業績指標（KPI）が明確に定義されておらず、経営状況の把握が困難な状態です。' },
          { title: 'データ分析体制の不足', description: 'データを分析し、経営に活用する体制やスキルが不足しており、データの価値を十分に引き出せていません。' },
          { title: '競合分析の不備', description: '市場動向や競合他社の分析が不十分で、戦略的な経営判断ができていない状態です。' },
          { title: '予算管理の非効率', description: '予算の策定や管理が非効率的で、リソースの最適配分ができていない状態です。' }
        ],
        advice: [
          { title: 'データ分析基盤の構築', description: 'BIツール（Tableau、Power BI等）の導入により、データの可視化と分析基盤を構築してください。初期投資は必要ですが、長期的な経営効率化に大きく貢献します。' },
          { title: 'KPI管理システムの導入', description: '重要な業績指標を明確に定義し、定期的なモニタリング体制を構築してください。' },
          { title: 'データ分析人材の育成', description: '社内にデータ分析スキルを持つ人材を育成するか、外部専門家の活用を検討してください。' }
        ]
      },
      4: {
        evaluation: '改善が必要な状態です。データ活用による経営判断の基盤が不十分です。',
        problems: [
          { title: 'データ活用の不足', description: '収集されたデータが経営判断に十分活用されておらず、データの価値が発揮されていません。' },
          { title: '分析体制の不備', description: 'データを分析し、経営に活用する体制が不十分で、効果的な意思決定ができていません。' }
        ],
        advice: [
          { title: 'データ分析ツールの導入', description: 'Excel以外の分析ツールを導入し、より高度なデータ分析を実施してください。' },
          { title: '定期的な経営会議の実施', description: 'データに基づく定期的な経営会議を実施し、データドリブンな意思決定を習慣化してください。' }
        ]
      },
      3: {
        evaluation: '部分的に改善が必要な状態です。基本的なデータ活用は行われていますが、さらなる強化が推奨されます。',
        problems: [
          { title: 'データ活用の限界', description: '基本的なデータ活用は行われているものの、より高度な分析や予測に活用できていません。' }
        ],
        advice: [
          { title: '高度な分析手法の導入', description: '統計分析や機械学習を活用した、より高度なデータ分析手法の導入を検討してください。' }
        ]
      },
      2: {
        evaluation: '良好な状態ですが、さらなる改善の余地があります。',
        problems: [
          { title: '高度なデータ活用の不足', description: '基本的なデータ活用は行われているものの、より高度な活用により、さらなる経営効率化が期待できます。' }
        ],
        advice: [
          { title: '高度なデータ分析の導入', description: '予測分析や機械学習を活用した、より高度なデータ分析の導入を検討してください。' }
        ]
      },
      1: {
        evaluation: '非常に良好な状態です。データ活用による経営判断が適切に実施されています。',
        problems: [
          { title: '継続的な改善の必要性', description: '現状は良好ですが、データ活用の手法は日々進化しているため、継続的な改善が必要です。' }
        ],
        advice: [
          { title: '継続的なデータ活用強化', description: '最新のデータ分析手法を継続的に学習し、経営判断の精度向上を図ってください。' }
        ]
      },
      0: {
        evaluation: '優秀な状態です。データ活用による経営判断が適切に実施されています。',
        problems: [
          { title: '維持・向上の継続', description: '現状のデータ活用レベルを維持し、さらなる向上を図る必要があります。' }
        ],
        advice: [
          { title: 'ベストプラクティスの共有', description: '他の企業との情報交換により、さらなる改善の機会を探してください。' }
        ]
      }
    },
    '業務プロセス・効率化': {
      5: {
        evaluation: '緊急の改善が必要な状態です。業務プロセスの非効率性が深刻です。',
        problems: [
          { title: '手作業による非効率性', description: '多くの業務が手作業に依存しており、人的ミスや作業時間の増大が発生しています。これにより、コスト増加と生産性の低下が深刻化しています。' },
          { title: 'プロセス標準化の不足', description: '業務プロセスが標準化されておらず、担当者によって作業方法が異なり、品質のばらつきが発生しています。' },
          { title: '情報共有の非効率', description: '部門間での情報共有が非効率的で、重複作業や情報の遅延が頻発しています。' },
          { title: '在庫管理の非効率', description: '在庫管理が非効率的で、過剰在庫や欠品が発生し、資金効率が悪化しています。' },
          { title: '顧客対応の非効率', description: '顧客対応プロセスが非効率的で、顧客満足度の低下や対応時間の増大が発生しています。' }
        ],
        advice: [
          { title: '業務プロセスの見直し', description: '全社的な業務プロセスの見直しを実施し、標準化と効率化を図ってください。RPAツールの導入により、定型業務の自動化を検討してください。' },
          { title: 'デジタルツールの導入', description: 'ERPシステムや業務管理システムの導入により、業務プロセスの統合と効率化を図ってください。' },
          { title: '段階的な改善計画', description: '優先度の高い業務から段階的に改善を進め、継続的な効率化を図ってください。' }
        ]
      },
      4: {
        evaluation: '改善が必要な状態です。業務プロセスの非効率性が存在します。',
        problems: [
          { title: '手作業の多さ', description: '多くの業務が手作業に依存しており、効率化の余地が大きい状態です。' },
          { title: 'プロセス標準化の不足', description: '業務プロセスが標準化されておらず、効率性にばらつきがあります。' }
        ],
        advice: [
          { title: '業務プロセスの標準化', description: '主要な業務プロセスを標準化し、効率性と品質の向上を図ってください。' },
          { title: 'デジタルツールの活用', description: '適切なデジタルツールを導入し、業務効率化を図ってください。' }
        ]
      },
      3: {
        evaluation: '部分的に改善が必要な状態です。基本的な効率化は行われていますが、さらなる改善が推奨されます。',
        problems: [
          { title: '高度な効率化の不足', description: '基本的な効率化は行われているものの、より高度な効率化により、さらなる生産性向上が期待できます。' }
        ],
        advice: [
          { title: '高度な効率化手法の導入', description: 'AIや機械学習を活用した、より高度な業務効率化手法の導入を検討してください。' }
        ]
      },
      2: {
        evaluation: '良好な状態ですが、さらなる改善の余地があります。',
        problems: [
          { title: '高度な効率化の余地', description: '基本的な効率化は行われているものの、より高度な効率化により、さらなる生産性向上が期待できます。' }
        ],
        advice: [
          { title: '高度な効率化の検討', description: '最新の効率化手法を調査し、適用可能なものを導入してください。' }
        ]
      },
      1: {
        evaluation: '非常に良好な状態です。業務プロセスの効率化が適切に実施されています。',
        problems: [
          { title: '継続的な改善の必要性', description: '現状は良好ですが、業務効率化の手法は日々進化しているため、継続的な改善が必要です。' }
        ],
        advice: [
          { title: '継続的な効率化', description: '最新の効率化手法を継続的に学習し、さらなる改善を図ってください。' }
        ]
      },
      0: {
        evaluation: '優秀な状態です。業務プロセスの効率化が適切に実施されています。',
        problems: [
          { title: '維持・向上の継続', description: '現状の効率化レベルを維持し、さらなる向上を図る必要があります。' }
        ],
        advice: [
          { title: 'ベストプラクティスの共有', description: '他の企業との情報交換により、さらなる改善の機会を探してください。' }
        ]
      }
    },
    'コミュニケーション・情報共有': {
      5: {
        evaluation: '緊急の改善が必要な状態です。コミュニケーション・情報共有の非効率性が深刻です。',
        problems: [
          { title: '情報共有の非効率性', description: '重要な情報が適切に共有されておらず、情報の遅延や重複が頻発しています。これにより、意思決定の遅れや機会損失が発生しています。' },
          { title: 'コミュニケーションツールの不足', description: '効果的なコミュニケーションツールが導入されておらず、情報伝達が非効率的です。' },
          { title: '会議の非効率性', description: '会議が非効率的で、時間の浪費と意思決定の遅延が発生しています。' },
          { title: '文書管理の不備', description: '重要な文書の管理が不適切で、必要な情報にアクセスできない状況が発生しています。' },
          { title: 'リモートワーク対応の不足', description: 'リモートワーク環境でのコミュニケーション体制が不十分で、生産性の低下が発生しています。' }
        ],
        advice: [
          { title: 'コミュニケーションツールの導入', description: 'Slack、Microsoft Teams、Zoomなどのコミュニケーションツールを導入し、効率的な情報共有体制を構築してください。' },
          { title: '会議効率化の実施', description: '会議の目的とアジェンダを明確化し、時間管理と意思決定の効率化を図ってください。' },
          { title: '文書管理システムの構築', description: 'Google WorkspaceやMicrosoft 365などの文書管理システムを導入し、情報の一元管理を図ってください。' }
        ]
      },
      4: {
        evaluation: '改善が必要な状態です。コミュニケーション・情報共有の非効率性が存在します。',
        problems: [
          { title: '情報共有の不足', description: '重要な情報が適切に共有されておらず、意思決定に支障をきたしています。' },
          { title: 'コミュニケーションツールの不備', description: '効果的なコミュニケーションツールが不足しており、情報伝達が非効率的です。' }
        ],
        advice: [
          { title: 'コミュニケーションツールの導入', description: '適切なコミュニケーションツールを導入し、情報共有の効率化を図ってください。' },
          { title: '情報共有ルールの策定', description: '情報共有のルールを策定し、全社的な情報共有体制を構築してください。' }
        ]
      },
      3: {
        evaluation: '部分的に改善が必要な状態です。基本的なコミュニケーションは行われていますが、さらなる改善が推奨されます。',
        problems: [
          { title: '高度なコミュニケーション手法の不足', description: '基本的なコミュニケーションは行われているものの、より高度な手法により、さらなる効率化が期待できます。' }
        ],
        advice: [
          { title: '高度なコミュニケーション手法の導入', description: 'ビデオ会議、チャット、プロジェクト管理ツールなどを活用した、より高度なコミュニケーション手法の導入を検討してください。' }
        ]
      },
      2: {
        evaluation: '良好な状態ですが、さらなる改善の余地があります。',
        problems: [
          { title: '高度なコミュニケーションの余地', description: '基本的なコミュニケーションは行われているものの、より高度な手法により、さらなる効率化が期待できます。' }
        ],
        advice: [
          { title: '高度なコミュニケーションの検討', description: '最新のコミュニケーションツールを調査し、適用可能なものを導入してください。' }
        ]
      },
      1: {
        evaluation: '非常に良好な状態です。コミュニケーション・情報共有が適切に実施されています。',
        problems: [
          { title: '継続的な改善の必要性', description: '現状は良好ですが、コミュニケーション手法は日々進化しているため、継続的な改善が必要です。' }
        ],
        advice: [
          { title: '継続的なコミュニケーション改善', description: '最新のコミュニケーション手法を継続的に学習し、さらなる改善を図ってください。' }
        ]
      },
      0: {
        evaluation: '優秀な状態です。コミュニケーション・情報共有が適切に実施されています。',
        problems: [
          { title: '維持・向上の継続', description: '現状のコミュニケーション・情報共有レベルを維持し、さらなる向上を図る必要があります。' }
        ],
        advice: [
          { title: 'ベストプラクティスの共有', description: '他の企業との情報交換により、さらなる改善の機会を探してください。' }
        ]
      }
    }
  };

  return analysisData[category][score] || analysisData[category][0];
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
    // 各カテゴリの前に改ページを挿入
    let isFirstCategory = true;
    geminiComment.split('\n').forEach(line => {
        const trimmedLine = line.trim();
        
        if (trimmedLine.startsWith('## ')) {
            // カテゴリ見出しの前に改ページを挿入（最初のカテゴリ以外）
            if (!isFirstCategory) {
                body.appendPageBreak();
            }
            isFirstCategory = false;
            
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
                // 箇条書き内の太字処理（記号のみ削除）
                const regex = /\*\*(.*?)\*\*/g;
                listText = listText.replace(regex, '$1');
            }
            body.appendListItem(listText).setGlyphType(DocumentApp.GlyphType.BULLET);
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
    
    // 改ページを挿入
    body.appendPageBreak();
    
    // 次のステップのご案内を追加
    body.appendParagraph('次のステップのご案内').setAttributes(h2Style);
    body.appendParagraph('診断で明らかになった課題を解決するため、専門家があなたの会社に合わせた最適な解決策をご提案します。\n\n【Step1：情報収集から始めたい方へ】\n「明日からできる！情報セキュリティ対策 最初の10のステップ」の資料をご用意しています。ご希望の場合はお問い合わせください。\n\n【Step2：具体的に相談したい方へ】\n「IT課題の壁打ち 30分無料オンライン相談会」を毎月3社様限定で実施中です。以下の連絡先までお気軽にご連絡ください。\n連絡先: xxx-xxxx-xxxx / email: info@example.com');
    
    doc.saveAndClose();
    const pdfFile = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
    const pdfInDrive = DriveApp.createFile(pdfFile).setName(docName + '.pdf');
    pdfInDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return pdfInDrive.getUrl();
}

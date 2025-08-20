// DX推進度チェックシート - メインスクリプト

const questions = [
    { category: 'コミュニケーション・情報共有', text: '社外や自宅からだと、社内にあるはずの必要なファイルにアクセスできない。' },
    { category: 'コミュニケーション・情報共有', text: '社内の主要な連絡手段がメールや電話で、急ぎの要件が伝わりにくいことがある。' },
    { category: 'コミュニケーション・情報共有', text: '会議のたびに大量の紙資料を印刷しており、ペーパーレス化が進んでいない。' },
    { category: 'コミュニケーション・情報共有', text: '「あの件、どうなった？」と担当者に聞かないと、仕事の進捗がわからない。' },
    { category: 'コミュニケーション・情報共有', text: '拠点間や部署間の情報共有がうまくいかず、何度も同じ説明をしている。' },
    { category: '業務プロセス・効率化', text: '見積書や稟議書など、いまだに「紙とハンコ」でのやり取りが必須となっている。' },
    { category: '業務プロセス・効率化', text: 'Excelへのデータ入力や、システム間の情報転記といった単純作業に時間を取られている。' },
    { category: '業務プロセス・効率化', text: '顧客情報や過去の取引履歴が、個々の営業担当者のExcelや手帳で管理されている。' },
    { category: '業務プロセス・効率化', text: '過去の資料やデータを探し出すのに、いつも5分以上かかっている。' },
    { category: '業務プロセス・効率化', text: '会社の売上や経費の状況を、複数の資料をかき集めないと把握できない。' },
    { category: 'セキュリティ・情報管理', text: '社員の私物のパソコンやスマホを、業務で使うことを黙認してしまっている。' },
    { category: 'セキュリティ・情報管理', text: '退職した社員のメールアドレスやアカウントが、そのまま放置されている可能性がある。' },
    { category: 'セキュリティ・情報管理', text: 'ウイルス対策ソフトは入れているが、それ以外のセキュリティ対策は特にしていない。' },
    { category: 'セキュリティ・情報管理', text: '重要なデータのバックアップを誰がいつ取っているか、明確なルールがない。' },
    { category: 'セキュリティ・情報管理', text: '社員がカフェの無料Wi-Fiなどを使い、重要なファイルをやり取りしている。' },
    { category: '経営・データ活用', text: '重要な経営判断を、社長や役員の「経験と勘」に頼ることがほとんどだ。' },
    { category: '経営・データ活用', text: 'Webサイトからの問い合わせや顧客データを、有効に活用できているとは言えない。' },
    { category: '経営・データ活用', text: 'ITの導入を「コスト（費用）」と捉えており、「投資」とは考えにくい。' },
    { category: '経営・データ活用', text: '社内にITに詳しい人材がおらず、パソコンのトラブルが起きると業務が止まる。' },
    { category: '経営・データ活用', text: '新しいツールを導入しようとすると、社員から「面倒だ」という反対の声が上がる。' }
];

// DOM要素の取得
const startBtn = document.getElementById('startBtn');
const introduction = document.getElementById('introduction');
const checkSheetForm = document.getElementById('checkSheetForm');
const questionsContainer = document.getElementById('questionsContainer');
const loading = document.getElementById('loading');
const result = document.getElementById('result');
const error = document.getElementById('error');
const downloadLink = document.getElementById('downloadLink');
const retryBtn = document.getElementById('retryBtn');
const progressSection = document.getElementById('progressSection');
const progressText = document.getElementById('progressText');
const progressBar = document.getElementById('progressBar');

/**
 * 設問を動的に生成する関数
 */
function renderQuestions() {
    let currentCategory = '';
    questions.forEach((q, index) => {
        if (q.category !== currentCategory) {
            currentCategory = q.category;
            const categoryTitle = document.createElement('h3');
            categoryTitle.className = 'text-xl font-bold text-gray-800 border-b-2 border-blue-500 pb-2 mb-4';
            categoryTitle.textContent = `【${currentCategory}】`;
            questionsContainer.appendChild(categoryTitle);
        }

        const questionBlock = document.createElement('div');
        questionBlock.className = 'p-4 border border-gray-200 rounded-lg';
        
        const questionText = document.createElement('p');
        questionText.className = 'font-medium text-gray-800 mb-3';
        questionText.textContent = `Q${index + 1}. ${q.text}`;
        
        const radioContainer = document.createElement('div');
        radioContainer.className = 'flex items-center space-x-6';
        
        const yesLabel = document.createElement('label');
        yesLabel.className = 'flex items-center cursor-pointer';
        const yesInput = document.createElement('input');
        yesInput.type = 'radio';
        yesInput.name = `q${index}`;
        yesInput.value = 'yes';
        yesInput.required = true;
        yesInput.className = 'h-5 w-5 text-blue-600 focus:ring-blue-500 border-gray-300';
        yesLabel.appendChild(yesInput);
        yesLabel.append(' はい');

        const noLabel = document.createElement('label');
        noLabel.className = 'flex items-center cursor-pointer';
        const noInput = document.createElement('input');
        noInput.type = 'radio';
        noInput.name = `q${index}`;
        noInput.value = 'no';
        noInput.required = true;
        noInput.className = 'h-5 w-5 text-blue-600 focus:ring-blue-500 border-gray-300';
        noLabel.appendChild(noInput);
        noLabel.append(' いいえ');
        
        radioContainer.appendChild(yesLabel);
        radioContainer.appendChild(noLabel);
        questionBlock.appendChild(questionText);
        questionBlock.appendChild(radioContainer);
        questionsContainer.appendChild(questionBlock);

        // 進捗更新のためのイベント
        yesInput.addEventListener('change', updateProgress);
        noInput.addEventListener('change', updateProgress);
    });
}

/**
 * フォーム送信処理
 */
async function handleFormSubmit(e) {
    e.preventDefault();
    
    // 全ての質問に回答したかチェック
    const formData = new FormData(checkSheetForm);
    const answers = [];
    let allAnswered = true;
    for (let i = 0; i < questions.length; i++) {
        const answer = formData.get(`q${i}`);
        if (!answer) {
            allAnswered = false;
            break;
        }
        answers.push(answer);
    }

    if (!allAnswered) {
        alert('すべての質問に回答してください。');
        return;
    }

    checkSheetForm.classList.add('hidden');
    loading.classList.remove('hidden');

    // ここにGASのウェブアプリURLを設定
    const GAS_URL = 'https://script.google.com/macros/s/AKfycbwbL8ChGCyWxdKi6lnKZFhZn4Frq-uSxv_8_WV6YqvN90LRec7evhA6INqccz3ufeAaRw/exec';

    try {
        const response = await fetch(GAS_URL, {
            method: 'POST',
            body: JSON.stringify({ answers: answers }),
            headers: {
                'Content-Type': 'text/plain;charset=utf-8', // GASでPOSTを受けるため
            },
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const resultData = await response.json();

        if (resultData.status === 'success') {
            downloadLink.href = resultData.pdfUrl;
            loading.classList.add('hidden');
            result.classList.remove('hidden');
        } else {
            throw new Error(resultData.message || 'PDFの生成に失敗しました。');
        }
    } catch (err) {
        console.error('Error:', err);
        loading.classList.add('hidden');
        error.classList.remove('hidden');
    }
}

/**
 * ページリロード処理
 */
function handleRetry() {
    window.location.reload();
}

/**
 * 診断開始処理
 */
function startDiagnosis() {
    introduction.classList.add('hidden');
    checkSheetForm.classList.remove('hidden');
    progressSection.classList.remove('hidden');
    renderQuestions();
    updateProgress();
}

/**
 * 進捗の更新
 */
function updateProgress() {
    const answeredCount = document.querySelectorAll('#questionsContainer input[type="radio"]:checked').length;
    const total = questions.length;
    const percentage = Math.round((answeredCount / total) * 100);
    progressText.textContent = `${answeredCount} / ${total}`;
    progressBar.style.width = `${percentage}%`;
}

// イベントリスナーの設定
document.addEventListener('DOMContentLoaded', function() {
    startBtn.addEventListener('click', startDiagnosis);
    checkSheetForm.addEventListener('submit', handleFormSubmit);
    retryBtn.addEventListener('click', handleRetry);
});

// HTMLドキュメントが完全に読み込まれてからJavaScriptを実行するための記述
document.addEventListener('DOMContentLoaded', () => {
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

    const startBtn = document.getElementById('startBtn');
    const introduction = document.getElementById('introduction');
    const checkSheetForm = document.getElementById('checkSheetForm');
    const loading = document.getElementById('loading');
    const result = document.getElementById('result');
    const error = document.getElementById('error');
    const errorDetails = document.getElementById('errorDetails');
    const downloadLink = document.getElementById('downloadLink');
    const retryBtn = document.getElementById('retryBtn');

    function renderQuestions() {
        let currentCategory = '';
        let content = '';
        questions.forEach((q, index) => {
            if (q.category !== currentCategory) {
                currentCategory = q.category;
                content += `<h3 class="text-xl font-bold text-gray-800 border-b-2 border-blue-500 pb-2 mb-4">【${currentCategory}】</h3>`;
            }
            content += `
                <div class="p-4 border border-gray-200 rounded-lg">
                    <p class="font-medium text-gray-800 mb-3">Q${index + 1}. ${q.text}</p>
                    <div class="flex items-center space-x-6">
                        <label class="flex items-center cursor-pointer">
                            <input type="radio" name="q${index}" value="yes" required class="h-5 w-5 text-blue-600 focus:ring-blue-500 border-gray-300">
                            <span class="ml-2">はい</span>
                        </label>
                        <label class="flex items-center cursor-pointer">
                            <input type="radio" name="q${index}" value="no" required class="h-5 w-5 text-blue-600 focus:ring-blue-500 border-gray-300">
                            <span class="ml-2">いいえ</span>
                        </label>
                    </div>
                </div>`;
        });
        content += `<button type="submit" id="submitBtn" class="w-full bg-green-600 text-white font-bold py-3 px-6 rounded-lg hover:bg-green-700 transition duration-300 text-lg">診断結果を見る</button>`;
        checkSheetForm.innerHTML = content;
    }

    startBtn.addEventListener('click', () => {
        introduction.classList.add('hidden');
        checkSheetForm.classList.remove('hidden');
        renderQuestions();
    });

    checkSheetForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const formData = new FormData(checkSheetForm);
        const answers = questions.map((_, i) => formData.get(`q${i}`));

        checkSheetForm.classList.add('hidden');
        loading.classList.remove('hidden');

        const GAS_URL = 'https://script.google.com/macros/s/AKfycbzl6MPeMw9Qrdj5aQSz6UpxiatEhVc7_w21sb1Jk42LFEP8V6Atgf6tREhcmAm24s1-Tg/exec';

        try {
            const response = await fetch(GAS_URL, {
                method: 'POST',
                body: JSON.stringify({ answers }),
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            });

            const resultData = await response.json();

            if (resultData.status === 'success') {
                downloadLink.href = resultData.pdfUrl;
                loading.classList.add('hidden');
                result.classList.remove('hidden');
            } else {
                // GASから返された詳細なエラー情報を表示
                errorDetails.textContent = JSON.stringify(resultData.details, null, 2);
                throw new Error(resultData.message || 'PDFの生成に失敗しました。');
            }
        } catch (err) {
            console.error('Fetch Error:', err);
            if (!errorDetails.textContent) {
                errorDetails.textContent = err.toString();
            }
            loading.classList.add('hidden');
            error.classList.remove('hidden');
        }
    });
    
    retryBtn.addEventListener('click', () => window.location.reload());
});
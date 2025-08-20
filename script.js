// 以下の関数をまるごと置き換えてください

checkSheetForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const formData = new FormData(checkSheetForm);
    const answers = questions.map((_, i) => formData.get(`q${i}`));

    checkSheetForm.classList.add('hidden');
    loading.classList.remove('hidden');

    // ★★★ ご自身のGASウェブアプリURLを設定してください ★★★
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
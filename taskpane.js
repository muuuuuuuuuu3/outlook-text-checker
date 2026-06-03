// Outlookアドインの準備ができたら実行
Office.onReady(() => {
  // 特に初期化処理は不要（ボタンはHTMLに直接書いてある）
});

// チェック結果をパネル内のdivに表示する
function showMessage(msg) {
  const resultDiv = document.getElementById("result");
  resultDiv.textContent = msg;
}

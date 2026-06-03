Office.onReady(() => {
  // 送信前チェックボタンを追加
});

// メール本文を取得してチェックする
function checkText() {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Text,
    function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const text = result.value;
        const issues = [];

        // 辞書ベースのチェック
        const typoMap = {
          "よろしくお願いします。": ["よろしくおねがいします", "宜しくお願いします"],
          "お世話になっております。": ["おせわになっております", "お世話になっています"],
          "ご確認ください。": ["ご確認下さい"],
        };

        for (const [correct, wrongs] of Object.entries(typoMap)) {
          for (const wrong of wrongs) {
            if (text.includes(wrong)) {
              issues.push(`「${wrong}」→「${correct}」に修正しましょう`);
            }
          }
        }

        // 結果を表示
        if (issues.length === 0) {
          showMessage("問題は見つかりませんでした！");
        } else {
          showMessage("修正候補:\n" + issues.join("\n"));
        }
      }
    }
  );
}

function showMessage(msg) {
  Office.context.mailbox.item.notificationMessages.addAsync("checker", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: "icon16",
    persistent: true,
  });
}

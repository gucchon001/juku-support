function processEmailsAndSaveAttachments() {
  var threads = GmailApp.search('subject:Gratitude has:attachment');
  var folder = DriveApp.getFolderById(settings.GoogleDriveID);; // GoogleドライブのフォルダIDを設定
  var sheet = SpreadsheetApp.openById(settings.SpreadSheetID).getSheetByName(settings.OCRシート);

  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var attachments = message.getAttachments();
      attachments.forEach(function(attachment) {
        if (/image/.test(attachment.getContentType())) {
          var file = folder.createFile(attachment);
          Utilities.sleep(3000); // OCR処理の完了を待つために少し待機
          var doc = DocumentApp.create(file.getName());
          var body = doc.getBody();
          body.clear();
          body.insertImage(0, file.getBlob());
          doc.saveAndClose();
          Utilities.sleep(10000); // OCR処理の完了を待つためにさらに待機
          var text = DocumentApp.openById(doc.getId()).getBody().getText();

          // 抽出したテキストから特定の情報を抽出
          var info = extractInformation(text);

          // スプレッドシートに情報を記録
          sheet.appendRow([text.slice(0, 50000), info.name, info.bankName, info.branchName, info.accountNumber]); // 文字数制限に注意
          // 不要になったドキュメントの削除
          DriveApp.getFileById(doc.getId()).setTrashed(true);
        }
      });
    });
  });
}


// 抽出したテキストから特定の情報を抽出する関数
function extractInformation(text) {
  var info = {};

  // 名前の抽出（例: "名前: 山田太郎"）
  var namePattern = /名前: ([^\n\r]+)/;
  var nameMatch = text.match(namePattern);
  info.name = nameMatch ? nameMatch[1].trim() : "名前が見つかりません";

  // 銀行名の抽出（例: "銀行名: 三井住友銀行"）
  var bankNamePattern = /銀行名: ([^\n\r]+)/;
  var bankNameMatch = text.match(bankNamePattern);
  info.bankName = bankNameMatch ? bankNameMatch[1].trim() : "銀行名が見つかりません";

  // 支店名の抽出（例: "支店名: 新宿支店"）
  var branchNamePattern = /支店名: ([^\n\r]+)/;
  var branchNameMatch = text.match(branchNamePattern);
  info.branchName = branchNameMatch ? branchNameMatch[1].trim() : "支店名が見つかりません";

  // 口座番号の抽出（例: "口座番号: 1234567"）
  var accountNumberPattern = /口座番号: (\d+)/;
  var accountNumberMatch = text.match(accountNumberPattern);
  info.accountNumber = accountNumberMatch ? accountNumberMatch[1].trim() : "口座番号が見つかりません";

  return info;
}
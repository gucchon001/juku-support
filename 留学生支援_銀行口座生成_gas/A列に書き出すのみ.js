function saveAttachmentsToDriveAndConvertToDoc_old() {
  // Gmailで「Gratitude」を件名に含み、添付ファイルがあるメールを検索
  var threads = GmailApp.search('label:inbox has:attachment subject:Gratitude');
  var folder = DriveApp.getFolderById(settings.GoogleDriveID);; // GoogleドライブのフォルダIDを設定
  var sheet = SpreadsheetApp.openById(settings.SpreadSheetID).getSheetByName(settings.OCRシート);
  //var sheetId = settings.SpreadSheetID; // スプレッドシートのIDを設定
  //var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var attachments = messages[j].getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        // 画像ファイルのみを対象とする
        if (attachment.getContentType().indexOf("image/") == 0) {
          // Googleドライブに添付ファイルを保存
          var file = folder.createFile(attachment);
          // 画像をGoogleドキュメントにOCR変換
          var blob = file.getBlob();
          var resource = {
            title: file.getName(),
            mimeType: file.getMimeType()
          };
          var options = {
            ocr: true
          };
          var docFile = Drive.Files.insert(resource, blob, options);
          var docId = docFile.id;
          // ドキュメントからテキストを抽出
          var doc = DocumentApp.openById(docId);
          var text = doc.getBody().getText();
          // スプレッドシートのA列にOCRテキストを転記
          sheet.appendRow([text]);
        }
      }
    }
  }
}

function processEmailsAndSaveAttachments1() {
  var threads = GmailApp.search('label:inbox has:attachment subject:Gratitude');
  var folder = DriveApp.getFolderById(settings.GoogleDriveID);; // GoogleドライブのフォルダIDを設定
  var sheet = SpreadsheetApp.openById(settings.SpreadSheetID).getSheetByName(settings.OCRシート);
  var existingFiles = getExistingFileNames(folder);

function getExistingFileNames(folder) {
  var files = folder.getFiles();
  var fileNames = [];
  while (files.hasNext()) {
    var file = files.next();
    fileNames.push(file.getName());
  }
  return fileNames;
}

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var attachments = messages[j].getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        if (attachment.getContentType().indexOf("image/") == 0 && !existingFiles.includes(attachment.getName())) {
          // 省略: 添付ファイルをドライブに保存し、OCRを使用してテキストを抽出する処理

          // 抽出したテキストを `text` 変数に格納
          var text = doc.getBody().getText();

         
        }
      }
    }
  }
}

// 指定した名前のフォルダを取得または作成する関数
function getOrCreateFolder(folderName, parentFolderId) {
  var query = `title = '${folderName}' and trashed = false and mimeType = 'application/vnd.google-apps.folder'`;
  var folders = Drive.Files.list({
    q: query,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    corpora: 'allDrives'
  }).items;

  if (folders.length > 0) {
    console.log("既存の'" + folderName + "'フォルダが見つかりました。");
    return folders[0];
  } else {
    console.log("'" + folderName + "'フォルダが見つからなかったため、新しく作成します。");
    return Drive.Files.insert({
      title: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [{ id: parentFolderId }]
    }, null, { supportsAllDrives: true });
  }
}

// 送信者のメールアドレスに対応するフォルダを作成する関数
function createSenderFolder(sender, parentFolderId) {
  var folderName = `${Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss")}_${sender.split('@')[0]}`;
  console.log("送信者'" + sender + "'用のフォルダ'" + folderName + "'を作成します。");
  return Drive.Files.insert({
    title: folderName,
    mimeType: 'application/vnd.google-apps.folder',
    parents: [{ id: parentFolderId }]
  }, null, { supportsAllDrives: true });
}

// 添付ファイルを処理し、OCR処理されたテキストを結合する関数
function processAttachments(attachments, attachmentFolder) {
  var combinedText = "";

  attachments.forEach(function(attachment, index) {
    if (attachment.getContentType().indexOf("image/") === 0) {
      var suffix = attachments.length > 1 ? `_${index + 1}` : '';
      var formattedName = `${attachment.getName().split('.').slice(0, -1).join('.')}${suffix}.${attachment.getName().split('.').pop()}`;
      
      // 添付ファイルをGoogleドライブに保存
      var newFile = DriveApp.createFile(attachment.copyBlob());
      newFile.setName(formattedName);
      var attachmentFolderDrive = DriveApp.getFolderById(attachmentFolder.id);
      attachmentFolderDrive.addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile);
      console.log("添付ファイル'" + formattedName + "'を保存しました。");
      
      // OCR処理を実行
      var resource = {
        title: formattedName,
        mimeType: attachment.getContentType()
      };
      var options = {
        ocr: true,
        ocrLanguage: 'ja'
      };
      
      var docFile = Drive.Files.insert(resource, newFile.getBlob(), options);
      var doc = DocumentApp.openById(docFile.id);
      var text = doc.getBody().getText();
      combinedText += text + "\n\n";
      Drive.Files.remove(docFile.id);
      console.log("OCR処理が完了しました。");
    }
  });

  return combinedText;
}

// 結合されたテキストを含むドキュメントを作成する関数
function createCombinedDoc(folderTitle, combinedText, parentFolderId) {
  var combinedDocName = `${folderTitle}_OCR.docx`;
  var combinedDocFile = DocumentApp.create(combinedDocName);
  combinedDocFile.getBody().setText(combinedText);
  combinedDocFile.saveAndClose();
  console.log("結合されたテキストを含むドキュメント'" + combinedDocName + "'を作成しました。");

  var combinedDocFileId = DriveApp.getFileById(combinedDocFile.getId()).moveTo(DriveApp.getFolderById(parentFolderId));
  return combinedDocFileId;
}

// スプレッドシートにデータを書き出す関数
function extractAndLogData(dataForProcessing) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var exportSheet = spreadsheet.getSheetByName(settings.書き出しシート);

  var headers = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn()).getValues()[0];

  // OCRから得たデータを使用して行のデータを処理
  dataForProcessing.forEach(function(data, index) {
    // 出力行のデータ設定
    var outputRow = headers.map(() => ''); // ヘッダー数に応じた空の配列を作成
    outputRow[headers.indexOf('メールアドレス')] = data.email;
    outputRow[headers.indexOf('元画像フォルダ')] = data.originalFolderUrl;
    outputRow[headers.indexOf('元OCRテキスト')] = data.ocrFileUrl;
    outputRow[headers.indexOf('転記内容')] = data.text;
    outputRow[headers.indexOf('タイムスタンプ')] = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");

    // パースしたデータを適切な列に配置
    Object.keys(data.parsedData).forEach(function(key) {
      var columnIndex = headers.indexOf(key);
      if (columnIndex > -1) {
        outputRow[columnIndex] = data.parsedData[key];
      }
    });

    exportSheet.appendRow(outputRow);
    console.log("スプレッドシートに転記しました。");
  });
}

// OpenAIからの応答をパースし、データオブジェクトに変換する関数
function parseResponse(response) {
  var data = {};
  try {
    // API応答からcontent部分を取得し、各フィールドを解析
    var content = response.split('\n');
    content.forEach(function(item) {
      // 先頭の数字と空白を除去してから分割
      item = item.trim().replace(/^\d+\s*/, '');
      var parts = item.split(':');
      if (parts.length === 2) {
        var key = parts[0].trim().replace(/^'|'$/g, ''); // キーから単一引用符を削除
        var value = parts[1].trim().replace(/^'|'$/g, ''); // 値から単一引用符を削除
        value = value.replace(/^"|"$/g, ''); // 値から二重引用符を削除
        value = value.replace(/^"|"$/g, ''); // 値から二重引用符を再度削除（無駄な二重引用符対応）
        value = value.replace(/,$/g, ''); // 値の末尾からカンマを削除
        data[key] = value;
      }
    });
  } catch (e) {
    console.error("応答解析中にエラーが発生しました: " + e.toString());
  }
  return data;
}

// パースされたデータを加工する関数
function processData(data) {
  // ダブルクォーテーションを削除
  Object.keys(data).forEach(function(key) {
    data[key] = data[key].replace(/"/g, '');
  });

  // 氏名（カナ）と銀行名（カナ）を半角カナに変換
  if (data['氏名（カナ）']) {
    data['氏名（カナ）'] = data['氏名（カナ）'].replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
  }
  if (data['銀行名（カナ）']) {
    data['銀行名（カナ）'] = data['銀行名（カナ）'].replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
  }

  return data;
}

function HIRAGANA_TO_KATAKANA(range) {
  var values = range.flat();
  var results = values.map(function(value) {
    if (typeof value !== 'string') {
      return value;
    }
    return value.replace(/[\u3041-\u3096]/g, function(match) {
      var chr = match.charCodeAt(0) + 0x60;
      return String.fromCharCode(chr);
    });
  });
  return results;
}

// スプレッドシートからプロンプトを取得する関数
function getPromptFromSheet(promptSheetName) {
  var promptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(promptSheetName);
  var prompt = promptSheet.getRange('G3').getValue();
  return prompt;
}
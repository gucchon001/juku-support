function saveAttachmentsToDriveAndConvertToDoc() {
 console.log("共有ドライブのフォルダIDを取得しています...");
 var folder;
 try {
   folder = Drive.Files.get(settings.GoogleDrivefolderID, { supportsAllDrives: true });
   console.log(folder)
   console.log("共有ドライブのフォルダが見つかりました: " + folder.title);
 } catch (e) {
   console.error("共有ドライブのフォルダの取得中にエラーが発生しました: " + e.toString());
   return;
 }

 var originalDataFolderName = '元データ';
 var originalDataFolder;
 var dataForProcessing = [];

 console.log("'元データ'フォルダを検索しています...");
 var query = `title = '${originalDataFolderName}' and trashed = false and mimeType = 'application/vnd.google-apps.folder'`;
  var folders = Drive.Files.list({
    q: query,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    corpora: 'allDrives'
  }).items;
  console.log("検索結果: ", folders);
 if (folders.length > 0) {
   originalDataFolder = folders[0];
   console.log("'元データ'フォルダが見つかりました: " + originalDataFolder.title);
 } else {
   console.log("'元データ'フォルダが見つかりませんでした。新しく作成します...");
   originalDataFolder = Drive.Files.insert({
     title: originalDataFolderName,
     mimeType: 'application/vnd.google-apps.folder',
     parents: [{ id: folder.id }]
   }, null, { supportsAllDrives: true });
   console.log("'元データ'フォルダが作成されました: " + originalDataFolder.title);
 }

 // 検索条件に一致するメールスレッドを検索
 var threads = GmailApp.search('label:inbox has:attachment subject:Gratitude');

 // 各メールスレッドを処理
 for (var i = 0; i < threads.length; i++) {
   var messages = threads[i].getMessages();
   for (var j = 0; j < messages.length; j++) {
     var attachments = messages[j].getAttachments();
     var sender = messages[j].getFrom().match(/<(.+?)>/)[1]; // 送信者のメールアドレスを抽出
     attachments.forEach(function(attachment, k) {
       // 添付ファイルが画像の場合のみ処理
       if (attachment.getContentType().indexOf("image/") === 0) {
         // ファイル名のフォーマットを設定
         var suffix = attachments.length > 1 ? `_${k + 1}` : '';
         var formattedName = `${Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss")}_${sender.split('@')[0]}${suffix}.${attachment.getName().split('.').pop()}`;
         console.log("添付ファイルを'元データ'フォルダに保存しています...");
         var newFile = Drive.Files.insert({
           title: formattedName,
           mimeType: attachment.getContentType(),
           parents: [{ id: originalDataFolder.id }]
         }, attachment, { supportsAllDrives: true });
         console.log("添付ファイルが'元データ'フォルダに保存されました: " + newFile.title);
         var fileId = newFile.id;
         console.log("添付ファイルのIDが取得されました: " + fileId);
         var resource = {
           title: formattedName,
           mimeType: attachment.getContentType()
         };
         var options = {
           ocr: true,
           ocrLanguage: 'ja'
         };
         console.log("OCR処理を開始しています...");
         var docFile = Drive.Files.insert(resource, attachment, options);
         console.log("OCR処理が完了し、新しいファイルが作成されました: " + docFile.title);
         var docId = docFile.id;
         console.log("OCR処理されたファイルのIDが取得されました: " + docId);
         var patchOptions = {
          supportsAllDrives: true,
          fields: 'id, parents'
        };
        Drive.Files.patch({ parents: [{ id: folder.id }] }, docId, patchOptions);
        console.log("OCR処理されたファイルが'銀行口座'フォルダに移動されました: " + docFile.title);
         console.log("OCR処理されたファイルが'銀行口座'フォルダに移動されました: " + docFile.title);
         // ドキュメントからテキストを抽出
         console.log("OCR処理されたファイルからテキストを抽出しています...");
         var doc = DocumentApp.openById(docId);
         var text = doc.getBody().getText();
         console.log("OCR処理されたファイルからテキストが抽出されました: " + text.slice(0, 50) + "...");
         // 処理したデータを配列に追加
         dataForProcessing.push({
           email: sender,
           originalFileId: fileId,
           originalFileUrl: `https://drive.google.com/open?id=${fileId}`,
           ocrFileId: docId,
           ocrFileUrl: `https://drive.google.com/open?id=${docId}`,
           text: text
         });
       }
     });
   }
 }
 // 抽出したデータを後続の処理に渡す
 console.log("抽出されたデータ: ", dataForProcessing);
 extractAndLogData(dataForProcessing);
}

function extractAndLogData(dataForProcessing) {
 console.log("スプレッドシートを取得中...");
 var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 
 console.log("出力シートを読み込み中...");
 var exportSheet = spreadsheet.getSheetByName(settings.書き出しシート);

 console.log("出力シートのヘッダーを取得中...");
 var headers = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn()).getValues()[0];
 console.log("出力シートのヘッダー: " + headers.join(", "));

 // OCRから得たデータを使用して行のデータを処理
 dataForProcessing.forEach(function(data, index) {
   console.log(`行 ${index + 2} のデータを処理中: テキスト: ${data.text}, ドキュメントID: ${data.docId}`);

   console.log(`ChatGPTにデータを送信中: ${data.text}`);
   var response = callOpenAI(data.text); // APIコール
   console.log(response);
   var parsedData = parseResponse(response); // 応答の解析

   // 出力行のデータ設定
   var outputRow = headers.map(() => ''); // ヘッダー数に応じた空の配列を作成
   outputRow[headers.indexOf('メールアドレス')] = data.email;
   outputRow[headers.indexOf('元画像')] = data.originalFileUrl;
   outputRow[headers.indexOf('元OCRテキスト')] = data.ocrFileUrl;
   outputRow[headers.indexOf('転記内容')] = data.text;
   outputRow[headers.indexOf('タイムスタンプ')] = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");

   // パースしたデータを適切な列に配置
   Object.keys(parsedData).forEach(function(key) {
     var columnIndex = headers.indexOf(key);
     if (columnIndex > -1) {
       outputRow[columnIndex] = parsedData[key];
     }
   });

   exportSheet.appendRow(outputRow);
   console.log(`行が書き込まれました: ${outputRow.join(", ")}`);
 });
}

function extractAndLogData(dataForProcessing) {
  console.log("スプレッドシートを取得中...");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log("出力シートを読み込み中...");
  var exportSheet = spreadsheet.getSheetByName(settings.書き出しシート);

  console.log("出力シートのヘッダーを取得中...");
  var headers = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn()).getValues()[0];
  console.log("出力シートのヘッダー: " + headers.join(", "));

  // OCRから得たデータを使用して行のデータを処理
  dataForProcessing.forEach(function(data, index) {
    console.log(`行 ${index + 2} のデータを処理中: テキスト: ${data.text}, ドキュメントID: ${data.docId}`);

    console.log(`ChatGPTにデータを送信中: ${data.text}`);
    var response = callOpenAI(data.text); // APIコール
    console.log(response);
    var parsedData = parseResponse(response); // 応答の解析

    // 出力行のデータ設定
    var outputRow = headers.map(() => ''); // ヘッダー数に応じた空の配列を作成
    outputRow[headers.indexOf('メールアドレス')] = data.email;
    outputRow[headers.indexOf('元画像')] = data.originalFileUrl;
    outputRow[headers.indexOf('元OCRテキスト')] = data.ocrFileUrl;
    outputRow[headers.indexOf('転記内容')] = data.text;
    outputRow[headers.indexOf('タイムスタンプ')] = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");

    // パースしたデータを適切な列に配置
    Object.keys(parsedData).forEach(function(key) {
      var columnIndex = headers.indexOf(key);
      if (columnIndex > -1) {
        outputRow[columnIndex] = parsedData[key];
      }
    });

    exportSheet.appendRow(outputRow);
    console.log(`行が書き込まれました: ${outputRow.join(", ")}`);
  });
}
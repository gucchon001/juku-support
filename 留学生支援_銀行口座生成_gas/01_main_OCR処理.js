// メインの処理
function saveAttachmentsToDriveAndConvertToDoc() {
  try {
    // 共有ドライブのフォルダを取得
    var folder = Drive.Files.get(settings.GoogleDrivefolderID, { supportsAllDrives: true });
    console.log("共有ドライブのフォルダが見つかりました: " + folder.title);
  } catch (e) {
    console.error("共有ドライブのフォルダの取得中にエラーが発生しました: " + e.toString());
    return;
  }

  // 元データフォルダを取得または作成
  var originalDataFolderName = '元データ';
  var originalDataFolder = getOrCreateFolder(originalDataFolderName, folder.id);
  var dataForProcessing = [];

  // 検索条件に一致するメールスレッドを検索
  var threads = GmailApp.search('label:inbox has:attachment subject:Gratitude');

  // 送信者のメールアドレスと対応するフォルダを格納するオブジェクト
  var senderFolders = {};

  // 各メールスレッドを処理
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var attachments = message.getAttachments();
      var sender = message.getFrom().match(/<(.+?)>/)[1]; // 送信者のメールアドレスを抽出
      console.log("処理中のメールアドレス: " + sender);

      // 送信者のメールアドレスに対応するフォルダを取得または作成
      var attachmentFolder = senderFolders[sender] || createSenderFolder(sender, originalDataFolder.id);
      senderFolders[sender] = attachmentFolder;

      // 添付ファイルを処理し、OCR処理されたテキストを結合
      var combinedText = processAttachments(attachments, attachmentFolder);

      // 結合されたテキストを含むドキュメントを作成
      var combinedDocFileId = createCombinedDoc(attachmentFolder.title, combinedText, folder.id);

      console.log("OCR処理されたテキスト: " + combinedText);

      // OpenAI APIを呼び出してテキストを変換
      var response = callOpenAI(combinedText);
      var parsedData = parseResponse(response);
      //console.log("OpenAIによる変換前のテキスト: " + combinedText);
      console.log("OpenAIによる変換後のテキスト: " + response);

      // 処理したデータを配列に追加
      dataForProcessing.push({
        email: sender,
        originalFolderId: attachmentFolder.id,
        originalFolderUrl: `https://drive.google.com/drive/folders/${attachmentFolder.id}`,
        ocrFileId: combinedDocFileId.getId(),
        ocrFileUrl: `https://docs.google.com/document/d/${combinedDocFileId.getId()}/edit`,
        text: combinedText,
        parsedData: parsedData
      });
    });
  });

  // 処理したデータをスプレッドシートに書き出す
  extractAndLogData(dataForProcessing);
}
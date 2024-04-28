

//OCRデータを元にスプレッドシートにデータ転記
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

//ChatGPTの処理
function callOpenAI(text, headers) {
  console.log("OpenAI APIに接続中...");
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  var apiURL = 'https://api.openai.com/v1/chat/completions';

  var instruction = getPromptFromSheet(settings.プロンプトシート);
  console.log("抽出指示: " + instruction);

  var payload = {
    model: GPT_MODEL,
    messages: [{
      role: "system",
      content: instruction
    }, {
      role: "user",
      content: text
    }],
    max_tokens: 150,
    temperature: 0
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(apiURL, options);
  var responseText = response.getContentText();
  var jsonResponse = JSON.parse(responseText);

  if (jsonResponse.choices && jsonResponse.choices.length > 0) {
    console.log("APIからの有効な応答を受信しました。");
    return jsonResponse.choices[0].message.content;
  } else {
    console.error("OpenAI APIからの応答に問題があります: ", jsonResponse);
    return "Error: No valid response from API. See logs for details.";
  }
}

//ChatGPTからのデータを正規表現処理して、データ形式を合わせる処理
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
        
        // 値をさらに分割して個別のフィールドに割り当てる
        var fields = value.split(/[\s"]+/);
        if (fields.length >= 7) {
          data['氏名'] = fields[0];
          data['氏名（カナ）'] = fields[1];
          data['銀行名'] = fields[2];
          data['銀行名（カナ）'] = fields[3];
          data['銀行コード'] = fields[4];
          data['口座番号'] = fields[5];
          data['店番号'] = fields[6];
          if (fields.length > 7) {
            data['日付'] = fields[7];
          }
        } else {
          data[key] = value;
        }
      }
    });
  } catch (e) {
    console.error("応答解析中にエラーが発生しました: " + e.toString());
  }
  console.log("解析されたデータ: ", data);
  return data;
}
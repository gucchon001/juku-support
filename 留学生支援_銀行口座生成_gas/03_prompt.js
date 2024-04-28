function getPromptFromSheet() {
  var promptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.プロンプトシート);
  var prompt = promptSheet.getRange('G3').getValue();
  console.log("スプレッドシートから取得されたプロンプト: " + prompt);
  return prompt;
}

function translateText() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.プロンプトシート);
  var textToTranslate = sheet.getRange('B3').getValue();

  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  var apiURL = 'https://api.openai.com/v1/chat/completions';
  var headers = {
    "Content-Type": "application/json",
    Authorization: "Bearer " + apiKey
  };

  var messages = [
    {
      role: "system",
      content: "You are a helpful assistant that translates Japanese to English. When translating, please keep any text enclosed in the following quotation marks as is, without translation:\n\n- 「」\n- \"\"\n- ''\n\nFor example, if the Japanese text contains 「名前」, \"名前\", or '名前', those should remain unchanged in the English translation. Please step by step"
    },
    {
      role: "user",
      content: "Translate the following Japanese text to English:\n\n" + textToTranslate
    }
  ];

  var payload = JSON.stringify({
    model: "gpt-4",
    messages: messages,
    max_tokens: 500,
    temperature: 0.3
  });

  console.log(payload); // APIに送信されるペイロードを出力

  var options = {
    method: 'post',
    headers: headers,
    payload: payload,
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(apiURL, options);
    console.log(response.getContentText()); // APIからの応答の内容を出力
    var jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.choices && jsonResponse.choices.length > 0) {
      var translatedText = jsonResponse.choices[0].message.content.trim();
      sheet.getRange('G3').setValue(translatedText);
      console.log("Translated text: " + translatedText);
    } else {
      console.error("Translation failed or no valid response from API.");
    }
  } catch (error) {
    console.error("Error occurred during API request:");
    console.error(error.message);
    console.error(error.stack);
  }
}
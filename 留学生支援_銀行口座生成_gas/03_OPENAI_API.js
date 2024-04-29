// OpenAIにテキストを送信し、応答を取得する関数
function callOpenAI(text) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  var apiURL = 'https://api.openai.com/v1/chat/completions';

  var instruction = getPromptFromSheet(settings.プロンプトシート);

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
    return jsonResponse.choices[0].message.content;
  } else {
    console.error("OpenAI APIからの応答に問題があります: ", jsonResponse);
    return "Error: No valid response from API. See logs for details.";
  }
}
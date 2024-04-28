//設定
var API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');;
var GPT_MODEL = "gpt-3.5-turbo"

//設定シート
function getSettings(sheet) {
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

//設定シート呼び出し＝＞グローバル変数に設定
var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('settings');
var settings = getSettings(settingsSheet);

//メニュー表示
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('実行メニュー')
    .addItem('実行', 'saveAttachmentsToDriveAndConvertToDoc')
    .addToUi();
}
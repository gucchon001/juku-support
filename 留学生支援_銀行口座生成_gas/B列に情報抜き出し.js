function extractAndWriteBankingDetails() {
  var sheet = SpreadsheetApp.openById(settings.SpreadSheetID).getSheetByName(settings.OCRシート);
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, 1, lastRow, 1); // A列のデータを取得
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    var text = values[i][0]; // A列の各行のテキスト
    // 銀行名、支店名、口座番号、口座名義を抽出する正規表現パターン（例）
    var bankNamePattern = /銀行名：(.+)/;
    var branchNamePattern = /【?店番】?\s*(\d+)/;
    var accountNumberPattern = /【?口座番号】?\s*(\d+)/;
    var accountHolderPattern = /(.+?)\s*(様|さま)/;

    // テキストから情報を抽出
    var bankName = text.match(bankNamePattern) ? text.match(bankNamePattern)[1] : "";
    var branchName = text.match(branchNamePattern) ? text.match(branchNamePattern)[1] : "";
    var accountNumber = text.match(accountNumberPattern) ? text.match(accountNumberPattern)[1] : "";
    var accountHolder = text.match(accountHolderPattern) ? text.match(accountHolderPattern)[1] : "";

    // B列に抽出した情報を入力（銀行名、支店名、口座番号、口座名義を結合）
    var bankingDetails = bankName + ", " + branchName + ", " + accountNumber + ", " + accountHolder;
    sheet.getRange(i + 1, 2).setValue(bankingDetails); // 2列目（B列）にセット
  }
}

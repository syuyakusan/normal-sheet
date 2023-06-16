/**
 * LINE公式アカウントのQRコードを表示させる関数
 */
function showQR() {
  let output = HtmlService.createTemplateFromFile('index');
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(400).setHeight(400);
  ss.show(html);    //メッセージボックスとしてを表示する
}

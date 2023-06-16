/**
 * 「ホーム」シートのデータから各メンバーのシートをつくる関数
 */
function createSheets(){

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = spreadsheet.getSheetByName('ホーム');
  const sheetNum = spreadsheet.getNumSheets();  //今のシート数
  const reqNum = setSheet.getRange('F6').getValue();  //メンバー人数
  let difNum = sheetNum-reqNum-2;


  while (difNum > 0) {  //シートが多い場合は削除
    const trashSheet = spreadsheet.getSheets()[2];
    spreadsheet.deleteSheet(trashSheet);
    difNum = difNum - 1;
  }

  for (j=0;j<(difNum*(-1));j++){  //シートが少ない場合は追加
      spreadsheet.getSheets()[sheetNum-1].copyTo(spreadsheet);
    }

  //この時点でシート数はreqNum+2個
  for(let i=0;i<reqNum;i++){  //シート名を初期化
    const sheet = spreadsheet.getSheets()[i+2];
    sheet.setName(i+1);
    }

  for(i=0;i<reqNum;i++){
    //シート名変更
    const sheet = spreadsheet.getSheets()[i+2];
    const range = setSheet.getRange(i+8,6,1,1);
    const name = range.getValue();
    sheet.setName(name);

    //セルの初期化
    const rows = sheet.getLastRow(); //行番号=行数
    const columns = sheet.getLastColumn();
    const tableRange = sheet.getRange(4,2,rows-3,columns-1) //各シートの初期化
    tableRange.clearContent();
    tableRange.clearFormat();
    tableRange.clearNote();
    tableRange.setHorizontalAlignment("center"); 


    //条件付き書式の設定
    sheet.clearConditionalFormatRules();
    const ruleGreen = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('○')
      .setBackground("#b7e1cd") //セル背景を設定（緑）
      .setRanges([tableRange])
      .build();

    const ruleYellow = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('△')
      .setBackground("#fce8b2") //セル背景を設定（黄）
      .setRanges([tableRange])
      .build();
    
    const ruleRed = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('×')
      .setBackground("#f4c7c3") //セル背景を設定（赤）
      .setRanges([tableRange])
      .build();

    const rules =[ruleGreen,ruleYellow,ruleRed];
    sheet.setConditionalFormatRules(rules);

  }

}
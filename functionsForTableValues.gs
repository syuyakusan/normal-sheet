/**
 * 各シートの一括選択のスペースで変更が加えられたとき、その日付の前セルを同じ値にする関数
 */
function overwriteAllColumnData() {

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const changedCell = sheet.getCurrentCell();
  const changedRow = changedCell.getRow()
  const changedColumn = changedCell.getColumn();

  if(sheet.getSheetName() != 'ホーム' && sheet.getSheetName() != '集約' && changedRow==4){
    const value = changedCell.getDisplayValue();
    const changeRange = sheet.getRange(5,changedColumn,31,1);
    changeRange.setValue(value);
  }

}

/**
 * 個人シートがどこまで入力されているかを調べる関数
 * @param name {String} メンバー名(かつシート名)
 * preturns lastDate {Date} 入力最終日 
 */
function checkAnswerStatus(name){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(name);
  const range = sheet.getRange(5,2,31,74);
  const lastColumn = getLastColumnInRange(range);
  const lastDate = sheet.getRange(2,lastColumn,1,1).getValue();

  return lastDate;
}

/**
 * 回答状況をシートに入力する関数
 */
function setAnswerStatus(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = spreadSheet.getSheetByName('ホーム');
  const reqNum = setSheet.getRange('F6').getValue();
  const memberList = setSheet.getRange(8,6,reqNum,1).getValues().flat();
  let outputArray = [];
  for (name of memberList){
    let lastDate = checkAnswerStatus(name);
    if(lastDate !== "日付"){
      Utilities.formatDate(lastDate, "Asia/Tokyo", "MM/dd");
    }else{
      lastDate = "";
    }
    outputArray.push([lastDate]);
  }
  setSheet.getRange(17,11,reqNum,1).setValues(outputArray);
}
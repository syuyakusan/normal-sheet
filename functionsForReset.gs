/**
 * すべて初期化する関数
 */
function reset() {

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = spreadSheet.getSheetByName('ホーム');
  const sumSheet = spreadSheet.getSheetByName('集約');


  // Google Formsの削除
  const formId = setSheet.getRange('N11').getValue();
  if(formId !== ""){
    const trahFile = DriveApp.getFileById(formId);
    trahFile.setTrashed(true);
  }

  // ホームシートの初期化
  const defaultNames = ['名前1','名前2','名前3','名前4','名前5','名前6','名前7'];
  for (i=0;i<7;i++) {
    setSheet.getRange(8+i,6,1,1).setValue(defaultNames[i]);
  }

  setSheet.getRange('F6').setValue(1);
  createSheets();

  const rows = sumSheet.getLastRow(); //行番号=行数
  const columns = sumSheet.getLastColumn();
  const tableRange = sumSheet.getRange(4,2,rows-3,columns-1) //集約シートの初期化
  tableRange.setValue(0);
  tableRange.clearFormat();
  tableRange.clearNote();
  tableRange.setHorizontalAlignment("center"); 

  setSheet.getRange('F7').setValue('2023/01/01');
  setSheet.getRange('K6').setValue('入力してください');
  setSheet.getRange('K7').setValue('');
  setSheet.getRange('K8').setValue('');
  setSheet.getRange('K12').setValue('');
  setSheet.getRange('K13').setValue('');
  setSheet.getRange('N11').setValue('');
  setSheet.getRange('F17').setValue('');
  setSheet.getRange('F18').setValue('');
  setSheet.getRange(17,11,7,1).clearContent();

  setSheet.getRange('B11').setFontColor('#f0f2f2');
  setSheet.getRange('B14').setFontColor('#f0f2f2');
  setSheet.getRange('J6').setFontColor('#484848');
  setSheet.getRange(6,5,3,1).setFontColor('#484848');
  setSheet.getRange('C11').setFontColor('#f0f2f2');
  setSheet.getRange('C14').setFontColor('#f0f2f2');

  // トリガーの削除
  deleteAllTrigger();

  // スクリプトプロパティの削除
  PropertiesService.getScriptProperties().deleteAllProperties();


}

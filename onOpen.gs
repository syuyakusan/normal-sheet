/**
 * スプレッドシート起動時に実行する関数
 */
function onOpen() {

  const setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ホーム");
  setSheet.getRange("O6").setValue(new Date());

  try{
    statusFunc('処理中...');
    showProcess('googleフォームの回答を反映中...');
    summarizeForms();

    statusFunc('処理中...');
    showProcess('集約結果をテキスト化...');
    setAnswerStatus();
  }catch(e){
    statusFunc('エラー');
    showProcess('やり直してください');
    return;
  }


  statusFunc('完了');
  showProcess('');

}


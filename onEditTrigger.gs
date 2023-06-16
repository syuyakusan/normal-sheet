/**
 * onEditトリガーに設定する関数
 * ※onEdit関数ではないのでトリガー設定が必要
 */
function onEditTriggerFunction() {
  
  statusFunc('処理中...');
  showProcess('一括選択を反映...');

  overwriteAllColumnData();

  showProcess('チェックボックスを確認...');

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = spreadSheet.getSheetByName('ホーム');

  const range1 = setSheet.getRange('F2');
  const range2 = setSheet.getRange('L2');
  const range3 = setSheet.getRange('J2');

  if (range1.getValue() == true) {
    let allow = Browser.msgBox('日程を集約します。\n※少し時間がかかる場合があります', Browser.Buttons.OK_CANCEL)
    if (allow === 'cancel') {
      statusFunc('完了');
      showProcess('');
      range1.setValue(false);
      return;
    } 

    statusFunc('処理中...');
    showProcess('全シートを集約...');
    try{
      checkDiff();
    }catch(e){
      Browser.msgBox('集約に失敗しました。'+e);
    }

    statusFunc('処理中...');
    showProcess('googleカレンダーの予定を反映...');
    try{
      checkCalendar();
    }catch(e){
      Browser.msgBox('カレンダーの予定の反映に失敗しました。'+e);
    }

    statusFunc('処理中...');
    showProcess('集約結果をテキスト化...');
    try{
      setSharedResultText();
    }catch(e){
      Browser.msgBox('集約結果のテキスト化に失敗しました。'+e);
    }


    Browser.msgBox("集約が完了しました");
    range1.setValue(false);
  }


  if (range2.getValue() == true) {
    let allow = Browser.msgBox('すべての情報が初期化されます。よろしいですか？', Browser.Buttons.OK_CANCEL)
    if (allow === 'cancel') {
      statusFunc('完了');
      showProcess('');
      range2.setValue(false);
      return;
    } 

    statusFunc('処理中...');
    showProcess('初期化中...');
    reset();
    Browser.msgBox('初期化が完了しました。');
    range2.setValue(false);
  }

  if (range3.getValue() == true) {
    let allow = Browser.msgBox('シートの表示期間を「開始日」セルの日付からに変更します。※今までの回答情報は削除されます。', Browser.Buttons.OK_CANCEL)
    if (allow === 'cancel') {
      statusFunc('完了');
      showProcess('');
      range3.setValue(false);
      return;
    } 

    statusFunc('処理中...');
    showProcess('表示期間を変更中...');
    createSheets();
    Browser.msgBox('表示期間を変更しました。');
    range3.setValue(false);
  }


  statusFunc('完了');
  showProcess('');

}




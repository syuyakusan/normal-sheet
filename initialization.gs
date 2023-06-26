/**
 * 初期設定をする関数1stStep
 * 「初期設定」ボタンに割り当てる
 * トリガーを設置した後、シート基本情報の入力を誘導する
 */
function initializeFunc1(){
  Browser.msgBox('初期設定を開始します。');
  
  // トリガーを設置
  try{
    setOnEditTrigger();
    setOnOpenTrigger();
  }catch(e){
    Browser.msgBox('トリガーの設定に失敗しました'+e);
    statusFunc('エラー');
    showProcess('やり直してください');
    return;
  }

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  // 共有設定を変更
  const access = DriveApp.Access.ANYONE_WITH_LINK;
  const permission = DriveApp.Permission.EDIT;
  DriveApp.getFileById(spreadSheet.getId()).setSharing(access, permission);

  // 基本情報を入力させるために文字色を赤に
  const setSheet = spreadSheet.getSheetByName("ホーム");

  // カレンダーIDを設定
  try{
    const calendarIdList = setSheet.getRange(14,14,5,1).getValues().flat();
    const calendarNameList = ['privateCalenderId','ensembleCalenderId','meetingCalenderId','otherCalenderId','companyCalenderId'];
    for(i=0;i<calendarNameList.length;i++){
      PropertiesService.getScriptProperties().setProperty(calendarNameList[i],calendarIdList[i]);
    }
  }catch(e){
    Browser.msgBox('カレンダーIDの設定に失敗しました'+e);
    statusFunc('エラー');
    showProcess('やり直してください');
    return;
  }

  setSheet.getRange(6,5,3,1).setFontColor('red');
  setSheet.getRange('B11').setFontColor('red');
  setSheet.getRange('B14').setFontColor('#484848');
  setSheet.getRange('C11').setFontColor('red');

  let allow = Browser.msgBox('赤文字の項目を入力し、「シートをつくる」ボタンを押して下さい', Browser.Buttons.OK_CANCEL)
    if (allow === 'cancel') {
      setSheet.getRange(6,5,3,1).setFontColor('#484848');
      setSheet.getRange('B11').setFontColor('#484848');
      setSheet.getRange('C11').setFontColor('#f0f2f2');
      showStatus.done();
      return;
    } 
    // showStatus.processing('赤文字の項目を入力後、「シートをつくる」ボタンを押して下さい。');
  // ユーザーは「シートをつくる」ボタンを押しinitializeFunc2へ
}

/**
 * 初期設定をする関数2ndStep
 * 「シートをつくる」ボタンに割り当てる
 * フォーム用の情報の入力を誘導する
 */
function initializeFunc2(){

  // シートを作成
  let allow = Browser.msgBox('シートを作成します。シートの情報は初期化されます。', Browser.Buttons.OK_CANCEL)
    if (allow === 'cancel') {
      statusFunc('完了');
      showProcess('');
      return;
    }

  statusFunc('処理中...');
  showProcess('シートを作成...');


  try{
    createSheets();
  }catch(e){
    Browser.msgBox('シートの作成に失敗しました'+e);
    statusFunc('エラー');
    showProcess('やり直してください');
    return;
  }

  // 変更済みの項目を赤文字から黒文字へ
  const setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ホーム");
  const reqNum = setSheet.getRange('F6').getValue();  //メンバー人数
  // メンバー以外の名前を削除
  setSheet.getRange(8+reqNum,6,7-reqNum,1).clearContent();
  setSheet.getRange(6,5,3,1).setFontColor('#484848');
  setSheet.getRange('B11').setFontColor('#484848');
  setSheet.getRange('C11').setFontColor('#f0f2f2');

  // Google Formsに関する情報を入力させるために文字色を赤に(ポップアップなし)
  setSheet.getRange('J6').setFontColor('red');
  setSheet.getRange('B14').setFontColor('red');
  setSheet.getRange('C14').setFontColor('red');

  allow = Browser.msgBox('赤文字の項目を入力し、「フォームをつくる」ボタンを押して下さい。', Browser.Buttons.OK_CANCEL)
    if (allow === 'cancel') {
      setSheet.getRange('J6').setFontColor('#484848');
      setSheet.getRange('B14').setFontColor('#484848');
      setSheet.getRange('C14').setFontColor('#f0f2f2');
      showStatus.done();
      return;
    } 
    // showStatus.processing('赤文字の項目を入力し、「フォームをつくる」ボタンを押して下さい。');
  statusFunc('完了');
  showProcess('');
}


/**
 * 初期設定をする関数3rdStep
 * 「フォームをつくる」ボタンに割り当てる
 */
function initializeFunc3(){
  // 変更済みの項目を赤文字から黒文字へ
  const setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ホーム");
  setSheet.getRange('J6').setFontColor('#484848');
  setSheet.getRange('B14').setFontColor('#484848');
  setSheet.getRange('C14').setFontColor('#f0f2f2');

  const userResponse = Browser.msgBox(`Googleフォームを作成します。`,Browser.Buttons.OK_CANCEL);
  if(userResponse == "cancel"){
  Browser.msgBox("処理を終了します。");
  return;
  }
  statusFunc('処理中...');
  showProcess('googleフォームを作成...');
  // Formの作成
  try{
    const formTitle = createForm();
    Browser.msgBox('「' + formTitle + '」のフォームを作成しました');
  }catch(e){
    Browser.msgBox('Googleフォームの作成に失敗しました。'+e);
    statusFunc('エラー');
    showProcess('やり直してください');
    return;
  }

  setSharedLinkText();
  
  // メッセージボックスの表示
  Browser.msgBox('初期設定が完了しました。');
  setSheet.getRange('B11').setFontColor('#f0f2f2');
  setSheet.getRange('B14').setFontColor('#f0f2f2');
  statusFunc('完了');
  showProcess('');
}
/**
 * onEditトリガーを設置する関数
 */
function setOnEditTrigger(){
 //現在の設定されているトリガーを取得
  triggers = ScriptApp.getProjectTriggers();

  for (i = 0; i < triggers.length; i++) {
   if (triggers[i].getHandlerFunction() == 'onEditTriggerFunction') {
     Logger.log('setOnEditTrigger: トリガー登録済');
     return 0;
   }
  }

  const sheet = SpreadsheetApp.getActive();
  //トリガーの設定をスクリプトで
  ScriptApp.newTrigger('onEditTriggerFunction')
  .forSpreadsheet(sheet)
  .onEdit()
  .create();
}

/**
 * onEditトリガーを設置する関数
 */
function setOnOpenTrigger(){
 //現在の設定されているトリガーを取得
  triggers = ScriptApp.getProjectTriggers();

  for (i = 0; i < triggers.length; i++) {
   if (triggers[i].getHandlerFunction() == 'onOpenTriggerFunction') {
     Logger.log('setOnOpenTrigger: トリガー登録済');
     return 0;
   }
  }

  const sheet = SpreadsheetApp.getActive();
  //トリガーの設定をスクリプトで
  ScriptApp.newTrigger('onOpenTriggerFunction')
  .forSpreadsheet(sheet)
  .onOpen()
  .create();
}

/**
 * トリガーをすべて削除する関数
 */
function deleteAllTrigger(){
  const triggers = ScriptApp.getProjectTriggers();
  for (i = 0; i < triggers.length; i++) {
    // トリガーの削除を実行
    ScriptApp.deleteTrigger(triggers[i]);

  }
}
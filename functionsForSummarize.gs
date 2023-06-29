/**
 * 各個人シートの値を比較して集約シートにまとめる関数
 */
function checkDiff(){

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sumSheet = spreadSheet.getSheetByName('集約');
  const setSheet = spreadSheet.getSheetByName('ホーム');
  //名前入りシート数を取得
  const memberNum = setSheet.getRange('F6').getValue();


  let values =[0];
  for (i=0;i<memberNum;i++) {  //各シートの値を配列に格納
    let hisSheet = spreadSheet.getSheets()[i+2];
    let hisRange = hisSheet.getRange(1, 1, hisSheet.getLastRow(), hisSheet.getLastColumn());
    values[i] = hisRange.getValues();
  }



  //表の大きさは縦5~35,横2~74
  let attendanceMatrix =[];
  let nameArray = [];
  let tmpRowArray = [];
  let tmpColumnArray = [];
  const nameList = setSheet.getRange(8,6,memberNum,1).getValues().flat();
  let sumSheetLastRow = sumSheet.getLastRow(); //行番号=行数
  let sumSheetLastColumn = sumSheet.getLastColumn();

  //表外に書き込みがある場合無視
  if (sumSheetLastRow > 35) {
    sumSheetLastRow = 35;
  }
  if (sumSheetLastColumn > 74) {
    sumSheetLastColumn = 74;
  }
  // 値を格納した配列で5Bにあたるところから判定を繰り返す
  // あるセルについて各シートを判定、終わったら次のセルへ
  // nameArray i行j列で〇だった人の配列
  // tmpRowArray[j] i行j列のnameArray
  // tmpColumnArray[i] i行のtmpRowArray 
  for (i=4;i<sumSheetLastRow;i++) {  //行の移動
    tmpRowArray =[];
    for (j=1;j<sumSheetLastColumn+1;j++) {  //列の移動
      nameArray = [];
      for (k=0;k < memberNum;k++) { //人の移動
        if (values[k][i][j] === "○") { //判定(漢数字ゼロではない)
          nameArray.push(nameList[k]); //〇だった人をnameArrayに追加
        }
      }
      tmpRowArray[j] = nameArray;
    }
    tmpColumnArray[i] = tmpRowArray;
    attendanceMatrix = tmpColumnArray;
    // attendanceMatrix[i][j] i行j列の出席者 [String,String,...]
  } 

  //配列を元に集約シートに書き込み
  const tableRange = sumSheet.getRange(5,2,sumSheetLastRow-4,sumSheetLastColumn-1);
  //集約シートの初期化
  tableRange.clearContent();
  tableRange.clearFormat();
  tableRange.clearNote();
  tableRange.setHorizontalAlignment("center"); 
  // 集約シートに書き込み
  let tmpList = [];
  let inputArray = sumSheet.getRange(5,2,35,74).getValues();
  for (i=4;i<sumSheetLastRow;i++) {
    for (j=1;j<sumSheetLastColumn+1
    ;j++) {
      let tmpCell = sumSheet.getRange(i+1,j+1);
      tmpList = attendanceMatrix[i][j];
      if (tmpList != []){
        // i行j列 name\nname\name...
        tmpCell.setValue(tmpList.length);
        // tmpCell.setNote(tmpList.join("\n"));
        inputArray[i-4][j-1] = attendanceMatrix[i][j].join("\n")+"\n\n--既存の予定--";
      }
    }
  }
  sumSheet.getRange(5,2,35,74).setNotes(inputArray);

  //条件付き書式を再設定
  const colorRange = sumSheet.getRange(5,2,sumSheetLastRow-4,sumSheetLastColumn);
  const ruleGreen = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(memberNum) 
    .setBackground("#b7e1cd") //セル背景を設定（緑）
    .setRanges([colorRange])
    .build();

  const ruleYellow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(memberNum-1) 
    .setBackground("#fce8b2") //セル背景を設定（黄）
    .setRanges([colorRange])
    .build();
  
  const ruleRed = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(memberNum-2) 
    .setBackground("#f4c7c3") //セル背景を設定（赤）
    .setRanges([colorRange])
    .build();
  
  const calendarRange = sumSheet.getRange(2,2,sumSheetLastRow-1,sumSheetLastColumn-1);
  const ruleGray = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("OR(B$2<today(),B$2>today()+28)") 
    .setBackground("#a6a6a6")
    .setRanges([calendarRange])
    .build();

  const ruleOrange = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("B$2=today()")
    .setBackground("#fce8b2")
    .setRanges([calendarRange])
    .build();

  const rules =[ruleGreen,ruleYellow,ruleRed,ruleGray,ruleOrange];
  sumSheet.setConditionalFormatRules(rules);

}


/**
 * 集約結果をテキストにして配列に格納して返す関数
 * @returns {Array} 集約結果
 * @example [["1/1 10:00~12:00"],[]]
 */
function summarizedResultToText(){
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sumSheet = spreadSheet.getSheetByName('集約');
  const setSheet = spreadSheet.getSheetByName('ホーム');
  //名前入りシート数を取得
  const memberNum = setSheet.getRange('F6').getValue();

  // 行列を転置する関数
  const transposeArray = array => array[0].map((_, colIndex) => array.map(row => row[colIndex]));

  const tableColumns = 74;
  const tableRows =31;

  const resultMatrix = sumSheet.getRange(5,2,tableRows,tableColumns).getValues();
  const timeHeaderArray = sumSheet.getRange(5,1,tableRows,1).getDisplayValues().flat();
  const dateHeaderArray = sumSheet.getRange(2,2,1,tableColumns).getValues().flat();


  // 転置する[[n行],[n+1行]]->[[n列],[n+1列]]
  const transposedResultMatrix = transposeArray(resultMatrix);

  let resultTextArray = [];
  // 列ごとに探索
  for(i=0;i<74;i++){
    let consecutiveCounter = 0;
    // 1行ずつ進める
    for(j=0;j<31;j++){
      const value = transposedResultMatrix[i][j];
      if(value === memberNum){
        consecutiveCounter++;
      }else if (value != memberNum && consecutiveCounter>0){
        let cellAmount = consecutiveCounter;
        let date = Utilities.formatDate(dateHeaderArray[i], 'JST', 'MM/dd');
        let startTime = timeHeaderArray[j-cellAmount];
        let endTime = timeHeaderArray[j-1];
        resultTextArray.push(`${date} ${startTime}~${endTime}`);
        consecutiveCounter=0;
      }
    }
  }


  return resultTextArray;
}
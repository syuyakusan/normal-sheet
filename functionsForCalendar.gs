/**
 * Google カレンダーの予定を『集約』シートに反映する関数
 */
function checkCalendar() {
  
  // 上から、「個人練(部室)」「合わせ練(部室)」「会議(部室)」「その他日程」「東北大MUSICA」

  const privateCalenderId = PropertiesService.getScriptProperties().getProperty('privateCalenderId');
  const ensembleCalenderId = PropertiesService.getScriptProperties().getProperty('ensembleCalenderId');
  const meetingCalenderId = PropertiesService.getScriptProperties().getProperty('meetingCalenderId');
  const otherCalenderId = PropertiesService.getScriptProperties().getProperty('otherCalenderId');
  const companyCalenderId = PropertiesService.getScriptProperties().getProperty('companyCalenderId');

  // getCalenderEvent(privateCalenderId);
  getCalenderEvent(ensembleCalenderId);
  getCalenderEvent(meetingCalenderId);
  getCalenderEvent(otherCalenderId);
  getCalenderEvent(companyCalenderId);


  statusFunc('完了');
  showProcess('');

}


/**
 * 指定されたidのカレンダーから予定を取得し、該当するセルにメモとして書き込む関数
 * @param {String} id カレンダーid
 */ 
function getCalenderEvent(id) {

  if(id === undefined){
    return;
  }

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = spreadSheet.getSheetByName('ホーム');

  //IDを指定して、Googleカレンダーを取得する
  let myCalendar = CalendarApp.getCalendarById(id);
  //今日のDateオブジェクトをつくる
  let startDate = new Date();
  const calendarFirstDate = setSheet.getRange('F7').getValue();;
  if(startDate < calendarFirstDate){
    startDate = calendarFirstDate;
  }
  //今日から28日後のDateオブジェクトをつくる
  let endDate = new Date();
  endDate.setFullYear(startDate.getFullYear())
  endDate.setMonth(startDate.getMonth())
  endDate.setDate(startDate.getDate() + 28);
  //開始日～終了日に存在するGoogleカレンダーのイベントを取得する
  let myEvent = myCalendar.getEvents(startDate, endDate);

  // 取得したmyEventの全件について整形し、配列に格納する
  let splitedTimeArray =[]; //[****,****,****][日付,開始時刻,終了時刻,予定名]
  for(i = 0 ; i < myEvent.length ; i++ ){

    //予定の日時
    const date = myEvent[i].getStartTime();
    startDate = Utilities.formatDate(date, "Asia/Tokyo", "MMdd");

    //予定の開始時刻
    let startHours = "0" + myEvent[i].getStartTime().getHours();
    startHours = startHours.slice(-2);
    let startMinutes = "0" + myEvent[i].getStartTime().getMinutes();
    startMinutes = startMinutes.slice(-2);
    let startTime = startHours + startMinutes; //データ型から文字列に変換

    //予定の終了時刻
    let endHours = "0" + myEvent[i].getEndTime().getHours();
    endHours = endHours.slice(-2);
    let endMinutes = "0" + myEvent[i].getEndTime().getMinutes();
    endMinutes = endMinutes.slice(-2);
    let endTime = endHours + endMinutes; //データ型から文字列に変換
    let title;
    if(startTime === "0000" && endTime === "0000"){
      startTime="0800"
      endTime="2300"
      // TODO:テーブルの時間帯の範囲が変更されても動的に変更できるように
      title = `${myEvent[i].getTitle()} (終日)`;
    }else{
      title = `${myEvent[i].getTitle()} (${startHours}:${startMinutes}~${endHours}:${endMinutes})`;
    }

    splitedTimeArray[i] =[startDate,startTime,endTime,title];

  }

  //セルの位置を求める関数
  const calcCellPosition = function(value) {
    let month = value[0].slice(0,2);
    let date = value[0].slice(2,5);
    let startH =value[1].slice(0,2);
    let startM = value[1].slice(2,5);
    let endH =value[2].slice(0,2);
    let endM = value[2].slice(2,5);
    let startRow = ((startH-7)*2 + (startM/30) +3);
    let endRow = ((endH-7)*2 + (endM/30) +3);
    let rowLength = endRow - startRow + 1;

    const fd = setSheet.getRange('F7').getValue();

    let firstDay = new Date(fd); //カレンダーの最初の日付

    let setDay = new Date(fd); //指定された日付
    setDay.setMonth(month-1);
    setDay.setDate(date);

    let difDays = (setDay - firstDay)/86400000;

    if (difDays < 0) { //年をまたぐ入力一年後の日付に
      setDay = setDay.setFullYear(setDay.getFullYear() + 1);
      difDays = (setDay - firstDay)/86400000;
    }


    const array = [difDays+2,startRow,rowLength,value[3]];

    return (array);
  }

  const calculatedTimeArray =splitedTimeArray.map(calcCellPosition); //[[列数,開始行数,開始行から終了行までの行数,タイトル]]

  //配列をもとにスプレッドシートに入力
  const sheet = spreadSheet.getSheetByName('集約');
  const notes = sheet.getRange(1,1,35,75).getNotes();  
  
  for (let i=0;i<calculatedTimeArray.length;i++) {
    const array = calculatedTimeArray[i];
    const column = array[0];
    let row = Number(array[1]);
    row = Math.trunc(row);
    let length = Number(array[2]);
    length = Math.ceil(length);
    const title = array[3];

    for (j=0;j<length;j++) {
      notes[row+j-1][column-1] = notes[row+j-1][column-1] +'\n'+title;
    }
  }
  sheet.getRange(1,1,35,75).setNotes(notes);

}


var Webhook_URL = "https://hooks.slack.com/services/********"

//カレンダーの情報を取得する
var calendars = CalendarApp.getCalendarById("********@gmail.com");


//全体LINEに送信機能
var group_push_ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/********/edit?usp=sharing")
var group_push_sheet = group_push_ss.getSheetByName("messages")
var GP_DATE_ROW = 1
var GP_TIME_ROW = 2
var GP_USER_NAME_ROW = 3
var GP_USER_ID_ROW = 4
var GP_TRIGGER_ID_ROW = 5
var GP_MESSAGE_ROW = 6

//userID管理
var user_id_ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/********/edit?usp=sharing")
var user_id_sheet = user_id_ss.getSheetByName("部員")
var UI_GRADE_ROW = 1
var UI_LAST_NAME_ROW = 2
var UI_FIRST_NAME_ROW = 3
var UI_ID_ROW = 4

//部練出欠表
var buren_ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/********/edit?usp=sharing")
var buren_sheet = buren_ss.getSheetByName("部練")
var BUREN_GRADE_ROW = 1
var BUREN_NAME_ROW = 2

//貸切出欠
var kashikiri_ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/********/edit?usp=sharing");
var kashikiri_sheet = kashikiri_ss.getSheetByName("貸切");
var KASHIKIRI_GRADE_ROW = 1;
var KASHIKIRI_NAME_ROW = 2;


function postSlack(text){
  var url = Webhook_URL;

  // 必要なオプションを指定
  var options = {
    "method" : "POST",
    "headers": {"Content-type": "application/json"},
    "payload" : '{"text":"' + text + '"}'
  };

  // Google Apps ScriptのFetchのためのAPIを叩く
  UrlFetchApp.fetch(url, options);
}



//Momentライブラリ曜日を日本語にセット
Moment.moment.lang('ja', {
    weekdays: ["日曜日","月曜日","火曜日","水曜日","木曜日","金曜日","土曜日"],
    weekdaysShort: ["日","月","火","水","木","金","土"],
});

//次の日の予定リマインド関数
function remindSchedule(is_morning){

    // Date型のオブジェクトの生成
    var date = new Moment.moment();

　　// dateを翌日にセットする
    date.add(1, "days");
　　//date.setDate(date.getDate() + 1);

    //翌日の予定を取得する
    var schedules = calendars.getEventsForDay(date.toDate());

    /*
    * 予定を出力する
    */
    var sendText = new String();
    //sendText += "\n【" + date.format("M月D日(ddd)") + " の予定】";
    //もし予定がない場合は
    if(schedules.length == 0) {
        //sendText += date.format("M月D日(ddd)")+"明日は何も予定がありません";
    } else if(schedules.length > 0 ) {
      // 予定を繰り返し出力する
      for(var i = 0; i < schedules.length; i++) {
        sendText = ""
        var startTime = Moment.moment(schedules[i].getStartTime());
        var endTime = Moment.moment(schedules[i].getEndTime());
        var title = schedules[i].getTitle();
        //var date_D = date.format("M/D");

        if (is_morning){
          if (title.match(/貸切/)!=null){
            //貸切イベント
            var yuushi_text = "";
            if (title.match(/有志/)!=null){ yuushi_text = "有志" }
            sendText += "【" + date.format("M月D日(ddd)") + yuushi_text + "貸切出欠確認】"
            sendText += "\n時間：" + startTime.format("HH:mm") + "〜" + endTime.format("HH:mm");
            sendText += "\n場所：";
            sendText += title.replace("有志","").replace("貸切@","");
            sendText += "\n出席者は";
            sendText += attend_member(startTime.toDate(),"kashikiri");
            sendText += "\nと伺っております。"
            sendText += "\nもし間違いや変更、音源変更がございましたら本日18:00までにご連絡ください。";
          }
        }else{
          if(title.match(/部練/)!=null) {
            sendText +=  "【" + date.format("M月D日(ddd)") + title +  "出席者】"
            sendText += attend_member(date,"buren");
          }else if(title.match(/バッジテスト/)!=null){
            sendText += "【" + date.format("M月D日(ddd)") + title + "】"
            sendText += "\n" + startTime.format("HH:mm") + "〜" + endTime.format("HH:mm")
            //sendText += "\n受験者は以下の通りです"
          }else{
            //その他のイベント
            sendText += "明日の予定です！\n"
            sendText += "" + date.format("M月D日(ddd)")
            if (startTime.isSame(endTime,'minute')){
            }else{
              sendText += startTime.format("HH:mm") + "〜" + endTime.format("HH:mm");
            }
            sendText += "\n"+title
          }

        }
        if(sendText!=""){
          postSlack(sendText);//部のSlackに投稿
        }
      }
    }

  return 0;
}

function attend_member(date,type){
  var attendance_row = 0;
  var sendText = new String();
  if (type == "kashikiri"){
    sheet = kashikiri_sheet
    GRADE_ROW = KASHIKIRI_GRADE_ROW
    NAME_ROW = KASHIKIRI_NAME_ROW
    is_same_key = 'minute'
  }else if(type == "buren"){
    sheet = buren_sheet
    GRADE_ROW = BUREN_GRADE_ROW
    NAME_ROW = BUREN_NAME_ROW
    is_same_key = 'day'
  }

  var days_on_sheet = sheet.getRange(1,1,1,50).getValues()
  for (var i = 2;i<100;i++){
    if(Moment.moment(sheet.getRange(1,i).getValue()).isSame(date,is_same_key)){
      attendance_row = i
      break;
    }
  }
  if (attendance_row == 0){
    sendText += "\n......未登録......";
    return sendText;
  }

  var grades = sheet.getRange(3,GRADE_ROW,50).getValues();
  var names = sheet.getRange(3,NAME_ROW,50).getValues();
  var attendance = sheet.getRange(3,attendance_row,50).getValues();
  var first = true;
  var tmp_grade = 0.0;
  var undecidedMembers ="";
  var undecidedFirst = true;
  for (i = 0; i<50;i++){
    if (names[i] != ""){
    if (attendance[i]=="出席" || attendance[i].toString().match(/早退/)!=null ||attendance[i].toString().match(/途中参加/)!=null){//　出席の人
      if (parseInt(grades[i])!=parseInt(tmp_grade)){
        first = true; //学年が切り替わった時改行
        tmp_grade = grades[i];
      }
      if (first){
        sendText += "\n"; //始めの人は改行
        first = false;
      }else{
        sendText += ","; //始め以外はコンマ
      };
      sendText += names[i];//名前を追加
      if(attendance[i].toString().match(/早退/)!=null ||attendance[i].toString().match(/途中参加/)!=null){
        sendText += "("+attendance[i]+")";//(遅刻or途中参加)を追加
      }
    }else if (names[i]=='他大生'){
      if (attendance[i]!=""){
      sendText += "\n[他大]";
      sendText +=attendance[i];
      }
    }else if (attendance[i]=="未定"||attendance[i]==""){
      if (undecidedFirst){
        undecidedMembers += "\n[未定]"; //始めの人は[未定]
        undecidedFirst = false;
      }else{
        undecidedMembers += ","; //始め以外はコンマ
      };
      undecidedMembers += names[i]
    }
    }
  }
  sendText += undecidedMembers;
  return sendText;
}


function morningRemind(){
  remindSchedule(true)
  return 0;
}

function eveningRemind(){
  remindSchedule(false)
  return 0;
}


function getUserLastName(userID){
  var user_id_data = user_id_sheet.getRange(2,2,50,3).getValues();
  var user_name = ""
  const USER_ID_ROW = UI_ID_ROW -2
  const USER_NAME_ROW = UI_LAST_NAME_ROW -2
  for(var i in user_id_data){
    if(user_id_data[i][USER_ID_ROW]==userID){
      user_name = user_id_data[i][USER_NAME_ROW]
    }
  }
  return user_name
}
function getUserFirstName(userID){
  var user_id_data = user_id_sheet.getRange(2,2,50,3).getValues();
  var user_name = ""
  const USER_ID_ROW = UI_ID_ROW -2
  const USER_NAME_ROW = UI_FIRST_NAME_ROW -2
  for(var i in user_id_data){
    if(user_id_data[i][USER_ID_ROW]==userID){
      user_name = user_id_data[i][USER_NAME_ROW]
    }
  }
  return user_name
}

function setUserName(user_message,userID){
  var user_name = user_message.split("#")[1]
  //スプレッドシートにユーザID登録
  for (var i = 1;i<50;i++){
    if(user_id_sheet.getRange(i,UI_LAST_NAME_ROW).getValue()==user_name){
      var registered_id = user_id_sheet.getRange(i,UI_ID_ROW).getValue()
      if(registered_id==""){
        user_id_sheet.getRange(i,UI_ID_ROW).setValue(userID)
        return [user_name + "さんとして登録しました！"]
      }else if(registered_id==userID) {
        return [user_name + "さんは登録済です"]
      }else{
        return [user_name + "さんは別のIDで登録済です",
               "変更したい場合は管理者に問い合わせて下さい"]
      }
    }
  }
  return [user_name + "さんを名簿から見つけることができませんでした"]
}

//メッセージ送信テスト
function createMessage() {
  //メッセージを定義する
  message = "よろしくお願いします！";
  return postSlack(message);//部のSlackに投稿
}

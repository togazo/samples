/*
- RcptMngSys.gs
- Author: TGA (@togazo / TGABOOK)
- Update: 2019.9.18
*/

// 定数
var ID_FORM_ = "[Form's ID]"; //フォームの固有ID（*要書換）
var ID_ST_ = "[Spreadsheet's ID]"; //スプレッドシートの固有ID（*要書換）
var ST_NAM_ENTRY_ = "フォームの回答 1"; //フォームの回答の入るシート名
var ST_NAM_ALLOW_ = "AllowUser"; //ホワイトリストのシート名

var MSG_FINAL_TITLE_ = "最終確認"; //フォームの最後の質問
var MSG_FINAL_ITEM_CANCEL_ = "参加申し込みを取り消します"; //キャンセルの時の回答項目

var EVENT_NAME_ = "ほげふがオンライン勉強会"; //今回のイベント名（*要書換）
var MAX_ATTENDEE = ??; //イベントの定員（任意の整数）
var URL_MTG_ = "https://zoom.us/j/ooooooooo"; //Zoomミーティング参加用URL（*要書換）
var TIME_IMMEDIATE_ST_ = "20??-??-??T??:??:00"; //イベント開始m分前の時刻（*要書換/書式はYYYY-MM-DDTHH:mm:00）
var TIME_EVT_BEFORE_  = ??; //イベント開始m分前のmの値（任意の整数）

var PRJ_NAME_ = "ほげほげふがふが"; //今回のイベントを主催する組織(?)名（*要書換）
var PRJ_URL_ = "http://fgfg.hoge2.dmy"; //組織のWebサイト（*要書換）
var MAIL_ADDR_ADMIN_ = "admin@hoge2.dmy"; //組織または担当者の連絡先（*要書換）
var ADMIN_NAME_ = PRJ_NAME_ + "管理人"; //担当者名


//-----------
/*
//アプリのアクセス許可をとるためだけの関数。
function myFunction() {
  FormApp.getActiveForm();
  SpreadsheetApp.getActive();
  Logger.log(MailApp.getRemainingDailyQuota());
}
*/

/*
//入力フォームの初期設定
function InitForm(){
  var f = FormApp.openById(ID_FORM_);
  f.setTitle(EVENT_NAME_ + " 参加受付フォーム");

  //送信後の確認メッセージ
  var message = "このたびは「" + EVENT_NAME_ +"」にお申込み頂きまして、誠にありがとうございます。\n\n";
  message += "配信用のURLは、開始時刻の約" + TIME_EVT_BEFORE_  + "分前にお申込みアドレス宛にお送りします。それ以降のお申込みの場合は都度連絡を差し上げます。\n\n";
  message += "5分以上経過しても当会からの折り返しの連絡が一切届かない場合は、入力したメールアドレスが間違っている可能性がございます。「回答を編集」を押して編集画面に戻り、念のためにメールアドレスが正しく入力されていることを御確認下さいませ。\n"
  message += "メールアドレスが正しい場合は、メーラーの迷惑メールフォルダなどに入っている可能性もございます。\n\n";
  message += "以上を御確認いただき、それでも見当たらない場合は、御手数でも下記まで御連絡下さいませ。\n\n";
  message += "▶問合せ・連絡先 " + MAIL_ADDR_ADMIN_;
  f.setConfirmationMessage(message);

  //回答を受け付けていないときの回答者へのメッセージ
  SetCustomClosedFormMessage_("申込み受付期間が終了したため");
}
*/
//-----------


//回答を受け付けていないときの回答者へのメッセージ
function SetCustomClosedFormMessage_(reason)
{
  var message = "このたびは「" + EVENT_NAME_ + "」申し込みページにアクセスいただきまして、誠にありがとうございます。\n\n";
  message += "申し訳ございませんが" + reason + "、受付を締め切らせていただきました。\nまたの参加をお待ちしております。\n\n";
  message += PRJ_NAME_ + "\n▶公式サイト " + PRJ_URL_ + "\n";
  FormApp.openById(ID_FORM_).setCustomClosedFormMessage(message);
}


//フォームからデータ編集があった
function EvtFormSubmit(e){
  if(e === undefined){ return; }
  var row = e.range.getRow();
  var addr = e.values[1];

  //キャンセル（編集可能な時だけ処理可能）
  if(e.namedValues[MSG_FINAL_TITLE_] == MSG_FINAL_ITEM_CANCEL_){
    SpreadsheetApp.getActiveSheet()
      .getRange("B" + row + ":C" + row)
      .clearContent(); //タイムスタンプとキャンセルの旨だけ残す
    return;
  }
  if(addr == ""){ return; } //メールアドレス以外の変更のみだったら何も処理をしない

  //身元の確認
  if(!ChkAllowUser_(addr)){
    SpreadsheetApp.getActiveSheet().getRange("E" + row).setValue("×");
    return;
  }
  SpreadsheetApp.getActiveSheet().getRange("E" + row).clearContent(); //再編集で承認済みメンバーに修正される場合もあるため

  DelDuplicate_(SpreadsheetApp.getActiveSheet(),row,addr); //重複を削除

  //定員ならばフォームを閉じる
  var num = SpreadsheetApp.getActiveSheet()
    .getRange("B2:E"+SpreadsheetApp.getActiveSheet().getLastRow())
    .getValues()
    .filter( //キャンセル、重複、ホワイトリスト非掲載を除く
      function(val){ return (val[0]!="" && val[3]=="");
     })
    .length;
  if(num >= MAX_ATTENDEE){
    SetCustomClosedFormMessage_("申込者数が定員に達したため");
    FormApp.openById(ID_FORM_).setAcceptingResponses(false);
  }
}


//スプレッドシート上の重複を削除
function DelDuplicate_(st,idx,mail){
  var row = st.getLastRow();
  var items = st.getRange("A2:E"+row).getValues();

Logger.log("idx="+idx);
  for(var i=0; i<items.length; i++){
Logger.log("[" + i + "]=" + i);
    if(i == idx-2 || items[i][1]!=mail){continue;}
Logger.log("削除");
    items[i] = ['','','','',''];
  }
  st.getRange("A2:E"+row).setValues(items);
}


//送信時間になったので一斉メールを送る
function EvtSendReminder(){
  //メールのリストを作成する
  var mailList = (function(){
    var st = SpreadsheetApp.getActive().getSheetByName(ST_NAM_ENTRY_);
    var row = st.getLastRow();
    var rtn = st.getRange("A2:E"+row)
      .getValues()
      .filter(function(val){ //キャンセル、重複、登録外の申込者、即時対応の時間以降を除外
        var timeSt = new Date(TIME_IMMEDIATE_ST_);
        return ((val[1] != "" && val[4] == "") && (val[0] < timeSt))
    });
    for(var i=0; i<rtn.length; i++){
      rtn.push(rtn.shift()[1]);
    }
    return rtn;
  }());

  //URLを通知する
  var body = "この度は「" + EVENT_NAME_ + "」へお申込み頂きまして、誠にありがとうございます。\n\n"
  body += "イベント開始" + TIME_EVT_BEFORE_  + "分前より次のアドレスからアクセスできます。\n";
  body += URL_MTG_ +"\n\n";
  body += "それでは、皆様とお話しできることを楽しみにしております。\n";
  SendAutoMail(mailList,"オンライン勉強会 参加URLのご案内",body);
}


//認証されたユーザかどうかチェックする
function ChkAllowUser_(email){
  var st = SpreadsheetApp.openById(ID_ST_).getSheetByName(ST_NAM_ALLOW_);
  var db = st.getRange("B2:B" + st.getLastRow()).getValues();
  for(var i=0; i<db.length; i++){
    db.push(db.shift()[0]);
  }
  return (db.indexOf(email)<0) ? false : true;
}

//フォームを閉じる
function EvtCloseForm(){
  SetCustomClosedFormMessage_("申込み受付期間が終了したため");
  FormApp.openById(ID_FORM_).setAcceptingResponses(false);
}

//---------------------
//フォーム側の処理
function EvtSubmit(e) {
  if(e === undefined){ return; }

  var atndId = e.response.getId();
  var atndMail = e.response.getRespondentEmail();
  var body = e.response.getItemResponses()[0].getResponse() + " 様\n\n";

  //キャンセル
  var bl = function(ir){
    for(var i=0; i<ir.length; i++){
      if(MSG_FINAL_TITLE_ != ir[i].getItem().getTitle()){continue;}
      if(MSG_FINAL_ITEM_CANCEL_ == ir[i].getResponse()){ return true; }
    }
    return false;
  }(e.response.getItemResponses());
  if(bl){
    e.source.deleteResponse(e.response.getId());
    //削除完了の旨を連絡（参加申し込みしていないのにキャンセルした場合も送る）
    body = "「" + EVENT_NAME_ + "」へのお申込みをキャンセルいたしました。\n";
    body += "それでは、またの参加をお待ちしております。\n";
    SendAutoMail(atndMail,"キャンセル手続きが完了しました",body);
    return;
  }

  //身元確認
  if(!ChkAllowUser_(atndMail)){
  body += "この度は「" + EVENT_NAME_ + "」に興味をお持ちいただきまして、誠にありがとうございます。\n\n"
  body += "申し訳ございませんが、頂いたメールアドレスに該当する登録者が見つからないため、受付を完了することができません。\n";
  body += "申し込み情報の変更は、下記URLより可能です。\n\n";
  body += "それでは、御手数ですが御確認のほど宜しくお願い致します;。\n\n";
  body += "▶登録情報の編集ページ\n" + e.response.getEditResponseUrl() + "\n"
  SendAutoMail(atndMail,"お手続きが完了しておりません",body);
Logger.log(body);
    return;
  }

  //重複削除
  var res = e.source.getResponses();
  for(var i=res.length-1;i>=0;i--){
    var compId = res[i].getId();
    if(atndId == compId){continue;}
    if(atndMail == res[i].getRespondentEmail()){
      e.source.deleteResponse(compId);
    }
  }

//即時対応の場合はURL発行、エラーの場合とそれ以外は編集アドレスありの確認メール
  var timeSt = new Date(TIME_IMMEDIATE_ST_);
  if(e.response.getTimestamp() >= timeSt){ //即時URL発行
    body += "この度は「"+ EVENT_NAME_ +"」へお申込み頂きまして、誠にありがとうございます。\n\n"
    body += "イベント開始" + TIME_EVT_BEFORE_  + "分前より次のアドレスからアクセスできます。\n";
    body += URL_MTG_ +"\n\n";
    body += "それでは、皆様とお話しできることを楽しみにしております!!\n";
    SendAutoMail(atndMail,"参加URLのご案内",body);
  }else{ //申し込み完了の通知
    body += "この度は「"+ EVENT_NAME_ +"」へお申込み頂きまして、誠にありがとうございます。\n\n"
    body += "当日、イベント開始およそ" + TIME_EVT_BEFORE_  + "分前に入室用のURLをメールにてお送りします。\n";
    body += "申し込み情報の変更、参加キャンセルのお手続きは下記URLより可能です。\n\n";
    body += "それでは、皆様とお話しできることを楽しみにしております!!\n\n";
    body += "▶登録情報の編集ページ\n" + e.response.getEditResponseUrl() + "\n";
    SendAutoMail(atndMail,"参加申し込みが完了しました",body);
  }
Logger.log(body);
}

//---------------------
//メール送信
function SendAutoMail(recipient,subject,body){
  var MAX_NUM_MAIL = 20;
Logger.log("SendAutoMail(recipient,subject,body)");
Logger.log("recipient=" + recipient);
Logger.log("subject=" + subject);
Logger.log("body=" + body);
  //自動送信メールのフッタ
  var MAIL_BODY_FOOTER_ = "--------------------\n";
  MAIL_BODY_FOOTER_ += "このメールは「" + PRJ_NAME_ + "」のメールシステムにより自動送信されました。\n";
  MAIL_BODY_FOOTER_ += "万が一、本メールに心当たりの無い場合は直ちに破棄して頂きますようお願いします。\n\n";
  MAIL_BODY_FOOTER_ += PRJ_NAME_ + "\n" + PRJ_URL_ + "\n";

  var param = { name: ADMIN_NAME_,
    replyTo: ADMIN_NAME_ + " <" + MAIL_ADDR_ADMIN_ + ">",
    subject: EVENT_NAME_ + " " + subject,
    body: body + MAIL_BODY_FOOTER_
  };

  if(typeof recipient != 'string' && !Array.isArray(recipient)){ return;}
  if(Array.isArray(recipient) && recipient.length == 1){recipient = recipient[0];}

  if(typeof recipient == 'string'){//1名の処理
    param['to'] = recipient;
Logger.log("param['to']="+param['to']);
return; //デバッグ用。メールを送っても問題なければ行削除。
    try{MailApp.sendEmail(param);}
    catch(e){Logger.log("メール送信失敗(1名)：" + e);}
  }
  else if(Array.isArray(recipient)){//複数名の処理（現状、GMailの上限を超える人数(100)は考慮していない）
    param['to']= MAIL_ADDR_ADMIN_;
    for(var i=0; i<recipient.length; i+=MAX_NUM_MAIL){
      param['bcc']=recipient.slice(1, i+MAX_NUM_MAIL).join();
Logger.log(i+" / param['bcc']="+param['bcc']);
continue; //デバッグ用。メールを送っても問題なければ行削除。
      try{MailApp.sendEmail(param);}
      catch(e){Logger.log("メール送信失敗(複数)：" + e );}
    }
  }
}

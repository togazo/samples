/*******************************
* main.gs
* Author: TGA (@togazo / TGABOOK)
* Create: 2020.4.28
* Update: 2021.1.3
*******************************/
const DEBUG_MODE = 0;  //デバッグモード(0=本番、1=デバッグ)

/* 定数 */
//フォームのID
const ID_FM = '[FORM ID 1]';
const ID_FM_MNG = '[FORM ID 2]'; //申込者管理
//シートのID
const ST_NAM_ENTRY = '[SHEET NAME 1]'; //申込フォームの回答が入るシート名
const ST_NAM_MNG = '[SHEET NAME 2]'; //管理フォームの回答が入るシート名
const ID_ST_DB = '[FORM ID 3]'; //ホワイトリスト、ブラックリストを管理しているDB
const ST_NAM_ALLOW = '[SHEET NAME Allow]'; //ホワイトリストのシート名
const ST_NAM_DISALLOW = '[SHEET NAME Disallow]'; //ブラックリストのシート名
const FI_MSG_ATND = '参加を申し込む';
const FI_MSG_CANCEL = '申し込みを取り消す';
const MSG_FINAL_TITLE = '最終確認'; //フォームの最後の質問
const MSG_FINAL_ITEM_CANCEL = '申し込みを取り消す'; //キャンセルの時の回答項目
const SEL_USER_OPEN = 'それ以外';
const SEL_USER_ALLOW = '参加経験あり';
const FLAG_CHK_ALLOW = '審査中'; //ゲスト参加の審査中
const FLAG_NUM_MAX = '満員';
const MLTI_INVITE = '招待状URL（ミーティングID）のご案内';
const MLTI_COMP = '参加申し込みが完了しました';
const MLTI_COMP_INVITE = '参加申込み完了のお知らせと招待URLのご案内';
const MLTI_CANCEL = '申し込みを取り消しました';
const MLTI_URVW = 'お申し込み内容を確認中です';
const MLTI_NOTADMIT = 'お申し込みを承認することができませんでした';
const PRJ_NAME = '[PROJECT NAME]'; //今回のイベントを主催する組織(?)名（*要書換）
const PRJ_URL = '[PROJECT URL]'; //組織のWebサイト（*要書換）
const MAIL_ADDR_ADMIN = '[ADMIN MAIL ADDR]'; //組織または担当者の連絡先（*要書換）
const ADMIN_NAME = `[ADMIN NAME]`; //担当者名
const WEBHOOK_SLK_URL = '[Slack WEBHOOK URL]'; //Slack Webhook通知
const WEBHK_SLK_CH = '#[Slack CHANNEL NAME]';
const WEBHOOK_DSC_URL = '[Discord WEBHOOK URL]'; //Discord Webhook通知

const SS = SpreadsheetApp.getActiveSpreadsheet();
const ST = SS.getSheetByName(ST_NAM_ENTRY);
const ST_MNG = SS.getSheetByName(ST_NAM_MNG);
const FM = FormApp.openById(ID_FM);
const FM_MNG = FormApp.openById(ID_FM_MNG);
const WD = ['日','月','火','水','木','金','土'];

//申込者フォームの回答のセルの列番号
const I_DT = 0; //タイムスタンプ
const I_ML = 1; //メールアドレス
const I_TY = 2; //手続き種別
const I_EXP = 3; //[???]への参加経験
const I_NM = 4; //参加者の名前
const I_CNCT = 5; //端末の台数（connection）
const I_AGE = 6; //ニンジャの年齢
const I_PN = 7; //電話番号
const I_PIC = 8; //画像・映像の公開可否について
const I_DJ = 9; //所属のDojo
const I_WK = 10; //Dojoあるいは日常生活の中での創作活動の状況について
const I_FR = 11; //自由記入欄
const I_AGR = 12; //最終確認（agree）
const I_WR = 13; //書込数
const I_ALW = 14; //参加許可(allow)
const COL_LAST = "O".charCodeAt(0) - "A".charCodeAt(0) + 1; //一番最後（参加許可）の列番号（忘れるので式にしてる…）

/* 定数（だけど継続的な運用中にちょいちょい変わる値） */
const ps = PropertiesService.getScriptProperties().getProperties();
const EVENT_NAME = ps.EVENT_NAME==null?'':ps.EVENT_NAME; //今回のイベント名
const MTG_URL = ps.MTG_URL==null?'':ps.MTG_URL; //ミーティング参加用URL（招待状URL/ミーティングID）
const MTG_PWD = ps.MTG_PWD==null?'':ps.MTG_PWD; //ミーティング参加用パスワード
const TIME_EVT_ST = ps.TIME_EVT_ST==null?'':ps.TIME_EVT_ST; //イベント開始時刻（書式はDate型）
const TIME_EVT_ED = ps.TIME_EVT_ED==null?'':ps.TIME_EVT_ED; //イベント終了時刻（書式はDate型）
const HOUR_EVT_ST_BFR = ps.HOUR_EVT_ST_BFR==null?'':ps.HOUR_EVT_ST_BFR; //ゲスト参加者の受付締切時刻 - イベント開始時刻のh時間前（h:任意の整数）、0=ゲストと常連を分けない
const MIN_EVT_ST_BFR = ps.MIN_EVT_ST_BFR==null?0:ps.MIN_EVT_ST_BFR; //イベント開始時刻のm分前（m:任意の整数）
const MIN_EVT_ED_BFR = ps.MIN_EVT_ED_BFR==null?0:ps.MIN_EVT_ED_BFR; //イベント終了時刻のm分前（m:任意の整数）
const MAX_RM= ps.MAX_RM==null?0:ps.MAX_RM; //リモート参加の定員

//---------------------
/*
* このスクリプトの設置時に一度だけ手動で実行する
*/
function init()
{
  ScriptApp.getProjectTriggers().forEach(i=>ScriptApp.deleteTrigger(i)); //不要なトリガーの除去
  ScriptApp.newTrigger("EvtFormSubmitSS").forSpreadsheet(SS).onFormSubmit().create(); //トリガーの作成:フォーム送信時（スプレッドシート）
  ScriptApp.newTrigger("addOrigMenu_").forSpreadsheet(SS).onOpen().create(); //トリガーの作成:フォームを開いた時
  UpdateSchedule_();
}

/*
*
* メニュー追加
*
*/
const ORIGMENU = "自動処理";
function addOrigMenu_(){
  SS.addMenu(ORIGMENU, [{name: "フォーム更新", functionName: "UpdateSchedule_"},
                        {name: "プロパティ更新", functionName: "UpdateScriptProperties_"},
                        {name: "メンテナンス中", functionName: "BlMaintenanceON_"},
                        {name: "メンテナンス解除", functionName: "BlMaintenanceOFF_"}]);
}

/*
*
* スクリプトのプロパティ値を更新する
* このスクリプトの プロジェクトのプロパティ>スクリプトのプロパティ の値をシート「properties」の値に更新する
* return: 0 更新失敗, 1 更新成功
*/
function UpdateScriptProperties_()
{
  let st = SS.getSheetByName('properties');
  if(!st){Logger.log('固定値の管理用シートがありません…'); return 0;}
  const row = st.getLastRow();
  if(row < 2){ SpreadsheetApp.getUi().alert('プロパティ用のデータが見つかりません'); return 0; }
  let po = new Object(); //property's object
  let items = st.getRange(2,1,row-1,2).getValues().forEach((e)=>{po[e[0]]=e[1]});
  PropertiesService.getScriptProperties().setProperties(po,true);
  return 1;
}


/*
* 募集要項の更新
*/
function UpdateSchedule_()
{
  if(!UpdateScriptProperties_()){return;} //プロパティの更新
  ScriptApp.getProjectTriggers().forEach(i=>{if(i.getHandlerFunction()!='EvtFormSubmitSS'&&i.getHandlerFunction()!='addOrigMenu_') ScriptApp.deleteTrigger(i)}); //不要なトリガーの除去
  //トリガーの作成（URL一斉通知）
  let trgDate = new Date(TIME_EVT_ST);
  trgDate.setMinutes(trgDate.getMinutes() - MIN_EVT_ST_BFR);
  ScriptApp.newTrigger('EvtSendReminder')
           .timeBased()
           .at(trgDate)
           .create();
  //トリガーの作成（フォームの回答を締め切る）
  trgDate = new Date(TIME_EVT_ED)
  trgDate.setMinutes(trgDate.getMinutes() - MIN_EVT_ED_BFR);
  ScriptApp.newTrigger('EvtCloseForm')
           .timeBased()
           .at(trgDate)
           .create();
  //トリガーの作成（ゲスト参加者の受付を締め切る）
  if(HOUR_EVT_ST_BFR>0){ //注:0の場合は常連の締め切りと同じ
    trgDate = new Date(TIME_EVT_ST);
    trgDate.setHours(trgDate.getHours() - HOUR_EVT_ST_BFR);
    ScriptApp.newTrigger('chAllowGuest_OFF')
             .timeBased()
             .at(trgDate)
             .create();
  }
  UpdateSpreadSheet(); //スプレッドシート更新
  UpdateForm(); //フォーム更新（データを一掃してから最後にUIを更新）
}



/*
 * メール送信
 */
function SendAutoMail(recipient,subject,body)
{
  if(typeof recipient != 'string'){ return;}
  body += TemplateMailFooter();
  subject = `【${PRJ_NAME}】${EVENT_NAME} ${subject}`;

  let param = { name: ADMIN_NAME,
    to: recipient,
    subject: subject,
    body: body
  };
  if(DEBUG_MODE){
    Logger.log(`SendAutoMail()\nrecipient=${recipient}\nsubject=${subject}\nbody=${body}`);
    return;
  }
  try{MailApp.sendEmail(param);}
  catch(e){Logger.log(`メール送信失敗：${e}`);}
}


/*
 * チャットツールに通知を飛ばす
 */
function SendChatNotification_(message)
{
  SendDiscordNotification_(message);
}

//Slack
function SendSlackNotification_(message) {
  message = '[' + EVENT_NAME + '] ' + message;
  const url = WEBHOOK_SLK_URL;
  const channel = WEBHK_SLK_CH;
  const payload = (undefined === channel || "" == channel) ? JSON.stringify({"text":message}):JSON.stringify({"text":message,"channel":channel});
  const options = {'method':'post','contentType':'application/json','muteHttpExceptions':true,'payload':payload};
  if(!DEBUG_MODE){
    const r = UrlFetchApp.fetch(url, options);
    Logger.log("Slack(Webhook): %s " + r.getContentText(), r.getResponseCode());
  }
  else{ Logger.log("Slack(Webhook): %s", message); }
}

//Discord
function SendDiscordNotification_(message)
{
  message = '[' + EVENT_NAME + '] ' + message;
  const url = WEBHOOK_DSC_URL;
  const payload = JSON.stringify({'content':message});
  const options = {'method':'post','contentType':'application/json','muteHttpExceptions':true,'payload':payload};
  if(!DEBUG_MODE){
    const r = UrlFetchApp.fetch(url, options);
    Logger.log("Discord(Webhook): %s %s", r.getResponseCode(), r.getContentText());
  }
  else{ Logger.log("Discord(Webhook): %s", message); }
}


//---------------------
/*
* 参加申込・取り消しの選択肢を切り替える…参加者受付をON/OFFする（満席時用）
* @param true:受付ON / false:受付OFF
*/
function SelectAtndCncl(bl)
{
  let form_type,pb_entry_sel,pb_send;
  FM.getItems().forEach(i=>{
    if(i.getType()==FormApp.ItemType.PAGE_BREAK && i.getTitle()=='参加申込み（詳細）')
      pb_entry_sel = i.asPageBreakItem();
    else if(i.getTitle()=='手続き種別')
      form_type = i.asMultipleChoiceItem();
  });
  if(bl && form_type.getChoices().length!=2){
    form_type.setHelpText('「参加を申し込む」を選ぶと、詳細情報の入力画面に移動します。誤って「申し込みを取り消す」を選んだ場合でも「送信」を選ぶまではキャンセル確定にはなりません。');
    form_type.setChoices([form_type.createChoice(FI_MSG_ATND,pb_entry_sel),form_type.createChoice(FI_MSG_CANCEL,FormApp.PageNavigationType.SUBMIT)]);
  }
  else if(!bl && form_type.getChoices().length!=1){
    form_type.setHelpText('満席のため、現在はキャンセル手続きのみ可能です。');
    form_type.setChoices([form_type.createChoice(FI_MSG_CANCEL,FormApp.PageNavigationType.SUBMIT)]);
  }
}

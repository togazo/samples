/*******************************
* main.gs
* Author: TGA (@togazo / TGABOOK)
* Create: 2020.4.28
* Update: 2020.12.30
*
* NOTE:
*   DojoCon Japan 2020 のワークショップ用にカスタマイズ
*
*******************************/
const DEBUG_MODE = 0;  //デバッグモード(0=本番、1=デバッグ)

/* 定数 */
//フォームのID
const ID_FORM = '[FORM_ID]';
const FM_ = FormApp.openById(ID_FORM);

//シートのID
const ST_NAM_ENTRY = '[SHEET NAME]'; //フォームの回答が入るシート名
const SS_ = SpreadsheetApp.getActiveSpreadsheet();
const ST_ = SS_.getSheetByName(ST_NAM_ENTRY);

//グループ管理用メールアドレス
const GRP_ID = '[GOOGLE GROUP ADDR]';

//
const FI_MSG_ATND = '参加を申し込む';
const FI_MSG_CANCEL = '申し込みを取り消す';
const MSG_FINAL_TITLE = '最終確認'; //フォームの最後の質問
const SEL_USER_OPEN = '登録なし';
const SEL_USER_ALLOW = '登録あり（フォームから申し込んで参加経験有）';
const NG_ = 'x';
const PRJ_NAME = '[PROJECT NAME]'; //今回のイベントを主催する組織(?)名（*要書換）
const PRJ_URL = '[PROJECT URL]'; //組織のWebサイト（*要書換）
const MAIL_ADDR_ADMIN = '[ADOMIN MAIL ADDR]'; //組織または担当者の連絡先（*要書換）
const ADMIN_NAME = `[ADOMIN NAME]`; //担当者名
const SBJ_COPY_REGIST = '[MAIL SUBJECT]'; //申し込み控えのメール件名
const EVENT_NAME = '[EVENT NAME]'; //今回のイベント名
const MTG_URL = '[SPREAD SHEET URL]'; //参加用URL
const TIME_EVT_ST = '[YYYY-MM-DD HH:mm:SS]'; //イベント開始時刻
const TIME_EVT_ED = '[YYYY-MM-DD HH:mm:SS]'; //イベント終了時刻
const LOG_ART='[FOLDER URL]'; //スプレッドシートのログが置いてあるフォルダ
const COL_LAST = "I".charCodeAt(0) - "A".charCodeAt(0) + 1; //一番最後（参加許可）の列番号（忘れるので式にしてる…）

const WEBHOOK_URL = '[SLACK WEBHOOK URL]'; //Slack通知
const WEBHK_CH = '#[SLACK CHANNEL NAME]';

//---------------------
/*
* 設置時に一度だけ手動で実行する
*/
function init()
{
  ScriptApp.getProjectTriggers().forEach(i=>ScriptApp.deleteTrigger(i)); //不要なトリガーの除去
  ScriptApp.newTrigger("EvtFormSubmitSS").forSpreadsheet(SS_).onFormSubmit().create(); //トリガーの作成:フォーム送信時（スプレッドシート）
  UpdateSchedule();
}

/*
* 募集要項の更新
*/
function UpdateSchedule()
{
  //トリガーの作成（フォームの回答を締め切る）
  ScriptApp.getProjectTriggers().forEach(i=>{if(i.getHandlerFunction()!='EvtFormSubmitSS') ScriptApp.deleteTrigger(i)}); //不要なトリガーの除去
  let trgDate = new Date(TIME_EVT_ED);
  trgDate.setMinutes(trgDate.getMinutes());
  ScriptApp.newTrigger("EvtCloseForm")
           .timeBased()
           .at(trgDate)
           .create();
  UpdateSpreadSheet(); //スプレッドシート更新
  UpdateForm(); //フォーム更新
}

//フォームを開ける（トリガーは手動）
function EvtOpenForm(){
  FormApp.openById(ID_FORM).setAcceptingResponses(true);
}


//メール送信
function SendAutoMail(recipient,subject,body)
{
  if(typeof recipient != 'string'){ return;}
  body += TemplateMailFooter();
  subject = `【${PRJ_NAME}】${EVENT_NAME} ${subject}`;
  
  let param = { name: `${PRJ_NAME}`,
    noReply:true,
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


/*------------------------------
チャットツールに通知を飛ばす
------------------------------*/
function SendChatNotification_(message)
{ //なんでわざわざラッピングしているかというと、Discordに通知するときがあるから
  SendSlackNotification_(message);
}

//Slack
function SendSlackNotification_(message) {
  const url = WEBHOOK_URL;
  const channel = WEBHK_CH;
  const payload = (undefined === channel || "" == channel) ? JSON.stringify({ "text": message }) : JSON.stringify({ "text": message, "channel": channel });
  const options = { 'method': 'post', 'contentType': 'application/json', 'muteHttpExceptions': true, 'payload': payload };
  if (!DEBUG_MODE) {
    const r = UrlFetchApp.fetch(url, options);
    Logger.log("Slack(Webhook): %s " + r.getContentText(), r.getResponseCode());
  }
  else { Logger.log("Slack(Webhook): %s", message); }
}


//Google グループにメンバーを追加
//return - true:追加成功 false:追加失敗
function AddGroupMember(usraddr) {
  //メンバー追加
  try {
    //グループにユーザーを追加
    let res = AdminDirectory.Members.insert({email: usraddr},GRP_ID);
    Logger.log(res);
  } catch(error) {
    //グルーブ設定
    Logger.log('name：'　+ error.name);
    Logger.log('message：'　+ error.message);      
    return false;
  }
  return true;
}

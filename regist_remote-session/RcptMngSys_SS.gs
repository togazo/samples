/*******************************
* RcptMngSys_SS.gs
* Author: TGA (@togazo / TGABOOK)
* Create: 2019.9.18
* Update: 2021.1.2
*
* NOTE:
*   回答はいつでも編集できる
*   重複申込みは削除
*   直ぐ反応を返す場合は後送URLを発行しない
*   確認メール送信時にNG垢に警告（エラー）メールを送る
*   エラーを返された場合も編集後に正常に受領可能
*   招待URLの一斉送信時にNG垢はスルーする
*   元々申し込んでいないのに「取り消し」した人にも取り消し手続き完了メールを送らない（メールの送信回数の制限を考慮して無駄に送信しない）
*   よく変更する変数（定数）の値は「プロジェクトのプロパティ」>「スクリプトのプロパティ」で管理
*
*   シート…
*   ・AllowUser（ホワイトリスト）のほか、DisallowUser（ブラックリスト）をつくった
*   ・キャンセルと重複、写真撮影可否が分かるように条件式を入れて色分け
*
*******************************/

/*
*
* 申込みデータを管理しているシートの更新
*
*/
function UpdateSpreadSheet()
{
  //申し込み回答の更新: シートの複製（データのバックアップ）、シートのデータのクリア
  let row = ST.getLastRow();
  let n = row - 1;
  if(n > 0){
    ST.copyTo(SS); //データのバックアップ
    ST.deleteRows(2,n); //完全に削除しないと過去に入力があった行を飛ばして入力されてしまう（フォーマットをクリアしてもダメ）
    ST.insertRows(2,n); //削除した分の行を追加
  }
  UpdateTempNotAdmit(); //テンプレートの更新

  //受付回答の更新
  row = ST_MNG.getLastRow();
  n = row - 1;
  if(n > 0){
    ST_MNG.deleteRows(2,n);
    ST_MNG.insertRows(2,n);
  }
}

//--------------------
/*
*
* イベント
*
*/
//フォーム送信時
function EvtFormSubmitSS(e)
{
  if(!e){ return; }
  if(!ST){Logger.log(`シート「${ST_NAM_ENTRY}」がないので処理続行不可`); return;} //デバッグ時など
  if(ST_NAM_MNG==SS.getActiveSheet().getName()) FormSubmitSS_manage(e);
  else FormSubmitSS_entry(e);
}

//フォーム送信（申込者管理）
function FormSubmitSS_manage(e)
{
  const nowIdx = e.range.getRow();
  if(nowIdx < 2){ return; } //記録見つからず→処理終了
  const mng_item = ST_MNG.getRange(`A${nowIdx}:C${nowIdx}`).getValues()[0];
  let timeSt = new Date(TIME_EVT_ST);
  timeSt.setMinutes(timeSt.getMinutes() - MIN_EVT_ST_BFR);
  const timeUser = new Date(mng_item[0]);
  if(DEBUG_MODE){ Logger.log(`timeUser >= timeSt → ${timeUser}>=${timeSt}(${timeUser >= timeSt})`); }
  const flg = (mng_item[2]=='承認')?'':'非承認';
  const r = ST.getLastRow();
  let db = ST.getRange(2,1,r-1,COL_LAST).getValues();
  String(mng_item[1]).split(', ').forEach(i=>{ //※アイテムがひとつの時にNumber型と判定されてsplit()が実行エラーを起こすので、念のためにString()にする。
    if(--i<0) return;
    db[i][I_ALW]=flg;
    if(flg) SendAutoMail(db[i][I_ML],MLTI_NOTADMIT,`${db[i][I_NM]} 様\n\n${TemplateMailNotAdmit()}`) //非承認のメールを送る
    else if(timeUser >= timeSt){
      if(DEBUG_MODE){Logger.log("初参加者の即時承認");}
      SendAutoMail(db[i][I_ML],MLTI_COMP_INVITE,`${db[i][I_NM]} 様\n\n${TemplateMailInvite()}`);
    }
  });
  db = db.map(i=>[i[I_ALW]]);
  ST.getRange(2,I_ALW+1,r-1,1).setValues(db); //「参加許可」列の更新
  FM.setDescription(TemplateFormDesc()); //フォームの申込み人数の更新
  UpdateMngForm(); //管理フォームの更新
}

//フォーム送信（申し込み処理）
function FormSubmitSS_entry(e)
{
  let a = new Array(COL_LAST).fill('');
  let nowIdx = e.range.getRow();
  if(nowIdx < 2){ return; } //記録見つからず→処理終了
  a = ST.getRange(nowIdx,1,1,COL_LAST).getValues()[0];
  a[I_WR]=Number(a[I_WR])+1; //書き込み回数の計上
  ST.getRange(nowIdx,I_WR+1,1,1).setValue(a[I_WR]);
  //スプレッドシート側の重複・取り消しデータの処理
  const delnum = DelItemSS(nowIdx,a[I_ML],(a[I_TY]==FI_MSG_ATND?false:true));
  //キャンセルの場合はメールを送って終了
  if(a[I_TY]==FI_MSG_CANCEL && (delnum>1||a[I_WR]>1)){ //取消手続きしてない書き込みが2件以上or2回以上書き込み
    SendAutoMail(a[I_ML],MLTI_CANCEL, TemplateMailCancel());
    if(!DEBUG_MODE){ SendChatNotification_(`キャンセル: ${a[I_NM]!=''?a[I_NM]:GetEntryUserName(a[I_ML])}`); }
  }
  //フォームの申込み人数の更新
  if(a[I_TY]==FI_MSG_CANCEL){
    FM.setDescription(TemplateFormDesc());
    SelectAtndCncl(true); //満席から人数が減ったら選択肢を戻す
    return;
  }
  //身元の確認（ブラックリストのユーザか）
  if(ChkDisallowUser_(a[I_ML]))
    return SetNotAdmit(nowIdx,a[I_ML],a[I_NM],'身元NG',TemplateMailError());
  //身元の確認（ゲスト枠）
  if(a[I_EXP]==SEL_USER_OPEN){
    ST.getRange(nowIdx,I_ALW+1,1,1).setValue(FLAG_CHK_ALLOW);
    SendAutoMail(a[I_ML],MLTI_URVW,`${a[I_NM]} 様\n\n${TemplateMailUReview(a)}`);
    if(!DEBUG_MODE){ SendChatNotification_(`要・受付内容確認: ${a[I_NM]}`); }
    UpdateMngForm(); //管理フォームの更新
    return;
  }
  //身元の確認（常連枠）
  if(!ChkAllowUser_(a[I_ML])) //ホワイトリストのユーザか
    return SetNotAdmit(nowIdx,a[I_ML],a[I_NM],'常連以外',TemplateMailError());
  //会場の定員の確認
  const r=ST.getLastRow();
  const nr = ST.getRange(2,1,r-1,COL_LAST).getValues()
                .filter(e=>(e[I_TY]==FI_MSG_ATND && e[I_CNCT]>0 && e[I_ALW]==""))
                .reduce((ac, v) => ac + Number(v[I_CNCT]), 0)
  if(MAX_RM < nr)
    return SetNotAdmit(nowIdx,a[I_ML],a[I_NM], FLAG_NUM_MAX, TemplateMailError());
  if(MAX_RM == nr){SelectAtndCncl(false);} //満席の場合はフォームの選択肢を変更する

  //即時対応の場合はURL発行、それ以外は編集アドレスありの確認メール
  const timeUser = new Date(a[I_DT]);
  let timeSt = new Date(TIME_EVT_ST);
  timeSt.setMinutes(timeSt.getMinutes() - MIN_EVT_ST_BFR);
  if(DEBUG_MODE){ Logger.log(`timeUser >= timeSt → ${timeUser}>=${timeSt}(${timeUser >= timeSt})`); }
  if(timeUser >= timeSt){ //URL即時発行
    if(DEBUG_MODE){Logger.log("即時URL発行");}
    SendAutoMail(a[I_ML],MLTI_COMP_INVITE,`${a[I_NM]} 様\n\n${TemplateMailInvite()}`);
  }else{ //申し込み完了の通知
    if(DEBUG_MODE){Logger.log("申し込み完了の通知");}
    SendAutoMail(a[I_ML],MLTI_COMP,`${a[I_NM]} 様\n\n${TemplateMailComp()}`);
  }
  FM.setDescription(TemplateFormDesc()); //フォームの申込み人数の更新
  if(!DEBUG_MODE){ SendChatNotification_(`申し込み（${a[I_CNCT]}台）: ${a[I_NM]}${a[I_FR]?`\n${a[I_FR]}`:''}`); }
}


/*
 * SetNotAdmit
 * 承認NG処理（シートに理由を書く、メールを送る）
 */
function SetNotAdmit(idx,ml,nm,rsn,msg)
{
  ST.getRange(idx,I_ALW+1,1, 1).setValue(rsn);
  SendAutoMail(ml,MLTI_NOTADMIT,`${nm} 様\n\n${msg}`);
  return 0;
}


/*
* EvtSendReminder
* 送信時間になったのでメールを送る
* 戻り値: 対応した件数
*/
function EvtSendReminder()
{
  //メールアドレスのリスト（配列）を作成する
  const row = ST.getLastRow();
  if(row < 2) return 0;
  let cmpD = new Date(TIME_EVT_ST);
  cmpD.setMinutes(cmpD.getMinutes() - MIN_EVT_ST_BFR);
  const addrs = ST.getRange(2,1,row-1,COL_LAST)
                   .getValues()
                   .filter(function(elm){ //申込取消、重複、登録外の申込者、即時対応の時間以降を除外
                     const d = new Date(elm[I_DT]);
                     return ((elm[I_ML] != "" && d < cmpD) && (elm[I_CNCT]>0 && (elm[I_AGR] != "" && elm[I_ALW] == "")));
                   })
  //URLを通知する
  const body = TemplateMailInvite();
  addrs.forEach(a=>SendAutoMail(a[I_ML],MLTI_INVITE,`${a[I_NM]} 様\n\n${body}`));
}


/*
 * UpdateTempNotAdmit
 * 手動でメール送信するとき用のテンプレート
 *
 */
function UpdateTempNotAdmit(){
  const smt = SS.getSheetByName('MailTemplate');
  if(smt){
    smt.getRange('B2:B3').setValues([[`【${PRJ_NAME}】${EVENT_NAME} ${MLTI_NOTADMIT}`],[`[申込者名] 様\n\n${TemplateMailNotAdmit()}`]]);
  }
}


//-----------
/*
* DelItemSS
* スプレッドシート上の特定の同じメールアドレスのデータを消す
* b=true: mailに合致するものを全て消す（申込取消）、false: IDがid以外のモノを消す（重複削除）
* 戻り値: 対応したデータ数
*/
function DelItemSS(i,mail,b)
{
  let n = 0;
  const row = ST.getLastRow();
  if(row < 2){ return 0; }
  i-=2; //シートの行番号は1〜なものの、実質2行目から取得→配列に入れたので2減らして調整
  let items = ST.getRange(2,1,row-1,COL_LAST).getValues().map(function(elm,idx,arr){
    const d1 = new Date(elm[I_DT]);
    const d2 = new Date(arr[i][0]);
    if(i==idx && elm[I_ALW]!=''){elm[I_ALW]='';} //1度該当データのフラグを消す
    if((elm[I_ML]==mail && (d1<=d2 && elm[I_ALW]=='')) && (b || (!b && idx != i)))
    {
      elm[I_ALW] = b ? '申込取消' : '重複削除';
      n++;
    }
    return elm;
  });
  ST.getRange(2,1,row-1,COL_LAST).setValues(items);
  return n;
}

/*
*
* 認証されたユーザかどうか（リストにあるかどうか）チェックする
*
*/
let list_allow_;
let list_disallow_;
//Allow(true=認証したユーザ)
function ChkAllowUser_(email)
{
  if(email === ''){return false;}
  if(list_allow_ === undefined){ list_allow_ = GetMlUserList(ST_NAM_ALLOW); }
  return (list_allow_.indexOf(email) < 0) ? false : true;
}
//Disallow(true=認証してないユーザ)
function ChkDisallowUser_(email){
  if(email === ''){return false;}
  if(list_disallow_ === undefined){ list_disallow_ = GetMlUserList(ST_NAM_DISALLOW); }
  return (list_disallow_.indexOf(email) < 0) ? false : true;
}
function GetMlUserList(stnam)
{
  let st = SpreadsheetApp.openById(ID_ST_DB).getSheetByName(stnam);
  return st.getRange(`B2:B${st.getLastRow()}`).getValues().flat();
}

/*
*
* キャンセル手続きの「だけ」の時に申込者名が取得できないので、過去の1番新しいデータから取得する
*/
function GetEntryUserName(adr)
{
  const r = ST.getLastRow();
  if(r < 2){ return '(名前なし)'; }
  const ans = ST.getRange(2,1,r-1,I_NM+1).getValues().filter(e=>e[I_NM]!=''&&adr==e[I_ML]).sort((a,b)=>(b[0].valueOf()-a[0].valueOf()));
  return (ans.length)?ans[0][I_NM]:'(名前なし)';
}

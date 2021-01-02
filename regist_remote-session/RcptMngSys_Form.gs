/*******************************
* RcptMngSys_Form.gs
* Author: TGA (@togazo / TGABOOK)
* Create: 2019.9.18
* Update: 2020.12.31
*******************************/

/*
* 入力フォームの更新
*/
function UpdateForm()
{
  //申し込みフォームをリセットする
  let fr = FM.deleteAllResponses() //全ての回答を削除
             .setAcceptingResponses(true); //フォーム再開
  fr.setTitle(`${EVENT_NAME} 参加受付フォーム`); //フォームの説明を更新
  fr.setDescription(TemplateFormDesc());
  fr.setConfirmationMessage(TemplateFormConfirm());
  SelectAtndCncl(true); //
  chAllowGuest_ON(); //常連・ゲストの両方が申込可にする
  BlMaintenanceON_(); //一旦メンテナンス中にする
  //管理用のフォームをリセットする
  FM_MNG.deleteAllResponses().setAcceptingResponses(true);
}

/*
* BlMaintenance
* メンテナンス中の表示
* false = メンテナンス終了（フォームオープン）、 true = メンテナンス中（フォームクローズ）
*/
//フォームのメンテナンス
function BlMaintenanceON_(){BlMaintenance(true);}
function BlMaintenanceOFF_(){BlMaintenance(false);}
//本体
function BlMaintenance(bl)
{
  if(!bl){
    FormApp.openById(ID_FM).setAcceptingResponses(true);
    SetCustomClosedFormMessage_("申込み受付期間が終了したため");
    return;
  }
  FormApp.openById(ID_FM).setAcceptingResponses(false).setCustomClosedFormMessage(TemplateFormCloseSysMaint()); //メンテナンス中
}


//回答を受け付けていないときの回答者へのメッセージ
function SetCustomClosedFormMessage_(reason)
{
  const o = {reason:reason};
  FormApp.openById(ID_FM).setCustomClosedFormMessage(TemplateFormClose(o));
}


//フォームを閉じる
function EvtCloseForm(){
  SetCustomClosedFormMessage_("申込み受付期間が終了したため");
  FormApp.openById(ID_FM).setAcceptingResponses(false);
}


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


/*
* ゲストの受付をON/OFFする
*/
function chAllowGuest_ON(){chAllowGuest_(true)}
function chAllowGuest_OFF(){chAllowGuest_(false)}
function chAllowGuest_(bl_guest)
{
  let user_type, pb_entry_oth, pb_entry_rgl;
  FM.getItems().forEach(i=>{
    if(i.getType()==FormApp.ItemType.PAGE_BREAK && i.getTitle()=='ゲスト参加を希望の方')
      pb_entry_oth = i.asPageBreakItem();
    else if(i.getType()==FormApp.ItemType.PAGE_BREAK && i.getTitle()=='')
      pb_entry_rgl = i.asPageBreakItem();
    else if(i.getTitle()=='小平道場への参加経験')
      user_type = i.asMultipleChoiceItem();
  });
  if(bl_guest){
    user_type.setHelpText(`過去2年間に小平道場に参加したことのある方（見学を含む）は「${SEL_USER_ALLOW}」をお選びください。それ以外の方は「${SEL_USER_OPEN}」をお選び下さい。`);
    user_type.setChoices([user_type.createChoice(SEL_USER_ALLOW,pb_entry_rgl),user_type.createChoice(SEL_USER_OPEN,pb_entry_oth)]);
  }
  else{
    user_type.setHelpText('ただいまの時間は、過去2年間に小平道場に参加したことのある方（見学を含む）のみの受付となります。');
    user_type.setChoices([user_type.createChoice(SEL_USER_ALLOW,pb_entry_rgl)]);
  }
}


/*
 * UpdateMngForm
 * 管理フォームの更新
 *
 */
function UpdateMngForm()
{
  const r = ST.getLastRow();
  let db = ST.getRange(2,1,r-1,COL_LAST)
                  .getValues()
                 .map((v,i)=>[++i,v[I_NM],v[I_DJ],v[I_WK],v[I_FR],v[I_ALW]])
                 .filter(i=>FLAG_CHK_ALLOW==i[5]);
  FM_MNG.setDescription('現在、承認待ちのデータが' + db.length + '件あります。' + db.map(v=>`\n\n【${v[0]}】\n名前: ${v[1]}（${v[2]}）\n${v[3]}${v[4]?`\n----------\n${v[4]}`:''}`).join(''));
  FM_MNG.getItems().forEach(i=>{ 
    if(i.getTitle()=='申込者一覧'){
      if(db.length){
        i.asCheckboxItem().setChoiceValues(db.map(v=>v[0]));
      } else {
        i.asCheckboxItem().setChoiceValues([0]);
      }
    }
  });
}


//-----------
//-----------
function test_MngFormReset()
{
  //管理用のフォームをリセットする
  FM_MNG.deleteAllResponses().setAcceptingResponses(true);
  //受付回答の更新
  row = ST_MNG.getLastRow();
  n = row - 1;
  if(n > 0){
    ST_MNG.deleteRows(2,n);
    ST_MNG.insertRows(2,n);
  }
  //管理用のフォームをリセットする
  FM_MNG.deleteAllResponses().setAcceptingResponses(true);
}

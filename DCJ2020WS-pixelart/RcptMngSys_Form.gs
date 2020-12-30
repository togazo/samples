/*******************************
* RcptMngSys_Form.gs
* Author: TGA (@togazo / TGABOOK)
* Create: 2019.9.18
* Update: 2020.12.21
*
* NOTE:
*   DojoCon Japan 2020 のワークショップ用にカスタマイズ
*
*******************************/

/*
* 入力フォームの更新
*/
function UpdateForm()
{
  let f = FM_.deleteAllResponses() //全ての回答を削除
             .setAcceptingResponses(true); //フォーム再開
  //フォームの説明を更新
  f.setTitle(`${EVENT_NAME} 入場口`);
  f.setDescription(TemplateFormDesc());
  f.setConfirmationMessage(TemplateFormConfirm());
  //回答を受け付けていないときの回答者へのメッセージ
  SetCustomClosedFormMessage_("申込み受付期間外のため");
}

/*
* BlMaintenance
* メンテナンス中の表示
* false = メンテナンス終了（フォームオープン）、 true = メンテナンス中（フォームクローズ）
*/
function BlMaintenance(bl)
{  
  if(!bl){
    FormApp.openById(ID_FORM).setAcceptingResponses(true);
    SetCustomClosedFormMessage_("申込み受付期間外のため");
    return;
  }
  FormApp.openById(ID_FORM).setAcceptingResponses(false);
  FormApp.openById(ID_FORM).setCustomClosedFormMessage(TemplateFormCloseSysMaint()); //メンテナンス中
}


//回答を受け付けていないときの回答者へのメッセージ
function SetCustomClosedFormMessage_(reason)
{
  const o = {reason:reason};
  FormApp.openById(ID_FORM).setCustomClosedFormMessage(TemplateFormClose(o));
}


//フォームを閉じる
function EvtCloseForm(){
  SetCustomClosedFormMessage_("申込み受付期間外のため");
  FormApp.openById(ID_FORM).setAcceptingResponses(false);
}

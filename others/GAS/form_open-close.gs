/*
 * ファイル名: form_open-close.gs
 * 内容: アンケートフォームの開閉
 * 使い方: 1. Google フォーム でアンケートを作成する
 *        2. 回答をスプレッドシートで収集するようにする。（任意）
 *        3. フォームのスクリプトにこれを丸コピする
 *        4. TM_OP と TM_CL にそれぞれ希望の時間を入れる。
 *        5. SS_ID にシートのIDを、ST_NM にシートの名前をそれぞれ設定する（任意）
 *        6. init() を実行する。必要に応じて許可をする。
 * ポイント: イベント当日、参加後のアンケートフォームの開閉の仕事を減らしたいがために作成した
 *         これでアンケートフォームのURLを事前に案内しても先回りして回答できないね☆
 *         作業後に回答を開きっぱなしにするウッカリさんのためにスクリプト実行5分後に閉じる親切設計!（主に作者用
 *
 * 作った人: TGA
 * ライセンス: CC0 1.0
 *           https://creativecommons.org/publicdomain/zero/1.0/
 *
 */

const TM_OP = 'YYYY-MM-DD HH:mm'; //フォームを開く時間
const TM_CL = 'YYYY-MM-DD HH:mm'; //フォームを閉じる時間
const SS_ID = '/* Spread Sheat's ID */'; //回答のシートID
const ST_NM = 'フォームの回答'; //回答のシートの名前

function setFormActive() {
  FormApp.getActiveForm().setAcceptingResponses(true);
}

function setFormInactive() {
  FormApp.getActiveForm().setAcceptingResponses(false);
}

function init()
{
  ScriptApp.getProjectTriggers().forEach(i=>ScriptApp.deleteTrigger(i)); //不要なトリガーの除去
  const d = new Date();
  d.setMinutes(d.getMinutes()+5);
  ScriptApp.newTrigger('setFormInactive') //トリガーの作成（フォームを閉じる）
           .timeBased()
           .at(d)
           .create();
  ScriptApp.newTrigger('setFormActive') //トリガーの作成（フォームを開く）
           .timeBased()
           .at(new Date(TM_OP))
           .create();
  ScriptApp.newTrigger('setFormInactive') //トリガーの作成（フォームを閉じる）
           .timeBased()
           .at(new Date(TM_CL))
           .create();
  try{
    const st = SpreadsheetApp.openById(SS_ID).getSheetByName(ST_NM);
    if(!st) return;
    st.hideRows(2,st.getLastRow()-1); 
  }catch(err){Logger.log(err);}
}

/*
 * ファイル名: check_registed_mail.gs
 * 内容: 登録されたメールアドレスか確認する
 * 使い方: 1. メアドを入力して送信するだけのGoogleフォームをつくる
 *        2. フォームのスクリプトにこれを丸コピする
 *        3. 登録アドレスの一覧…要するにホワイトリスト（とNGアドレス一覧…ブラックリスト）をつくる
 *        4. フォーム送信時のトリガを作成する
 * ポイント: 別作成のイベント申し込みフォーム（登録メアド以外の人は申し込みできない）で申し込みの際、
 *          自分のメアドが登録されているのか申込者自身が確認できるようにするために作成した。
 *          なお、登録は身元確認を厳重に行うため、現在自動化してない。
 *          Qiitaでの解説は https://qiita.com/togazo/items/96dfa3979e2b81853522
 *
 * 作った人: TGA
 * ライセンス: CC0 1.0
 *           https://creativecommons.org/publicdomain/zero/1.0/
 *
 */
const ID_ST_DB = '[ID]'; //ホワイトリスト、ブラックリストを管理
const ST_NAM_ALLOW = 'AllowUser'; //ホワイトリストのシート名
const ST_NAM_DISALLOW = 'DisallowUser'; //ブラックリストのシート名
const PRJ_NAME = '[Projcet Name]'; //今回のイベントを主催する組織(?)名（*要書換）
const ADMIN_NAME = `${PRJ_NAME} [Name]`; //担当者名
const ML_HEADER = ` 様`;
const ML_FOOTER = `ーーーーーーーーーーーーーーーーーーーーーーーーー
このメールは自動送信です。
お心当たりのない場合は、他の方がメールアドレスを間違えて送信した
可能性がございますので、お手数ですが破棄して下さい。
ーーーーーーーーーーーーーーーーーーーーーーーーー`;
const ML_BODY_ALW = `${ML_HEADER}\n\nこのメールアドレスは登録されています。\n\n${ML_FOOTER}`;
const ML_BODY_DAL = `${ML_HEADER}\n\nこのメールアドレスは現在凍結されています。\n\n${ML_FOOTER}`

function FormAction(e) {
  if(!e) return;
  const em = e.response.getRespondentEmail();
  e.source.deleteResponse(e.response.getId());
  if(ChkDisallowUser_(em)){
    try{MailApp.sendEmail({ name: ADMIN_NAME,
                            to: em,
                            subject: `【${PRJ_NAME}】メールアドレスの登録確認`,
                            body: `${em}${ML_BODY_DAL}`});
    }catch(e){Logger.log(`メール送信失敗（凍結）：${e}`);}
    return;
  }
  if(!ChkAllowUser_(em)) return;
  try{MailApp.sendEmail({ name: ADMIN_NAME,
                          to: em,
                          subject: `【${PRJ_NAME}】メールアドレスの登録確認`,
                          body: `${em}${ML_BODY_ALW}`});
  }catch(e){Logger.log(`メール送信失敗（登録済）：${e}`);}
}

function ChkAllowUser_(em)
{
  let l;
  if(em === '') return false;
  if(l === undefined){
    const st = SpreadsheetApp.openById(ID_ST_DB).getSheetByName(ST_NAM_ALLOW);
    l = st.getRange(`B2:B${st.getLastRow()}`)
                    .getValues().flat();
  }
  return (l.indexOf(em) < 0) ? false : true;
}

function ChkDisallowUser_(em){
  let l;
  if(em === ''){return false;}
  if(l === undefined){
    const st = SpreadsheetApp.openById(ID_ST_DB).getSheetByName(ST_NAM_DISALLOW);
    l = st.getRange(`B2:C${st.getLastRow()}`).getValues();
  }
  const i = l.map(e=>e[0]).indexOf(em);
  return (i < 0)? 0 : l[i][1];
}

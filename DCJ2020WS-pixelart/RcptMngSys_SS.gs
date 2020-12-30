/*******************************
* RcptMngSys_SS.gs
* Author: TGA (@togazo / TGABOOK)
* Create: 2019.9.18
* Update: 2020.12.27
*
* NOTE:
*   DojoCon Japan 2020 のワークショップ用にカスタマイズ
*
*******************************/

const I_NM = 1;//おなまえ
const I_GA = 2;//Googleアカ
const I_CD = 3;//チェックインコード
const I_DBNM = 5;//DB名前
const I_DBML = 6;//DBメール
const I_WR = 7;//参加許可
const I_ALW = 8;//参加許可


/*
* 設置時に一度だけ手動で実行する
*/
function InitSS()
{
  //トリガーの作成
  let st = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("EvtFormSubmitSS").forSpreadsheet(st).onFormSubmit().create(); //フォーム送信時（スプレッドシート）
}

/*
*
* メニュー追加
*
*/
//----------------------

/*
*
* 申込みデータを管理しているシートの更新
*
*/
function UpdateSpreadSheet()
{
  //シートのデータのクリア
  let row = ST_.getLastRow();
  let n = row - 1;
  if(n > 0){
    ST_.deleteRows(2,n); //完全に削除しないと過去に入力があった行を飛ばして入力されてしまう（フォーマットをクリアしてもダメ）
    ST_.insertRows(2,n); //削除した分の行を追加
  }
}

//フォームのメンテナンス
function BlMaintenanceON(){BlMaintenance(true);}
function BlMaintenanceOFF(){BlMaintenance(false);}

//--------------------
/*
*
* イベント
*
*/
//フォーム送信時
function EvtFormSubmitSS(e)
{
  let a = new Array(COL_LAST).fill('');
  if(!ST_){
    Logger.log(`シート「${ST_NAM_ENTRY}」がないので処理続行不可`); return;
  }
  if(!e){ return; }
  let nowIdx = e.range.getRow();
  if(nowIdx < 2){ return; } //記録見つからず→処理終了 
  
  a = ST_.getRange(nowIdx,1,1,COL_LAST).getValues()[0];
  a[I_WR]=Number(a[I_WR])+1; //書き込み回数の計上
  ST_.getRange(nowIdx,I_ALW,1,1).setValue(a[COL_LAST-2]);
  //身元の確認
  const alwusr = ChkAllowUser_(a[I_CD]);
  let cellAllow = (!alwusr)? NG_ : '';
  if(a[I_ALW] != cellAllow){
    a[I_ALW] = cellAllow;
    ST_.getRange(nowIdx,COL_LAST,1, 1).setValue(cellAllow);
  }
  if(!alwusr){ //身元NG
    SendAutoMail(a[I_GA],"入場手続きが完了していません",`${a[I_NM]} 様\n\n${TemplateMailError()}`);
    return;
  }
  ST_.getRange(nowIdx,I_DBNM+1,1, 2).setValues([[alwusr[0],alwusr[1]]]); //取得情報の書き込み

  //グループ追加
  if(!AddGroupMember(a[I_GA])){ //追加失敗
    ST_.getRange(nowIdx,I_ALW+1,1, 1).setValue("G追加失敗");
    SendAutoMail(a[I_GA],"入場手続きが完了していません",`${a[I_NM]} 様\n\n${TemplateMailError()}`);
    return;
  }
  //URL発行
  if(DEBUG_MODE){Logger.log("即時URL発行");}
  SendAutoMail(a[I_GA],"入場手続きが完了しました",`${a[I_NM]} 様\n\n${TemplateMailInvite()}`);
  if(!DEBUG_MODE){ SendChatNotification_(`ドット絵参加: ${a[I_NM]}`); }
}

/*
*
* 認証されたユーザかどうか（リストにあるかどうか）チェックする
* return: true:[Array]ユーザデータ、false:取得失敗
*
*/
const ID_ST_DB = '[DOORKEEPER EXCEL]'; //ホワイトリスト（Doorkeeperの申込みデータ）
const ST_NAM_ALLOW = '参加者'; //A-C列
function ChkAllowUser_(chkcode)
{  
  if(!chkcode){return;}
  const st = SpreadsheetApp.openById(ID_ST_DB).getSheetByName(ST_NAM_ALLOW);
  const arr = st.getRange(`A2:C${st.getLastRow()}`).getValues(); //[Array] 0:名前,1:メアド,2:チェックインコード
  const idx = arr.findIndex(i=>i[2]==chkcode);
  if(idx<0){return;}
  return arr[idx];
}

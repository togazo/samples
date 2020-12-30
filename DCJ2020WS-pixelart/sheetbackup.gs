/*******************************
* sheetbackup.gs
* Author: TGA (@togazo)
* Create: 2020.11.14
* Update: 2020.12.27
*
* NOTE:
* ・任意のスプレッドシートを、任意のフォルダー内に定期的にゴッソリまるごとバックアップする
* ・バックアップ処理は、決められた日時から決められた日時までの間、n分おきに実行する
*
* HOW TO USE:
* 1. []の部分に適切な値を設定する。
* 2. SetTrigger() を実行する。あとはテキトーに放置する。
*
*******************************/
/*
 * 特定の期間に、定期的に（分単位で）特定の関数を実行する
 * （特定の時間から定期的にトリガーを回し、特定の時間になったら不要になったトリガーを削除する）
 */
const CRN_FNC = 'BackupSheet'; //定期実行したい関数名
const CRN_MIN = 5; //定期実行の間隔(分) ※注:選べる値は 1/5/10/15/30
const CRN_DATE_ST = new Date('[YYYY-MM-DD HH:mm:ss]'); //定期実行の開始日時
const CRN_DATE_ED = new Date('[YYYY-MM-DD HH:mm:ss]'); //定期実行の終了日時
function SetTrigger() { //指定時刻に開始・終了するトリガーをセット
  ScriptApp.newTrigger('SetTrigger4EveryMin').timeBased().at(CRN_DATE_ST).create();
  ScriptApp.newTrigger('DeleteTrigger4EveryMin').timeBased().at(CRN_DATE_ED).create();
}
function SetTrigger4EveryMin() { //トリガーをセットし、そのIDをスクリプトプロパティに登録する
  PropertiesService.getScriptProperties().setProperty('SPROP_TRGID', ScriptApp.newTrigger(CRN_FNC).timeBased().everyMinutes(CRN_MIN).create().getUniqueId());
}
function DeleteTrigger4EveryMin() { //特定のIDのトリガーを削除する
  const uid = PropertiesService.getScriptProperties().getProperty('SPROP_TRGID');
  if(uid===null){Logger.log('DeleteTrigger4EveryMin(): トリガーIDなし、実行できず。'); return;}
  ScriptApp.getProjectTriggers().filter(i=>i.getUniqueId()==uid).forEach(i=>ScriptApp.deleteTrigger(i));
}

/*
 * シートのバックアップ
 */
const SSID = '[sheet's ID]'; //複製元のスプレッドシート
const STNM = '[sheet's name]'; //（同上）の複製したいシート
const FID='[folder's name]'; //バックアップ保存フォルダ
function BackupSheet()
{
  let nm = (new Date()).valueOf().toString(16); //バックアップファイル名は時間を16進数表記にしたもの
  let fld = DriveApp.getFolderById(FID);
  let ssNew = SpreadsheetApp.create(nm);
  let del_st = ssNew.getActiveSheet();
  SpreadsheetApp.openById(SSID).getSheetByName(STNM).copyTo(ssNew).setName(nm); //ファイルの複製とシート名変更
  DriveApp.getFileById(ssNew.getId()).moveTo(fld); //ファイルの移動
  ssNew.deleteSheet(del_st); //不要なシートの削除
}



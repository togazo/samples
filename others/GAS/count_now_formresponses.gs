/*
 * ファイル名: count_now_formresponses.gs
 * 内容: 現在のフォームの回答数を定期的にDiscordに投げ飛ばす
 * 使い方: 1. Google フォーム でアンケートを作成する
 *        2. フォームのスクリプトにこれを丸コピする
 *        3. DiscordのWebhookのURLを貼り付ける
 *        4. このフォームを何日まで実行するか決めて END_DATE に設定する（書式は YYYY-MM-DD HH:mm）
 *        5. トリガーを「日付ベースのタイマ」で登録する（時間設定は任意）
 *
 * 作った人: TGA
 * ライセンス: CC0 1.0
 *           https://creativecommons.org/publicdomain/zero/1.0/
 *
 */
const END_DATE = '[YYYY-MM-DD HH:mm]'; //このコードの実行をやめる日時
function myFunction() {
  const m = `現在の回答数 ${FormApp.getActiveForm().getResponses().length}`; //回答数
    const u = '[Webhook url]'; //Discord の Webhook URL を貼付
  const o = {'method':'post',
              'contentType':'application/json',
              'payload':JSON.stringify({'content':m})};
  try{
    const r = UrlFetchApp.fetch(u, o);
    Logger.log(`code=${r.getResponseCode()}`); //204ならヨシ!!
  }catch(e){ Logger.log(e); }
  const nd = new Date(); //now date
  const cd = new Date(END_DATE); //for compire
  if(nd>=cd)ScriptApp.getProjectTriggers().forEach(i=>ScriptApp.deleteTrigger(i));
}

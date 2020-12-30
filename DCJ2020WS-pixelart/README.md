# ドット絵おえかき参加者管理ツールのスクリプト集
## なにこれ?
去る2020年12月27日に開催された [DojoCon Japan 2020](https://dojocon2020.coderdojo.jp/) 内のワークショップ企画「[ドット絵おえかき](https://dojocon-japan.doorkeeper.jp/events/114698)」で使ったスクリプト集です。

## 構成

- main.gs
- RcptMngSys_Form.gs
- RcptMngSys_SS.gs
- template.gs
- sheetbackup.gs

sheetbackup.gs は別ファイルでも他と一緒に管理してもOK

## 使い方
ざっくり説明するとこんなところ。

1. Google Form を新規作成する（先に質問項目を一通り作成する）。
2. Form の回答をスプレッドシートに転記するように設定する
3. 回答を収集するスプレッドシートを開き、[ツール]>[スクリプトエディタ]で開いたIDEに上記スクリプトを貼り付ける。

![フォームの項目](https://raw.githubusercontent.com/togazo/samples/master/DCJ2020WS-pixelart/form_setting.png)

### 備考
「チェックインコード」って何?それのデータの作り方は？など、詳細は一旦割愛。

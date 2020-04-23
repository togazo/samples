# 技術書典7合同誌のサンプルコード
## 概要（これなに）
これは2019年9月22日に開催の「技術書典7」で頒布した合同誌「11人のピーターパンによるコーディング芋煮会」の記事「サーバー費用0円!Webイベントの受付管理のためのGoogle Apps Script入門」内で紹介したソースコードのサンプルを公開したモノである。

- Readme.md : 本ドキュメント
- RcptMngSys_Form.gs : フォーム側のプロジェクト「RcptMngSys_Form」のソースコード
- RcptMngSys.gs : スプレッドシート側のプロジェクト「RcptMngSys」のソースコード

## 使い方
1. 記事に従い、RcptMngSys_Form.gs と RcptMngSys.gs を適宜コピー&ペーストする
2. RcptMngSys.gsの、「定数」の一部の値を、自分の環境に合わせて書きかえる
3. 記事に従い、実行する。

### 補足
- RcptMngSys.gs 内の myFunction() や InitForm() はよしなに…
- トリガの設定方法などは記事参照
- 筆者自身まだまだ大いに未熟ゆえにソースコードは随時改善する予定
- もし、ソースコードを見て気になる点などが見つかった場合は、togazo/samplesのIssueから提案いただいたりPull&Request頂けると幸いです!!

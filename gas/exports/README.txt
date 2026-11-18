貼り付け用GAS全文（2プロジェクトに分離）

01_receiver_save.gs.txt
  既存のフォーム受信・保存用GASへ、コード全文を置き換えてください。
  既存のTOKEN・FORM_IDと、processSubmissionsの1分おきのトリガーを維持します。
  ダッシュボード用プロジェクトには貼り付けません。

02_dashboard.gs.txt
  CC第二ダッシュボードのGASへ、コード全文を置き換えてください。
  既存のHTML（indexなど）と、autoRefreshのトリガーを維持します。
  このコードにはフォーム受信処理を含みません。
  同じプロジェクト内に以前の受信用コードが別ファイルで残っている場合は重複を解消してください。

反映後の確認
1. 受信用GASのsetupStatisticsDataを一度実行し、非表示シートを準備します。
2. 受信用GASのdiagnoseStatisticsReceiverで、フォーム送信先URLとシートの存在を確認できます。
3. TelephoneToolの「自分の集計を送信」で当日のデータを再送します。
4. 受信用GASのprocessSubmissionsを実行するか、既存の1分おきの実行を待ちます。
5. 受信用GASの実行ログで、schemaVersion:2、intervalSheet、intervalCountを確認します。
6. TelephoneToolの「全員の集計」を更新します。

既存フォームが稼働している場合、setupでフォームを作り直す必要はありません。
フォームの操作権限は受信用GASのトリガー作成者が承認します。
ダッシュボード側に誤って追加したprocessSubmissionsのトリガーだけ削除し、
元から受信用GASにあるprocessSubmissionsと、ダッシュボードのautoRefreshは維持してください。

このフォルダーは貼り付け用の出力です。受信用の実装元はgas/call_statistics.gsです。
元のdashboard_with_call_statistics.gs.txtは互換名として残していますが、現在の内容は02と同じです。

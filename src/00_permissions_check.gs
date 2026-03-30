/**
 * 権限チェック専用関数。
 * この関数を実行することで、プロジェクト全体が必要とする
 * すべてのOAuthスコープの認証ダイアログを発生させる。
 * 副作用ゼロ（読み取り・外部通信はいずれも無害な操作のみ）。
 */
(function checkAllPermissions() {
  console.time("全権限チェック");
  try{

  // 1. Googleドキュメント（read）
  DocumentApp.getActiveDocument();

  // 2. Googleスプレッドシート（read/write スコープ要求）
  // getActiveSpreadsheet()はnull返却でも権限要求は発生する
  SpreadsheetApp.getActiveSpreadsheet();

  // 3. 外部サービスへの接続
  // Googleの204エンドポイント：レスポンスボディなし・ログなし・完全無害
  UrlFetchApp.fetch("https://www.gstatic.com/generate_204", {
    muteHttpExceptions: true
  });

  // 4. スクリプトプロパティ（読み取りのみ・値は使わない）
  PropertiesService.getScriptProperties().getKeys();

  // 5. HtmlService（スコープ不要だが念のため疎通確認）
  HtmlService.createHtmlOutput("");
  }catch(e){
    console.error(e.message);
  }

  console.timeEnd("全権限チェック");
})();// 無駄に実行して権限を確保
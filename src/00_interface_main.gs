//00_interface_main.gs : 個別の年間計画レビューから直接呼ばれる関数を置いておく場所

/**
 * @typedef {Object} RunFunctionRequest
 * @property {string} id - 実行したい関数を指定するid。
 * @property {Object} [context] - 実行されるときに関数に渡したい引数を詰め込んだオブジェクト。
 */
/**
 * 任意の関数を実行できるようにするための抽象関数。
 * オブジェクト requestを受け取り、idとcontextを取り出してFunctionRunnerServiceに渡す。
 *
 * 呼び出せる関数のレパートリーを変更する方法はDocumentを参照して下さい。
 * @param {RunFunctionRequest} request - id{string} と context{object} (contextはなくてもよい)が入った、関数実行用のオブジェクト。
 */
function abstractFunction(request) {
  //requestのバリデーション
  if (!request || typeof request !== "object") {
    throw new Error(
      "abstractFunctionの引数はオブジェクトである必要があります。",
    );
  }

  //requestの中身のバリデーション
  const id = request["id"]; //関数特定用のid
  const context = request["context"]; //関数に渡す用のコンテキストオブジェクト

  if (!id || typeof id !== "string") {
    throw new Error(
      `渡されたオブジェクトに適切なid(string)が入っていません。:\n${JSON.stringify(request)}`,
    );
  }

  if (context !== undefined && typeof context !== "object") {
    throw new Error(
      `渡されたオブジェクトに不適切なcontext{object}が入っていました。contextはオブジェクトにするか、そもそも入れないかのいずれかにしてください。:\n${JSON.stringify(request)}`,
    );
  }
  // NOTE(2026-03-14): 今後機能が拡張されるようなら「ガントチャートテンプレート」ライブラリのような柔軟な実装にするが、いまはそこまで不要なのでswitchで簡易的に実装
  try {
    return runById(id);
  } catch (e) {
    console.error(e);
    GASRefferenceSheetLogService.error(e.message + "\n" + e.stack);
    throw e;
  }
}

/**
 * idごとに関数を振り分ける関数。
 * @param {idArg} 関数を特定する識別子
 */
function runById(idArg) {
  const id = String(idArg).trim();
  switch (id) {
    case "genNewNotForNextMeeting":
      return genNewNotForNextMeeting();
    default:
      throw new Error(`不明なid : ${id}`);
  }
}

/**
 * 生徒の年間計画レビューのonOpenで呼ばれる関数。
 * ここで呼ばれる関数はこのライブラリ内のものなので違和感があるかもしれませんが、呼び出し元の目線で定義する必要があるので、
 * きちんと「YearyPlanReviewLib.」接頭辞をつけて関数名を登録する必要があります。これは各生徒の年間計画レビューのスクリプト内でのこのライブラリのインポート名です。
 * このインポート名はどの呼び出し元でも共通しているはずです。複数種類への対応は考慮する必要はありません。
 */
function onOpenAction() {
  try {
    const ui = DocumentApp.getUi();
    ui.createMenu("小見出し生成")
      .addItem("面談メモを生成", "YearlyPlanReviewLib.createMeetingMemo")
      .addItem("特訓メモを生成", "YearlyPlanReviewLib.createTrainingMemo")
      .addItem(
        "年間計画レビューを生成",
        "YearlyPlanReviewLib.createAnnualReview",
      )
      .addItem("月間面談を生成", "YearlyPlanReviewLib.createMonthlyMeeting")
      .addToUi();
    ui.createMenu("次回面談時注意抽出")
      .addItem("確定", "YearlyPlanReviewLib.openSidebar")
      .addToUi();
  } catch (e) {
    const message = e.message || String(e); // ユーザー向け（簡潔）
    const detail = e.stack || "No stack trace"; // 開発者向け（詳細）
    // ユーザーにはトーストで簡潔に（エラーがあることだけは伝える）
    ToastNotificationService.send(
      "メニューの読み込みに失敗しました: " + message,
    );
    // 開発者はログ（コンソール）で詳細を確認
    console.error("onOpenAction Error Details:\n" + detail);
  }
}

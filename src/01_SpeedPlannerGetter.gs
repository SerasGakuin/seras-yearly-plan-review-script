/**
 * 指定されたドキュメントID（省略時は現在開いているドキュメント）に対応する
 * 生徒のスピードプランナーSSインスタンスを返す。
 * エラーが発生した場合（対応生徒なし・複数ヒット・SS取得失敗など）は null を返す。
 *
 * @param {string | null} [docId=null] - ドキュメントID。省略またはnullの場合はアクティブなドキュメントを使用。
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet | null}
 */
function getSpeedPlannerSsForDoc(docId = null) {
  try {
    // docIdが渡されなかった場合はアクティブなドキュメントから取得
    const fileId = docId
      ? String(docId).trim()
      : String(DocumentApp.getActiveDocument().getId());

    // 生徒マスターから全生徒情報を取得
    const studentMaster = StudentMasterLib.getStudentMaster_V2();
    const allStudentsInfoArr = studentMaster.getAllStudentsDataRecordsArray();

    // DriveファイルのURLからIDだけ取り出すヘルパー
    const extractDriveFileId = (url) => {
      const str = String(url || "");
      const match = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
      return match ? match[1] : null;
    };

    // 現在のドキュメントIDに一致する生徒を絞り込む
    const matched = allStudentsInfoArr.filter((rec) => {
      return extractDriveFileId(rec.yearlyPlanReviewUrl) === fileId;
    });

    // 0件・複数件はエラー扱い
    if (matched.length === 0) {
      throw new Error(
        `このドキュメント(id: ${fileId}) に対応する生徒の情報を見つけられませんでした。`
      );
    }
    if (matched.length >= 2) {
      throw new Error(
        `このドキュメント(id: ${fileId}) に対応する生徒の情報が複数発見されました！`
      );
    }

    // 対応するスピードプランナーSSを開いて返す
    const speedPlannerUrl = matched[0].speedPlannerUrl;
    return SpreadsheetApp.openByUrl(speedPlannerUrl);

  } catch (e) {
    const errMsg = `getSpeedPlannerSsForDoc: SSの取得に失敗しました。\n${e.message}\n${e.stack}`;
    console.error(errMsg);
    GASRefferenceSheetLogService.error(errMsg);
    return null;
  }
}
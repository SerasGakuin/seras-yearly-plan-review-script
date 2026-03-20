// 02_createMeetingMemo_main.gs
/**
 * # MeetingMemoTemplateCreator
 *
 * 面談メモの型紙的なテキストをドキュメントの上に追記するマクロ。
 *
 * ###　処理の流れ
 *
 * 1. 日付からタイトルの日付表示を生成
 * 2. 個別の生徒に合わせて教科名の一覧を配列で取得
 * 3. ドキュメントを捜索してメモを挿入すべき位置を特定
 * 4. メモの挿入を実行
 *
 * ### 運用
 *
 * - 運用中に小さなエラーがあれば人間が対応する
 *
 *
 * ### 実装の概要
 *
 * 教科一覧を下にメモの項目を生成する。
 * - 教科一覧生成器オブジェクト
 * - メモ生成オブジェクト
 * この2つの提携によってメモを生成している。
 *
 *
 * ### 課題
 *
 * - メモの生成方法の柔軟性が低い。全く新しい形式(チェックボックスの追加など)に耐えられない。（しかし安易に対応させると過度な抽象化に）
 *
 *
 *
 * ### 生成例
 *
 * 2026/03/14　面談メモ　指導講師：
 * 【英語】
 *
 * 【国語】
 *
 * 【数学】
 *
 * 【理科】
 *
 * 【社会】
 *
 * 【その他】
 *
 * 以上。
 */

function createMeetingMemo() {
  genericMemoCreationFunction("面談メモ");
}
function createTrainingMemo() {
  genericMemoCreationFunction("特訓メモ");
}
function createAnnualReview() {
  genericMemoCreationFunction("年間計画レビュー");
}
function createMonthlyMeeting() {
  genericMemoCreationFunction("月間面談");
}

/**
 * 各種メモを生成する関数。エラーハンドリングなどの共通部分をここに抽出している
 * @param {string} memoTitle - タイトル
 */
function genericMemoCreationFunction(memoTitle = "面談メモ") {
  try {
    // 設定初期化
    const document = DocumentApp.getActiveDocument();
    // 生成用クラス初期化
    const subjectsCollector = new MeetingMemoSubjectService();
    const generator = new MeetingMemoTemplateCreator(
      document,
      memoTitle,
      subjectsCollector,
    );
    // 生成開始
    generator.generate();
  } catch (e) {
    const msg = `${e.message}\n${e.stack}`;
    ToastNotificationService.send(msg);
    GASRefferenceSheetLogService.error(msg);
    throw e;
  }
}

/**
 * @typedef {Object} MeetingMemoSubjectsCollector
 * @property {(doc: GoogleAppsScript.Document.Document) => string[]} fetchSubjectsArray
 */

class MeetingMemoTemplateCreator {
  /**
   * @param {Object} document ドキュメントのインスタンス。
   * @param {string} memoTitle - メモのタイトル
   * @param {MeetingMemoSubjectService} subjectArrFetcher - ドキュメントのインスタンスから教科一覧を配列で返すメソッド、fetchSubjectsArrayを持つクラス
   * @public
   */
  constructor(document, memoTitle, subjectArrFetcher) {
    // バリデーション
    if (
      !document ||
      typeof document != "object" ||
      !document.getBody ||
      typeof document.getBody != "function"
    ) {
      throw new Error(
        `グーグルドキュメントのインスタンスを渡してください！(urlではないです)\n渡されたもの(表示は参考程度に)：${JSON.stringify(document)}`,
      );
    }
    if (!memoTitle || typeof memoTitle != "string") {
      throw new Error(
        `型が違います。：${typeof memoTitle}。\nstring型の引数をメモのタイトルとしてください。`,
      );
    }
    if (
      !subjectArrFetcher ||
      typeof subjectArrFetcher != "object" ||
      !subjectArrFetcher.fetchSubjectsArray ||
      typeof subjectArrFetcher.fetchSubjectsArray != "function"
    ) {
      throw new Error(
        `subjectArrFetcherには、教科一覧を取得するメソッドを持つオブジェクトを渡してください！\n実際に渡されたもの：${JSON.stringify(subjectArrFetcher)}`,
      );
    }
    // 代入
    this._document = document;
    this._memoTitle = memoTitle;
    this._subjectArrFetcher = subjectArrFetcher;
  }

  /**
   * 生成開始メソッド。
   * @public
   */
  generate() {
    const document = this._document;
    const memoTitle = this._memoTitle;
    const subjectArrFetcher = this._subjectArrFetcher;

    // タイトルのテキストを構築
    const formattedDate = this._getDateView();
    const headingText = formattedDate + "　" + memoTitle + "　指導講師：";

    const body = document.getBody();

    // ドキュメント内の最初の見出しHEADING1の前に挿入する
    const foundHeadingIndex = this._findPosToInsertNewHeading(body);

    let headingParagraph;
    if (foundHeadingIndex != -1) {
      // 既存のHEADING1が見つかった場合、その前に挿入
      headingParagraph = body.insertParagraph(foundHeadingIndex, headingText);
    } else {
      // HEADING1が見つからない場合、ドキュメントの最初に挿入
      headingParagraph = body.insertParagraph(0, headingText);
    }
    headingParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    // 以降、追加した見出しの位置を基準にする。
    let insertIndex = body.getChildIndex(headingParagraph);

    // 項目を追加
    const subjects = subjectArrFetcher.fetchSubjectsArray(document);
    for (let j = 0; j < subjects.length; j++) {
      body.insertParagraph(++insertIndex, "【" + subjects[j] + "】");
      body.insertParagraph(++insertIndex, ""); // 空白行を追加
    }

    // 最後に「以上。」を追加
    body.insertParagraph(++insertIndex, "以上。");
    body.insertParagraph(++insertIndex, "");
  }

  /**
   * 今回の日付の表示テキストを得る。
   * @returns {string} 今回の日付の表示テキスト
   * @private
   */
  _getDateView() {
    const date = new Date();
    const year = date.getFullYear();
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    const day = ("0" + date.getDate()).slice(-2);
    return `${year}/${month}/${day}`;
  }

  /**
   * 次に挿入する見出しの位置のインデックスを特定
   * Bodyのchild要素を直接走査してHEADING1を探す
   * @param {GoogleAppsScript.Document.Body} body
   * @returns {number} 挿入位置。見つからなければ -1
   * @private
   */
  _findPosToInsertNewHeading(body) {
    const childCount = body.getNumChildren();

    for (let i = 0; i < childCount; i++) {
      const element = body.getChild(i);

      // Paragraph のみ対象
      if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const paragraph = element.asParagraph();

        if (paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          return i;
        }
      }
    }

    return -1;
  }
}

/**
 * 面談メモを生成したいドキュメントのインスタンスを受け取り、そのメモに入れる項目の配列を返すクラス。
 */
class MeetingMemoSubjectService {
  /**
   * デフォルトの下位項目の教科一覧。エラー発生時に利用する。
   * @private
   */
  get _defaultSubjectsArray() {
    return ["英語", "国語", "数学", "理科", "社会", "その他"];
  }

  /**
   * どの生徒でも必ず項目に含めるべき科目の配列。現在はこれらは最後に追記される仕様。
   * @private
   */
  get _mustIncludeSubjectsArr() {
    return ["その他"];
  }
  /**
   * 科目名（メモの下位項目）の一覧を配列で取得する
   * @param {Object} document
   * @returns {string[]} 一覧
   * @public
   */
  fetchSubjectsArray(document) {
    if (!document)
      throw new Error(
        "_fetchSubjectsArray(document): 引数がnullまたはundefinedです！",
      );
    // try-catchで包むことでエラーが起きても処理を止めない
    try {
      // ドキュメントに対応する生徒の情報を取得
      const studentInfo = this._getStudentInfoForDoc(document);
      // その生徒のスピードプランナーのurlを取得
      const speedPlannerUrl = studentInfo.speedPlannerUrl;
      // そのスピードプランナーのIOマネージャーを初期化
      const bookObj = SpreadsheetApp.openByUrl(speedPlannerUrl);
      const spIOManager =
        SpeedPlannerIOManagerLib.getSpeedPlannerIOManagerReadOnly(bookObj);
      // 今月の教材の科目を取得
      const activeSubjectsArr = spIOManager
        .getActiveMaterialsSubjects()
        .map((row) => String(row[0]).trim())
        .filter((item) => !!item && item != "");
      // 必ず必要なのは追加しておく
      const mustIncludeSubjects = this._mustIncludeSubjectsArr;
      mustIncludeSubjects.forEach((item) => activeSubjectsArr.push(item));
      // 重複排除して返す
      return [...new Set(activeSubjectsArr)];
    } catch (e) {
      // エラーが発生した場合は、デフォルト設定を返却
      const errMsg = `教材一覧の取得中にエラーが発生しました。デフォルト設定を適用します。\n${e.message}\n${e.stack}`;
      ToastNotificationService.send(errMsg, 60);
      GASRefferenceSheetLogService.error(errMsg);
      return this._defaultSubjectsArray;
    }
  }

  /**
   * @param {Object} ドキュメントのインスタンス
   * @returns {Object} 生徒情報
   * @private
   */
  _getStudentInfoForDoc(document) {
    // ドキュメントのファイルidを取得
    const fileId = String(document.getId());
    // 生徒マスターの検索機構を起動
    const studentMaster = StudentMasterLib.getStudentMaster_V2();
    // 全生徒の情報の入った配列を取得
    const allStudentsInfoArr = studentMaster.getAllStudentsDataRecordsArray();
    // このドキュメントidの生徒を探す
    const extractDriveFileId = (url) => {
      const str = String(url || "");
      const match = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
      return match ? match[1] : null;
    };
    const docStudentInfoArr = allStudentsInfoArr.filter((rec) => {
      return extractDriveFileId(rec.yearlyPlanReviewUrl) === fileId;
    });
    // 該当が0または複数ならエラー！
    if (docStudentInfoArr.length === 0) {
      throw new Error(
        `このドキュメント(id: ${fileId}) に対応する生徒の情報を見つけられませんでした。`,
      );
    } else if (docStudentInfoArr.length >= 2) {
      throw new Error(
        `このドキュメント(id: ${fileId}) に対応する生徒の情報が複数発見されました！`,
      );
    }
    // 一つだけヒットしているならそれをとりだし
    return docStudentInfoArr[0];
  }
}

//01_ToastNotificationService.gs
/**
 * Toast通知の送信を担当するクラス。
 * Toast通知とは、画面に比較的小さく表示される、一定時間で自動的に表示がきえる通知方法のこと。軽微な内容の通知にどうぞ。
 * スプレッドシートでないと正規のtoast通知は利用できないのでサイドバーで代用している
 */
class ToastNotificationService {
  /** @returns {string} @private */
  static get LOG_TAG() {
    return "ToastNotificationService:";
  }

  /** @returns {number} @private */
  static get DEFAULT_TIME() {
    return 16;
  }

  /** 通知スタック */
  static get STACK() {
    if (!this._stack) this._stack = [];
    return this._stack;
  }

  /** @returns {DocumentApp.Document | null} @private */
  static get targetDoc() {
    if (!this._targetDoc) this._targetDoc = DocumentApp.getActiveDocument();
    if (!this._targetDoc)
      console.error(`${this.LOG_TAG}\nActiveDocumentが取得できません。`);
    return this._targetDoc;
  }

  /** @returns {GoogleAppsScript.Base.Ui | null} @private */
  static get targetUi() {
    try {
      return DocumentApp.getUi();
    } catch (e) {
      console.error(`${this.LOG_TAG}\nDocumentUiが取得できません。`, e);
      return null;
    }
  }

  /**
   * Toast通知を送信する。
   * @param {string} message - 送信したい通知内容
   * @param {number} [time = 16] - toastの表示時間(秒単位。ミリ秒ではない)
   * @public
   */
  static send(message, time = 16) {
    try {
      const targetDoc = this.targetDoc;
      const targetUi = this.targetUi;
      if (!targetDoc || !targetUi) return;

      const sanitizedMsg = !message ? "" : String(message);
      const sanitizedTime = this._sanitizeTimeInput(time);

      // 通知をスタックに追加（新しいものが上）
      this.STACK.unshift({
        message: sanitizedMsg,
        time: sanitizedTime,
        ts: new Date().toLocaleTimeString(),
      });

      // HTML生成
      const html = HtmlService.createHtmlOutput(
        this._buildToastHtml_(),
      ).setTitle("通知");

      targetUi.showSidebar(html);

      console.log(
        `${this.LOG_TAG}\nToast通知として以下の内容を${sanitizedTime}秒間表示します:\n${sanitizedMsg}`,
      );
    } catch (e) {
      console.error(`${this.LOG_TAG}\ntoastの送信に失敗しました！`, e);
    }
  }

  /**
   * Toastの表示時間の入力を正規化する関数。
   * @param {any} baseTime
   * @returns {number}
   * @private
   */
  static _sanitizeTimeInput(baseTime) {
    const timeNum = Math.ceil(Number(baseTime));

    if (!Number.isFinite(timeNum)) {
      console.warn(
        `${this.LOG_TAG}\n渡されたtime:${baseTime}秒はToast表示時間として不正です。Toastの表示時間をデフォルト値${this.DEFAULT_TIME}秒に設定します。`,
      );
      return this.DEFAULT_TIME;
    } else if (timeNum <= 0) {
      console.warn(
        `${this.LOG_TAG}\n渡されたtime:${baseTime}秒はToast表示時間として不正です。Toastの表示時間をデフォルト値${this.DEFAULT_TIME}秒に設定します。`,
      );
      return this.DEFAULT_TIME;
    } else if (timeNum > 600) {
      console.warn(
        `${this.LOG_TAG}\nToastの表示時間が${baseTime}秒に指定されました。異常に長い値です。間違いでないか確認してください。`,
      );
      return timeNum;
    } else {
      return timeNum;
    }
  }

  /**
   * HTML生成
   * @returns {string}
   * @private
   */
  static _buildToastHtml_() {
    const items = this.STACK.map((n, i) => {
      const border =
        i === 0
          ? ""
          : "border-top:1px solid #ddd;margin-top:10px;padding-top:10px;";

      return `
      <div style="${border}">
        <div style="font-size:12px;color:#888">${n.ts}</div>
        <div style="
          background:#323232;
          color:#fff;
          padding:10px;
          border-radius:6px;
          margin-top:4px;
          font-size:13px;
          line-height:1.5;
        ">
          ${this._escapeHtml_(n.message)}
        </div>
      </div>
      `;
    }).join("");

    return `
<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
body{
  margin:0;
  padding:12px;
  font-family:Arial;
  background:#fafafa;
  overflow-x:hidden;
  word-break:break-word;
  overflow-wrap:anywhere;
}

*{
  box-sizing:border-box;
  max-width:100%;
}

pre,
code,
td,
th,
div,
span,
p{
  word-break:break-word;
  overflow-wrap:anywhere;
}
</style>
</head>
<body>

${items}

</body>
</html>
`;
  }

  /**
   * HTMLエスケープ
   * @private
   */
  static _escapeHtml_(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
}

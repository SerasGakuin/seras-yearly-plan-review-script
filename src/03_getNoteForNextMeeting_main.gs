// 02_GetNoteForNextMeeting.gs
/**
 * 次回以降の面談のためのヒントを生成する機能。
 */
/**
 * サイドバー展開用。
 */
function openSidebar() {
  const ui = DocumentApp.getUi();
  const template = HtmlService.createTemplateFromFile("03_index");

  let existingNotes;
  let errMsg;
  try{
    existingNotes = getExistingNotes(8).slice(0, 8).map(note => ({
      date: !!note.date && (typeof note.date === 'object') && !!(note.date.toISOString) ?
       note.date.toISOString() : String(note.date), // 後者は苦肉の策。文字列化しないとhtmlでオブジェクトをうけとれない。
      content: note.content || ""
    }));
    errMsg = null;
  }catch(e){
    existingNotes = [];
    errMsg = e.message;
  }
  
  template.existingNotes = existingNotes;
  template.errMsg = errMsg;

  const htmlOutput = template.evaluate();
  htmlOutput.setTitle("次回面談時注意抽出");
  ui.showSidebar(htmlOutput);
}

/**
* @typedef {Object} MeetingNote - 次回面談のためのメモオブジェクト
* @property {Date} date - 日付
* @property {string} content - 内容
*/
function getExistingNotesWithContext(context) {
  const maxN = (context && context.maxN) ? context.maxN : 8;
  const result = getExistingNotes(maxN);
  console.log("getExistingNotes result:", JSON.stringify(result));
  return result.map(note => ({
    date: note.date.toISOString(),
    content: note.content
  }));
}


/**
 * 既に存在しているメモをスピードプランナーから取得
 * @param {number} [maxN=8] 取得する最大件数
 * @returns {MeetingNote[]}
 * @throws スピードプランナーが見つからなかった場合・IOでエラーが発生した場合
 */
function getExistingNotes(maxN = 8) {
  const speedPlannerSs = getSpeedPlannerSsForCurrentDoc();

  if (!speedPlannerSs) {
    const currDoc = DocumentApp.getActiveDocument();
    const docId = currDoc.getId();
    const docName = currDoc.getName();
    throw new Error(`ドキュメントid:${docId}の${docName}に対応するスピードプランナーは見つかりませんでした。`);
  }

  const spIoManager = SpeedPlannerIOManagerLib.getSpeedPlannerIOManagerReadOnly(speedPlannerSs);
  return spIoManager.getLatestMeetingNotes(maxN);
}

/**
 * htmlから呼び出すための、contextを受け取るラッパー
 */
function saveNewNoteWithContext({ note }) {
  note.date = new Date(note.date);
  console.log("note.date:", note.date, typeof note.date);
  console.log("note.content:", note.content, typeof note.content);
  console.log("note全体:", JSON.stringify(note));
  saveNewNote(note);
}
/**
* 新規メモをセーブする
 * @param {MeetingNote} note -新規メモオブジェクト
*/
function saveNewNote(note) {
  // 引数チェック
  const isNoteValid = !!note && (typeof note === 'object') && !!(note.content) && !!(note.date);
  if (!isNoteValid) {
    const errMsg = `渡されたメモオブジェクトが不正です！
参考：
/**
* @typedef {Object} MeetingNote - 次回面談のためのメモオブジェクト
* @property {Date} date - 日付
* @property {string} content - 内容
*/
渡された引数：
${JSON.stringify(note)}`;
    GASRefferenceSheetLogService.error(errMsg);
    ToastNotificationService.send('新規メモを保存できませんでした。', 60);
    return;
  }
  // 現在の生徒のioManager起動
  const speedPlannerSs = getSpeedPlannerSsForCurrentDoc();
  if (!speedPlannerSs) {
    throw new Error(`このドキュメントに対応するスピードプランナーを生徒マスターから見つけられませんでした。
  docId: ${docId}`);
  }
  const spIoManager = SpeedPlannerIOManagerLib.getSpeedPlannerIOManager(speedPlannerSs);
  // セーブ
  spIoManager.appendNewMeetingNote(note);
  SpreadsheetApp.flush();
}

/**
 * 生成用
 */
function genNewNoteForNextMeeting() {
  const SEARCH_STRING = "面談メモ";
  const END_MARK = "以上。";
  const MODEL = "gpt-4o";

  const doc = DocumentApp.getActiveDocument();
  const docBody = doc.getBody().getText();
  var lines = docBody.split("\n");
  var tmpmeetingmemo = [];
  var ismeetingmemo = false;
  var count = 0;
  Logger.log(lines);
  for (i = 0; i < lines.length; i++) {
    if (lines[i].includes(SEARCH_STRING)) {
      tmpmeetingmemo.push(lines[i]);
      ismeetingmemo = true;
    } else if (ismeetingmemo && lines[i].includes(END_MARK)) {
      tmpmeetingmemo.push(lines[i]);
      count++;
      ismeetingmemo = false;
    } else if (ismeetingmemo) {
      tmpmeetingmemo.push(lines[i]);
    }
    if (count == 3) {
      break;
    }
  }
  var meetingnotes = tmpmeetingmemo.join("\n");
  // 1) 生徒名を使ってプロンプト文を作成
  const prompt = `以下は、生徒への指導報告です。この生徒の学習面・態度面の課題や問題点を洗い出し、指導者が今後特に注意を払うべきポイントを具体的に教えてください。その際、指摘する各問題点に対して、必ず指導報告の内容を根拠として引用しながら示してください。また、具体的な改善策や効果的なアドバイスについても、指導報告の内容を元に理由を添えて提示してください。

【指導報告】
${meetingnotes}

【指摘すべきポイントと根拠】
- 学習面の問題（理解不足、学習方法の問題、進捗状況など）
- 態度面の問題（学習態度、モチベーション、集中力など）

【改善策やアドバイスとその理由】
- 具体的な指導方法の提案
- 生徒との接し方やコミュニケーションのコツ
- 学習習慣定着のための実践的な工夫`;

  // 2) OpenAI API 呼び出し用のペイロードを構築
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("CHAT_GPT_API_KEY");
  Logger.log("apiKey exists: " + !!apiKey);
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: MODEL,
    messages: [
      {
        role: "system",
        content:
          "あなたは優秀な塾講師です。生徒の問題点を的確に指摘してください。",
      },
      { role: "user", content: prompt },
    ],
    temperature: 0.8,
    max_tokens: 1500,
  };

  // 3) UrlFetchApp で API を呼び出し
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${apiKey}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  // エラー処理
  if (
    response.getResponseCode() !== 200 ||
    !result.choices ||
    !result.choices[0]
  ) {
    Logger.log(response.getContentText());
    throw new Error("OpenAI API 呼び出しに失敗しました。");
  }

  // 4) 応答メッセージを取り出して返却
  var reply = result.choices[0].message.content.trim();
  return reply;
}

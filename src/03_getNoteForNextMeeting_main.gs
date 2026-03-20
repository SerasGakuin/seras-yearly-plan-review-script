// 02_GetNoteForNextMeeting.gs
/**
 * 次回以降の面談のためのヒントを生成する機能。
 */
/**
 * サイドバー展開用
 */
function openSidebar() {
  const ui = DocumentApp.getUi();
  const template = HtmlService.createTemplateFromFile("03_index");
  const htmlOutput = template.evaluate();
  htmlOutput.setTitle("次回面談時注意抽出");
  ui.showSidebar(htmlOutput);
}
/**
 * 生成用
 */
function get_note_for_next_meeting() {
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

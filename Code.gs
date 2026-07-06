/**
 * UTOLから届いた更新通知を検索し、不要な行を削除して自分宛に転送するスクリプト。
 * 2026.06.15 作成
*/

// ===== 設定 =====

// 送信者のメールアドレス
const SENDER_EMAIL = "sender@gmail.com";

// 転送先メールアドレス
const FORWARD_TO = "myadress@gmail.com";

// 毎回メールに含まれていて削除したい文言（カンマで複数指定可）
const REMOVED_TEXTS = [
  "============================================================================",
  "このメールは学習管理システム (UTOL) から自動的に送信しています。",
  "このメールへの返信はできませんのでご注意ください。",
  "履修者: xxxxxxxxxxxxx",
  "曜日, 時限：",
  "授業名: ",
  "===================================================================",
  "-システムの利用方法については、UTOLのHelpやManual または 下記URLをご確認ください。",
  "https://utelecon.adm.u-tokyo.ac.jp/utol/",
  "-メールの詳細については、各授業の担当教員にお伺いください。",
  "=================== 東京大学情報基盤センター　UTOL 担当 ======================"
];

// 処理済みのメールに付けるラベル名
const PROCESSED_LABEL = "UTOL_転送済";

// 転送時の件名
const FOWARD_SUBJECT = "UTOL通知転送";

// 転送時のアカウント名
const FOWARD_NAME = "UTOL通知転送";

// トリガーに設定する時間間隔(分)
const Trigger_Time = 10;

// 連続する空白行をどう扱うか
//   true  : 空白行をすべて削除する
//   false : 空白行は1行までにまとめる（完全には消さない）
const REMOVE_ALL_BLANK_LINES = true;
// =============================================


// main関数
function forwardFilteredEmails() {
  const label = GmailApp.getUserLabelByName(PROCESSED_LABEL);

  // 指定送信者からの、まだ処理していないメールを検索
  const query =
    'from:"' + SENDER_EMAIL + '" -label:"' + PROCESSED_LABEL + '"';

  // 現在時刻から"Trigger_Time+2"分前のタイムスタンプ
  const Trigger_Time_MS = (Trigger_Time + 2) * 60 * 1000; 
  const cutoff = new Date(Date.now() - Trigger_Time_MS);

  const threads = GmailApp.search(query, 0, 20); // 一度に最大20スレッド

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      // 受信時刻がcutoffより古い場合はスキップ
      if (message.getDate() < cutoff) continue;

      // 本文を取得
      let plainBody = message.getPlainBody();
      let htmlBody = message.getBody(); // HTML本文
      // Logger.log(htmlBody);

      // 指定文言をすべて削除
      plainBody = removeTexts(plainBody, REMOVED_TEXTS);
      htmlBody = removeTexts(htmlBody, REMOVED_TEXTS);
      // Logger.log(htmlBody);

      // 空白行を整理
      plainBody = cleanBlankLinesPlain(plainBody);
      htmlBody = cleanBlankLinesHtml(htmlBody);
      // Logger.log(htmlBody);

      // 添付ファイルがあれば一緒に転送
      const attachments = message.getAttachments();

      GmailApp.sendEmail(
        FORWARD_TO,
        FOWARD_SUBJECT,
        plainBody,
        {
          htmlBody: htmlBody,
          attachments: attachments,
          name: FOWARD_NAME,
        }
      );
    }

    // このスレッドにラベルを付与
    thread.addLabel(label);
    thread.moveToTrash();
  }
}


// target配列で指定した文言をtextからすべて削除する
function removeTexts(text, targets) {
  let result = text;
  for (const target of targets) {
    result = result.split(target).join("");
  }
  return result;
}


// プレーンテキストの空白行を削除する
function cleanBlankLinesPlain(text) {
  // 行末・行頭の空白（スペース・タブ）を除去
  let result = text
    .split("\n")
    .map((line) => line.replace(/[ \t\u00A0]+$/g, "").replace(/^[ \t\u00A0]+$/g, ""))
    .join("\n");
 
  if (REMOVE_ALL_BLANK_LINES) {
    // 空行（何もない行）をすべて削除
    result = result
      .split("\n")
      .filter((line) => line.trim() !== "")
      .join("\n");
  } else {
    // 連続する空行を1行にまとめる
    result = result.replace(/\n{3,}/g, "\n\n");
  }
 
  return result.trim();
}


// HTML本文の空行・空要素を整理する
function cleanBlankLinesHtml(html) {
  let result = html;
 
  // 連続する <br> タグをまとめる／削除する
  if (REMOVE_ALL_BLANK_LINES) {
    result = result.replace(/(<br\s*\/?>\s*){2,}/gi, "<br>");
  } else {
    result = result.replace(/(<br\s*\/?>\s*){3,}/gi, "<br><br>");
  }
 
  // 中身が空（または&nbsp;だけ）の <p> や <div> を削除
  result = result.replace(
    /<(p|div)[^>]*>(\s|&nbsp;|<br\s*\/?>)*<\/\1>/gi,
    ""
  );
 
  // タグの間にできた余分な改行・空白行を削除
  result = result.replace(/>\s*\n\s*</g, ">\n<");
  result = result.replace(/\n{2,}/g, "\n");
 
  return result;
}












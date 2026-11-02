/**
 * こころノート → Googleスプレッドシート 自動保存
 *
 * このコードを、記録を残したいスプレッドシートの
 * 「拡張機能 → Apps Script」に貼り付けて、ウェブアプリとして公開します。
 * こころノート（GitHub版）から送られた記録を、
 * 「受けとめる」「伝える」の2つのシートに1行ずつ書き込みます。
 * 同じ記録（同じID）がもう一度送られたときは、その行を上書きします。
 */

const HEADERS = {
  seiri: ["ID", "日付", "状況", "気分", "認知", "身体", "行動", "別の認知", "気分の変化", "次の行動", "問いへの答え", "考え方のクセ", "更新日時"],
  assert: ["ID", "日付", "種類", "トレーニング対象場面", "相手の言動（事実）", "自分の感じた気持ち", "実際にとった対応", "タイプ",
    "結果（状況・関係性への影響）", "D（事実）", "E（気持ち）", "S（提案）", "C（結果）", "トレーニング後の振り返り・気づき", "今後に向けた改善目標", "更新日時"]
};
const SHEET_NAMES = { seiri: "受けとめる", assert: "伝える" };

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    (data.rows || []).forEach(function (r) {
      if (!HEADERS[r.kind] || !Array.isArray(r.values)) return;
      upsert_(r.kind, r.values.slice(0, HEADERS[r.kind].length).map(String));
    });
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

// ブラウザでURLを開いたときの確認用
function doGet() {
  return ContentService.createTextOutput("こころノートの保存先は動いています。");
}

function upsert_(kind, values) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES[kind]);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES[kind]);
    sh.appendRow(HEADERS[kind]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS[kind].length).setFontWeight("bold");
    sh.hideColumns(1); // IDの列は見えなくてよいので隠す
  }
  const id = values[0];
  const last = sh.getLastRow();
  if (last > 1) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
    const i = ids.indexOf(id);
    if (i >= 0) {
      sh.getRange(i + 2, 1, 1, values.length).setValues([values]);
      return;
    }
  }
  sh.appendRow(values);
}

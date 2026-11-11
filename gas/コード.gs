/**
 * スマホからアクセスされた際にHTML画面（入力フォーム）を返します。
 */
function doGet() {
  // index.htmlを読み込み、スマホ用レイアウト（viewport）を適用して表示
  return HtmlService.createHtmlOutputFromFile('index')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * フロント（HTML）のフォームからデータを受け取り、スプレッドシートに書き込みます。
 *
 * @param {string} mealType - 時間帯（朝、昼、夜、間食）
 * @param {string} menuName - メニュー名
 * @param {number} calories - カロリー
 * @param {number|string} steps - 歩数（未入力の場合は空欄）
 * @param {string} [memo] - メモ（任意。送られてこない場合は空文字）
 * @return {string} スマホ画面に表示する完了メッセージ
 */
function addLog(mealType, menuName, calories, steps, memo) {
  // アクティブなスプレッドシートの「食事ログ」シートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('食事ログ');

  // ヘッダー（1行目）のF列が空なら「メモ」を入れる
  var memoHeader = sheet.getRange(1, 6);
  if (memoHeader.getValue() === "") {
    memoHeader.setValue('メモ');
  }

  // タイムゾーンを日本時間に合わせて現在の日付（yyyy/MM/dd）を取得
  var today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

  // 歩数の入力チェック（数値があれば変換、空欄ならハイフンに）
  var stepsValue = steps ? Number(steps) : "-";

  // メモの入力チェック（未入力、または古い画面から送られてこない場合は空文字）
  var memoValue = memo ? String(memo) : "";

  // シートの最終行の下に［日付, 時間帯, メニュー名, カロリー, 歩数, メモ］の順で追加
  sheet.appendRow([
    today,
    mealType,
    menuName,
    Number(calories),
    stepsValue,
    memoValue
  ]);

  // スマホ側に返す完了通知メッセージ
  return menuName + "（" + calories + "kcal）/ 歩数: " + stepsValue + " を記録しました！";
}

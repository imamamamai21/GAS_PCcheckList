///////////////////////////////////////////////////////////////////////
/////////////////// PC台帳から特定のデータをチェックします //////////////////
///////////////////////////////////////////////////////////////////////

/**
 * 平日：毎朝10:00~11:00にチェックを走らせる
 * トリガー登録: 池田
 */
function morningCheck() {
  var day = new Date().getDay();
  if (day != 0 && day != 6) {
    updateStokeData();
    updateReservationData();
    updateWaitingDeliveryData();
    updateWaitingReturnData();
    updateCantUseData();
    updateDiscardData();
    updateStockListData();
    sendRemindMail_ReturnPc();
  }
  if (day === 5) botAtSummary(); // 毎週金曜日はまとめを投稿する
}

function botAtSummary() {
  var sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('はじめに');
  
  postBotForArms('# 台帳の状況のお知らせ\n現在の状況は以下の通りです。' + 
    '\nそれぞれ未対応のもの台数を表示いたします。' +
    '\n```\n' + sheet.getRange('D2').getValue() + 
    '\n' + sheet.getRange('D3').getValue() + 
    '\n' + sheet.getRange('D4').getValue() + 
    '\n' + sheet.getRange('D5').getValue() + 
    '\n' + sheet.getRange('D6').getValue() + 
    '\n' + sheet.getRange('D7').getValue() + '\n```' +
    '\n毎週金曜日にお知らせします。\n数を減らせるよう、ご対応をお願いいたします。' +
    '\n\n▶[PC台帳チェックリストへ](' + SHEET_URL_BASE + MY_SHEET_ID + '/)' +
    '\n#ARMS台帳状況');
}

function showTitleError(key) {
  Browser.msgBox('データが見つかりません', '表のタイトル名を変えていませんか？ : ' + key, Browser.Buttons.OK);
}

/**
 * 指定カラムから本日からの差分を返す式を返す
 * @param range {stirng} 指定カラム('A1'など)
 * @return {string} 条件式
 */
function getDataIf(range) {
  var ifText = function(row) { return 'DATEDIF(' + row + ',today(), "d")'; };
  return '=if(' + range + '<>"", ' + ifText(range) + ',' + ifText(range) + ')'; // 本日からの差分を取る
}
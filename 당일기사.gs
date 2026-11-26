/*****************************************************
 * 「세탁물 현황」 당일 배송기사 자동 기록 (매일 자정 직후 실행)
 *
 * Code.gs(현황지 출력 API)는 읽기 전용이고, 시트에 쓰는 것은 이 파일뿐이다.
 * 쓰는 칸은 정확히 세 곳이다.
 *   - I4:I(마지막 행)  : 그날 배송 나가는 업체 줄에 배송기사 이름(전화번호 제외), 나머지는 빈칸
 *   - L2(L2:M2 병합)   : 오늘 날짜 ("11월 26일")
 *   - N1(N1:O2 병합)   : 오늘 요일 ("목요일")
 * 예전에는 L2·N1이 =H1(주차 숫자 1)을 날짜로 보여 "12월 31일 / 일요일"로 표시됐다.
 *
 * 기사 이름 출처·당일 배송 판정은 현황지 출력 프로그램(Code.gs buildDayData_)과 똑같다.
 *   - 그 요일 T~Z 플래그가 1이고, A~F 그 요일 열에 코스코드가 있는 줄 = 당일 배송
 *   - 기사 = 「거래처」 시트 요일별 배송기사 열(AD~AI)에서 업체번호로 찾은 값
 * 월요일(휴무)은 I열을 전부 비운다.
 *
 * 처음 한 번: 편집기에서 setupTodayDriverTrigger 실행(권한 허용) → 매일 0시대 자동 실행.
 *****************************************************/

const TODAY_DRIVER_TZ = "Asia/Seoul";
const TODAY_DRIVER_COL = 9; // I열 (머리글 I3 "당일 기사")

function updateTodayDrivers() {
  var now = new Date();
  var iso = Utilities.formatDate(now, TODAY_DRIVER_TZ, "yyyy-MM-dd");
  var dayIdx = Number(Utilities.formatDate(now, TODAY_DRIVER_TZ, "u")); // 1=월 ... 7=일
  var weekday = ["월", "화", "수", "목", "금", "토", "일"][dayIdx - 1];
  var dateObj = parseIsoDate_(iso);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEB_SHEET_NAME);
  if (!sheet) throw new Error("'" + WEB_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  // 날짜·요일 표시칸 (병합 셀의 왼쪽 위 칸에만 쓴다)
  sheet.getRange("L2").setValue(dateObj).setNumberFormat('M"월" d"일"');
  sheet.getRange("N1").setValue(WEEKDAY_FULL_NAME[weekday]);
  sheet.getRange(3, TODAY_DRIVER_COL).setValue("당일 기사");

  var lastRow = sheet.getLastRow();
  var numRows = lastRow - DATA_START_ROW + 1;
  if (numRows <= 0) return { date: iso, weekday: weekday, written: 0 };

  var out = [];
  for (var r = 0; r < numRows; r++) out.push([""]);

  var codeCol = findDayCodeColumn_(sheet, WEEKDAY_FULL_NAME[weekday]);
  var written = 0;
  if (codeCol !== -1) { // 월요일 등 휴무일이면 전부 빈칸
    var flagCol = WEEKDAY_FLAG_COL[weekday];
    var tradeMap = buildTradeAssignmentMap_(weekday);
    var block = sheet.getRange(DATA_START_ROW, 1, numRows, READ_LAST_COL).getValues();
    for (var i = 0; i < numRows; i++) {
      var row = block[i];
      if (row[flagCol - 1] != 1) continue;
      var code = cellStr_(row[codeCol - 1]).replace(/^-+/, "");
      if (!/^[A-Za-z]+/.test(code)) continue;
      var trade = tradeMap[cellStr_(row[COL_COMPANY_NO - 1])];
      var name = driverNameOnly_(trade && trade.driver);
      if (name) { out[i][0] = name; written++; }
    }
  }
  sheet.getRange(DATA_START_ROW, TODAY_DRIVER_COL, numRows, 1).setValues(out);
  return { date: iso, weekday: weekday, written: written };
}

// "이병찬 010-1234-5678" → "이병찬". "---"처럼 글자가 없는 미배정 표시는 빈칸으로.
function driverNameOnly_(raw) {
  var name = String(raw || "").replace(/\s*01[016789][-\s.]?\d{3,4}[-\s.]?\d{4}.*$/, "").trim();
  return /[가-힣A-Za-z]/.test(name) ? name : "";
}

// 처음 한 번 실행: 기존 같은 트리거를 지우고 매일 0시대(한국 시간) 트리거를 새로 단 뒤 바로 한 번 기록한다.
function setupTodayDriverTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "updateTodayDrivers") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("updateTodayDrivers")
    .timeBased().everyDays(1).atHour(0).nearMinute(5).inTimezone(TODAY_DRIVER_TZ)
    .create();
  return updateTodayDrivers();
}

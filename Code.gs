/*****************************************************
 * 유성클리닝 현황지 출력 프로그램 — Apps Script 데이터 API
 *
 * 이 스크립트는 순수한 읽기(+코멘트 저장) API 역할만 한다.
 * 시트 복제, 정렬, 서식/색상/테두리 변경, 행높이·열너비 조정, 인쇄영역
 * 설정, PDF 생성은 전부 제거했다 — 표 렌더링과 인쇄는 브라우저(HTML/CSS/
 * window.print())가 담당한다. 원본 스프레드시트는 절대 쓰지 않는다
 * (코멘트 저장만 예외, 지정된 셀에만 값을 쓴다).
 *****************************************************/

/***** 지역 ↔ 색상 맵 (프런트엔드 색상 표시용) *****/
const REGION_COLOR_MAP = [
  ["강남구","#2D7FF9"],["서초구","#24427C"],["송파구","#10B981"],["광진구","#8B5CF6"],
  ["분당구","#6AA84F"],["기흥구","#F59E0B"],["수지구","#60A5FA"],["수정구","#7CB342"],
  ["의왕","#14B8A6"],["과천","#84CC16"],["군포","#F97316"],["광명시","#EAB308"],
  ["권선구","#B45309"],["단원구","#06B6D4"],["동안구","#64748B"],["만안구","#DC2626"]
];

/***** 기본 설정 *****/
const WEB_SHEET_NAME = "세탁물 현황";
const DATA_START_ROW = 4;

/***** 열 위치 (실제 시트 헤더 행을 읽어 확인한 값) *****
 * H=구역 I=코스 J=기사명 K=업체번호 L=약칭 M=업종 N=포대 배송 O=수건 배송
 * P=수거량(중요!) Q=사이즈 R=출고 S=포장 담당자 T~Z=요일별 플래그
 * AD=배송 주소 AE=수거방식 AF=배송 코멘트 AG=거래처 번호(전화번호)
 * (I·R은 예전 출력 스크립트가 코스코드/출고번호를 임시로 써넣던 칸이라
 * 원본 값을 신뢰하지 않고 이 스크립트가 직접 계산한다.)
 */
const COL_REGION = 8;      // H
const COL_DRIVER = 10;     // J
const COL_COMPANY_NO = 11; // K
const COL_NICKNAME = 12;   // L
const COL_BIZ_TYPE = 13;   // M
const COL_BAG = 14;        // N  포대 배송
const COL_TOWEL = 15;      // O  수건 배송
const COL_QTY = 16;        // P  수거량(중요!)
const COL_ADDRESS = 30;    // AD
const COL_METHOD = 31;     // AE 수거방식
const COL_COMMENT = 32;    // AF 배송 코멘트
const COL_PHONE = 33;      // AG 거래처 번호(전화번호)
const READ_LAST_COL = 33;  // A~AG 한 번에 읽음

const ORIGIN_ROW = 3;
const ORIGIN_NAME_COL = 29;    // AC — "유성(출발지)"
const ORIGIN_ADDRESS_COL = 30; // AD
const ORIGIN_COMMENT_COL = 32; // AF

/***** 요일 설정 *****
 * A~F열(헤더에 "OO요일" 문구가 있는 열)에 그 요일의 코스코드가 미리 입력되어 있고,
 * T~Z열은 그 행이 해당 요일에 나가는지를 나타내는 1/공백 플래그다.
 */
const WEEKDAY_LIST = ["월", "화", "수", "목", "금", "토", "일"];
const WEEKDAY_FULL_NAME = {
  "월": "월요일", "화": "화요일", "수": "수요일", "목": "목요일",
  "금": "금요일", "토": "토요일", "일": "일요일"
};
const WEEKDAY_FLAG_COL = { "월": 20, "화": 21, "수": 22, "목": 23, "금": 24, "토": 25, "일": 26 }; // T~Z

// =============================================
// 날짜/요일 헬퍼
// =============================================
function getDayName_(dateObj) {
  return ["일요일","월요일","화요일","수요일","목요일","금요일","토요일"][dateObj.getDay()];
}

function getIsoWeekNumber_(dateObj) {
  var date = new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
  var dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}

function getTodayWeekday_() {
  var map = ["일", "월", "화", "수", "목", "금", "토"];
  return map[new Date().getDay()];
}

// 이번 주(월~일) 안에서 해당 요일에 해당하는 날짜를 반환한다.
function getDateForWeekday_(weekday) {
  var idx = WEEKDAY_LIST.indexOf(weekday); // 월=0 ... 일=6
  var today = new Date();
  var todayIdx = (today.getDay() + 6) % 7; // 월=0 ... 일=6
  return new Date(today.getFullYear(), today.getMonth(), today.getDate() + (idx - todayIdx));
}

function toIsoDate_(dateObj) {
  var y = dateObj.getFullYear();
  var m = String(dateObj.getMonth() + 1).padStart(2, "0");
  var d = String(dateObj.getDate()).padStart(2, "0");
  return y + "-" + m + "-" + d;
}

function findDayCodeColumn_(sheet, fullDayName) {
  var headers = sheet.getRange("A1:F1").getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && String(headers[i]).indexOf(fullDayName) !== -1) {
      return i + 1;
    }
  }
  return -1;
}

// 셀 값을 항상 문자열로. "30_3" 같은 값이 날짜/숫자로 바뀌는 일이 없게 하고,
// 빈 셀은 빈 문자열로 통일한다.
function cellStr_(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v).trim();
}

// =============================================
// 요일 데이터 조회 (읽기 전용, 원본 시트는 절대 쓰지 않음)
// 한 번의 범위 조회(getValues)로 필요한 모든 열을 읽고, 이후는 전부
// 메모리(JS 배열)에서 가공한다 — 셀 단위 반복 호출 없음.
// =============================================
function buildDayData_(weekday) {
  var fullName = WEEKDAY_FULL_NAME[weekday];
  if (!fullName) throw new Error("요일 값이 올바르지 않습니다: " + weekday);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEB_SHEET_NAME);
  if (!sheet) throw new Error("'" + WEB_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  var codeCol = findDayCodeColumn_(sheet, fullName);
  if (codeCol === -1) throw new Error("헤더(A1:F1)에서 '" + fullName + "' 열을 찾을 수 없습니다.");

  var flagCol = WEEKDAY_FLAG_COL[weekday];
  var lastRow = sheet.getLastRow();
  var dataEndRow = lastRow - 2; // 하단 요약행 제외 (기존 로직과 동일)
  var numRows = dataEndRow - DATA_START_ROW + 1;

  var origin = { name: "", address: "", comment: "", phone: "" };
  if (lastRow >= ORIGIN_ROW) {
    var originRow = sheet.getRange(ORIGIN_ROW, 1, 1, READ_LAST_COL).getValues()[0];
    origin.name = cellStr_(originRow[ORIGIN_NAME_COL - 1]);
    origin.address = cellStr_(originRow[ORIGIN_ADDRESS_COL - 1]);
    origin.comment = cellStr_(originRow[ORIGIN_COMMENT_COL - 1]);
  }

  var groups = {};
  var order = [];

  if (numRows > 0) {
    var block = sheet.getRange(DATA_START_ROW, 1, numRows, READ_LAST_COL).getValues();

    for (var i = 0; i < numRows; i++) {
      var row = block[i];
      var flagVal = row[flagCol - 1];
      if (flagVal != 1) continue;

      var rawCode = cellStr_(row[codeCol - 1]);
      var cleaned = rawCode.replace(/^-+/, "");
      var match = cleaned.match(/^([A-Za-z]+)/);
      if (!match) continue;

      var alpha = match[1].toUpperCase();
      var suffixNum = parseInt(cleaned.replace(/^[A-Za-z]+-?/, ""), 10) || 0;

      if (!groups[alpha]) {
        groups[alpha] = [];
        order.push(alpha);
      }

      groups[alpha].push({
        sheetRow: DATA_START_ROW + i,
        sortKey: suffixNum,
        driver: cellStr_(row[COL_DRIVER - 1]),
        companyNumber: cellStr_(row[COL_COMPANY_NO - 1]),
        nickname: cellStr_(row[COL_NICKNAME - 1]),
        businessType: cellStr_(row[COL_BIZ_TYPE - 1]),
        bagDelivery: cellStr_(row[COL_BAG - 1]),
        towelDelivery: cellStr_(row[COL_TOWEL - 1]),
        collectionQuantity: cellStr_(row[COL_QTY - 1]),
        region: cellStr_(row[COL_REGION - 1]),
        address: cellStr_(row[COL_ADDRESS - 1]),
        collectionMethod: cellStr_(row[COL_METHOD - 1]),
        deliveryComment: cellStr_(row[COL_COMMENT - 1]),
        phone: cellStr_(row[COL_PHONE - 1])
      });
    }
  }

  order.sort();

  var drivers = {};
  var items = [];
  var shipSeq = 0;

  order.forEach(function(alpha) {
    groups[alpha].sort(function(a, b) { return a.sortKey - b.sortKey; });
    drivers[alpha] = groups[alpha].length ? groups[alpha][0].driver : "";

    groups[alpha].forEach(function(item, idx) {
      shipSeq++;
      items.push({
        course: alpha,
        deliveryOrder: alpha + "-" + String(idx + 1).padStart(2, "0"),
        shippingNumber: shipSeq,
        row: item.sheetRow,
        driver: item.driver,
        companyNumber: item.companyNumber,
        nickname: item.nickname,
        businessType: item.businessType,
        bagDelivery: item.bagDelivery,
        towelDelivery: item.towelDelivery,
        collectionQuantity: item.collectionQuantity,
        region: item.region,
        address: item.address,
        collectionMethod: item.collectionMethod,
        deliveryComment: item.deliveryComment,
        phone: item.phone
      });
    });
  });

  var dateObj = getDateForWeekday_(weekday);

  return {
    success: true,
    weekday: weekday,
    dayOfWeek: fullName,
    requestedDate: toIsoDate_(dateObj),
    weekNumber: getIsoWeekNumber_(dateObj) + "주차",
    origin: origin,
    courses: order,
    drivers: drivers,
    items: items,
    count: items.length,
    fetchedAt: new Date().toISOString()
  };
}

// =============================================
// 배송 코멘트 저장 (지정된 행의 AF열만 갱신 — 그 외에는 시트를 절대 건드리지 않음)
// updates: [{ row: 4, deliveryComment: "..." }, ...]
// =============================================
function saveComments_(updates) {
  if (!updates || !updates.length) return { success: true, updated: 0 };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEB_SHEET_NAME);
  if (!sheet) throw new Error("'" + WEB_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  var lastRow = sheet.getLastRow();
  var updated = 0;

  updates.forEach(function(u) {
    var row = Number(u.row);
    if (!row || row < DATA_START_ROW || row > lastRow) return;
    sheet.getRange(row, COL_COMMENT).setValue(String(u.deliveryComment || ""));
    updated++;
  });

  return { success: true, updated: updated };
}

// =============================================
// 웹 앱 진입점
// =============================================
function doPost(e) {
  var startedAt = Date.now();

  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === "getData") {
      var weekday = String(body.weekday || getTodayWeekday_()).trim();
      if (!WEEKDAY_FLAG_COL[weekday]) {
        return jsonOutput_({ success: false, message: "요일 값이 올바르지 않습니다." });
      }
      var data = buildDayData_(weekday);
      data.queryMs = Date.now() - startedAt;
      return jsonOutput_(data);
    }

    if (action === "saveComments") {
      var result = saveComments_(body.updates);
      result.queryMs = Date.now() - startedAt;
      return jsonOutput_(result);
    }

    return jsonOutput_({ success: false, message: "알 수 없는 요청입니다: " + action });

  } catch (err) {
    return jsonOutput_({ success: false, message: err.message });
  }
}

function doGet(e) {
  return jsonOutput_({
    success: true,
    message: "유성클리닝 현황지 출력 프로그램 데이터 API. POST로 { action: 'getData', weekday: '월'..'일' } 호출."
  });
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

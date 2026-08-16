/*****************************************************
 * 유성클리닝 현황지 출력 프로그램 — Apps Script 데이터 API
 *
 * 이 스크립트는 순수한 읽기 전용 API다. 원본 스프레드시트에는 어떤 경우에도
 * 절대 쓰지 않는다 (setValue/setValues/appendRow/insertRow/deleteRow/sort/
 * 시트 복제·서식 변경 등 모든 쓰기 계열 기능이 이 파일에는 존재하지 않는다).
 * 사용자가 화면에서 변경하는 모든 값(배송순서, 기사, 코스, 업체 추가/제외,
 * 배송 코멘트 등)은 브라우저 localStorage에 날짜별로만 저장되고, 원본 시트에는
 * 절대 반영되지 않는다.
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
const COL_BAGSIZE = 17;    // Q  사이즈(포대 규격)
const COL_ADDRESS = 30;    // AD
const COL_METHOD = 31;     // AE 수거방식
const COL_COMMENT = 32;    // AF 배송 코멘트
const COL_PHONE = 33;      // AG 거래처 번호(전화번호)
const READ_LAST_COL = 33;  // A~AG 한 번에 읽음

const ORIGIN_ROW = 3;
const ORIGIN_NAME_COL = 29;    // AC — "유성(출발지)"
const ORIGIN_ADDRESS_COL = 30; // AD
const ORIGIN_COMMENT_COL = 32; // AF

/***** "거래처" 시트 — 요일별 코스/기사 배정의 진짜 원본 *****
 * "세탁물 현황" 시트의 J열(기사명)은 업체마다 개별적으로 적혀 있어 기사가
 * 바뀌어도 모든 행이 함께 갱신되지 않고 오래된 이름이 섞여 남는 경우가 있다
 * (실사례: 8/13 목요일 E코스 — 세탁물 현황 J열은 이병찬/홍민선/공백이 섞여
 * 있었지만, 거래처 시트의 "요일별 배송기사" 열은 8개 업체 전부 고종원으로
 * 일관되게 적혀 있었다). 그래서 코스 대표 기사명은 이 "거래처" 시트의
 * 요일별 배송기사 열을 우선 신뢰하고, 세탁물 현황 J열은 그 값이 비어 있을
 * 때만 보조로 사용한다.
 */
const TRADE_SHEET_NAME = "거래처";
const TRADE_COL_COMPANY_CODE = 1; // A
const TRADE_COURSE_COL = { "화": 24, "수": 25, "목": 26, "금": 27, "토": 28, "일": 29 };  // X~AC 요일별 배송코스
const TRADE_DRIVER_COL = { "화": 30, "수": 31, "목": 32, "금": 33, "토": 34, "일": 35 };  // AD~AI 요일별 배송기사
const TRADE_READ_LAST_COL = 35;

/***** 요일 설정 *****
 * A~F열(헤더에 "OO요일" 문구가 있는 열)에 그 요일의 코스코드가 미리 입력되어 있고,
 * T~Z열은 그 행이 해당 요일에 나가는지를 나타내는 1/공백 플래그다.
 * (월요일은 휴무 — 세탁물 현황/거래처 시트 모두 월요일 열이 없다.)
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
// 프런트엔드가 명시적 date를 안 보냈을 때만 쓰는 예전 방식의 보조 수단이다 —
// 실제 요청 날짜는 buildDayData_에 넘어온 explicitDate를 우선한다.
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

// "YYYY-MM-DD" 문자열을 로컬 Date로 안전하게 파싱한다 (new Date("YYYY-MM-DD")는
// UTC 자정으로 해석되어 시간대에 따라 하루 밀릴 수 있어 직접 파싱한다).
function parseIsoDate_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || "").trim());
  if (!m) return null;
  var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (isNaN(d.getTime())) return null;
  return d;
}

function weekdayOfDate_(dateObj) {
  var map = ["일", "월", "화", "수", "목", "금", "토"];
  return map[dateObj.getDay()];
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
// "거래처" 시트에서 요일별 코스/기사 배정표를 한 번에 읽어
// { 업체코드: { course, driver } } 형태로 반환한다 (해당 요일 컬럼만).
// =============================================
function buildTradeAssignmentMap_(weekday) {
  var courseCol = TRADE_COURSE_COL[weekday];
  var driverCol = TRADE_DRIVER_COL[weekday];
  var map = {};
  if (!courseCol || !driverCol) return map; // 월요일 등 휴무 요일은 배정표가 없음

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRADE_SHEET_NAME);
  if (!sheet) return map;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  var block = sheet.getRange(1, 1, lastRow, TRADE_READ_LAST_COL).getValues();
  for (var i = 1; i < block.length; i++) { // 1행은 헤더이므로 건너뜀
    var row = block[i];
    var code = cellStr_(row[TRADE_COL_COMPANY_CODE - 1]);
    if (!code) continue;
    map[code] = {
      course: cellStr_(row[courseCol - 1]),
      driver: cellStr_(row[driverCol - 1])
    };
  }
  return map;
}

// =============================================
// 요일 데이터 조회 (읽기 전용, 원본 시트는 절대 쓰지 않음)
// 한 번의 범위 조회(getValues)로 필요한 모든 열을 읽고, 이후는 전부
// 메모리(JS 배열)에서 가공한다 — 셀 단위 반복 호출 없음.
//
// explicitDate: 프런트엔드가 넘긴 "YYYY-MM-DD" 날짜 문자열(선택). 넘어오면
// 이 날짜를 그대로 requestedDate/weekNumber 계산에 쓴다 — "이번 주 안에서
// 해당 요일을 역산"하는 예전 방식은 오늘이 속한 주 밖의 날짜(예: 다음 주,
// 다음 달의 같은 요일)를 선택했을 때 실제 선택 날짜와 다른 날짜를 돌려주는
// 문제가 있었다. 날짜와 요일이 서로 다른 요일을 가리키면 오류로 처리한다.
// =============================================
function buildDayData_(weekday, explicitDate) {
  var fullName = WEEKDAY_FULL_NAME[weekday];
  if (!fullName) throw new Error("요일 값이 올바르지 않습니다: " + weekday);

  var dateObj;
  if (explicitDate) {
    dateObj = parseIsoDate_(explicitDate);
    if (!dateObj) throw new Error("날짜 형식이 올바르지 않습니다: " + explicitDate);
    if (weekdayOfDate_(dateObj) !== weekday) {
      throw new Error("날짜(" + explicitDate + ")와 요일(" + weekday + ")이 일치하지 않습니다.");
    }
  } else {
    dateObj = getDateForWeekday_(weekday);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEB_SHEET_NAME);
  if (!sheet) throw new Error("'" + WEB_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  var codeCol = findDayCodeColumn_(sheet, fullName);
  if (codeCol === -1) {
    // 월요일처럼 헤더 자체가 없는 휴무 요일 — 빈 결과를 정상 반환한다.
    return {
      success: true,
      weekday: weekday,
      dayOfWeek: fullName,
      requestedDate: toIsoDate_(dateObj),
      weekNumber: getIsoWeekNumber_(dateObj) + "주차",
      origin: { name: "", address: "", comment: "" },
      courses: [],
      drivers: {},
      items: [],
      count: 0,
      closedDay: true,
      fetchedAt: new Date().toISOString()
    };
  }

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
        bagSize: cellStr_(row[COL_BAGSIZE - 1]),
        region: cellStr_(row[COL_REGION - 1]),
        address: cellStr_(row[COL_ADDRESS - 1]),
        collectionMethod: cellStr_(row[COL_METHOD - 1]),
        deliveryComment: cellStr_(row[COL_COMMENT - 1]),
        phone: cellStr_(row[COL_PHONE - 1])
      });
    }
  }

  order.sort();

  // "거래처" 시트의 요일별 배정표(업체코드 → {course, driver}) — 기사명의
  // 진짜 출처. 세탁물 현황 J열은 이 값이 없을 때만 보조로 쓴다.
  var tradeMap = buildTradeAssignmentMap_(weekday);

  var drivers = {};
  var items = [];
  var shipSeq = 0;

  order.forEach(function(alpha) {
    groups[alpha].sort(function(a, b) { return a.sortKey - b.sortKey; });

    // 각 업체의 "확정 기사명"을 거래처 시트 배정표로 먼저 정하고, 없으면
    // 세탁물 현황의 그 행 자체 값을 쓴다.
    groups[alpha].forEach(function(item) {
      var trade = tradeMap[item.companyNumber];
      item.resolvedDriver = (trade && trade.driver) ? trade.driver : item.driver;
    });

    // 코스 대표 기사명은 확정 기사명들의 최빈값(가장 많이 등장하는 이름)으로 정한다.
    var driverCounts = {};
    groups[alpha].forEach(function(it) {
      if (it.resolvedDriver) driverCounts[it.resolvedDriver] = (driverCounts[it.resolvedDriver] || 0) + 1;
    });
    var bestDriver = "", bestCount = 0;
    Object.keys(driverCounts).forEach(function(name) {
      if (driverCounts[name] > bestCount) { bestDriver = name; bestCount = driverCounts[name]; }
    });
    drivers[alpha] = bestDriver;

    groups[alpha].forEach(function(item, idx) {
      shipSeq++;
      items.push({
        course: alpha,
        deliveryOrder: alpha + "-" + String(idx + 1).padStart(2, "0"),
        orderSuffix: item.sortKey, // 원본 시트에 적힌 배송순서 번호(가공 전) — 누락/중복 검사용
        shippingNumber: shipSeq,
        row: item.sheetRow,
        driver: item.resolvedDriver,
        companyNumber: item.companyNumber,
        nickname: item.nickname,
        businessType: item.businessType,
        bagDelivery: item.bagDelivery,
        towelDelivery: item.towelDelivery,
        collectionQuantity: item.collectionQuantity,
        bagSize: item.bagSize,
        region: item.region,
        address: item.address,
        collectionMethod: item.collectionMethod,
        deliveryComment: item.deliveryComment,
        phone: item.phone
      });
    });
  });

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
    closedDay: false,
    fetchedAt: new Date().toISOString()
  };
}

// =============================================
// 웹 앱 진입점 — getData 한 가지 요청만 지원하는 순수 읽기 전용 API.
// 원본 시트에 값을 쓰는 기능(과거의 saveComments 등)은 전부 제거했다.
// =============================================
function doPost(e) {
  var startedAt = Date.now();

  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === "getData") {
      var weekday = String(body.weekday || getTodayWeekday_()).trim();
      if (!WEEKDAY_FULL_NAME[weekday]) {
        return jsonOutput_({ success: false, message: "요일 값이 올바르지 않습니다." });
      }
      var data = buildDayData_(weekday, body.date ? String(body.date).trim() : "");
      data.queryMs = Date.now() - startedAt;
      return jsonOutput_(data);
    }

    return jsonOutput_({ success: false, message: "알 수 없는 요청입니다: " + action });

  } catch (err) {
    return jsonOutput_({ success: false, message: err.message });
  }
}

function doGet(e) {
  return jsonOutput_({
    success: true,
    message: "유성클리닝 현황지 출력 프로그램 데이터 API (읽기 전용). POST로 { action: 'getData', date: 'YYYY-MM-DD', weekday: '월'..'일' } 호출."
  });
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

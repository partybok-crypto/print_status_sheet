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
 * H=구역 I=코스 J=합세탁(가능/체크 등, 세탁공정 문서용 — 기사명 아님) K=업체번호 L=약칭 M=업종 N=포대 배송 O=수건 배송
 * P=수거량(중요!) Q=사이즈 R=출고 S=포장 담당자 T~Z=요일별 플래그
 * AD=배송 주소 AE=수거방식 AF=배송 코멘트 AG=거래처 번호(전화번호)
 * (I·R은 예전 출력 스크립트가 코스코드/출고번호를 임시로 써넣던 칸이라
 * 원본 값을 신뢰하지 않고 이 스크립트가 직접 계산한다.)
 */
const COL_REGION = 8;      // H
const COL_DRIVER = 10;     // J (과거 기사명 열 — 이제 기사명 보조로 읽지 않는다)
const COL_COMBINED_WASH = 10; // J 합세탁 (A/B/C/X 직접 지정 / 가능 / 체크·신규·확인 / 공백)
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

/***** 세탁공정 작성자 — "세탁공장-일정표"의 주차별 쉬프트 탭에서 그 날짜 일정표의 첫 출근자 *****/
const SCHEDULE_SPREADSHEET_ID = "1kqYgHyvl-k1tfKk7qmVXBYfFwh3SVLNJNp41Otgl-u0";
const SCHEDULE_SHEET_NAME = "주차별 쉬프트";

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

// 합세탁 그룹: A 헬스장·필라테스 / B 미용실 / C 피부관리·마사지 / X 단독 / N 신규·미확인(단독+체크)
// J열 직접 입력(A/B/C/X)이 최우선, 그다음 '가능'+업종 자동, 체크·신규·확인은 N
function washGroupOf_(washRaw, bizRaw) {
  var w = String(washRaw === null || washRaw === undefined ? "" : washRaw).replace(/\s+/g, "");
  var biz = String(bizRaw === null || bizRaw === undefined ? "" : bizRaw).replace(/\s+/g, "");
  // 1단계: J열에 A/B/C(또는 X)를 직접 적으면 그게 우선 — 업종표에 없는 곳·예외 업체는 시트에서 직접 지정
  var direct = w.toUpperCase();
  if (/^[ABC]\d{0,2}$/.test(direct) || direct === "X") return direct; // B1·B2처럼 번호를 붙이면 같은 B색이지만 따로 묶음
  if (w === "체크" || w === "신규" || w === "확인") return "N";
  // 2단계: 가능이면 업종으로 자동
  if (w !== "가능") return "X";
  if (/헬스|필라테스/.test(biz)) return "A";
  if (/미용실/.test(biz)) return "B";
  if (/피부관리|마사지/.test(biz)) return "C";
  return "X";
}

// 일정표 시간 값(6, 9.5, "06:30")을 시(소수)로 바꾼다. 읽을 수 없으면 null.
function scheduleHour_(v) {
  var t = String(v === null || v === undefined ? "" : v).trim();
  if (!t) return null;
  var m = /^(\d{1,2})\s*:\s*(\d{1,2})$/.exec(t);
  if (m) return Number(m[1]) + Number(m[2]) / 60;
  var n = Number(t.replace(",", "."));
  return (isNaN(n) || n < 0 || n > 24) ? null : n;
}

// 세탁은 수거일 다음 영업일(월요일은 휴무라 건너뜀)에 하므로, 작성자는 그 세탁일 아침 근무자 기준이다.
function washDateOf_(dateObj) {
  var w = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate() + 1);
  if (w.getDay() === 1) w = new Date(w.getFullYear(), w.getMonth(), w.getDate() + 1);
  return w;
}

// 그 날짜 일정표 블록(날짜 칸 "10-31-토" + 바로 아래 머리글 출근/종료/포지션/이름/작업)에서
// 휴무가 아닌 사람 중 출근이 가장 이른 사람의 이름을 돌려준다(같으면 시트 순서). 못 찾거나 접근이 안 되면 "".
function getFirstWorkerName_(dateObj) {
  // 다른 스프레드시트를 열어 읽는 데 몇 초 걸리므로 날짜별로 30분간 기억한다(찾은 이름만 기억 - 못 찾으면 매번 다시 시도).
  var cache = null, ckey = "washWriter:" + toIsoDate_(dateObj);
  try { cache = CacheService.getScriptCache(); var hit = cache.get(ckey); if (hit) return hit; } catch (e0) {}
  var found = getFirstWorkerNameUncached_(dateObj);
  if (found) { try { cache.put(ckey, found, 1800); } catch (e1) {} }
  return found;
}

function getFirstWorkerNameUncached_(dateObj) {
  try {
    var sh = SpreadsheetApp.openById(SCHEDULE_SPREADSHEET_ID).getSheetByName(SCHEDULE_SHEET_NAME);
    if (!sh) return "";
    var v = sh.getDataRange().getDisplayValues();
    var wd = ["일", "월", "화", "수", "목", "금", "토"][dateObj.getDay()];
    var mm = dateObj.getMonth() + 1, dd = dateObj.getDate();
    for (var r = 0; r + 1 < v.length; r++) {
      for (var c = 0; c < v[r].length; c++) {
        var m = /^\s*(\d{1,2})\s*-\s*(\d{1,2})\s*-\s*([월화수목금토일])\s*$/.exec(v[r][c]);
        if (!m || Number(m[1]) !== mm || Number(m[2]) !== dd || m[3] !== wd) continue;
        var cols = {};
        for (var k = c; k < Math.min(v[r + 1].length, c + 8); k++) {
          var lab = String(v[r + 1][k]).trim();
          if ((lab === "출근" || lab === "이름" || lab === "작업" || lab === "포지션") && cols[lab] === undefined) cols[lab] = k;
        }
        if (cols["출근"] === undefined || cols["이름"] === undefined) continue; // 일정표 블록이 아님
        var best = "", bestHour = null;
        for (var i = r + 2; i < Math.min(v.length, r + 2 + 80); i++) {
          var row = v[i];
          var name = String(row[cols["이름"]] || "").trim();
          if (!name) continue;
          var task = cols["작업"] !== undefined ? String(row[cols["작업"]] || "") : "";
          var pos = cols["포지션"] !== undefined ? String(row[cols["포지션"]] || "") : "";
          if (task.indexOf("휴무") !== -1 || pos.indexOf("휴무") !== -1 || name.indexOf("휴무") !== -1) continue;
          var h = scheduleHour_(row[cols["출근"]]);
          if (h === null) continue;
          if (bestHour === null || h < bestHour) { bestHour = h; best = name; }
        }
        return best;
      }
    }
  } catch (err) {}
  return "";
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
  var buildStart = Date.now(), writerMs = 0;
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
        combinedWash: cellStr_(row[COL_COMBINED_WASH - 1]),
        washGroup: washGroupOf_(row[COL_COMBINED_WASH - 1], row[COL_BIZ_TYPE - 1]),
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
  var tTrade0 = Date.now();
  var tradeMap = buildTradeAssignmentMap_(weekday);
  var tradeMs = Date.now() - tTrade0;

  var drivers = {};
  var items = [];
  var shipSeq = 0;

  order.forEach(function(alpha) {
    groups[alpha].sort(function(a, b) { return a.sortKey - b.sortKey; });

    // 각 업체의 "확정 기사명"을 거래처 시트 배정표로 먼저 정하고, 없으면
    // 세탁물 현황의 그 행 자체 값을 쓴다.
    groups[alpha].forEach(function(item) {
      var trade = tradeMap[item.companyNumber];
      // 세탁물 현황 J열은 다른 용도(합세탁 표시)로 쓰므로 기사명 보조로 읽지 않는다.
      item.resolvedDriver = (trade && trade.driver) ? trade.driver : "";
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
        combinedWash: item.combinedWash,
        washGroup: item.washGroup,
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
    washWriter: (function () { var t0 = Date.now(); var w = getFirstWorkerName_(washDateOf_(dateObj)); writerMs = Date.now() - t0; return w; })(),
    drivers: drivers,
    items: items,
    count: items.length,
    closedDay: false,
    timing: { tradeMs: tradeMs, writerMs: writerMs, totalMs: Date.now() - buildStart },
    fetchedAt: new Date().toISOString()
  };
}

// =============================================
// 세탁공정 수량 이력 (읽기 전용)
// "업체별 수거량" 탭: D열=날짜("11-06-금"), 3행=업체번호, 5행=거래처명, E열 이후=날짜별 수거량.
// 거래처가 가로로 길게(수백 열) 늘어서 있으므로 머리글 행 전체를 읽어 업체번호 → 열 위치 표를 만든다.
// 같은 업체번호가 뒤쪽 정산 영역에 또 나오므로 가장 앞 열만 쓴다.
// 날짜 칸에는 연도가 없어서, 선택한 날짜("MM-DD-요일")가 있는 행을 찾고 거기서 7행씩 거슬러 올라가며 같은 요일을 읽는다.
// 값이 비어 있으면 null(기록 없음), 0이면 0(실제로 0포대) — 둘을 반드시 구분한다.
// =============================================
const COLLECT_SHEET_NAME = "업체별 수거량";
const COLLECT_FIRST_DATA_ROW = 6;
const COLLECT_LOOKBACK_WEEKS = 8;
const COLLECT_MAX_VALID_QTY = 30;

function collectQty_(v) {
  if (v === "" || v === null || v === undefined) return null;
  var n = (typeof v === "number") ? v : (/^\d+$/.test(String(v).trim()) ? Number(String(v).trim()) : NaN);
  if (isNaN(n) || n < 0 || n > COLLECT_MAX_VALID_QTY || Math.floor(n) !== n) return null;
  return n;
}

// 최신 수거량은 "세탁공장 현장입력" 시트에 먼저 들어온다(업체별 수거량 탭은 며칠 늦게 채워짐):
//   기사링크_수거량 = 기사 앱(수거 당일 저녁부터, 방문완료만) / 포대수량 = 배송입력 앱(다음날 아침 반영)
// 같은 날 같은 업체가 둘 다 있으면 기사링크_수거량 > 포대수량 순으로 쓴다.
const FIELD_SPREADSHEET_ID = "1efKO9y_-00VyBqdqOEJM_QeNgwQsrra6V5QrO5eM444";

function fieldRead_(sheetName, wanted) {
  var sh = SpreadsheetApp.openById(FIELD_SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sh) throw new Error("현장입력 시트에 '" + sheetName + "' 탭이 없습니다.");
  var v = sh.getDataRange().getDisplayValues();
  if (v.length < 2) return [];
  var idx = {};
  v[0].forEach(function (h, i) { if (idx[String(h).trim()] === undefined) idx[String(h).trim()] = i; });
  wanted.forEach(function (w) { if (idx[w] === undefined) throw new Error("'" + sheetName + "' 탭에 '" + w + "' 열이 없습니다."); });
  return { idx: idx, rows: v.slice(1) };
}

// { link: {날짜: {업체번호: 포대수}}, pq: {...} } — 같은 날·업체가 여러 줄이면 마지막 줄
function readFieldQty_() {
  var out = { link: {}, pq: {}, info: {} };   // info: {날짜: {업체번호: {name, course}}} (목록에 없는 거래처를 추가할 때 쓴다)
  var l = fieldRead_("기사링크_수거량", ["업체번호", "포대수", "운행일", "방문상태", "거래처명", "코스"]);
  l.rows.forEach(function (r) {
    if (String(r[l.idx["방문상태"]]).trim() !== "방문완료") return;
    var q = collectQty_(r[l.idx["포대수"]]); var d = String(r[l.idx["운행일"]]).trim(); var no = String(r[l.idx["업체번호"]]).trim();
    if (q === null || !d || !no) return;
    (out.link[d] = out.link[d] || {})[no] = q;
    (out.info[d] = out.info[d] || {})[no] = { name: String(r[l.idx["거래처명"]] || "").trim(), course: String(r[l.idx["코스"]] || "").trim() };
  });
  var p = fieldRead_("포대수량", ["날짜", "업체번호", "포대수", "거래처명"]);
  p.rows.forEach(function (r) {
    var q = collectQty_(r[p.idx["포대수"]]); var d = String(r[p.idx["날짜"]]).trim(); var no = String(r[p.idx["업체번호"]]).trim();
    if (q === null || !d || !no) return;
    (out.pq[d] = out.pq[d] || {})[no] = q;
    if (!(out.info[d] && out.info[d][no])) (out.info[d] = out.info[d] || {})[no] = { name: String(r[p.idx["거래처명"]] || "").trim(), course: "" };
  });
  return out;
}

function getWashHistory_(dateIso, weekday, companyNos) {
  var t0 = Date.now();
  var dateObj = parseIsoDate_(dateIso);
  if (!dateObj) throw new Error("날짜 형식이 올바르지 않습니다: " + dateIso);
  if (weekdayOfDate_(dateObj) !== weekday) throw new Error("날짜(" + dateIso + ")와 요일(" + weekday + ")이 일치하지 않습니다.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(COLLECT_SHEET_NAME);
  if (!sh) throw new Error("'" + COLLECT_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();

  // 머리글(3행 업체번호)로 업체번호 → 열 위치. 같은 번호가 뒤쪽 정산 영역에 또 있어 가장 앞 열만 쓴다.
  var head = sh.getRange(1, 1, 5, lastCol).getDisplayValues();
  var colOf = {};
  for (var c = 4; c < lastCol; c++) {
    var no = String(head[2][c]).trim();
    if (no && colOf[no] === undefined) colOf[no] = c;
  }

  // 업체별 수거량 탭에서 선택 날짜 행("MM-DD-요일")을 찾는다
  var mm = dateObj.getMonth() + 1, dd = dateObj.getDate();
  var dates = sh.getRange(COLLECT_FIRST_DATA_ROW, 4, Math.max(lastRow - COLLECT_FIRST_DATA_ROW + 1, 1), 1).getDisplayValues();
  var targetIdx = -1;
  for (var i = 0; i < dates.length; i++) {
    var m = /^\s*(\d{1,2})\s*-\s*(\d{1,2})\s*-\s*([월화수목금토일])/.exec(dates[i][0]);
    if (m && Number(m[1]) === mm && Number(m[2]) === dd && m[3] === weekday) { targetIdx = i; break; }
  }
  var block = null, firstIdx = 0;
  if (targetIdx >= 0) {
    firstIdx = Math.max(0, targetIdx - 7 * COLLECT_LOOKBACK_WEEKS);
    block = sh.getRange(COLLECT_FIRST_DATA_ROW + firstIdx, 1, targetIdx - firstIdx + 1, lastCol).getValues();
  }
  function sheetQty(k, col) {  // 업체별 수거량 탭의 k주 전 같은 요일 값(없으면 null)
    if (!block || col === undefined) return null;
    var idx = targetIdx - 7 * k;
    if (idx < firstIdx) return null;
    var lbl = String(dates[idx][0]).trim();
    var d2 = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate() - 7 * k);
    var mk = /^(\d{1,2})-(\d{1,2})-([월화수목금토일])/.exec(lbl);
    if (!mk || Number(mk[1]) !== d2.getMonth() + 1 || Number(mk[2]) !== d2.getDate() || mk[3] !== weekday) return null;
    return collectQty_(block[idx - firstIdx][col]);
  }

  // 현장입력 시트(기사 앱·배송입력 앱)
  var field = null, fieldError = "";
  try { field = readFieldQty_(); } catch (e) { fieldError = String(e && e.message || e); }
  function fieldQty(d, no) {
    if (!field) return null;
    if (field.link[d] && field.link[d][no] !== undefined) return { qty: field.link[d][no], from: "기사 앱(기사링크_수거량)" };
    if (field.pq[d] && field.pq[d][no] !== undefined) return { qty: field.pq[d][no], from: "배송입력 앱(포대수량)" };
    return null;
  }

  var isoDates = [];
  for (var k = 1; k <= COLLECT_LOOKBACK_WEEKS; k++) isoDates.push(toIsoDate_(new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate() - 7 * k)));

  var result = {
    success: true, date: dateIso, weekday: weekday, rowFound: targetIdx >= 0, fieldOk: !!field, fieldError: fieldError,
    fieldCount: field ? { link: Object.keys(field.link[dateIso] || {}).length, pq: Object.keys(field.pq[dateIso] || {}).length } : null,
    companies: {}, notFound: [], rows: isoDates
  };
  var wanted = companyNos && companyNos.length ? companyNos.slice() : Object.keys(colOf);

  // 그날 목록에 없는데 실제로 수거된 거래처(기사 앱·배송입력 앱 기록)는 목록에 덧붙인다 — 포대가 있는데 배정표에 빠지면 안 된다.
  var extras = [];
  if (field && companyNos && companyNos.length) {
    var listed = {}; companyNos.forEach(function (n) { listed[String(n).trim()] = true; });
    var cand = {};
    [field.link[dateIso] || {}, field.pq[dateIso] || {}].forEach(function (m) { Object.keys(m).forEach(function (n) { if (m[n] > 0 && !listed[n]) cand[n] = true; }); });
    var candNos = Object.keys(cand);
    if (candNos.length) {
      var stSheet = ss.getSheetByName(WEB_SHEET_NAME);
      var stRows = stSheet ? stSheet.getRange(DATA_START_ROW, 1, Math.max(stSheet.getLastRow() - 2 - DATA_START_ROW + 1, 1), READ_LAST_COL).getValues() : [];
      candNos.forEach(function (n) {
        var row = null;
        for (var q = 0; q < stRows.length; q++) { if (cellStr_(stRows[q][COL_COMPANY_NO - 1]) === n) { row = stRows[q]; break; } }
        var inf = (field.info[dateIso] && field.info[dateIso][n]) || { name: "", course: "" };
        extras.push({
          companyNumber: n,
          nickname: row ? cellStr_(row[COL_NICKNAME - 1]) : inf.name,
          businessType: row ? cellStr_(row[COL_BIZ_TYPE - 1]) : "",
          course: inf.course || "?",
          combinedWash: row ? cellStr_(row[COL_COMBINED_WASH - 1]) : "",
          washGroup: row ? washGroupOf_(row[COL_COMBINED_WASH - 1], row[COL_BIZ_TYPE - 1]) : "X"
        });
        wanted.push(n);
      });
    }
  }
  result.extras = extras;
  wanted.forEach(function (rawNo) {
    var no = String(rawNo).trim();
    var col = colOf[no];
    // 그날 실수량: 기사 앱 > 업체별 수거량 시트(사람이 고친 값) > 배송입력 앱(포대수량)
    var f0 = fieldQty(dateIso, no), sv0 = sheetQty(0, col);
    var actual = null, actualFrom = "";
    if (f0 && f0.from.indexOf("기사 앱") === 0) { actual = f0.qty; actualFrom = f0.from; }
    else if (sv0 !== null) { actual = sv0; actualFrom = "업체별 수거량 시트"; }
    else if (f0) { actual = f0.qty; actualFrom = f0.from; }
    var hist = isoDates.map(function (d, i) {
      var v = sheetQty(i + 1, col);
      if (v === null) { var f = fieldQty(d, no); if (f) v = f.qty; }
      return { date: d, qty: v };
    });
    if (col === undefined && actual === null && hist.every(function (h) { return h.qty === null; })) { result.notFound.push(no); return; }
    result.companies[no] = { actual: actual, actualFrom: actualFrom, history: hist };
  });
  result.timing = { totalMs: Date.now() - t0 };
  return result;
}

// =============================================
// 웹 앱 진입점 — 읽기 전용 API (getData, getWashHistory).
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

    if (action === "getWashHistory") {
      var wd = String(body.weekday || "").trim();
      if (!WEEKDAY_FULL_NAME[wd]) return jsonOutput_({ success: false, message: "요일 값이 올바르지 않습니다." });
      var hist = getWashHistory_(String(body.date || "").trim(), wd, body.companies || []);
      hist.queryMs = Date.now() - startedAt;
      return jsonOutput_(hist);
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

/***** 지역 ↔ 색상 맵 (외부 웹 대시보드용) *****/
const REGION_COLOR_MAP = [
  ["강남구","#2D7FF9"],["서초구","#24427C"],["송파구","#10B981"],["광진구","#8B5CF6"],
  ["분당구","#6AA84F"],["기흥구","#F59E0B"],["수지구","#60A5FA"],["수정구","#7CB342"],
  ["의왕","#14B8A6"],["과천","#84CC16"],["군포","#F97316"],["광명시","#EAB308"],
  ["권선구","#B45309"],["단원구","#06B6D4"],["동안구","#64748B"],["만안구","#DC2626"]
];

/***** 기본 설정 *****/
const OUTPUT_FOLDER_NAME = "현황지_출력";
const LOG_SHEET_NAME = "출력로그";
const WEB_SHEET_NAME = "세탁물 현황";

const DATA_START_ROW = 4;

const FETCH_BASE_DELAY_MS = 0;
const FETCH_RETRY_DELAY_MS = 3000;
const FETCH_MAX_RETRY = 3;
const FETCH_PARALLEL_BATCH_SIZE = 20; // UrlFetchApp.fetchAll 배치당 요청 수 (429 방지용 안전선)

const FORCE_FONT_ON_TEMP = true;
const PDF_FONT_FAMILY = "Noto Sans KR";

/***** 출력 범위 설정 *****/
// 전체 출력
const TOTAL_STATUS_START_COL = 8;    // H
const TOTAL_STATUS_END_COL   = 19;   // S
const TOTAL_ADDRESS_START_COL = 28;  // AB
const TOTAL_ADDRESS_END_COL   = 34;  // AH

// 알파벳별 출력: H/J/Q/S 삭제 후
// 삭제 대상: H 구역 / J 기사명 / Q 사이즈 / S 포장담당자
// 남는 현황지: I,K,L,M,N,O,P,R → 삭제 후 H:O
// 주소지: AB:AH → 삭제 후 X:AD
const ALPHA_STATUS_START_COL = 8;    // H
const ALPHA_STATUS_END_COL   = 15;   // O
const ALPHA_ADDRESS_START_COL = 24;  // X
const ALPHA_ADDRESS_END_COL   = 30;  // AD

/***** 요일 설정 *****
 * A~F열(헤더에 "OO요일" 문구가 있는 열)에 그 요일의 코스코드가 미리 입력되어 있고,
 * T~Z열은 그 행이 해당 요일에 나가는지를 나타내는 1/공백 플래그다.
 * 예전에는 H1에 날짜를 입력하면 이 요일을 계산해서 시트 자체를 정렬/숨김/재번호 처리했지만,
 * 그 과정이 느리고 여러 출력 단계에서 반복 재계산되면서 값이 흔들리는 원인이었다.
 * 이제는 날짜 대신 요일을 직접 선택하고, 원본 시트는 절대 건드리지 않고
 * 필요한 값만 한 번 읽어 메모리에서 계산한다.
 */
const WEEKDAY_LIST = ["월", "화", "수", "목", "금", "토", "일"];
const WEEKDAY_FULL_NAME = {
  "월": "월요일", "화": "화요일", "수": "수요일", "목": "목요일",
  "금": "금요일", "토": "토요일", "일": "일요일"
};
const WEEKDAY_FLAG_COL = { "월": 20, "화": 21, "수": 22, "목": 23, "금": 24, "토": 25, "일": 26 }; // T~Z

// =============================================
// 메뉴 생성
// =============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 현황지 출력 📋")
    .addItem("현장 출력센터 열기", "showOutputCenter")
    .addSeparator()
    .addItem("오늘 전체 출력 (마감자+코스별)", "exportTodayPackage")
    .addItem("코스별 출력 (오늘)", "exportCoursesToday")
    .addItem("마감자 출력 (오늘)", "exportCloserToday")
    .addSeparator()
    .addItem("수정모드 복구", "resetViewModeMenu_")
    .addToUi();
}

function resetViewModeMenu_() {
  resetViewMode();
}

// =============================================
// 날짜/요일 라벨 헬퍼
// =============================================
function getDayName_(dateObj) {
  return ["일요일","월요일","화요일","수요일","목요일","금요일","토요일"][dateObj.getDay()];
}

function getDateLabel_(dateObj) {
  return (dateObj.getMonth() + 1) + "월" + dateObj.getDate() + "일_" + getDayName_(dateObj);
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

// 원본 시트의 인쇄 헤더(H1 및 이를 참조하는 날짜/요일 표시 셀)를 채우기 위한 날짜 계산.
// 이번 주(월~일) 안에서 해당 요일에 해당하는 날짜를 반환한다.
function getDateForWeekday_(weekday) {
  var idx = WEEKDAY_LIST.indexOf(weekday); // 월=0 ... 일=6
  var today = new Date();
  var todayIdx = (today.getDay() + 6) % 7; // 월=0 ... 일=6
  return new Date(today.getFullYear(), today.getMonth(), today.getDate() + (idx - todayIdx));
}

// =============================================
// Google Drive 폴더 자동 생성
// 현황지_출력 / 2026년_27주차 / 7월4일_토요일
// =============================================
function getOrCreateSubFolder_(parent, name) {
  var folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function getOutputRootFolder() {
  var folders = DriveApp.getFoldersByName(OUTPUT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(OUTPUT_FOLDER_NAME);
}

function getOutputFolderForDate(dateObj) {
  var root = getOutputRootFolder();
  var weekNum = getIsoWeekNumber_(dateObj);
  var weekFolderName = dateObj.getFullYear() + "년_" + weekNum + "주차";
  var dateFolderName = getDateLabel_(dateObj);

  var weekFolder = getOrCreateSubFolder_(root, weekFolderName);
  return getOrCreateSubFolder_(weekFolder, dateFolderName);
}

function getOutputFolderForToday_() {
  return getOutputFolderForDate(new Date());
}

// =============================================
// 중복 파일 처리
// 같은 폴더에 같은 이름 있으면 기존 파일 휴지통 이동
// =============================================
function trashDuplicateFiles_(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

// 폴더 파일 목록을 한 번만 훑어서 이름→파일 맵을 만든다.
// (fetchPdfsInParallel_에서 파일마다 getFilesByName으로 검색하는 대신 이 맵을 재사용해
// Drive API 왕복 횟수를 줄인다.)
function buildExistingFileMap_(folder) {
  var map = {};
  var files = folder.getFiles();

  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName();
    if (!map[name]) map[name] = [];
    map[name].push(f);
  }

  return map;
}

// =============================================
// 출력 로그
// =============================================
function getOrCreateLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 7).setValues([[
      "출력시간", "메뉴", "요일", "코스", "파일구분", "파일명", "URL"
    ]]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function logOutputLinks_(links) {
  if (!links || links.length === 0) return;

  var logSheet = getOrCreateLogSheet_();
  var now = new Date();

  var rows = links.map(function(item) {
    return [
      now,
      item.menuName || "",
      item.weekday || "",
      item.course || "",
      item.fileType || "",
      item.fileName || "",
      item.url || ""
    ];
  });

  logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
}

// =============================================
// HTML 링크창
// =============================================
function escapeHtml_(s) {
  return String(s || "").replace(/[&<>"']/g, function(m) {
    return {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    }[m];
  });
}

function showCompletionDialog_(title, links) {
  var html = [];
  html.push('<div style="font-family:Arial, sans-serif; padding:16px;">');
  html.push('<h2 style="margin:0 0 12px;">✅ ' + escapeHtml_(title) + '</h2>');
  html.push('<p style="margin:0 0 14px; color:#555;">총 ' + links.length + '개 파일이 생성되었습니다.</p>');

  html.push('<div style="max-height:420px; overflow:auto; border:1px solid #ddd; padding:10px; border-radius:8px;">');

  links.forEach(function(item) {
    html.push(
      '<div style="margin:0 0 8px; padding:8px; border-bottom:1px solid #eee;">' +
      '<div style="font-weight:bold; margin-bottom:4px;">' + escapeHtml_(item.fileName) + '</div>' +
      '<a href="' + escapeHtml_(item.url) + '" target="_blank" style="color:#0b57d0;">파일 열기</a>' +
      '</div>'
    );
  });

  html.push('</div>');

  html.push('<div style="margin-top:16px;">');
  html.push('<button onclick="google.script.host.close()" style="width:100%; padding:10px 12px; background:#0b57d0; color:white; border:0; border-radius:6px; font-weight:700; cursor:pointer;">닫기</button>');
  html.push('</div>');
  html.push('</div>');

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html.join("")).setWidth(620).setHeight(600),
    title
  );
}

// =============================================
// 수정모드 복구 (모든 숨김 해제)
// 새 출력 로직은 원본 시트를 숨기거나 정렬하지 않으므로 평소에는 쓸 일이 없지만,
// 다른 이유로 행/열이 숨겨졌을 때를 대비해 남겨둠.
// =============================================
function resetViewMode(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.showRows(1, sheet.getMaxRows());
  sheet.showColumns(1, sheet.getMaxColumns());
  SpreadsheetApp.getActiveSpreadsheet().toast("모든 숨김이 해제되었습니다.", "📋 현황지 출력", 5);
}

// =============================================
// 요일별 코스 데이터 수집 (원본 시트는 읽기만 함, 절대 쓰지 않음)
// =============================================
function findDayCodeColumn_(sheet, fullDayName) {
  var headers = sheet.getRange("A1:F1").getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && String(headers[i]).indexOf(fullDayName) !== -1) {
      return i + 1;
    }
  }

  return -1;
}

function collectWeekdayCourseData_(sheet, weekday) {
  var fullName = WEEKDAY_FULL_NAME[weekday];
  if (!fullName) throw new Error("요일 값이 올바르지 않습니다: " + weekday);

  var codeCol = findDayCodeColumn_(sheet, fullName);
  if (codeCol === -1) throw new Error("헤더(A1:F1)에서 '" + fullName + "' 열을 찾을 수 없습니다.");

  var flagCol = WEEKDAY_FLAG_COL[weekday];

  var lastRow = sheet.getLastRow();
  var dataEndRow = lastRow - 2;
  var numRows = dataEndRow - DATA_START_ROW + 1;

  var groups = {};
  var order = [];

  if (numRows <= 0) return { groups: groups, order: order, dataEndRow: dataEndRow };

  var codeVals = sheet.getRange(DATA_START_ROW, codeCol, numRows, 1).getValues();
  var flagVals = sheet.getRange(DATA_START_ROW, flagCol, numRows, 1).getValues();
  var jVals = sheet.getRange(DATA_START_ROW, 10, numRows, 1).getValues(); // J 기사명
  var kVals = sheet.getRange(DATA_START_ROW, 11, numRows, 1).getValues(); // K 업체번호
  var lVals = sheet.getRange(DATA_START_ROW, 12, numRows, 1).getValues(); // L 약칭

  for (var i = 0; i < numRows; i++) {
    if (flagVals[i][0] != 1) continue;

    var rawCode = String(codeVals[i][0] || "").trim();
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
      row: DATA_START_ROW + i,
      rawCode: rawCode,
      sortKey: suffixNum,
      driver: String(jVals[i][0] || "").trim(),
      customerNo: String(kVals[i][0] || "").trim(),
      shortName: String(lVals[i][0] || "").trim()
    });
  }

  order.sort();

  var shipSeq = 0;

  order.forEach(function(alpha) {
    groups[alpha].sort(function(a, b) { return a.sortKey - b.sortKey; });

    groups[alpha].forEach(function(item, idx) {
      item.displayCode = alpha + "-" + String(idx + 1).padStart(2, "0");
      shipSeq++;
      item.shipNo = shipSeq;
    });
  });

  return { groups: groups, order: order, dataEndRow: dataEndRow };
}

// =============================================
// 출력 전 오류 검사
// =============================================
function validateCourseData_(groups, order) {
  var warnings = [];

  order.forEach(function(alpha) {
    groups[alpha].forEach(function(item) {
      if (!item.driver) warnings.push(item.row + "행: 기사명 없음");
      if (!item.customerNo) warnings.push(item.row + "행: 업체번호 없음");
      if (!item.shortName) warnings.push(item.row + "행: 약칭 없음");

      [item.rawCode, item.driver, item.customerNo, item.shortName].forEach(function(v) {
        if (String(v).indexOf("#REF!") !== -1) {
          warnings.push(item.row + "행: #REF! 포함");
        }
      });
    });
  });

  if (warnings.length > 40) {
    warnings = warnings.slice(0, 40).concat(["외 오류 다수..."]);
  }

  return warnings;
}

function confirmValidationWarnings_(warnings) {
  if (!warnings || warnings.length === 0) return true;

  var ui = SpreadsheetApp.getUi();
  var msg =
    "⚠️ 출력 전 확인할 항목이 있습니다.\n\n" +
    warnings.join("\n") +
    "\n\n그래도 계속 출력할까요?";

  var result = ui.alert("출력 전 오류 검사", msg, ui.ButtonSet.YES_NO);
  return result === ui.Button.YES;
}

// =============================================
// PDF fetch (순차 재시도 — 병렬 fetch 실패분 재처리용)
// =============================================
function fetchPdfWithRetry(url, token, fileName, folder) {
  var retryDelay = FETCH_RETRY_DELAY_MS;

  var opts = {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  };

  for (var attempt = 1; attempt <= FETCH_MAX_RETRY; attempt++) {
    if (FETCH_BASE_DELAY_MS > 0) {
      Utilities.sleep(FETCH_BASE_DELAY_MS);
    }

    var res = UrlFetchApp.fetch(url, opts);
    var code = res.getResponseCode();

    if (code === 200) {
      trashDuplicateFiles_(folder, fileName);
      var blob = res.getBlob().setName(fileName);
      return folder.createFile(blob).getUrl();
    }

    if (code === 429) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "⏳ 요청 과다 429 — " + retryDelay / 1000 + "초 대기 후 재시도 " +
        attempt + "/" + FETCH_MAX_RETRY,
        "📋 현황지 출력",
        10
      );

      Utilities.sleep(retryDelay);
      retryDelay += 2000;
      continue;
    }

    throw new Error("PDF 생성 실패 [HTTP " + code + "] " + fileName);
  }

  throw new Error("PDF 생성 실패 — 최대 재시도 초과: " + fileName);
}

// =============================================
// PDF export URL 생성 (fetch는 하지 않음)
// =============================================
function buildPdfExportUrl_(ss, sheet, rangeA1, portrait) {
  var range = sheet.getRange(rangeA1);

  var r1 = range.getRow() - 1;
  var c1 = range.getColumn() - 1;
  var r2 = range.getLastRow();
  var c2 = range.getLastColumn();

  return "https://docs.google.com/spreadsheets/d/" + ss.getId() +
    "/export?exportFormat=pdf&format=pdf" +
    "&size=A4" +
    "&portrait=" + (portrait ? "true" : "false") +
    "&fitw=true" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false" +
    "&top_margin=0.25" +
    "&bottom_margin=0.25" +
    "&left_margin=0.25" +
    "&right_margin=0.25" +
    "&gid=" + sheet.getSheetId() +
    "&range=" + encodeURIComponent(rangeA1) +
    "&r1=" + r1 +
    "&c1=" + c1 +
    "&r2=" + r2 +
    "&c2=" + c2;
}

// =============================================
// 여러 PDF를 UrlFetchApp.fetchAll로 병렬 요청
// jobs: [{ url, fileName, folder, course, fileType }]
// 반환: jobs와 같은 순서의 Drive 파일 URL 배열
// =============================================
function fetchPdfsInParallel_(jobs, existingFileMap) {
  var token = ScriptApp.getOAuthToken();
  var results = new Array(jobs.length);
  var retryIdx = [];

  for (var b = 0; b < jobs.length; b += FETCH_PARALLEL_BATCH_SIZE) {
    var chunk = jobs.slice(b, b + FETCH_PARALLEL_BATCH_SIZE);

    var requests = chunk.map(function(job) {
      return {
        url: job.url,
        headers: { Authorization: "Bearer " + token },
        muteHttpExceptions: true
      };
    });

    var responses = UrlFetchApp.fetchAll(requests);

    for (var k = 0; k < responses.length; k++) {
      var res = responses[k];
      var job = chunk[k];
      var idx = b + k;
      var code = res.getResponseCode();

      if (code === 200) {
        var existing = existingFileMap && existingFileMap[job.fileName];
        if (existing) {
          existing.forEach(function(f) { f.setTrashed(true); });
        }
        var blob = res.getBlob().setName(job.fileName);
        results[idx] = job.folder.createFile(blob).getUrl();
      } else {
        retryIdx.push(idx);
      }
    }
  }

  // 병렬 요청에서 실패한 것(주로 429)만 기존 순차 재시도 로직으로 처리
  retryIdx.forEach(function(idx) {
    var job = jobs[idx];
    results[idx] = fetchPdfWithRetry(job.url, token, job.fileName, job.folder);
  });

  return results;
}

// =============================================
// 임시시트 준비 (메모리에서 계산한 값만 반영, 원본은 읽기만 함)
// =============================================
function prepareTempSheetForRows_(ss, sourceSheet, dataEndRow, items, isAlpha, k1Name, headerDate) {
  var tempSheet = sourceSheet.copyTo(ss);
  tempSheet.setName("__TEMP__" + Utilities.getUuid().slice(0, 8));

  if (headerDate) {
    // H1은 날짜/요일 표시 셀(L2 등)이 "=H1"으로 참조하는 값이다.
    // 아래 PASTE_VALUES가 수식을 그 순간의 계산값으로 굳혀버리므로,
    // H1을 먼저 바꾸고 재계산까지 끝낸 뒤에 값으로 고정해야 날짜가 제대로 찍힌다.
    tempSheet.getRange(1, 8).setValue(headerDate); // H1
    SpreadsheetApp.flush();
  }

  var usedLastRow = tempSheet.getLastRow();
  var usedLastCol = tempSheet.getLastColumn();

  if (usedLastRow > 0 && usedLastCol > 0) {
    var usedRange = tempSheet.getRange(1, 1, usedLastRow, usedLastCol);

    usedRange.copyTo(
      usedRange,
      SpreadsheetApp.CopyPasteType.PASTE_VALUES,
      false
    );

    if (FORCE_FONT_ON_TEMP) {
      usedRange.setFontFamily(PDF_FONT_FAMILY);
      usedRange.setVerticalAlignment("middle");
    }
  }

  if (k1Name) {
    tempSheet.getRange(1, 11).setValue(k1Name); // K1
  }

  var keepMap = {};
  items.forEach(function(item) { keepMap[item.row] = item; });

  var dataRowCount = dataEndRow - DATA_START_ROW + 1;

  if (dataRowCount > 0) {
    var iVals = [], rVals = [], aaVals = [], abVals = [], acVals = [];
    var deleteArr = [];

    for (var r = DATA_START_ROW; r <= dataEndRow; r++) {
      var item = keepMap[r];

      if (item) {
        iVals.push([item.displayCode]);
        rVals.push([item.shipNo]);
        aaVals.push([item.displayCode]);
        abVals.push([item.customerNo]);
        acVals.push([item.shortName]);
      } else {
        iVals.push([""]);
        rVals.push([""]);
        aaVals.push([""]);
        abVals.push([""]);
        acVals.push([""]);
        deleteArr.push(r);
      }
    }

    tempSheet.getRange(DATA_START_ROW, 9, dataRowCount, 1).setValues(iVals);   // I 배송순서(코스코드)
    tempSheet.getRange(DATA_START_ROW, 18, dataRowCount, 1).setValues(rVals);  // R 출고번호
    tempSheet.getRange(DATA_START_ROW, 27, dataRowCount, 1).setValues(aaVals); // AA
    tempSheet.getRange(DATA_START_ROW, 28, dataRowCount, 1).setValues(abVals); // AB
    tempSheet.getRange(DATA_START_ROW, 29, dataRowCount, 1).setValues(acVals); // AC

    deleteArr.sort(function(a, b) { return b - a; });

    var di = 0;
    while (di < deleteArr.length) {
      var blockEnd = deleteArr[di];
      var dj = di + 1;

      while (dj < deleteArr.length && deleteArr[dj] === blockEnd - (dj - di)) {
        dj++;
      }

      tempSheet.deleteRows(deleteArr[dj - 1], dj - di);
      di = dj;
    }
  }

  // 남은 행을 I열(코스코드) 기준으로 정렬 → 코스별로 묶이고 코스 내 번호순 정렬됨
  var newLastRow = tempSheet.getLastRow();
  var newDataEndRow = newLastRow - 2;
  var remaining = newDataEndRow - DATA_START_ROW + 1;

  if (remaining > 1) {
    tempSheet.getRange(DATA_START_ROW, 1, remaining, tempSheet.getLastColumn())
      .sort([{ column: 9, ascending: true }]);
  }

  if (isAlpha) {
    // 삭제 대상: S 포장담당자 / Q 사이즈 / J 기사명 / H 구역
    tempSheet.deleteColumns(19, 1); // S 포장담당자
    tempSheet.deleteColumns(17, 1); // Q 사이즈
    tempSheet.deleteColumns(10, 1); // J 기사명
    tempSheet.deleteColumns(8, 1);  // H 구역
  }

  tempSheet.showColumns(1, tempSheet.getMaxColumns());

  return tempSheet;
}

// =============================================
// 코스별 요약표 추가 - 기존 하단행 덮어쓰기 버전
// =============================================
function appendCourseSummaryToTempSheet_(tempSheet) {
  var lastRow = tempSheet.getLastRow();

  var writeRow = Math.max(4, lastRow - 1);

  var dataStartRow = 4;
  var dataEndRow = writeRow - 1;
  var numRows = dataEndRow - dataStartRow + 1;

  if (numRows <= 0) return null;

  var vals = tempSheet.getRange(dataStartRow, 9, numRows, 2).getValues(); // I=코스, J=기사명

  var summary = {};
  var order = [];

  vals.forEach(function(row) {
    var courseRaw = String(row[0] || "").trim().replace(/^-+/, "");
    var match = courseRaw.match(/^([A-Za-z]+)/);
    if (!match) return;

    var alpha = match[1].toUpperCase();
    var driver = String(row[1] || "").trim();

    if (!summary[alpha]) {
      summary[alpha] = { driver: driver, count: 0 };
      order.push(alpha);
    }

    if (!summary[alpha].driver && driver) {
      summary[alpha].driver = driver;
    }

    summary[alpha].count++;
  });

  if (order.length === 0) return null;

  var total = 0;
  var pieces = [];

  order.forEach(function(alpha) {
    var driver = summary[alpha].driver;
    var count = summary[alpha].count;
    total += count;

    if (driver) {
      pieces.push(alpha + " " + driver + " " + count + "곳");
    } else {
      pieces.push(alpha + " " + count + "곳");
    }
  });

  pieces.push("총 " + total + "곳");

  var summaryText = "코스별 요약  |  " + pieces.join("  |  ");

  var startCol = 8;   // H
  var colCount = 13;  // H:T

  var targetRange = tempSheet.getRange(writeRow, startCol, 1, colCount);
  targetRange.clearContent();

  targetRange.merge()
    .setValue(summaryText)
    .setFontFamily(PDF_FONT_FAMILY)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setBackground("#E8F0FE")
    .setBorder(true, true, true, true, true, true)
    .setWrap(false);

  tempSheet.setRowHeight(writeRow, 30);

  return null;
}

// =============================================
// 출력 링크 객체 생성 / PDF 요청(job) 생성
// =============================================
function makeOutputItem_(menuName, weekday, course, fileType, fileName, url) {
  return {
    menuName: menuName,
    weekday: weekday,
    course: course,
    fileType: fileType,
    fileName: fileName,
    url: url
  };
}

function addPairJobs_(ss, tempSheet, prefix, folder, isAlpha, courseLabel, jobs) {
  var lastRow = tempSheet.getLastRow();

  var statusStartCol = isAlpha ? ALPHA_STATUS_START_COL : TOTAL_STATUS_START_COL;
  var statusEndCol   = isAlpha ? ALPHA_STATUS_END_COL   : TOTAL_STATUS_END_COL;

  var addressStartCol = isAlpha ? ALPHA_ADDRESS_START_COL : TOTAL_ADDRESS_START_COL;
  var addressEndCol   = isAlpha ? ALPHA_ADDRESS_END_COL   : TOTAL_ADDRESS_END_COL;

  var statusRangeA1 = tempSheet
    .getRange(1, statusStartCol, lastRow, statusEndCol - statusStartCol + 1)
    .getA1Notation();

  var addressRangeA1 = tempSheet
    .getRange(1, addressStartCol, lastRow, addressEndCol - addressStartCol + 1)
    .getA1Notation();

  var statusFileName = prefix + "_현황지.pdf";
  var addressFileName = prefix + "_주소지_가로.pdf";

  jobs.push({
    url: buildPdfExportUrl_(ss, tempSheet, statusRangeA1, true),
    fileName: statusFileName,
    folder: folder,
    course: courseLabel,
    fileType: "현황지"
  });

  jobs.push({
    url: buildPdfExportUrl_(ss, tempSheet, addressRangeA1, false),
    fileName: addressFileName,
    folder: folder,
    course: courseLabel,
    fileType: "주소지"
  });
}

// =============================================
// 출력 실행 (핵심)
// 1) 요일 데이터를 한 번만 읽어서 메모리에서 계산
// 2) 필요한 임시시트를 모두 준비
// 3) fetchAll로 PDF를 병렬 생성
// 4) 임시시트 정리
// =============================================
function buildExportJobs_(ss, sheet, weekday, options) {
  options = options || {};

  var includeTotal = options.includeTotal !== false;
  var includeCourses = options.includeCourses !== false;
  var skipValidation = options.skipValidation === true;
  var label = options.label || "출력";

  var data = collectWeekdayCourseData_(sheet, weekday);
  var order = data.order;
  var groups = data.groups;
  var dataEndRow = data.dataEndRow;

  if (order.length === 0) {
    throw new Error(weekday + "요일에 해당하는 코스 데이터가 없습니다.");
  }

  if (!skipValidation) {
    var warnings = validateCourseData_(groups, order);
    if (!confirmValidationWarnings_(warnings)) return null;
  }

  var folder = getOutputFolderForToday_();
  var prefix = weekday + "요일";
  var headerDate = getDateForWeekday_(weekday);

  var totalUnits = (includeTotal ? 1 : 0) + (includeCourses ? order.length : 0);
  var doneUnits = 0;
  progressStart_(label, Math.max(totalUnits, 1));

  var tempSheets = [];
  var jobs = [];

  try {
    if (includeTotal) {
      progressStep_(doneUnits, totalUnits, label + " - 전체(마감자) 준비 중");

      var allItems = [];
      order.forEach(function(a) { allItems = allItems.concat(groups[a]); });

      var totalTemp = prepareTempSheetForRows_(ss, sheet, dataEndRow, allItems, false, "마감자", headerDate);
      appendCourseSummaryToTempSheet_(totalTemp);
      tempSheets.push(totalTemp);
      addPairJobs_(ss, totalTemp, prefix + "_전체", folder, false, "전체", jobs);

      doneUnits++;
      progressStep_(doneUnits, totalUnits, label + " - 전체(마감자) 준비 완료");
    }

    if (includeCourses) {
      order.forEach(function(alpha) {
        progressStep_(doneUnits, totalUnits, label + " - " + alpha + " 코스 준비 중");

        var items = groups[alpha];
        var driverName = items.length ? items[0].driver : "";

        var courseTemp = prepareTempSheetForRows_(ss, sheet, dataEndRow, items, true, driverName, headerDate);
        tempSheets.push(courseTemp);
        addPairJobs_(ss, courseTemp, prefix + "_" + alpha, folder, true, alpha, jobs);

        doneUnits++;
        progressStep_(doneUnits, totalUnits, label + " - " + alpha + " 코스 준비 완료");
      });
    }

    SpreadsheetApp.flush();

    progressStep_(doneUnits, totalUnits, label + " - PDF 생성(병렬) 중...");
    var existingFileMap = buildExistingFileMap_(folder);
    var urls = fetchPdfsInParallel_(jobs, existingFileMap);

    var links = [];
    for (var i = 0; i < jobs.length; i++) {
      links.push(makeOutputItem_(label, weekday, jobs[i].course, jobs[i].fileType, jobs[i].fileName, urls[i]));
    }

    progressFinish_(label + " 완료 (" + jobs.length + "개 파일)");

    return links;

  } catch (e) {
    progressFail_(e.message);
    throw e;
  } finally {
    tempSheets.forEach(function(ts) {
      try { ss.deleteSheet(ts); } catch (e2) {}
    });
  }
}

function runWeekdayExport_(weekday, options, label) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  SpreadsheetApp.getActiveSpreadsheet().toast(weekday + "요일 데이터 확인 중...", "📋 현황지 출력", 30);

  var links = buildExportJobs_(ss, sheet, weekday, {
    includeTotal: options.includeTotal,
    includeCourses: options.includeCourses,
    skipValidation: options.skipValidation,
    label: label
  });

  if (!links) return null;

  logOutputLinks_(links);
  showCompletionDialog_(label + " 완료 (" + weekday + "요일)", links);

  return links;
}

// =============================================
// 메뉴 액션 (항상 "오늘" 요일 기준)
// =============================================
function exportCoursesToday() {
  runWeekdayExport_(getTodayWeekday_(), { includeTotal: false, includeCourses: true }, "코스별 출력");
}

function exportCloserToday() {
  runWeekdayExport_(getTodayWeekday_(), { includeTotal: true, includeCourses: false }, "마감자 출력");
}

function exportTodayPackage() {
  runWeekdayExport_(getTodayWeekday_(), { includeTotal: true, includeCourses: true }, "오늘 전체 출력");
}

// =============================================
// 현장 출력센터 GUI 사이드바
// =============================================
function showOutputCenter() {
  var html = HtmlService
    .createHtmlOutput(buildOutputCenterHtml_())
    .setTitle("현장 출력센터");

  SpreadsheetApp.getUi().showSidebar(html);
}

function buildOutputCenterHtml_() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: Arial, 'Noto Sans KR', sans-serif;
      margin: 0;
      padding: 16px;
      background: #f8fafd;
      color: #202124;
    }

    .title {
      font-size: 22px;
      font-weight: 800;
      margin-bottom: 4px;
    }

    .subtitle {
      font-size: 13px;
      color: #5f6368;
      margin-bottom: 16px;
    }

    .card {
      background: #ffffff;
      border: 1px solid #e0e0e0;
      border-radius: 12px;
      padding: 14px;
      margin-bottom: 12px;
      box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }

    label {
      display: block;
      font-size: 13px;
      font-weight: 700;
      margin-bottom: 6px;
      color: #3c4043;
    }

    .day-row {
      display: flex;
      gap: 6px;
      margin-bottom: 8px;
    }

    .day-btn {
      flex: 1;
      padding: 10px 0;
      font-size: 14px;
      font-weight: 700;
      border: 1px solid #dadce0;
      border-radius: 8px;
      background: #fff;
      color: #202124;
      cursor: pointer;
    }

    .day-btn.selected {
      background: #0b57d0;
      color: #fff;
      border-color: #0b57d0;
    }

    .btn {
      width: 100%;
      border: 0;
      border-radius: 10px;
      padding: 13px 10px;
      font-size: 15px;
      font-weight: 800;
      cursor: pointer;
      margin-bottom: 9px;
    }

    .blue { background:#0b57d0; color:white; }
    .green { background:#188038; color:white; }
    .gray { background:#5f6368; color:white; }
    .red { background:#d93025; color:white; }

    .btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }

    .status {
      background:#eef3fe;
      color:#174ea6;
      border-left:4px solid #1a73e8;
      padding:10px;
      font-size:13px;
      line-height:1.5;
      border-radius:8px;
      margin-top:10px;
      white-space:pre-line;
    }

    .small {
      font-size:12px;
      color:#5f6368;
      line-height:1.5;
    }

    .divider {
      height:1px;
      background:#e0e0e0;
      margin:12px 0;
    }

    .folder-link {
      display:block;
      text-align:center;
      background:#e6f4ea;
      color:#137333;
      text-decoration:none;
      padding:11px;
      border-radius:10px;
      font-weight:800;
      margin-top:8px;
    }
  </style>
</head>

<body>
  <div class="title">📋 현장 출력센터</div>
  <div class="subtitle">요일을 선택하고 버튼을 누르세요.</div>

  <div class="card">
    <label>출력 요일</label>
    <div class="day-row" id="daySelector"></div>
    <div class="small">오늘 요일이 자동으로 선택되어 있습니다. 미리 준비하려면 다른 요일을 눌러 바꾸세요.</div>
  </div>

  <div class="card">
    <button class="btn blue" onclick="runAction('package')">오늘 전체 출력 (마감자+코스별)</button>
    <button class="btn green" onclick="runAction('courses')">코스별 출력</button>
    <button class="btn gray" onclick="runAction('closer')">마감자 출력</button>

    <div class="divider"></div>

    <button class="btn red" onclick="runAction('reset')">수정모드 복구</button>
  </div>

  <div class="card">
    <label>진행상황</label>

    <div style="background:#e0e0e0; border-radius:10px; height:18px; overflow:hidden;">
      <div id="progressBar" style="height:18px; width:0%; background:#0b57d0; transition:width 0.4s;"></div>
    </div>

    <div id="progressText" style="font-size:13px; margin-top:8px; line-height:1.5; white-space:pre-line;">
      대기 중
    </div>
  </div>

  <div class="card">
    <a id="folderLink" class="folder-link" href="#" target="_blank" style="display:none;">📁 오늘 출력 폴더 열기</a>
  </div>

  <div id="status" class="status">준비되었습니다.</div>

<script>
  var WEEKDAYS = ['월', '화', '수', '목', '금', '토', '일'];
  var selectedDay = '';
  var progressTimer = null;

  function setStatus(msg) {
    document.getElementById('status').innerText = msg;
  }

  function formatSec(sec) {
    if (sec === null || sec === undefined) return '계산 중';

    sec = Number(sec);
    if (sec < 60) return sec + '초';

    var min = Math.floor(sec / 60);
    var remain = sec % 60;

    return min + '분 ' + remain + '초';
  }

  function startProgressPolling() {
    stopProgressPolling();

    progressTimer = setInterval(function() {
      google.script.run
        .withSuccessHandler(renderProgress)
        .withFailureHandler(function(err) { console.log(err); })
        .getOutputProgress();
    }, 2000);
  }

  function stopProgressPolling() {
    if (progressTimer) {
      clearInterval(progressTimer);
      progressTimer = null;
    }
  }

  function renderProgress(p) {
    if (!p) return;

    var percent = p.percent || 0;

    var bar = document.getElementById('progressBar');
    var textBox = document.getElementById('progressText');

    if (!bar || !textBox) return;

    bar.style.width = percent + '%';

    var lines = [];
    lines.push('현재 작업: ' + (p.label || '-'));
    lines.push('진행률: ' + (p.current || 0) + ' / ' + (p.total || 0) + ' (' + percent + '%)');
    lines.push('경과 시간: ' + formatSec(p.elapsedSec || 0));
    lines.push('예상 남은 시간: ' + formatSec(p.remainSec));

    if (p.message) lines.push(p.message);

    textBox.innerText = lines.join(String.fromCharCode(10));

    if (p.done || percent >= 100) {
      stopProgressPolling();
    }
  }

  function setDisabled(disabled) {
    document.querySelectorAll('button').forEach(function(btn) {
      btn.disabled = disabled;
    });
  }

  function buildDaySelector(today) {
    var box = document.getElementById('daySelector');
    box.innerHTML = '';

    WEEKDAYS.forEach(function(d) {
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'day-btn';
      b.id = 'day-' + d;
      b.textContent = d;
      b.onclick = function() { selectDay(d); };
      box.appendChild(b);
    });

    selectDay(today);
  }

  function selectDay(d) {
    selectedDay = d;

    WEEKDAYS.forEach(function(w) {
      var btn = document.getElementById('day-' + w);
      if (!btn) return;
      btn.classList.toggle('selected', w === d);
    });
  }

  function init() {
    setStatus('오늘 요일을 확인하는 중입니다.');

    google.script.run
      .withSuccessHandler(function(state) {
        buildDaySelector(state.today);

        if (state.folderUrl) {
          var link = document.getElementById('folderLink');
          link.href = state.folderUrl;
          link.style.display = 'block';
        }

        setStatus('준비되었습니다. 요일을 확인하고 버튼을 누르세요.');
      })
      .withFailureHandler(function(err) {
        setStatus('오류: ' + err.message);
      })
      .getOutputCenterState();
  }

  function runAction(type) {
    if (type !== 'reset' && !selectedDay) {
      setStatus('요일을 먼저 선택해주세요.');
      return;
    }

    setDisabled(true);
    startProgressPolling();

    var labelMap = {
      package: '오늘 전체 출력 중입니다. PDF 생성 시간이 걸릴 수 있습니다.',
      courses: '코스별 출력 중입니다.',
      closer: '마감자 출력 중입니다.',
      reset: '수정모드 복구 중입니다.'
    };

    setStatus(labelMap[type] || '실행 중입니다.');

    google.script.run
      .withSuccessHandler(function(state) {
        stopProgressPolling();

        if (state && state.folderUrl) {
          var link = document.getElementById('folderLink');
          link.href = state.folderUrl;
          link.style.display = 'block';
        }

        setStatus((state && state.message) ? state.message : '완료되었습니다.');
        setDisabled(false);
      })
      .withFailureHandler(function(err) {
        stopProgressPolling();
        setStatus('오류: ' + err.message);
        setDisabled(false);
      })
      .sidebarRunAction(type, selectedDay);
  }

  init();
</script>
</body>
</html>
`;
}

// =============================================
// 출력센터 서버 함수
// =============================================
function getOutputCenterState() {
  var folderUrl = "";
  try { folderUrl = getOutputFolderForToday_().getUrl(); } catch (e) {}

  return {
    today: getTodayWeekday_(),
    folderUrl: folderUrl
  };
}

function sidebarRunAction(type, weekday) {
  if (type === "reset") {
    resetViewMode();
    return { message: "수정모드 복구 완료" };
  }

  if (!WEEKDAY_FLAG_COL[weekday]) {
    throw new Error("요일을 먼저 선택해주세요.");
  }

  var confMap = {
    package: { includeTotal: true, includeCourses: true, label: "오늘 전체 출력" },
    closer:  { includeTotal: true, includeCourses: false, label: "마감자 출력" },
    courses: { includeTotal: false, includeCourses: true, label: "코스별 출력" }
  };

  var conf = confMap[type];
  if (!conf) throw new Error("알 수 없는 요청입니다: " + type);

  var links = runWeekdayExport_(weekday, {
    includeTotal: conf.includeTotal,
    includeCourses: conf.includeCourses
  }, conf.label);

  var folderUrl = "";
  try { folderUrl = getOutputFolderForToday_().getUrl(); } catch (e) {}

  return {
    message: links
      ? conf.label + "이(가) 완료되었습니다. (" + links.length + "개 파일)"
      : "검증에서 취소되었습니다.",
    folderUrl: folderUrl
  };
}

// =============================================
// 출력 진행상황 저장 / 조회
// =============================================
const OUTPUT_PROGRESS_KEY = "OUTPUT_CENTER_PROGRESS";

function progressStart_(title, total) {
  var now = new Date().getTime();

  var data = {
    running: true,
    done: false,
    title: title || "출력 작업",
    current: 0,
    total: total || 1,
    label: "작업 준비 중...",
    startedAt: now,
    updatedAt: now,
    elapsedSec: 0,
    remainSec: null,
    percent: 0,
    message: ""
  };

  PropertiesService.getDocumentProperties()
    .setProperty(OUTPUT_PROGRESS_KEY, JSON.stringify(data));
}

function progressStep_(current, total, label) {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty(OUTPUT_PROGRESS_KEY);
  var data = raw ? JSON.parse(raw) : {};

  var now = new Date().getTime();
  var startedAt = data.startedAt || now;

  current = current || 0;
  total = total || data.total || 1;

  var elapsedSec = Math.max(1, Math.floor((now - startedAt) / 1000));
  var percent = Math.min(100, Math.round((current / total) * 100));

  var remainSec = null;

  if (current > 0 && current < total) {
    var avgSec = elapsedSec / current;
    remainSec = Math.round(avgSec * (total - current));
  }

  data.running = true;
  data.done = false;
  data.current = current;
  data.total = total;
  data.label = label || "";
  data.updatedAt = now;
  data.elapsedSec = elapsedSec;
  data.remainSec = remainSec;
  data.percent = percent;

  props.setProperty(OUTPUT_PROGRESS_KEY, JSON.stringify(data));
}

function progressFinish_(message) {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty(OUTPUT_PROGRESS_KEY);
  var data = raw ? JSON.parse(raw) : {};

  var now = new Date().getTime();
  var startedAt = data.startedAt || now;

  data.running = false;
  data.done = true;
  data.current = data.total || data.current || 1;
  data.percent = 100;
  data.label = "완료";
  data.message = message || "출력이 완료되었습니다.";
  data.elapsedSec = Math.floor((now - startedAt) / 1000);
  data.remainSec = 0;
  data.updatedAt = now;

  props.setProperty(OUTPUT_PROGRESS_KEY, JSON.stringify(data));
}

function progressFail_(message) {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty(OUTPUT_PROGRESS_KEY);
  var data = raw ? JSON.parse(raw) : {};

  data.running = false;
  data.done = true;
  data.percent = data.percent || 0;
  data.label = "오류 발생";
  data.message = message || "출력 중 오류가 발생했습니다.";
  data.remainSec = null;
  data.updatedAt = new Date().getTime();

  props.setProperty(OUTPUT_PROGRESS_KEY, JSON.stringify(data));
}

function getOutputProgress() {
  var raw = PropertiesService.getDocumentProperties()
    .getProperty(OUTPUT_PROGRESS_KEY);

  if (!raw) {
    return {
      running: false,
      done: false,
      current: 0,
      total: 0,
      percent: 0,
      label: "대기 중",
      elapsedSec: 0,
      remainSec: null,
      message: ""
    };
  }

  return JSON.parse(raw);
}

/*****************************************************
 * 웹 앱 연동 (외부 대시보드 ↔ 이 스프레드시트)
 * WEB_SHEET_NAME으로 고정된 시트를 대상으로 하며,
 * 요일 기준으로 동작하는 것 외에는 사이드바와 동일한 로직을 재사용한다.
 *****************************************************/
function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action;

  if (action === "getProgress") {
    return jsonOutput_(getOutputProgress());
  }

  var lock = LockService.getScriptLock();
  var gotLock = false;

  try {
    gotLock = lock.tryLock(25000);
    if (!gotLock) {
      return jsonOutput_({ status: "error", message: "다른 출력 작업이 진행 중입니다. 잠시 후 다시 시도해주세요." });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(WEB_SHEET_NAME);

    if (!sheet) {
      return jsonOutput_({ status: "error", message: "'" + WEB_SHEET_NAME + "' 시트를 찾을 수 없습니다." });
    }

    if (action === "resetViewMode") {
      resetViewMode(sheet);
      return jsonOutput_({ status: "ok", links: [] });
    }

    var weekday = String(body.weekday || getTodayWeekday_()).trim();
    if (!WEEKDAY_FLAG_COL[weekday]) {
      return jsonOutput_({ status: "error", message: "요일 값이 올바르지 않습니다." });
    }

    var confMap = {
      exportTodayPackage: { includeTotal: true, includeCourses: true, label: "오늘 전체 출력" },
      exportCloserByDate: { includeTotal: true, includeCourses: false, label: "마감자 출력" },
      exportByDateDriversOnly: { includeTotal: false, includeCourses: true, label: "코스별 출력" }
    };

    var conf = confMap[action];
    if (!conf) {
      return jsonOutput_({ status: "error", message: "알 수 없는 요청입니다: " + action });
    }

    var links = buildExportJobs_(ss, sheet, weekday, {
      includeTotal: conf.includeTotal,
      includeCourses: conf.includeCourses,
      label: conf.label,
      skipValidation: true
    });

    logOutputLinks_(links || []);

    var folderUrl = "";
    try { folderUrl = getOutputFolderForToday_().getUrl(); } catch (e2) {}

    return jsonOutput_({ status: "ok", links: links || [], folderUrl: folderUrl });

  } catch (err) {
    progressFail_(err.message);
    return jsonOutput_({ status: "error", message: err.message });
  } finally {
    if (gotLock) lock.releaseLock();
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

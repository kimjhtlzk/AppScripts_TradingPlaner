/**
 * 자동 수식 적용 스크립트
 * A열에 종목코드 입력시 자동으로 B, C, D열 수식 적용
 */

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // A열 또는 E열이 수정되었는지 확인
  if (range.getColumn() !== 1 && range.getColumn() !== 5) return;

  var row = range.getRow();

  // 헤더 행(1행)은 제외
  if (row === 1) return;

  var stockCode = sheet.getRange(row, 1).getValue();
  var market = sheet.getRange(row, 5).getValue();

  // A열에 값이 있고 E열에 시장구분이 있을 때만 수식 적용
  if (stockCode !== "" && market !== "") {
    applyFormulas(sheet, row);
  } else if (stockCode === "") {
    // A열이 비어있으면 B, C, D열도 비우기
    sheet.getRange(row, 2, 1, 3).clearContent();
  }
}

function applyFormulas(sheet, row) {
  // B열: 종목명
  var nameFormula = '=IF(A' + row + '="","",IF(E' + row + '="NASDAQ",IFERROR(GOOGLEFINANCE(A' + row + ',"name"),A' + row + '),IFERROR(REGEXEXTRACT(IMPORTXML("https://finance.naver.com/item/main.nhn?code="&A' + row + ',"//title"),"^([^:]+)"),"")))';

  // C열: 현재가
  var priceFormula = '=IF(A' + row + '="","",IF(E' + row + '="NASDAQ",IFERROR(GOOGLEFINANCE(A' + row + ',"price"),"N/A"),IFERROR(VALUE(SUBSTITUTE(IMPORTXML("https://finance.naver.com/item/sise.naver?code="&A' + row + ',"//strong[@id=\'_nowVal\']"),",","")),"로딩 실패")))';

  // D열: 등락률
  var changeFormula = '=IF(A' + row + '="","",IF(E' + row + '="NASDAQ",IFERROR(GOOGLEFINANCE(A' + row + ',"changepct")/100,"N/A"),IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(IMPORTXML("https://finance.naver.com/item/sise.naver?code="&A' + row + ',"//strong[@id=\'_rate\']//span"),"%",""),",",""))/100,"로딩 실패")))';

  // 수식 적용
  sheet.getRange(row, 2).setFormula(nameFormula);
  sheet.getRange(row, 3).setFormula(priceFormula);
  sheet.getRange(row, 4).setFormula(changeFormula);

  // 시장구분에 따라 C열 통화 서식 적용
  var market = sheet.getRange(row, 5).getValue();
  var priceCell = sheet.getRange(row, 3);

  if (market === "NASDAQ") {
    // 미국 주식: 달러 표시, 소수점 2자리
    priceCell.setNumberFormat("$#,##0.00");
  } else if (market === "KOSPI" || market === "KOSDAQ") {
    // 한국 주식: 원화 표시, 소수점 없음
    priceCell.setNumberFormat("₩#,##0");
  }
}

/**
 * 기존 데이터에 일괄 수식 적용
 */
function applyFormulasToAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  // 2행부터 마지막 행까지 수식 적용
  for (var row = 2; row <= lastRow; row++) {
    var stockCode = sheet.getRange(row, 1).getValue();
    var market = sheet.getRange(row, 5).getValue();

    if (stockCode !== "" && market !== "") {
      applyFormulas(sheet, row);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 모든 행에 수식 적용 완료', '완료', 3);
}

/**
 * 메뉴에 추가
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 주식 트래커')
    .addItem('🔄 수동 새로고침', 'manualRefresh')
    .addItem('⏰ 1분 자동 새로고침 시작', 'startAutoRefresh')
    .addItem('⏹ 자동 새로고침 중지', 'stopAutoRefresh')
    .addItem('🔧 오류 수정', 'fixFormulas')
    .addSeparator()
    .addItem('📝 모든 행에 수식 적용', 'applyFormulasToAll')
    .addItem('💱 통화 서식 일괄 적용', 'applyCurrencyFormatToAll')
    .addToUi();
}

// 기존 auto_refresh_script.gs의 함수들도 포함
function manualRefresh() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("B2:F50");
  var formulas = dataRange.getFormulas();

  // 현재 시간을 F1 셀에 표시
  var now = new Date();
  var timeString = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.getRange("F1").setValue("마지막 새로고침: " + timeString);

  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j] && formulas[i][j].includes("IMPORTXML")) {
        var cell = dataRange.getCell(i + 1, j + 1);
        var originalFormula = formulas[i][j];
        cell.setValue("");
        SpreadsheetApp.flush();
        cell.setFormula(originalFormula);
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 새로고침 완료', '완료', 2);
}

function fixFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("B2:F50");
  var formulas = dataRange.getFormulas();
  var fixedCount = 0;

  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j]) {
        var formula = formulas[i][j];
        if (formula.includes("))&t=") || formula.includes("))&r=")) {
          formula = formula.replace(/\)\)&[tr]=\d+/, "))");
          dataRange.getCell(i + 1, j + 1).setFormula(formula);
          fixedCount++;
        }
        if (formula.match(/\)\)[^)]*$/)) {
          formula = formula.replace(/(\)\))[^)]*$/, "$1");
          dataRange.getCell(i + 1, j + 1).setFormula(formula);
          fixedCount++;
        }
      }
    }
  }

  if (fixedCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ ' + fixedCount + '개 수식 수정 완료', '수정 완료', 3);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('👍 수정할 오류가 없습니다', '확인', 2);
  }
}

function startAutoRefresh() {
  stopAutoRefresh();
  ScriptApp.newTrigger('safeRefresh')
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast('⏰ 1분마다 자동 새로고침 시작', '설정 완료', 3);
}

function safeRefresh() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("B2:F50");
  var formulas = dataRange.getFormulas();

  // 현재 시간을 F1 셀에 표시
  var now = new Date();
  var timeString = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.getRange("F1").setValue("마지막 새로고침: " + timeString);

  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      var formula = formulas[i][j];
      if (formula && formula.includes("IMPORTXML") && !formula.includes("&t=")) {
        var cell = dataRange.getCell(i + 1, j + 1);
        cell.setFormula(formula);
      }
    }
  }
}

function stopAutoRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() == 'safeRefresh' ||
        trigger.getHandlerFunction() == 'manualRefresh' ||
        trigger.getHandlerFunction() == 'forceRefresh') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('⏹ 자동 새로고침 중지', '중지', 2);
}

/**
 * 모든 행에 통화 서식 일괄 적용
 */
function applyCurrencyFormatToAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var updatedCount = 0;

  // 2행부터 마지막 행까지 통화 서식 적용
  for (var row = 2; row <= lastRow; row++) {
    var market = sheet.getRange(row, 5).getValue();
    var priceCell = sheet.getRange(row, 3);

    if (market === "NASDAQ") {
      // 미국 주식: 달러 표시, 소수점 2자리
      priceCell.setNumberFormat("$#,##0.00");
      updatedCount++;
    } else if (market === "KOSPI" || market === "KOSDAQ") {
      // 한국 주식: 원화 표시, 소수점 없음
      priceCell.setNumberFormat("₩#,##0");
      updatedCount++;
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ ' + updatedCount + '개 행에 통화 서식 적용 완료', '완료', 3);
}

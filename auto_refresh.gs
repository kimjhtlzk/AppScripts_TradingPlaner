/**
 * 수정된 자동 새로고침 스크립트
 * #NAME? 오류 해결 버전
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 주식 트래커')
    .addItem('🔄 수동 새로고침', 'manualRefresh')
    .addItem('⏰ 1분 자동 새로고침 시작', 'startAutoRefresh')
    .addItem('⏹ 자동 새로고침 중지', 'stopAutoRefresh')
    .addItem('🔧 오류 수정', 'fixFormulas')
    .addToUi();
}

/**
 * 수동 새로고침 - 수식을 건드리지 않고 재계산만 유도
 */
function manualRefresh() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("B2:F50");
  var formulas = dataRange.getFormulas();

  // 각 수식을 그대로 다시 설정 (캐시 클리어)
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j] && formulas[i][j].includes("IMPORTXML")) {
        var cell = dataRange.getCell(i + 1, j + 1);
        var originalFormula = formulas[i][j];

        // 임시로 빈 값 설정 후 원래 수식 그대로 재입력
        cell.setValue("");
        SpreadsheetApp.flush();
        cell.setFormula(originalFormula);
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 새로고침 완료', '완료', 2);
}

/**
 * 잘못된 수식 수정 (기존 오류 해결)
 */
function fixFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("B2:F50");
  var formulas = dataRange.getFormulas();
  var fixedCount = 0;

  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j]) {
        var formula = formulas[i][j];

        // &t= 이후 부분 제거 (잘못 추가된 타임스탬프)
        if (formula.includes("))&t=") || formula.includes("))&r=")) {
          formula = formula.replace(/\)\)&[tr]=\d+/, "))");
          dataRange.getCell(i + 1, j + 1).setFormula(formula);
          fixedCount++;
        }

        // 수식 끝에 잘못 추가된 파라미터 제거
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

/**
 * 1분마다 자동 새로고침
 */
function startAutoRefresh() {
  // 기존 트리거 삭제
  stopAutoRefresh();

  // 새 트리거 생성 (1분마다)
  ScriptApp.newTrigger('safeRefresh')
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast('⏰ 1분마다 자동 새로고침 시작', '설정 완료', 3);
}

/**
 * 안전한 새로고침 (수식 변경 없이)
 */
function safeRefresh() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("B2:F50");
  var formulas = dataRange.getFormulas();
  var values = dataRange.getValues();

  // IMPORTXML이 포함된 셀만 재계산
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      var formula = formulas[i][j];

      // IMPORTXML 수식이 있고 오류가 아닌 경우만
      if (formula && formula.includes("IMPORTXML") && !formula.includes("&t=")) {
        var cell = dataRange.getCell(i + 1, j + 1);

        // 수식 그대로 재설정 (변경 없이)
        cell.setFormula(formula);
      }
    }
  }
}

/**
 * 자동 새로고침 중지
 */
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
 * 대안: 캐시 무효화가 필요한 경우 (URL 내부에 파라미터 추가)
 */
function refreshWithCacheBust() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("B2:F50");
  var formulas = dataRange.getFormulas();
  var timestamp = new Date().getTime();

  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j] && formulas[i][j].includes("IMPORTXML")) {
        var formula = formulas[i][j];
        var cell = dataRange.getCell(i + 1, j + 1);

        // URL 내부에 파라미터 추가 (수식 안에서)
        if (formula.includes("?code=")) {
          // 기존 타임스탬프 제거
          formula = formula.replace(/&t=\d+/g, "");
          // 새 타임스탬프를 URL 안에 추가
          formula = formula.replace(/("https:\/\/[^"]+\?code=[^"]+)"/g, '$1&t=' + timestamp + '"');
        }

        cell.setFormula(formula);
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 캐시 무효화 새로고침 완료', '완료', 2);
}
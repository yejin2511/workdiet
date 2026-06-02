// ============================================
// 업무 다이어트 - Google Apps Script 코드
// ============================================
// 사용법:
// 1. 구글 시트 새로 만들기
// 2. 상단 메뉴 → 확장 프로그램 → Apps Script
// 3. 이 코드 전체 붙여넣기 (기존 코드 전부 삭제 후)
// 4. 저장 (Ctrl+S)
// 5. 배포 → 새 배포 → 유형: 웹 앱
//    - 설명: 업무다이어트
//    - 다음 사용자로 실행: 나(모모님 계정)
//    - 액세스 권한: 모든 사용자
// 6. 배포 → 배포 URL 복사
// 7. 복사한 URL을 앱 HTML 파일의 SCRIPT_URL 변수에 붙여넣기
// ============================================

var SHEET_NAME = '업무목록';

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    // 시트가 없으면 새로 만들고 헤더 추가
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      var headers = [
        '제출일시', '이름', '부서', '직무', '업무명',
        '주기', '1회 소요시간', '월 빈도',
        '중요한 이유', '왜 하나요', '없애도 될까요',
        'HR 판정(초안)', '비고'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1A1917')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // 데이터 파싱
    var data = JSON.parse(e.postData.contents);
    var rows = data.rows;
    var now = new Date();
    var timestamp = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    // 각 업무를 한 행씩 추가
    rows.forEach(function(row) {
      sheet.appendRow([
        timestamp,
        row.name,
        row.dept,
        row.job,
        row.task,
        row.period,
        row.time,
        row.freq,
        row.importance,
        row.why,
        row.keep,
        row.verdict,
        row.reason
      ]);
    });

    // 버리기 검토 행 색상 표시
    colorRows(sheet);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', count: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function colorRows(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var verdictCol = 12; // HR 판정 컬럼
  var data = sheet.getRange(2, verdictCol, lastRow - 1, 1).getValues();
  data.forEach(function(row, i) {
    var rowNum = i + 2;
    var verdict = row[0];
    var range = sheet.getRange(rowNum, 1, 1, 13);
    if (verdict === '버리기 검토') {
      range.setBackground('#FEF0EC');
    } else if (verdict === '축소 검토') {
      range.setBackground('#FEF8EC');
    } else {
      range.setBackground('#FFFFFF');
    }
  });
}

// GET 요청 테스트용 (브라우저에서 URL 접속하면 확인 가능)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '업무 다이어트 API 정상 작동 중' }))
    .setMimeType(ContentService.MimeType.JSON);
}

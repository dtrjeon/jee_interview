/**
 * 62기 전도폭발 인터뷰 체크리스트 - 구글시트 제출 백엔드 (폼 리다이렉트 방식)
 *
 * [설치 방법]
 * 1. 새 구글시트를 만들고, 시트 이름을 "인터뷰기록" 으로 변경 (또는 아래 SHEET_NAME 수정)
 * 2. 확장 프로그램 > Apps Script 에 들어가서 기본 코드를 지우고 이 코드를 붙여넣기
 * 3. 아래 RETURN_URL을 실제 체크리스트가 배포될 GitHub Pages 주소로 수정
 * 4. 저장 후 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자 (Anyone)
 * 5. 배포된 웹 앱 URL을 복사해서 ee_interview_checklist.html의
 *    <form id="submitForm" method="POST" action="여기에_배포된_웹앱_URL_붙여넣기"> 부분에 붙여넣기
 */

const SHEET_NAME = "인터뷰기록";
const RETURN_URL = "https://dtrjeon.github.io/여기에_체크리스트_경로";

const HEADERS = [
  "제출시각", "기수", "이름", "담당자", "일시", "시간", "인터뷰방식",
  "일반생활", "교회내직분", "봉사활동", "신청반", "현장실습가능시간",
  "훈련기간지속가능", "삼단계지속가능", "팀책임응답", "순종서약",
  "필기시험형식", "인터뷰의견", "체크현황"
];

function doGet(e) {
  if (e.parameter.action === "list") {
    const sheet = getOrCreateSheet_();
    const lastRow = sheet.getLastRow();
    let records = [];
    if (lastRow > 1) {
      const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      records = values.map(row => {
        const obj = {};
        HEADERS.forEach((h, i) => {
          obj[h] = row[i] instanceof Date ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : row[i];
        });
        return obj;
      });
    }
    return ContentService
      .createTextOutput(JSON.stringify({ records }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutput("이 웹앱은 체크리스트 제출 전용입니다.");
}

function doPost(e) {
  const sheet = getOrCreateSheet_();
  const data = e.parameter; // 폼 제출 (application/x-www-form-urlencoded)

  const row = HEADERS.map(h => {
    if (h === "제출시각") return new Date();
    return data[h] !== undefined ? data[h] : "";
  });
  sheet.appendRow(row);

  return HtmlService.createHtmlOutput(
    `<html><head><meta charset="utf-8">
     <meta http-equiv="refresh" content="2;url=${RETURN_URL}">
     <style>
       body{font-family:sans-serif;background:#233150;color:#F6F2E9;
            display:flex;align-items:center;justify-content:center;
            height:100vh;margin:0;text-align:center;}
       div{padding:20px;}
     </style></head>
     <body><div>제출이 완료되었습니다.<br>잠시 후 이동합니다...</div></body></html>`
  );
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}


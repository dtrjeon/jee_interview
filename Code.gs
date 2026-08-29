/**
 * 62기 전도폭발 인터뷰 체크리스트 - 구글시트 제출 백엔드 (폼 리다이렉트 방식 + 사진 OCR)
 *
 * [설치 방법]
 * 1. 새 구글시트를 만들고, 시트 이름을 "인터뷰기록" 으로 변경 (또는 아래 SHEET_NAME 수정)
 * 2. 확장 프로그램 > Apps Script 에 들어가서 기본 코드를 지우고 이 코드를 붙여넣기
 * 3. OCR 엔진: Google Cloud Vision API 사용 (Drive OCR보다 인식률이 좋습니다)
 *    a. https://console.cloud.google.com 에서 프로젝트 생성(또는 기존 프로젝트 사용)
 *    b. "API 및 서비스 > 라이브러리"에서 "Cloud Vision API" 사용 설정
 *    c. "API 및 서비스 > 사용자 인증 정보"에서 API 키 생성 (Vision API로 제한 권장)
 *    d. Apps Script 편집기 좌측 톱니바퀴(프로젝트 설정) > "스크립트 속성"에
 *       속성 이름: VISION_API_KEY / 값: 발급받은 API 키 를 추가
 * 4. 아래 RETURN_URL을 실제 체크리스트가 배포될 GitHub Pages 주소로 수정
 * 5. 저장 후 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자 (Anyone)
 * 6. 배포된 웹 앱 URL을 복사해서 ee_interview_checklist.html의
 *    <form id="submitForm" method="POST" action="여기에_배포된_웹앱_URL_붙여넣기"> 부분에 붙여넣기
 *
 * ※ API 키는 이 스크립트 속성에만 저장되며(서버 측), 클라이언트(HTML)에는 절대 노출되지 않습니다.
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
  // 1) 카메라 촬영 사진 OCR 요청 (JSON body: {action:'ocr', image, mimeType})
  if (e.postData && e.postData.contents && e.postData.contents.trim().startsWith("{")) {
    try {
      const body = JSON.parse(e.postData.contents);
      if (body.action === "ocr") {
        return handleOcr_(body);
      }
    } catch (err) {
      // JSON이 아니면 아래 폼 제출 처리로 진행
    }
  }

  // 2) 체크리스트 폼 제출 (application/x-www-form-urlencoded)
  const sheet = getOrCreateSheet_();
  const data = e.parameter;

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

function handleOcr_(body) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty("VISION_API_KEY");
    if (!apiKey) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "VISION_API_KEY가 스크립트 속성에 설정되어 있지 않습니다." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const base64 = body.image.split(",")[1]; // "data:image/jpeg;base64,...." 에서 앞부분 제거

    const payload = {
      requests: [{
        image: { content: base64 },
        features: [{ type: "DOCUMENT_TEXT_DETECTION" }],
        imageContext: { languageHints: ["ko"] }
      }]
    };

    const response = UrlFetchApp.fetch(
      "https://vision.googleapis.com/v1/images:annotate?key=" + apiKey,
      {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    const result = JSON.parse(response.getContentText());
    const annotation = result.responses && result.responses[0];

    if (!annotation || annotation.error) {
      const msg = annotation && annotation.error ? annotation.error.message : "인식 결과가 없습니다.";
      return ContentService
        .createTextOutput(JSON.stringify({ error: msg }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const text = (annotation.fullTextAnnotation && annotation.fullTextAnnotation.text) || "";

    return ContentService
      .createTextOutput(JSON.stringify({ text: text }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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


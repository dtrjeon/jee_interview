/**
 * 62기 전도폭발 인터뷰 체크리스트 - 구글시트 제출 백엔드 (폼 리다이렉트 방식 + 사진 OCR)
 *
 * [설치 방법]
 * 1. 새 구글시트를 만들고, 시트 이름을 "인터뷰기록" 으로 변경 (또는 아래 SHEET_NAME 수정)
 * 2. 확장 프로그램 > Apps Script 에 들어가서 기본 코드를 지우고 이 코드를 붙여넣기
 * 3. OCR 엔진: Anthropic API(Claude, 이미지 인식) 사용 — 손글씨 인식에 강합니다
 *    a. console.anthropic.com 에서 API 키 발급 (sk-ant-... 형식)
 *    b. Apps Script 편집기 좌측 톱니바퀴(프로젝트 설정) > "스크립트 속성"에
 *       속성 이름: ANTHROPIC_API_KEY / 값: 발급받은 API 키 를 추가
 * 4. 아래 RETURN_URL을 실제 체크리스트가 배포될 GitHub Pages 주소로 수정
 * 5. 저장 후 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자 (Anyone)
 * 6. 배포된 웹 앱 URL을 복사해서 ee_interview_checklist.html의
 *    <form id="submitForm" method="POST" action="여기에_배포된_웹앱_URL_붙여넣기"> 부분에 붙여넣기
 *
 * ※ API 키는 이 스크립트 속성에만 저장되며(서버 측), 클라이언트(HTML)에는 절대 노출되지 않습니다.
 * ※ Anthropic API는 사용량에 따라 과금됩니다(console.anthropic.com에서 요금 확인 가능).
 */

const SHEET_NAME = "인터뷰기록";
const RETURN_URL = "https://dtrjeon.github.io/jee_interview/";

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
    const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
    if (!apiKey) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "ANTHROPIC_API_KEY가 스크립트 속성에 설정되어 있지 않습니다." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const base64 = body.image.split(",")[1]; // "data:image/jpeg;base64,...." 에서 앞부분 제거
    const mediaType = body.mimeType || "image/jpeg";

    const payload = {
      model: "claude-sonnet-5",
      max_tokens: 2000,
      messages: [{
        role: "user",
        content: [
          {
            type: "image",
            source: { type: "base64", media_type: mediaType, data: base64 }
          },
          {
            type: "text",
            text: "이 이미지 안의 모든 텍스트를 빠짐없이 그대로 옮겨 적어줘. 인쇄된 글자뿐 아니라 손글씨로 적힌 내용, 체크(✓) 표시, 동그라미 표시도 최대한 정확하게 판독해서 포함해줘. 설명이나 요약 없이 원문 텍스트만 출력해줘."
          }
        ]
      }]
    };

    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      contentType: "application/json",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());

    if (result.error) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: result.error.message || "Anthropic API 오류" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const text = (result.content && result.content[0] && result.content[0].text) || "";

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


/**
 * 편집기에서 doPost를 직접 테스트하기 위한 함수.
 * 이 함수를 선택해서 "실행" 버튼을 누르면 폼 제출을 흉내내서 테스트할 수 있습니다.
 */
function testDoPost() {
  const fakeEvent = {
    parameter: {
      기수: "62",
      이름: "테스트",
      담당자: "테스트담당자",
      일시: "2026-08-30",
      시간: "10:00"
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}

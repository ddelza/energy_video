const SS_MAIN_ID = '1RqwApuuD52wKIUo1ZcRsQaerU2epRqwGyGjAfWmhlHg'; // 학생 정보 / 설문지 응답 시트1 (이 스크립트가 바인딩된 스프레드시트)
const SS_PREP_ID = '17HV5qLlOhlxsll8ZZdAUxe3pqAmwA7n1ysmg4gCukU4';
const SS_GRADING_ID = SS_PREP_ID; // '채점' 탭이 같은 스프레드시트에 있음
const SS_PW_ID = '1JcgoufQUypJR6ItEBGWR7xVE-e1vBqXqfYddkEgXlJg';

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;

  if (!action) {
    if (e.parameter.page === 'youtube') {
      return HtmlService.createHtmlOutputFromFile('youtube')
        .setTitle('발표 유튜브 링크 모아보기')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (e.parameter.page === 'student') {
      return HtmlService.createHtmlOutputFromFile('student')
        .setTitle('내 발표 현황 조회')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('발표회 청중 보고서 조회')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  let result;
  if (action === 'getStudents') {
    result = getStudents();
  } else if (action === 'getSubmissions') {
    result = getSubmissions(e.parameter.반, e.parameter.번호);
  } else if (action === 'getPrepStatus') {
    result = getPrepStatus(e.parameter.반, e.parameter.성명);
  } else if (action === 'getYoutubeByClass') {
    result = getYoutubeByClass(e.parameter.반);
  } else if (action === 'getMyData') {
    try {
      result = getMyData(e.parameter.반, e.parameter.번호, e.parameter.성명, e.parameter.비밀번호);
    } catch (err) {
      result = { error: err.toString() };
    }
  } else {
    result = { error: 'unknown action' };
  }

  const json = JSON.stringify(result);

  // JSONP: callback 파라미터가 있으면 함수 호출 형태로 반환
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getStudents() {
  const ss = SpreadsheetApp.openById(SS_MAIN_ID);
  const infoSheet = ss.getSheetByName('학생 정보');
  const respSheet = ss.getSheetByName('설문지 응답 시트1');

  // 제출 횟수 집계: key = "반_번호"
  const countMap = {};
  const respData = respSheet.getDataRange().getValues();
  for (let i = 1; i < respData.length; i++) {
    const row = respData[i];
    if (!row[0]) continue;
    const key = String(row[1]) + '_' + String(row[2]);
    countMap[key] = (countMap[key] || 0) + 1;
  }

  const data = infoSheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const key = String(row[1]) + '_' + String(row[2]);
    students.push({
      학번: String(row[0]),
      반: String(row[1]),
      번호: String(row[2]),
      성명: String(row[3]),
      모둠: String(row[4]),
      제출수: countMap[key] || 0
    });
  }
  return students;
}

function norm(s) {
  return String(s || '').replace(/\s+/g, '').trim();
}

function getPrepStatus(반, 성명) {
  const sheet = SpreadsheetApp.openById(SS_PREP_ID).getSheetByName('시트1');
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const members = [row[1], row[2], row[3], row[4]].map(String).filter(Boolean);
    const isMatch = String(row[0]) === String(반) &&
      members.some(m => norm(m) === norm(성명));
    if (isMatch) {
      results.push({
        반: String(row[0]),
        모둠원: members,
        주제: String(row[5] || ''),
        대본: String(row[6] || ''),
        캔바링크: String(row[7] || ''),
        예상질문: [row[8], row[9], row[10], row[11]].map(String).filter(Boolean)
      });
    }
  }
  return results;
}

function getYoutubeByClass(반) {
  const sheet = SpreadsheetApp.openById(SS_PREP_ID).getSheetByName('시트1');
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[0]) !== String(반)) continue;
    const members = [row[1], row[2], row[3], row[4]].map(String).filter(Boolean);
    results.push({
      반: String(row[0]),
      모둠원: members,
      주제: String(row[5] || ''),
      유튜브링크: String(row[12] || '')
    });
  }
  return results;
}

// 빈 칸이면 null(채점중), 숫자면 Number 반환
function scoreOrNull(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function getMyData(반, 번호, 성명, 비밀번호) {
  // 1. 학생 정보 확인 (반+번호+이름 일치해야 함)
  const infoSheet = SpreadsheetApp.openById(SS_MAIN_ID).getSheetByName('학생 정보');
  const infoData = infoSheet.getDataRange().getValues();
  let student = null;
  for (let i = 1; i < infoData.length; i++) {
    const row = infoData[i];
    if (!row[0]) continue;
    if (String(row[1]) === String(반) && String(row[2]) === String(번호) && norm(row[3]) === norm(성명)) {
      student = { 학번: String(row[0]), 반: String(row[1]), 번호: String(row[2]), 성명: String(row[3]), 모둠: String(row[4]) };
      break;
    }
  }
  if (!student) return { error: '학생 정보를 찾을 수 없습니다. 반/번호/이름을 다시 확인해주세요.' };

  // 2. 비밀번호 확인
  const pwSheet = SpreadsheetApp.openById(SS_PW_ID).getSheets()[0];
  const pwData = pwSheet.getDataRange().getValues();
  let correctPw = null;
  for (let i = 1; i < pwData.length; i++) {
    const row = pwData[i];
    if (!row[0]) continue;
    if (String(row[0]) === student.학번) {
      correctPw = String(row[2]);
      break;
    }
  }
  if (correctPw === null) return { error: '비밀번호 정보를 찾을 수 없습니다. 선생님께 문의해주세요.' };
  const MASTER_PW = 'byung0703!';
  if (String(비밀번호) !== correctPw && String(비밀번호) !== MASTER_PW) return { error: '비밀번호가 일치하지 않습니다.' };

  // 3. 발표 준비 현황 (시트1)
  const prepSheet = SpreadsheetApp.openById(SS_PREP_ID).getSheetByName('시트1');
  const prepData = prepSheet.getDataRange().getValues();
  let prep = null;
  for (let i = 1; i < prepData.length; i++) {
    const row = prepData[i];
    if (!row[0]) continue;
    const members = [row[1], row[2], row[3], row[4]].map(String).filter(Boolean);
    if (String(row[0]) === student.반 && members.some(m => norm(m) === norm(student.성명))) {
      prep = {
        모둠원: members,
        주제: String(row[5] || ''),
        대본: String(row[6] || ''),
        캔바링크: String(row[7] || ''),
        예상질문: [row[8], row[9], row[10], row[11]].map(String).filter(Boolean)
      };
      break;
    }
  }

  // 4. 채점 (반 + 이름 일치하는 행)
  const gradeSheet = SpreadsheetApp.openById(SS_GRADING_ID).getSheetByName('채점');
  const gradeData = gradeSheet.getDataRange().getValues();
  let scoreRow = null;
  for (let i = 1; i < gradeData.length; i++) {
    const row = gradeData[i];
    if (!row[0]) continue;
    if (String(row[0]) === student.반 && norm(row[1]) === norm(student.성명)) {
      scoreRow = row;
      break;
    }
  }

  const scores = {
    대본: { 항목: '발표 대본 (핵심 개념 작성하기)', 점수: scoreRow ? scoreOrNull(scoreRow[3]) : null, 만점: 3 },
    예상질문: { 항목: '예상 질문 작성하기', 점수: scoreRow ? scoreOrNull(scoreRow[4]) : null, 만점: 4 },
    자료: { 항목: '발표 자료 준비·제작하기', 점수: scoreRow ? scoreOrNull(scoreRow[5]) : null, 만점: 3 },
    메모: { 항목: '청중보고서 - 발표내용 메모하기', 점수: scoreRow ? scoreOrNull(scoreRow[6]) : null, 만점: 6 },
    질문점수합: { 항목: '청중보고서 - 질문 점수 (질문 만들기 + 보너스)', 점수: scoreRow ? scoreOrNull(scoreRow[9]) : null, 만점: 4 }
  };

  const items = [scores.대본, scores.예상질문, scores.자료, scores.메모, scores.질문점수합];
  const allGraded = items.every(s => s.점수 !== null);
  const total = allGraded
    ? items.reduce((sum, s) => sum + s.점수, 0)
    : (scoreRow ? scoreOrNull(scoreRow[10]) : null);
  scores.총점 = { 항목: '총점', 점수: total, 만점: 20 };

  // 5. 청중보고서 제출 내역
  const submissions = getSubmissions(student.반, student.번호);

  return {
    학생: student,
    발표준비: prep,
    점수: scores,
    청중보고서: submissions
  };
}

function getSubmissions(반, 번호) {
  const sheet = SpreadsheetApp.openById(SS_MAIN_ID).getSheetByName('설문지 응답 시트1');
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[1]) === String(반) && String(row[2]) === String(번호)) {
      results.push({
        타임스탬프: row[0] instanceof Date
          ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
          : String(row[0]),
        주제: String(row[3] || ''),
        메모: String(row[4] || ''),
        질문: String(row[5] || '')
      });
    }
  }
  return results;
}

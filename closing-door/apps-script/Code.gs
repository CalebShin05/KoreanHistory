/** 「닫히는 문」 학생 기록 받는 곳 — 스프레드시트 '닫히는 문 학생 기록'의 '제출' 탭에 한 줄씩 저장
 *  배포: 웹 앱 / 실행 사용자: 나 / 액세스: 모든 사용자 (2026.11.26 배포)
 */
const SHEET_ID = '1R38DhZYbwVoqX3B-N-IKKDgD9Gz4PP4iN5c2wxtIBhk';
const HEADERS = ['제출시각','학번','이름','차시','질문','의견','근거카드','평균신뢰도','사료수집','미니게임재시도','생각기록','기사기록'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('제출') || ss.insertSheet('제출');
    if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
    const safe = v => { const s = (v === undefined || v === null) ? '' : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; };
    sh.appendRow(HEADERS.map(h => safe(d[h])));
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('닫히는 문 기록함이 연결되어 있습니다.');
}

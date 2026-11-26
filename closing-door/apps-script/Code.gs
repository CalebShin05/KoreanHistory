/** 「닫히는 문」 학생 기록 받는 곳 + 세특 초안 도우미
 *  - 학생 제출 → '제출' 탭에 한 줄씩 저장 (1~4차시)
 *  - 저장할 때마다 '세특' 탭을 학번별로 다시 정리하고 자동 초안을 만듦
 *  - '세특' 탭의 '최종 문구(교사 작성)' 칸은 다시 정리해도 지워지지 않음 (학번 기준 보존)
 *  - 바이트: NEIS 기준 (한글 등 3바이트, 영문·숫자·공백·엔터 1바이트)
 *  배포: 웹 앱 / 실행 사용자: 나 / 액세스: 모든 사용자
 */
const SHEET_ID = '1R38DhZYbwVoqX3B-N-IKKDgD9Gz4PP4iN5c2wxtIBhk';
const HEADERS = ['제출시각','학번','이름','차시','질문','의견','근거카드','평균신뢰도','사료수집','미니게임재시도','생각기록','기사기록'];

const SE_HEADERS = ['학번','이름','참여 차시','평균 신뢰도','1차시 의견(원문)','2차시 의견(원문)','3차시 의견(원문)','4차시 의견(원문)','취재 수첩 입장(참고)','자동 초안','초안 바이트','최종 문구(교사 작성)','최종 바이트','목표 대비'];
const LESSONS = ['1차시', '2차시', '3차시', '4차시'];
const TOPIC = {
  '1차시': '3·15 부정 선거와 4·19 혁명의 전개를 탐구하며',
  '2차시': '4·19 혁명으로 열린 민주주의가 1년 만에 다시 위기를 맞은 원인을 탐구하며',
  '3차시': '유신 체제의 긴급 조치를 분석하며 법으로 비판을 막는 것의 한계를 탐구하고',
  '4차시': '5·18 민주화 운동과 6월 민주 항쟁의 흐름을 탐구하며'
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('제출') || ss.insertSheet('제출');
    if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
    const safe = v => { const s = (v === undefined || v === null) ? '' : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; };
    sh.appendRow(HEADERS.map(h => safe(d[h])));
    try { buildSeteuk_(ss); } catch (err) { console.error(err); }  // 세특 정리가 실패해도 제출은 저장됨
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('닫히는 문 기록함이 연결되어 있습니다.');
}

/** 편집기에서 직접 실행해 '세특' 탭을 다시 만들 때 사용 */
function rebuildSeteuk() {
  buildSeteuk_(SpreadsheetApp.openById(SHEET_ID));
}

/* ---------------- 세특 정리 ---------------- */
function buildSeteuk_(ss) {
  const src = ss.getSheetByName('제출');
  if (!src || src.getLastRow() < 2) return;
  const rows = src.getRange(2, 1, src.getLastRow() - 1, HEADERS.length).getValues();
  const col = name => HEADERS.indexOf(name);

  // 학번별·차시별 최신 제출만 사용
  const by = {};
  rows.forEach(r => {
    const sid = String(r[col('학번')]).trim();
    if (!sid) return;
    const L = String(r[col('차시')]).trim();
    const s = by[sid] = by[sid] || { sid, name: '', lessons: {} };
    s.name = String(r[col('이름')]).trim() || s.name;
    s.lessons[L] = r;  // 아래쪽(최신) 줄이 덮어씀
  });

  const settings = ensureSettings_(ss);
  let se = ss.getSheetByName('세특');
  const fresh = !se;
  if (fresh) se = ss.insertSheet('세특');

  // 교사가 쓴 최종 문구 보존
  const keep = {};
  if (se.getLastRow() >= 2) {
    se.getRange(2, 1, se.getLastRow() - 1, SE_HEADERS.length).getValues().forEach(r => {
      if (r[0] !== '' && r[11] !== '') keep[String(r[0]).trim()] = r[11];
    });
  }

  const sids = Object.keys(by).sort();
  const out = sids.map((sid, i) => {
    const s = by[sid];
    const Ls = LESSONS.filter(L => s.lessons[L]);
    const scores = Ls.map(L => Number(s.lessons[L][col('평균신뢰도')])).filter(n => !isNaN(n) && n > 0);
    const avg = scores.length ? Math.round(scores.reduce((a, b) => a + b, 0) / scores.length) : '';
    const op = L => s.lessons[L] ? String(s.lessons[L][col('의견')]) : '';
    const stances = Ls.map(L => stanceOf_(String(s.lessons[L][col('생각기록')]))).filter(Boolean).join('\n');
    const draft = draftOf_(s, Ls, avg, col);
    const n = i + 2;
    const bytes = c => `=IF(${c}${n}="","",LEN(${c}${n})+2*LEN(REGEXREPLACE(${c}${n},"[[:ascii:]]","")))`;
    return [
      s.sid, s.name, Ls.join(', '), avg, op('1차시'), op('2차시'), op('3차시'), op('4차시'), stances,
      draft, bytes('J'), keep[sid] || '', bytes('L'),
      `=IF(M${n}="",IF(K${n}="","","초안 "&IF(K${n}>설정!$B$1,K${n}-설정!$B$1&"바이트 초과","OK")),IF(M${n}>설정!$B$1,"⚠ "&(M${n}-설정!$B$1)&"바이트 초과","✓ "&(설정!$B$1-M${n})&"바이트 남음"))`
    ];
  });

  se.clear();
  se.getRange(1, 1, 1, SE_HEADERS.length).setValues([SE_HEADERS])
    .setFontWeight('bold').setBackground('#efe4c8').setWrap(true);
  se.setFrozenRows(1);
  se.setFrozenColumns(2);
  if (out.length) {
    se.getRange(2, 1, out.length, SE_HEADERS.length).setValues(out);
    se.getRange(2, 1, out.length, 1).setNumberFormat('@');
    se.getRange(2, 1, out.length, SE_HEADERS.length).setWrap(true).setVerticalAlignment('top');
    se.getRange(2, 12, out.length, 1).setBackground('#fffbe6');
  }
  const widths = [70, 70, 110, 70, 220, 220, 220, 220, 200, 380, 70, 380, 70, 120];
  widths.forEach((w, i) => se.setColumnWidth(i + 1, w));
  // 목표 초과 빨간색
  const rng = se.getRange('N2:N');
  se.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('초과').setFontColor('#b3261e').setBackground('#fde7e4').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✓').setFontColor('#2e6b2e').setRanges([rng]).build()
  ]);
}

function ensureSettings_(ss) {
  let st = ss.getSheetByName('설정');
  if (st) return st;
  st = ss.insertSheet('설정');
  st.getRange('A1:B1').setValues([['목표 바이트', 1500]]).setFontWeight('bold');
  st.getRange('A2').setValue('한국사2(공통과목)는 1·2학기 합산 1,500바이트입니다. 이 활동 문장에 쓸 분량으로 B1을 바꾸면 세특 탭의 목표 대비가 바로 다시 계산됩니다.');
  st.getRange('A4').setValue('바이트 계산기 — 아래 A5 칸에 문장을 붙여 넣으세요').setFontWeight('bold');
  st.getRange('B4').setValue('NEIS 바이트').setFontWeight('bold');
  st.getRange('B5').setFormula('=IF(A5="","",LEN(A5)+2*LEN(REGEXREPLACE(A5,"[[:ascii:]]","")))');
  st.getRange('A6').setValue('기준: 한글·한자·특수 기호 3바이트 / 영문·숫자·공백·엔터 1바이트 (학교생활기록부 기재요령)');
  st.setColumnWidth(1, 520); st.setColumnWidth(2, 110);
  st.getRange('A5').setWrap(true).setBackground('#fffbe6');
  return st;
}

/* ---------------- 초안 문장 만들기 ---------------- */
function draftOf_(s, Ls, avg, col) {
  if (!Ls.length) return '';
  const parts = [];
  parts.push('‘닫히는 문’ 역사 게임 수업에서 신문 기자가 되어 1960~1980년대 민주주의의 흐름을 사료로 취재함.');
  Ls.forEach(L => {
    const r = s.lessons[L];
    const cards = String(r[col('근거카드')]).split('/').map(cardShort_).filter(Boolean);
    const claim = claimOf_(String(r[col('의견')]));
    const cardText = cards.length >= 2 ? `「${cards[0]}」${josa_(cards[0] + '」', '와', '과')} 「${cards[1]}」` : cards.length ? `「${cards[0]}」` : '';
    const topic = TOPIC[L] || '';
    if (claim && cardText) parts.push(`${topic} ${cardText} 사료를 근거로 ‘${claim}’${/이다$/.test(claim) ? '라는' : '는'} 주장을 펼침.`);
    else if (cardText) parts.push(`${topic} ${cardText} 사료를 근거로 자신의 견해를 서술함.`);
  });
  const retried = Ls.some(L => /([2-9])번째 송고/.test(String(s.lessons[L][col('기사기록')])));
  if (retried) parts.push('편집장의 피드백을 반영해 기사를 다시 송고하며 사료 해석을 다듬음.');
  else if (avg !== '' && avg >= 85) parts.push('사료에 맞는 헤드라인과 근거를 정확히 골라 신뢰도 높은 기사를 작성함.');
  else if (avg !== '') parts.push('사료를 근거로 사실에 맞는 기사를 쓰는 활동을 성실히 수행함.');
  return parts.join(' ');
}

/** 카드 제목에서 괄호 속 연도·설명을 뺌: "긴급 조치 1호 (일부, 1974)" → "긴급 조치 1호" */
function cardShort_(t) {
  return String(t).replace(/\s*\([^)]*\)\s*/g, ' ').replace(/[—–-].*$/, '').trim();
}

/** 학생 의견에서 "…(이)라고/다고 생각한다" 앞의 주장만 뽑음. 못 찾으면 빈칸 */
function claimOf_(text) {
  const sents = String(text).replace(/\s+/g, ' ').split(/(?<=[.!?])\s+/).filter(x => x.indexOf('생각') >= 0);
  if (!sents.length) return '';
  let s = sents[sents.length - 1];
  const i = s.search(/(이?라고|다고|고)\s*생각/);
  if (i < 0) return '';
  let pre = s.slice(0, i);
  const m = s.slice(i).match(/^(이?라고|다고|고)/)[1];
  if (m === '다고') pre += '다';
  else if (m === '라고' || m === '이라고') pre += (m === '이라고' ? '이다' : '다');
  pre = pre.replace(/^(그래서|따라서|그러므로|결국|즉)\s*,?\s*/, '').replace(/^(나는|저는|내 생각에는)\s*/, '').trim();
  pre = pre.replace(/[.,\s]+$/, '');
  if (!/다$/.test(pre) || pre.length < 8 || pre.length > 90) return '';
  return pre;
}

/** 생각기록에서 취재 수첩 입장 줄만 추려 짧게 */
function stanceOf_(log) {
  return String(log).split('\n').filter(l => l.indexOf(' / 근거:') >= 0).map(l => {
    const m = l.match(/→\s*(.+?)\s*\/\s*근거:\s*(.+?)\s*\//);
    return m ? `${m[1]} (근거: ${cardShort_(m[2])})` : l;
  }).join('\n');
}

/** 받침에 따라 와/과, 이/가 등 고르기 */
function josa_(word, a, b) {
  const w = String(word).replace(/[^가-힣0-9A-Za-z]/g, '');
  const c = w.charCodeAt(w.length - 1);
  if (c >= 0xAC00 && c <= 0xD7A3) return ((c - 0xAC00) % 28) ? b : a;
  if (/[0-9]$/.test(w)) return /[013678]$/.test(w) ? b : a;
  return a;
}

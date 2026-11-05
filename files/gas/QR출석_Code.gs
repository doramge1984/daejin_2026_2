/**
 * QR 출석·출구 티켓 — 서버 코드(Code.gs)
 * 시트 탭: '설정' (B1 오늘의 단어 | B2 교수 비밀번호)
 *          '출석' (A 출석 시각 | B 날짜 | C 학번 | D 이름)
 * 학생: 웹 앱 주소(QR) → 학번·이름·오늘의 단어 → 출석.  같은 학번은 하루 한 번만.
 * 교수: 주소 끝에 ?admin=비밀번호 를 붙이면 오늘 출석 인원이 10초마다 바뀌어 보여요.
 * 배포: 배포 → 새 배포 → 웹 앱 → 실행 사용자 '나' · 액세스 '모든 사용자'
 */
const SET_SHEET = '설정';
const LOG_SHEET = '출석';
const PAGE_TITLE = '오늘 수업 출석';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(PAGE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function settings_() {
  const v = SpreadsheetApp.getActive().getSheetByName(SET_SHEET).getRange('B1:B2').getDisplayValues();
  return { word: String(v[0][0]).trim(), pw: String(v[1][0]).trim() };
}
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

// 학생 출석
function checkIn(form) {
  const sid = String(form.sid || '').trim();
  const name = String(form.name || '').trim();
  const word = String(form.word || '').trim();
  if (!sid || !name || !word) return { ok: false, msg: '학번, 이름, 오늘의 단어를 모두 넣어 주세요.' };
  const set = settings_();
  if (word.replace(/\s/g, '') !== set.word.replace(/\s/g, '')) return { ok: false, msg: '오늘의 단어가 맞지 않아요. 화면의 단어를 다시 확인해 주세요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
    const day = today_();
    const rows = sh.getDataRange().getDisplayValues().slice(1);
    if (rows.some(function (r) { return String(r[1]).trim() === day && String(r[2]).trim() === sid; })) {
      return { ok: false, msg: '오늘은 이미 출석했어요.' };
    }
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
    sh.appendRow(["'" + now, "'" + day, "'" + sid, name]);
    return { ok: true, msg: name + '님, ' + now.slice(0, 5) + ' 출석됐어요!' };
  } finally {
    lock.releaseLock();
  }
}

// 교수 화면: 비밀번호가 맞을 때만 오늘 인원과 최근 출석을 돌려줍니다.
function getToday(pw) {
  if (String(pw || '').trim() !== settings_().pw) return { ok: false, msg: '비밀번호가 맞지 않아요.' };
  const day = today_();
  const rows = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET).getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return String(r[1]).trim() === day; });
  const recent = rows.slice(-8).reverse().map(function (r) { return { time: r[0], sid: r[2], name: r[3] }; });
  return { ok: true, day: day, count: rows.length, recent: recent };
}

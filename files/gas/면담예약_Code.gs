/**
 * 교수 면담 시간 예약 보드 — 서버 코드(Code.gs)
 * 시트 탭: '시간' (A 날짜 | B 시각)          예: 11/10(화) | 14:00
 *          '예약' (A 예약 시각 | B 날짜 | C 시각 | D 학번 | E 이름 | F 면담 내용)
 * 누가 고른 시간은 다른 학생 화면에서 사라져요. 같은 학번은 한 번만 예약할 수 있어요.
 * 배포: 배포 → 새 배포 → 웹 앱 → 실행 사용자 '나' · 액세스 '모든 사용자'
 */
const SLOT_SHEET = '시간';
const BOOK_SHEET = '예약';
const PAGE_TITLE = '교수 면담 시간 예약';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(PAGE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 아직 아무도 고르지 않은 시간만 날짜별로 묶어 돌려줍니다.
function getSlots() {
  const ss = SpreadsheetApp.getActive();
  const slots = ss.getSheetByName(SLOT_SHEET).getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); });
  const taken = {};
  ss.getSheetByName(BOOK_SHEET).getDataRange().getDisplayValues().slice(1).forEach(function (r) {
    taken[String(r[1]).trim() + '|' + String(r[2]).trim()] = true;
  });
  const days = [];
  const byDay = {};
  let free = 0;
  slots.forEach(function (r) {
    const d = String(r[0]).trim(), t = String(r[1]).trim();
    if (!byDay[d]) { byDay[d] = { date: d, times: [] }; days.push(byDay[d]); }
    if (!taken[d + '|' + t]) { byDay[d].times.push(t); free++; }
  });
  return { title: PAGE_TITLE, days: days, free: free, total: slots.length };
}

// 예약을 받습니다. 동시에 눌러도 한 칸에 두 명이 들어가지 않게 잠금을 겁니다.
function book(form) {
  const sid = String(form.sid || '').trim();
  const name = String(form.name || '').trim();
  const date = String(form.date || '').trim();
  const time = String(form.time || '').trim();
  const memo = String(form.memo || '').trim().slice(0, 200);
  if (!sid || !name || !date || !time) return { ok: false, msg: '시간을 고르고 학번과 이름을 넣어 주세요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(BOOK_SHEET);
    const rows = sh.getDataRange().getDisplayValues().slice(1);
    const mine = rows.filter(function (r) { return String(r[3]).trim() === sid; })[0];
    if (mine) return { ok: false, msg: '이미 ' + mine[1] + ' ' + mine[2] + '에 예약했어요. 바꾸려면 교수님께 말씀해 주세요.', board: getSlots() };
    if (rows.some(function (r) { return String(r[1]).trim() === date && String(r[2]).trim() === time; })) {
      return { ok: false, msg: '방금 다른 학생이 예약했어요. 다른 시간을 골라 주세요.', board: getSlots() };
    }
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    sh.appendRow([now, "'" + date, "'" + time, "'" + sid, name, memo]);
    return { ok: true, msg: date + ' ' + time + ' 면담이 예약됐어요!', board: getSlots() };
  } finally {
    lock.releaseLock();
  }
}

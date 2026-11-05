/**
 * 실습실 장비 대여·반납 대장 — 서버 코드(Code.gs)
 * 시트 탭: '장비' (A 장비 | B 보유 수량)
 *          '대여' (A 대여 시각 | B 장비 | C 학번 | D 이름 | E 반납 예정일 | F 반납 시각)
 * F(반납 시각)가 비어 있는 줄 = 지금 빌려 간 장비.  남은 수량 = 보유 수량 − 빌려 간 수.
 * 배포: 배포 → 새 배포 → 웹 앱 → 실행 사용자 '나' · 액세스 '모든 사용자'
 */
const ITEM_SHEET = '장비';
const RENT_SHEET = '대여';
const PAGE_TITLE = '실습실 장비 대여·반납';
const DAYS = 7; // 대여 기간(일)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(PAGE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function fmt_(d, p) { return Utilities.formatDate(d, Session.getScriptTimeZone(), p); }
function mask_(name) { name = String(name).replace(/\(.*\)/, '').trim(); return name.length <= 1 ? name : name.charAt(0) + '*' + (name.length > 2 ? name.slice(-1) : ''); }

// 장비별 남은 수량과 지금 빌려 간 목록(이름은 가려서)을 돌려줍니다.
function getBoard() {
  const ss = SpreadsheetApp.getActive();
  const items = ss.getSheetByName(ITEM_SHEET).getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return String(r[0]).trim(); });
  const out = ss.getSheetByName(RENT_SHEET).getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return String(r[1]).trim() && !String(r[5]).trim(); });
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  return {
    title: PAGE_TITLE, today: today,
    items: items.map(function (r) {
      const n = out.filter(function (o) { return String(o[1]).trim() === String(r[0]).trim(); }).length;
      return { name: String(r[0]).trim(), total: Number(r[1]) || 0, out: n };
    }),
    out: out.map(function (o) {
      return { item: o[1], who: mask_(o[3]), due: o[4], late: String(o[4]) < today };
    })
  };
}

// 대여: 남은 수량이 있을 때만, 같은 학번은 같은 장비를 한 개만.
function rent(form) {
  const sid = String(form.sid || '').trim(), name = String(form.name || '').trim(), item = String(form.item || '').trim();
  if (!sid || !name || !item) return { ok: false, msg: '장비를 고르고 학번과 이름을 넣어 주세요.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(RENT_SHEET);
    const rows = sh.getDataRange().getDisplayValues().slice(1);
    if (rows.some(function (r) { return String(r[2]).trim() === sid && String(r[1]).trim() === item && !String(r[5]).trim(); })) {
      return { ok: false, msg: '이미 ' + item + '을(를) 빌려 갔어요. 반납한 뒤 다시 빌려 주세요.', board: getBoard() };
    }
    const it = getBoard().items.filter(function (x) { return x.name === item; })[0];
    if (!it) return { ok: false, msg: '장비를 찾을 수 없어요. 화면을 새로고침해 주세요.', board: getBoard() };
    if (it.out >= it.total) return { ok: false, msg: '방금 다 빌려 갔어요. 반납되면 다시 빌릴 수 있어요.', board: getBoard() };
    const now = new Date();
    const due = fmt_(new Date(now.getTime() + DAYS * 86400000), 'yyyy-MM-dd');
    sh.appendRow(["'" + fmt_(now, 'yyyy-MM-dd HH:mm'), item, "'" + sid, name, "'" + due, '']);
    return { ok: true, msg: item + ' 대여 완료! 반납 예정일은 ' + due + '이에요.', board: getBoard() };
  } finally {
    lock.releaseLock();
  }
}

// 반납: 학번과 장비가 맞는, 아직 반납 안 한 줄에 반납 시각을 적어요.
function giveBack(form) {
  const sid = String(form.sid || '').trim(), item = String(form.item || '').trim();
  if (!sid || !item) return { ok: false, msg: '장비를 고르고 학번을 넣어 주세요.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(RENT_SHEET);
    const rows = sh.getDataRange().getDisplayValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][2]).trim() === sid && String(rows[i][1]).trim() === item && !String(rows[i][5]).trim()) {
        sh.getRange(i + 1, 6).setValue("'" + fmt_(new Date(), 'yyyy-MM-dd HH:mm'));
        return { ok: true, msg: item + ' 반납됐어요. 고마워요!', board: getBoard() };
      }
    }
    return { ok: false, msg: '이 학번으로 빌려 간 ' + item + '이(가) 없어요.', board: getBoard() };
  } finally {
    lock.releaseLock();
  }
}

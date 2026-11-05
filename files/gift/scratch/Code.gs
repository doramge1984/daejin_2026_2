/**
 * 🎟️ 행운 복권 — 서버 코드(Code.gs)
 * 시트 탭: '선물' (A 선물 이름 | B 개수)
 *          '복권' (A 시각 | B 이름 | C 학과 | D 결과)
 * 참가: 웹 앱 주소(QR) → 이름 → 손가락으로 긁기. 한 사람당 한 장.
 * 결과는 서버에서 정해요. 선물 개수보다 많이 당첨되지 않아요(동시에 긁어도).
 * 강사: 주소 끝에 ?admin=비밀번호 → 남은 선물과 당첨자 현황판
 */
const ADMIN_PW = '1106';      // ← 강사 화면 비밀번호 (꼭 바꾸세요)
const EXPECTED = 30;          // ← 예상 참가 인원. 선물이 끝까지 고르게 나가도록 당첨 확률을 맞춰요
const TITLE = '대진대 교수 연수 · 행운 복권';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no');
}

function sh_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function now_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'); }
function state_() {
  const plays = sh_('복권').getDataRange().getDisplayValues().slice(1).filter(function (r) { return String(r[1]).trim(); });
  const gifts = sh_('선물').getDataRange().getDisplayValues().slice(1).filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      const used = plays.filter(function (p) { return p[3] === r[0]; }).length;
      return { name: r[0], total: Number(r[1]) || 0, left: Math.max(0, (Number(r[1]) || 0) - used) };
    });
  return { plays: plays, gifts: gifts };
}

// 복권 한 장 (같은 이름+학과는 처음 결과를 그대로 다시 보여 줘요)
function play(form) {
  const name = String(form.name || '').trim().slice(0, 20);
  const dept = String(form.dept || '').trim().slice(0, 30);
  if (!name) return { ok: false, msg: '이름을 넣어 주세요.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const st = state_();
    const mine = st.plays.filter(function (p) { return p[1] === name && p[2] === dept; })[0];
    if (mine) return { ok: true, again: true, win: mine[3] !== '꽝', gift: mine[3] };
    const left = st.gifts.reduce(function (a, g) { return a + g.left; }, 0);
    const remainPeople = Math.max(1, EXPECTED - st.plays.length);
    let result = '꽝';
    if (left > 0 && Math.random() < Math.min(1, left / remainPeople)) {
      let r = Math.random() * left;
      for (let i = 0; i < st.gifts.length; i++) { r -= st.gifts[i].left; if (r < 0) { result = st.gifts[i].name; break; } }
    }
    sh_('복권').appendRow(["'" + now_(), name, dept, result]);
    return { ok: true, win: result !== '꽝', gift: result };
  } finally { lock.releaseLock(); }
}

// 강사 현황판
function getBoard(pw) {
  if (String(pw) !== ADMIN_PW) return { ok: false, msg: '비밀번호가 맞지 않아요.' };
  const st = state_();
  return {
    ok: true, title: TITLE, url: ScriptApp.getService().getUrl(), played: st.plays.length, expected: EXPECTED, gifts: st.gifts,
    winners: st.plays.filter(function (p) { return p[3] !== '꽝'; }).map(function (p) { return { time: p[0], name: p[1], dept: p[2], gift: p[3] }; })
  };
}

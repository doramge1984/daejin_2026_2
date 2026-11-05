/**
 * 🎁 행운의 선물 추첨 — 서버 코드(Code.gs)
 * 시트 탭: '선물' (A 선물 이름 | B 개수)
 *          '참가자' (A 등록 시각 | B 이름 | C 학과)
 *          '당첨' (A 당첨 시각 | B 선물 | C 이름 | D 학과)
 * 참가: 웹 앱 주소(QR)로 들어와 이름 등록
 * 무대: 주소 끝에 ?admin=비밀번호  → 큰 화면에서 추첨 (같은 사람은 한 번만 당첨)
 */
const ADMIN_PW = '1106';            // ← 강사 화면 비밀번호 (꼭 바꾸세요)
const TITLE = '대진대 교수 연수 · 행운의 선물 추첨';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sh_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function now_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'); }
function rows_(name) { return sh_(name).getDataRange().getDisplayValues().slice(1).filter(function (r) { return String(r[1] || r[0]).trim(); }); }

// 참가 등록 (같은 이름+학과는 한 번만)
function register(form) {
  const name = String(form.name || '').trim().slice(0, 20);
  const dept = String(form.dept || '').trim().slice(0, 30);
  if (!name) return { ok: false, msg: '이름을 넣어 주세요.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const list = rows_('참가자');
    if (list.some(function (r) { return r[1] === name && r[2] === dept; })) {
      return { ok: true, again: true, msg: name + ' 교수님은 이미 등록됐어요. 행운을 빌어요! 🍀', count: list.length };
    }
    sh_('참가자').appendRow(["'" + now_(), name, dept]);
    return { ok: true, msg: name + ' 교수님, 등록 완료! 행운을 빌어요 🍀', count: list.length + 1 };
  } finally { lock.releaseLock(); }
}

// 무대 화면 상태
function getStage(pw) {
  if (String(pw) !== ADMIN_PW) return { ok: false, msg: '비밀번호가 맞지 않아요.' };
  const people = rows_('참가자').map(function (r) { return { name: r[1], dept: r[2] }; });
  const wins = rows_('당첨').map(function (r) { return { time: r[0], gift: r[1], name: r[2], dept: r[3] }; });
  const gifts = sh_('선물').getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      const used = wins.filter(function (w) { return w.gift === r[0]; }).length;
      return { name: r[0], total: Number(r[1]) || 0, left: Math.max(0, (Number(r[1]) || 0) - used) };
    });
  return { ok: true, title: TITLE, url: ScriptApp.getService().getUrl(), people: people, wins: wins, gifts: gifts };
}

// 추첨: 아직 당첨 안 된 사람 중에서 무작위로 한 명
function draw(pw, giftName) {
  if (String(pw) !== ADMIN_PW) return { ok: false, msg: '비밀번호가 맞지 않아요.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const st = getStage(pw);
    const g = st.gifts.filter(function (x) { return x.name === giftName; })[0];
    if (!g) return { ok: false, msg: '선물을 골라 주세요.' };
    if (g.left <= 0) return { ok: false, msg: g.name + '은(는) 모두 나갔어요.' };
    const won = {};
    st.wins.forEach(function (w) { won[w.name + '|' + w.dept] = true; });
    const pool = st.people.filter(function (p) { return !won[p.name + '|' + p.dept]; });
    if (!pool.length) return { ok: false, msg: '추첨할 사람이 없어요. 모두 당첨됐어요!' };
    const pick = pool[Math.floor(Math.random() * pool.length)];
    sh_('당첨').appendRow(["'" + now_(), g.name, pick.name, pick.dept]);
    return { ok: true, winner: pick, gift: g.name, stage: getStage(pw) };
  } finally { lock.releaseLock(); }
}

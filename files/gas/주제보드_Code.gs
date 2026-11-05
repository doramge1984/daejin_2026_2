/**
 * 팀 프로젝트 주제 선착순 보드 — 서버 코드(Code.gs)
 * 시트 탭: '주제' (A 주제번호 | B 주제 | C 설명 | D 정원(팀 수))
 *          '신청' (A 신청 시각 | B 팀 이름 | C 대표 학번 | D 대표 이름 | E 주제번호 | F 주제)
 * 배포: 배포 → 새 배포 → 웹 앱 → 실행 사용자 '나' · 액세스 '모든 사용자'
 */
const TOPIC_SHEET = '주제';
const APPLY_SHEET = '신청';
const BOARD_TITLE = '팀 프로젝트 주제 신청 보드';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(BOARD_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 주제별 정원과 지금까지 신청한 팀 목록을 돌려줍니다.
function getBoard() {
  const ss = SpreadsheetApp.getActive();
  const topics = ss.getSheetByName(TOPIC_SHEET).getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return String(r[0]).trim() !== ''; });
  const applies = ss.getSheetByName(APPLY_SHEET).getDataRange().getDisplayValues().slice(1);
  const list = topics.map(function (r) {
    const id = String(r[0]).trim();
    const teams = applies.filter(function (a) { return String(a[4]).trim() === id; })
      .map(function (a) { return String(a[1]); });
    return { id: id, title: String(r[1]), desc: String(r[2]), cap: Number(r[3]) || 1, teams: teams };
  });
  return { title: BOARD_TITLE, topics: list };
}

// 신청을 받습니다. 동시에 눌러도 정원을 넘지 않게 잠금을 겁니다.
function apply(form) {
  const team = String(form.team || '').trim();
  const sid = String(form.sid || '').trim();
  const name = String(form.name || '').trim();
  const topicId = String(form.topicId || '').trim();
  if (!team || !sid || !name || !topicId) return { ok: false, msg: '팀 이름, 대표 학번, 대표 이름, 주제를 모두 넣어 주세요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(APPLY_SHEET);
    const rows = sh.getDataRange().getDisplayValues().slice(1);
    if (rows.some(function (r) { return String(r[2]).trim() === sid; })) {
      return { ok: false, msg: '이 학번으로 이미 신청했어요. 바꾸려면 교수님께 말씀해 주세요.', board: getBoard() };
    }
    const board = getBoard();
    const t = board.topics.filter(function (x) { return x.id === topicId; })[0];
    if (!t) return { ok: false, msg: '주제를 찾을 수 없어요. 화면을 새로고침해 주세요.', board: board };
    if (t.teams.length >= t.cap) return { ok: false, msg: '방금 마감됐어요. 다른 주제를 골라 주세요.', board: board };
    sh.appendRow([new Date(), team, "'" + sid, name, t.id, t.title]);
    return { ok: true, msg: '"' + t.title + '" 주제로 신청됐어요!', board: getBoard() };
  } finally {
    lock.releaseLock();
  }
}

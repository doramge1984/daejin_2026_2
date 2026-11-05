/**
 * 내 점수·피드백 조회 페이지 — 서버 코드(Code.gs)
 * 시트 탭: '점수' (A 학번 | B 조회 코드 | C 이름 | D~G 기준별 점수 | H 총점 | I 피드백)
 *   - D~G·H 열의 제목(1행)이 그대로 화면에 나와요. 기준 이름을 바꾸려면 1행만 고치세요.
 * 학번과 조회 코드가 둘 다 맞는 한 줄만 돌려줍니다. 다른 학생 정보는 화면으로 나가지 않아요.
 * 배포: 배포 → 새 배포 → 웹 앱 → 실행 사용자 '나' · 액세스 '모든 사용자'
 */
const SCORE_SHEET = '점수';
const PAGE_TITLE = '내 점수·피드백 확인';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(PAGE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function lookup(sid, code) {
  sid = String(sid || '').trim();
  code = String(code || '').trim();
  if (!sid || !code) return { ok: false, msg: '학번과 조회 코드를 모두 넣어 주세요.' };
  Utilities.sleep(500); // 무작위로 계속 넣어 보는 것을 늦추기 위한 짧은 대기
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SCORE_SHEET) || ss.getSheets()[0]; // 탭 이름이 달라도 첫 탭을 읽어요
  const v = sh.getDataRange().getDisplayValues();
  const head = v[0];
  const row = v.slice(1).filter(function (r) {
    return String(r[0]).trim() === sid && String(r[1]).trim() === code;
  })[0];
  if (!row) return { ok: false, msg: '학번 또는 조회 코드가 맞지 않아요. 다시 확인해 주세요.' };
  const items = [];
  for (let c = 3; c <= 6; c++) items.push({ name: head[c], score: row[c] });
  return { ok: true, title: PAGE_TITLE, name: row[2], items: items, totalName: head[7], total: row[7], feedback: row[8] };
}

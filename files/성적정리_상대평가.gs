/**
 * 성적 정리 · 상대평가 버전 (가상 데이터 연습용)
 * 시트 준비
 *  - '성적' 탭: A 학번 | B 이름 | C 출석률 | D 중간 | E 과제 | F 기말 (1행은 제목)
 *  - '기준' 탭: 반영 비율, 학점별 누적 비율, 출석 F 기준  ← 모두 [학교 기준 확인 필요] · 예시 값
 * 하는 일
 *  ① 총점 계산 → ② 석차(동점은 같은 석차) → ③ 누적 비율로 A+~D 배정, 출석 미달은 F
 *  ④ 학점 경계 ±1점 학생 노란색 + '경계선' 탭 목록 → ⑤ 경계에 걸친 동점자 표시
 *  ⑥ '요약' 탭: 학점별 인원·비율·기준 비교 + 막대그래프
 * 메일 발송·외부 접근은 없습니다. 학점은 초안이며, 확정은 교수님이 직접 합니다.
 */
function 상대평가정리() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('성적') || ss.getSheets()[0];
  const cfg = ss.getSheetByName('기준');
  if (!cfg) { ss.toast("'기준' 탭이 없어요. 사본 시트를 그대로 쓰세요."); return; }
  const last = sh.getLastRow();
  if (last < 2) { ss.toast('성적 데이터가 없어요.'); return; }

  const c = readConfig_(cfg);
  const raw = sh.getRange(2, 1, last - 1, 6).getValues();
  const res = computeGrades(raw, c);

  // 결과 쓰기 (G 총점 | H 석차 | I 학점(초안) | J 확인 메모) 후 총점순 정렬
  sh.getRange('G1:J1').setValues([['총점', '석차', '학점(초안)', '확인 메모']]).setFontWeight('bold');
  const out = res.rows.map(function (r) { return r.src.concat([r.total, r.rank, r.grade, r.memo]); });
  sh.getRange(2, 1, out.length, 10).setValues(out).setBackground(null);
  res.rows.forEach(function (r, i) {
    if (r.border) sh.getRange(i + 2, 1, 1, 10).setBackground('#fff2a8');
    if (r.tie) sh.getRange(i + 2, 10).setFontColor('#b45309').setFontWeight('bold');
    if (r.grade === 'F') sh.getRange(i + 2, 9).setFontColor('#c0392b').setFontWeight('bold');
  });

  writeSummary_(ss, res, c);
  writeBorder_(ss, res);
  ss.toast('정리 완료! 노란 줄은 학점 경계 ±1점 학생이에요. 요약·경계선 탭을 확인하세요.', '상대평가 정리', 8);
}

/* ---------- 기준 읽기 ---------- */
function readConfig_(cfg) {
  const v = cfg.getDataRange().getValues();
  const get = function (label) {
    for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === label) return Number(v[i][1]);
    return NaN;
  };
  const grades = ['A+', 'A0', 'B+', 'B0', 'C+', 'C0', 'D'];
  return {
    w: [get('출석 반영(%)'), get('중간 반영(%)'), get('과제 반영(%)'), get('기말 반영(%)')].map(function (x) { return x / 100; }),
    grades: grades,
    cum: grades.map(function (g) { return get(g + ' 누적(%)') / 100; }),
    attMin: get('출석률 F 기준(미만)'),
    band: get('경계 범위(±점)') || 1
  };
}

/* ---------- 핵심 계산 (시트와 상관없는 순수 함수) ---------- */
function computeGrades(raw, c) {
  const n = raw.length;
  let rows = raw.map(function (r) {
    const t = Math.round((Number(r[2]) * c.w[0] + Number(r[3]) * c.w[1] + Number(r[4]) * c.w[2] + Number(r[5]) * c.w[3]) * 10) / 10;
    return { src: r.slice(0, 6), total: t, att: Number(r[2]), memo: '', border: false, tie: false };
  });
  rows.sort(function (a, b) { return b.total - a.total; });
  // 석차: 동점은 같은 석차
  rows.forEach(function (r, i) { r.rank = (i > 0 && r.total === rows[i - 1].total) ? rows[i - 1].rank : i + 1; });

  // 출석 미달은 F, 나머지(m명)로 상대평가
  const pool = rows.filter(function (r) { return !(r.att < c.attMin); });
  rows.forEach(function (r) { if (r.att < c.attMin) { r.grade = 'F'; r.memo = '출석 미달 → F'; } });
  const m = pool.length;
  const quota = c.cum.map(function (p) { return Math.floor(p * m + 1e-9); }); // 누적 허용 인원
  let gi = 0;
  pool.forEach(function (r, i) {
    while (gi < c.grades.length - 1 && i + 1 > quota[gi]) gi++;
    r.grade = c.grades[gi];
  });
  // 동점자: 같은 점수는 같은 학점(윗 학점)으로 맞추고, 경계에 걸쳤으면 표시
  for (let i = 1; i < m; i++) {
    if (pool[i].total === pool[i - 1].total && pool[i].grade !== pool[i - 1].grade) {
      const g = pool[i - 1].grade, j0 = i;
      for (let j = i; j < m && pool[j].total === pool[i - 1].total; j++) { pool[j].grade = g; }
      for (let j = j0 - 1; j >= 0 && pool[j].total === pool[j0].total; j--) pool[j].tie = true;
      for (let j = j0; j < m && pool[j].total === pool[j0 - 1].total; j++) pool[j].tie = true;
    }
  }
  pool.forEach(function (r) { if (r.tie) r.memo = '동점자 · 학점 경계에 걸침 (기준 비율 초과 여부 확인)'; });

  // 경계 ±band점: 위 학점 마지막 학생과 아래 학점 첫 학생의 가운데 점수 기준
  const borders = [];
  for (let i = 1; i < m; i++) {
    if (pool[i].grade !== pool[i - 1].grade) {
      const mid = (pool[i - 1].total + pool[i].total) / 2;
      const label = pool[i - 1].grade + ' ↔ ' + pool[i].grade;
      borders.push({ label: label, upper: pool[i - 1].total, lower: pool[i].total });
      pool.forEach(function (r) {
        if (Math.abs(r.total - mid) <= c.band) {
          r.border = true;
          if (r.memo.indexOf(label) < 0) r.memo = (r.memo ? r.memo + ' / ' : '') + label + ' 경계 ±' + c.band + '점';
        }
      });
    }
  }
  // 학점별 인원
  const all = c.grades.concat(['F']);
  const count = all.map(function (g) { return rows.filter(function (r) { return r.grade === g; }).length; });
  return { rows: rows, n: n, m: m, grades: all, count: count, quota: quota, borders: borders };
}

/* ---------- 요약 탭 + 그래프 ---------- */
function writeSummary_(ss, res, c) {
  let sum = ss.getSheetByName('요약');
  if (sum) ss.deleteSheet(sum);
  sum = ss.insertSheet('요약');
  const head = [['학점', '인원', '비율(%)', '실제 누적(%)', '기준 누적(%)', '확인']];
  let acc = 0;
  const body = res.grades.map(function (g, i) {
    const k = res.count[i]; acc += (g === 'F') ? 0 : k;
    const pct = Math.round(k / res.n * 1000) / 10;
    if (g === 'F') return [g, k, pct, '', '', '출석 미달 등'];
    const real = Math.round(acc / res.m * 1000) / 10, std = Math.round(c.cum[i] * 1000) / 10;
    return [g, k, pct, real, std, real > std ? '⚠ 기준 초과 (동점자 확인)' : '✓'];
  });
  sum.getRange(1, 1, 1, 6).setValues(head).setFontWeight('bold').setBackground('#eef3fb');
  sum.getRange(2, 1, body.length, 6).setValues(body);
  sum.getRange(body.length + 3, 1, 2, 2).setValues([['수강 인원', res.n], ['상대평가 대상(출석 미달 제외)', res.m]]);
  sum.getRange(body.length + 6, 1).setValue('※ 기준 비율은 예시 값이에요. 소속 학교 학칙·학사 규정을 확인해 기준 탭을 고치세요. [학교 기준 확인 필요]').setFontColor('#a5372f');
  sum.autoResizeColumns(1, 6);
  const chart = sum.newChart().asColumnChart()
    .addRange(sum.getRange(1, 1, body.length + 1, 2)).setNumHeaders(1)
    .setOption('title', '학점별 인원').setOption('legend', { position: 'none' })
    .setPosition(2, 8, 0, 0).build();
  sum.insertChart(chart);
}

/* ---------- 경계선 탭 ---------- */
function writeBorder_(ss, res) {
  let b = ss.getSheetByName('경계선');
  if (b) ss.deleteSheet(b);
  b = ss.insertSheet('경계선');
  const list = res.rows.filter(function (r) { return r.border || r.tie; })
    .map(function (r) { return [r.src[0], r.src[1], r.total, r.rank, r.grade, r.memo]; });
  b.getRange(1, 1, 1, 6).setValues([['학번', '이름', '총점', '석차', '학점(초안)', '왜 다시 볼까요?']]).setFontWeight('bold').setBackground('#fff2a8');
  if (list.length) b.getRange(2, 1, list.length, 6).setValues(list);
  b.getRange(list.length + 3, 1).setValue('이 학생들은 채점 실수 1~2점으로 학점이 바뀔 수 있어요. 답안·과제를 한 번 더 확인하세요.').setFontColor('#5b6475');
  b.autoResizeColumns(1, 6);
}

/**
 * 가상 성적 시트 한 방 정리
 * 전제: 첫 번째 시트에 A~F열 = 학번, 이름, 출석률, 중간, 과제, 기말 (1행은 제목)
 * 하는 일: 총점·학점(초안) 계산 → 출석 80 미만/총점 60 미만 표시 → 총점순 정렬 → '요약' 시트와 그래프 생성
 * 이미 손으로 G·H열을 만들었어도 덮어씁니다. 메일 발송·외부 접근은 없습니다.
 */
function 성적정리() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets()[0];
  const last = sh.getLastRow();
  if (last < 2) { ss.toast('데이터가 없습니다. A1에 표를 먼저 붙여넣으세요.'); return; }
  const n = last - 1;

  // 1) 총점·학점(초안) — 출석 10%, 중간 30%, 과제 20%, 기말 40%
  sh.getRange('G1:H1').setValues([['총점', '학점(초안)']]);
  const f = [];
  for (let r = 2; r <= last; r++) {
    f.push([
      '=ROUND(C' + r + '*0.1+D' + r + '*0.3+E' + r + '*0.2+F' + r + '*0.4,1)',
      '=IF(G' + r + '>=90,"A",IF(G' + r + '>=80,"B",IF(G' + r + '>=70,"C","D")))'
    ]);
  }
  sh.getRange(2, 7, n, 2).setFormulas(f);

  // 2) 서식: 제목 줄, 출석 80 미만(빨강), 총점 60 미만(노랑)
  sh.getRange('A1:H1').setFontWeight('bold').setBackground('#23303f').setFontColor('#ffffff');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(80)
      .setBackground('#f4c7c3').setRanges([sh.getRange(2, 3, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(60)
      .setBackground('#fce8b2').setRanges([sh.getRange(2, 7, n, 1)]).build()
  ];
  sh.setConditionalFormatRules(rules);

  // 3) 총점 높은 순 정렬, 열 너비, 1행 고정
  sh.getRange(2, 1, n, 8).sort({ column: 7, ascending: false });
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 8);

  // 4) 요약 시트 (있으면 지우고 다시 만듦)
  let sum = ss.getSheetByName('요약');
  if (sum) { sum.clear(); sum.getCharts().forEach(function (c) { sum.removeChart(c); }); }
  else { sum = ss.insertSheet('요약'); }
  const name = sh.getName().replace(/'/g, "''");
  sum.getRange('A1:B1').setValues([['학점', '인원']]);
  const grades = ['A', 'B', 'C', 'D'];
  sum.getRange(2, 1, 4, 2).setFormulas(grades.map(function (g, i) {
    return [g, "=COUNTIF('" + name + "'!H2:H" + last + ",A" + (i + 2) + ")"];
  }));
  sum.getRange('D1:E1').setValues([['전체 평균 총점', '출석 80 미만 인원']]);
  sum.getRange('D2:E2').setFormulas([[
    "=ROUND(AVERAGE('" + name + "'!G2:G" + last + "),1)",
    "=COUNTIF('" + name + "'!C2:C" + last + ",\"<80\")"
  ]]);
  sum.getRange('A1:E1').setFontWeight('bold');
  sum.insertChart(sum.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sum.getRange('A1:B5')).setPosition(4, 1, 0, 0)
    .setOption('title', '학점별 인원(초안)').build());

  ss.setActiveSheet(sum);
  ss.toast('정리 완료! 학점 기준과 확정은 교수님이 직접 확인하세요.');
}

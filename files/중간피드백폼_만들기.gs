/**
 * 중간 강의 피드백 · 구글 폼 + 응답 시트 자동 만들기
 * 사용법: script.new → 이 코드를 모두 붙여 넣기 → 저장 → 위쪽에서 makeForm 고르고 ▶ 실행
 *         (처음 한 번 권한 허용) → 아래 '실행 로그'에 나온 주소를 열어 보세요.
 * 문항 이름은 오후반 7 라이브 대시보드와 같아요. 이 폼의 응답 시트로 대시보드를 바로 만들 수 있어요.
 */
function makeForm() {
  const form = FormApp.create('[과목명] 중간 강의 피드백 (익명 · 2분)');
  form.setDescription('남은 수업을 더 좋게 만들려고 해요. 이름·학번은 받지 않아요.');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setConfirmationMessage('응답 고마워요! 다음 수업에 반영할게요.');
  try { form.setRequireLogin(false); } catch (e) {} // 개인 계정에서는 그냥 넘어가요

  form.addMultipleChoiceItem().setTitle('학년')
    .setChoiceValues(['1학년', '2학년', '3학년', '4학년']).setRequired(true);
  form.addMultipleChoiceItem().setTitle('전공 구분')
    .setChoiceValues(['주전공', '복수·부전공', '교양으로 수강']).showOtherOption(true).setRequired(true);
  form.addScaleItem().setTitle('이해도')
    .setBounds(1, 5).setLabels('거의 모르겠어요', '잘 이해했어요').setRequired(true);
  form.addMultipleChoiceItem().setTitle('수업 속도')
    .setChoiceValues(['빠름', '적당', '느림']).setRequired(true);
  form.addMultipleChoiceItem().setTitle('가장 어려웠던 개념')
    .setChoiceValues(['[개념 1]', '[개념 2]', '[개념 3]', '[개념 4]', '[개념 5]']).showOtherOption(true).setRequired(true);
  form.addCheckboxItem().setTitle('도움이 된 활동 (여러 개 골라도 돼요)')
    .setChoiceValues(['강의 설명', '예제 풀이', '실습', '퀴즈', '조별 토론']).showOtherOption(true);
  form.addParagraphTextItem().setTitle('개선 제안 (선택)');

  const ss = SpreadsheetApp.create('[과목명] 중간 강의 피드백 · 응답');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('① 학생에게 줄 응답 주소: ' + form.getPublishedUrl());
  Logger.log('② 폼 수정 주소(교수용): ' + form.getEditUrl());
  Logger.log('③ 응답 시트 주소: ' + ss.getUrl());
}

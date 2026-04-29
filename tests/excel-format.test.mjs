import test from 'node:test';
import assert from 'node:assert/strict';
import {
  formatCareerForTemplate,
  formatEvaluationCareerForTemplate,
} from '../src/lib/profileExcelFormatter.js';

test('formatCareerForTemplate splits inline current and previous career entries', () => {
  const formatted = formatCareerForTemplate({
    affiliation: '세명대학교 보건안전학과 교수',
    careerDetails: '現 세명대학교 보건안전학과 교수 前 경희대학교 건설안전경영학과 교수 前 미국 캘리포니아 건설업체 CEO 前 ㈜ 피엠씨엠 부장 前 동아건설산업㈜ 대리 (면접, 프로젝트 등)',
  });

  assert.equal(formatted, [
    '現) 세명대학교 보건안전학과 교수',
    '前) 경희대학교 건설안전경영학과 교수',
    '前) 미국 캘리포니아 건설업체 CEO',
    '前) ㈜ 피엠씨엠 부장',
  ].join('\n'));
});

test('formatCareerForTemplate summarizes date-ranged careers to four rows', () => {
  const formatted = formatCareerForTemplate({
    affiliation: '시야 건축사사무소 대표',
    careerDetails: '(2019. 06. 11~현재) 시야 건축사사무소 대표 (건축계획 · 설계) (2012. 01. 01~2019. 02. 21) ㈜강호엔지니어링건축사사무소 이사 (건축계획 · 설계) (2007. 07. 09~2011. 12. 31) ㈜건축사사무소 뷰 이사 (건축계획 · 설계)',
  });

  assert.equal(formatted, [
    '現) (2019. 06. 11~현재) 시야 건축사사무소 대표 (건축계획 · 설계)',
    '前) (2012. 01. 01~2019. 02. 21) ㈜강호엔지니어링건축사사무소 이사 (건축계획 · 설계)',
    '前) (2007. 07. 09~2011. 12. 31) ㈜건축사사무소 뷰 이사 (건축계획 · 설계)',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate ignores parenthetical interview noise', () => {
  assert.equal(
    formatEvaluationCareerForTemplate({
      careerDetails: '現 세명대학교 교수 前 경희대학교 교수 (면접, 프로젝트 등)',
    }),
    ''
  );
});

test('formatEvaluationCareerForTemplate separates explicit document and interview blocks', () => {
  const formatted = formatEvaluationCareerForTemplate({
    careerDetails: '면접: KDB 산업은행, 금융감독원, 국민은행, 기업은행, 하나캐피탈 서류: 기술보증기금, 한국수출입은행, 국민은행, 한국남부발전, 서민금융진흥원',
  });

  assert.equal(formatted, [
    '[서류] 기술보증기금, 한국수출입은행, 국민은행, 한국남부발전 등',
    '[면접] KDB 산업은행, 금융감독원, 국민은행, 기업은행 등',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate handles bracket labels without colons', () => {
  const formatted = formatEvaluationCareerForTemplate({
    careerDetails: '[면접평가] KB 국민은행, 금융감독원, IBK 기업은행, 한국원자력공단, 한국무역보험공사 [서류평가] KB 국민은행, 한국산업기술진흥원, 한국언론진흥재단, 국립해양생물자원관, 한국수자원공사',
  });

  assert.equal(formatted, [
    '[서류] KB 국민은행, 한국산업기술진흥원, 한국언론진흥재단, 국립해양생물자원관 등',
    '[면접] KB 국민은행, 금융감독원, IBK 기업은행, 한국원자력공단 등',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate handles bare interview and document labels', () => {
  const formatted = formatEvaluationCareerForTemplate({
    careerDetails: '면접 한국수출입은행, 국립해양생물자원관, 제주문화예술재단, 서울물재생시설공단, 우체국물류지원단 서류 국가철도공단, KB 국민은행, 한국우편사업진흥원, 한국산림복지진흥원, 한국무역보험공사',
  });

  assert.equal(formatted, [
    '[서류] 국가철도공단, KB 국민은행, 한국우편사업진흥원, 한국산림복지진흥원 등',
    '[면접] 한국수출입은행, 국립해양생물자원관, 제주문화예술재단, 서울물재생시설공단 등',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate cleans category prefixes from evaluation lists', () => {
  const formatted = formatEvaluationCareerForTemplate({
    careerDetails: '면접전형 - [금융-한국은행, 금융감독원, KDB 산업은행, 한국수출입은행 서류전형 - [금융-한국은행, 금융결제원, KDB 산업은행산업은행, 한국무역보험공사',
  });

  assert.equal(formatted, [
    '[서류] 한국은행, 금융결제원, KDB 산업은행, 한국무역보험공사',
    '[면접] 한국은행, 금융감독원, KDB 산업은행, 한국수출입은행',
  ].join('\n'));
});

test('formatCareerForTemplate excludes explicit evaluation blocks from career summary', () => {
  const formatted = formatCareerForTemplate({
    affiliation: 'HR 임팩트 대표',
    careerDetails: 'HR 임팩트 대표: (2024~현재) ㈜ 임팩트그룹코리아 이사: 조직개발센터 (2024~현재) ㈜선연그룹 이사: 컨설팅 사업본부 (2021~2024) ㈜ 에스티유니타스 전략기획팀장: 전략 및 신사업 기획 (2015~2017) 면접관 경력 면접: KDB 산업은행, 금융감독원 서류: 기술보증기금',
  });

  assert.equal(formatted, [
    '現) HR 임팩트 대표: (2024~현재) ㈜ 임팩트그룹코리아 이사: 조직개발센터 (2024~현재) ㈜선연그룹 이사: 컨설팅 사업본부 (2021~2024) ㈜ 에스티유니타스 전략기획팀장: 전략 및 신사업 기획 (2015~2017)',
  ].join('\n'));
});

import test from 'node:test';
import assert from 'node:assert/strict';
import { EMPTY_VALUE, __testing, tagSuspiciousProfile } from '../src/lib/pptProfileParser.js';

test('extractPhone keeps labeled compact phone numbers', () => {
  assert.equal(__testing.extractPhone('01012345678'), '010-1234-5678');
});

test('extractPhone rejects unlabeled digit-only matches in full-document fallback', () => {
  const text = '생년월일 1973.07.07 경력 2010.01~2020.02 평가번호 01012345678';
  assert.equal(__testing.extractPhone(text, { requireContext: true }), '');
});

test('extractPhone requires nearby context for compact full-document matches', () => {
  const text = `연락처\n\n주요경력 ${'평가 '.repeat(12)} 평가번호 01012345678`;
  assert.equal(__testing.extractPhone(text, { requireContext: true }), '');
});

test('extractPhone accepts formatted phones in full-document fallback', () => {
  assert.equal(__testing.extractPhone('비상 연락 010-1234-5678', { requireContext: true }), '010-1234-5678');
});

test('extractGender ignores 남/여 choice templates', () => {
  assert.equal(__testing.extractGender('성별(남/여)'), '');
  assert.equal(__testing.extractGender('성별: 남'), '남');
});

test('extractGenderFromFileName reads delimited gender markers only', () => {
  assert.equal(__testing.extractGenderFromFileName('이상훈 프로필_HR_5월 7일_남.pptx'), '남');
  assert.equal(__testing.extractGenderFromFileName('이다솜_프로필_IT_5월 6일_여.pptx'), '여');
  assert.equal(__testing.extractGenderFromFileName('남기정 프로필.pptx'), '');
});

test('extractNameFromFileName prefers the leading person name', () => {
  assert.equal(__testing.extractNameFromFileName('김미애_프로필.pptx'), '김미애');
});

test('chooseAffiliation refuses generic document text without organization signal', () => {
  assert.equal(__testing.chooseAffiliation('전문분야 경영전략 주요경력 평가위원', ''), '');
});

test('sanitizeAffiliation removes contact and open-ended date tails', () => {
  assert.equal(__testing.sanitizeAffiliation('서브레인 대표 ( Tel ) 010- 2579'), '서브레인 대표');
  assert.equal(__testing.sanitizeAffiliation('메드소프트 대표 (2010 년 2 월 ~'), '메드소프트 대표');
  assert.equal(__testing.sanitizeAffiliation('이씨에스텔레콤 PM 팀 프로젝트매니저 파트장 ( 부장 , 2023 년 03 월 ~'), '이씨에스텔레콤 PM 팀 프로젝트매니저 파트장');
  assert.equal(__testing.sanitizeAffiliation('이루다에이치알 대표 : 채용컨설팅 ( 20 20 년 07 월 ~'), '이루다에이치알 대표 : 채용컨설팅');
  assert.equal(__testing.sanitizeAffiliation('엔지니어링그룹 에이원 / 전무 010'), '엔지니어링그룹 에이원 / 전무');
  assert.equal(__testing.sanitizeAffiliation('한국중소기업금융협회 본부장'), '한국중소기업금융협회 본부장');
});

test('splitEducationRecords removes combined mixed-degree fallback records', () => {
  const records = __testing.splitEducationRecords('한국외국어대학교 이란어 / 경영학 (부) 학사 한양사이버대학원 IT MBA 석사');

  assert.deepEqual(records, [
    '한국외국어대학교 이란어 경영학 (부) 학사',
    '한양사이버대학원 IT MBA 석사',
  ]);
  assert.equal(__testing.extractHighestEducation(records), '한양사이버대학원 IT MBA 석사');
});

test('extractHighestEducation prefers the highest degree over a longer foreign-school lower degree', () => {
  const records = __testing.splitEducationRecords(
    '연세대학원 정보미디어 석사 Macarthur Community College (ITTI) / Information Technology Diploma 학사'
  );

  assert.deepEqual(records, [
    '연세대학원 정보미디어 석사',
    'Macarthur Community College (ITTI) Information Technology Diploma 학사',
  ]);
  assert.equal(__testing.extractHighestEducation(records), '연세대학원 정보미디어 석사');
});

test('splitEducationRecords treats leading degree labels as labels, not extra degrees', () => {
  const records = __testing.splitEducationRecords('박사 : 한양대학교 컴퓨터공학 박사수료');

  assert.deepEqual(records, ['한양대학교 컴퓨터공학 박사수료']);
  assert.equal(__testing.extractHighestEducation(records), '한양대학교 컴퓨터공학 박사수료');
});

test('tagSuspiciousProfile flags missing fields for review', () => {
  const tags = tagSuspiciousProfile({
    phone: EMPTY_VALUE,
    affiliation: EMPTY_VALUE,
    education: EMPTY_VALUE,
    gender: EMPTY_VALUE,
    educationList: [],
    error: false,
  });

  assert.deepEqual(tags, ['phone_missing', 'affiliation_missing', 'education_review', 'gender_missing']);
});

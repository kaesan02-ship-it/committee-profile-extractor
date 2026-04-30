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

test('splitEducationRecords extracts parenthetical degree records and doctorate courses', () => {
  const records = __testing.splitEducationRecords(
    '이화여자대학교 심리학과 ( 학사 ) , 이화여자대학교 교육학 ( 석사 ), 숭실대학교 경영학과 ( 박사과정 )'
  );

  assert.deepEqual(records, [
    '이화여자대학교 심리학과 학사',
    '이화여자대학교 교육학 석사',
    '숭실대학교 경영학과 박사과정',
  ]);
  assert.equal(__testing.extractHighestEducation(records), '숭실대학교 경영학과 박사과정');
});

test('splitEducationRecords carries slash shorthand degree context', () => {
  const records = __testing.splitEducationRecords('경희대학교 관광경영학 학사 / 석사 / 박사 졸');

  assert.deepEqual(records, [
    '경희대학교 관광경영학 학사',
    '경희대학교 관광경영학 석사',
    '경희대학교 관광경영학 박사 졸',
  ]);
  assert.equal(__testing.extractHighestEducation(records), '경희대학교 관광경영학 박사 졸');
});

test('splitEducationRecords removes numbering and keeps education context', () => {
  const records = __testing.splitEducationRecords('1. 경희대학교 컴퓨터공학과 학사 2. 경희대학교 컴퓨터공학과 석사 3. 경희대학교 컴퓨터공학과 박사 졸');

  assert.deepEqual(records, [
    '경희대학교 컴퓨터공학과 학사',
    '경희대학교 컴퓨터공학과 석사',
    '경희대학교 컴퓨터공학과 박사 졸',
  ]);
  assert.equal(__testing.extractHighestEducation(records), '경희대학교 컴퓨터공학과 박사 졸');
});

test('splitEducationRecords moves leading degree labels after the school context', () => {
  const records = __testing.splitEducationRecords('학사 경희대학교 컴퓨터공학과 졸업 석사 경희대학교 컴퓨터공학과 졸업 박사 경희대학교 컴퓨터공학과 졸업');

  assert.deepEqual(records, [
    '경희대학교 컴퓨터공학과 학사',
    '경희대학교 컴퓨터공학과 석사',
    '경희대학교 컴퓨터공학과 박사',
  ]);
  assert.equal(__testing.extractHighestEducation(records), '경희대학교 컴퓨터공학과 박사');
});

test('splitEducationRecords keeps context when doctorate completion appears before school', () => {
  const records = __testing.splitEducationRecords('박사 수료 서울대학교 컴퓨터공학과');

  assert.deepEqual(records, ['서울대학교 컴퓨터공학과 박사수료']);
  assert.equal(__testing.extractHighestEducation(records), '서울대학교 컴퓨터공학과 박사수료');
});

test('splitEducationRecords does not cross-apply degrees across separate education rows', () => {
  assert.deepEqual(
    __testing.splitEducationRecords('한국방송통신대학교 영어영문학 학사\n세종대학교 관광경영학 석사\n세종대학교 호텔관광경영학 박사'),
    [
      '한국방송통신대학교 영어영문학 학사',
      '세종대학교 관광경영학 석사',
      '세종대학교 호텔관광경영학 박사',
    ]
  );

  assert.deepEqual(
    __testing.splitEducationRecords('연세대학교 행정학과 학사\n연세대학교 경제학과 석사\n항공대학교 경영학과 박사'),
    [
      '연세대학교 행정학과 학사',
      '연세대학교 경제학과 석사',
      '항공대학교 경영학과 박사',
    ]
  );
});

test('splitEducationRecords removes degree-only fragments and duplicate bracket degree labels', () => {
  assert.deepEqual(
    __testing.splitEducationRecords('송원대학교 아동보육학 학사\n세종대학교 산업대학원 석사과정'),
    [
      '송원대학교 아동보육학 학사',
      '세종대학교 산업대학원 석사과정',
    ]
  );

  assert.deepEqual(
    __testing.splitEducationRecords('인하대 토목공학 공학사, 중앙대 토목시공관리 공학석사\n[박사] 경희대 건설관리 공학박사'),
    [
      '인하대 토목공학 공학사',
      '중앙대 토목시공관리 공학석사',
      '경희대 건설관리 공학박사',
    ]
  );
});

test('findSectionBody matches nested performance section headers instead of parent titles', () => {
  const allText = [
    '경력사항 및 주요실적',
    '주요이력',
    '現 한국전문면접평가인증원 전문위원',
    '주요실적',
    '[서류] 한국수출입은행, 국민카드',
    '[면접] IBK 기업은행, 금융감독원',
    '기타',
    '자격 및 이수',
  ].join('\n');

  assert.equal(
    __testing.findSectionBody(allText, ['주요실적'], ['기타']),
    '[서류] 한국수출입은행, 국민카드\n[면접] IBK 기업은행, 금융감독원'
  );
});

test('chooseAffiliation keeps current-career position when explicit affiliation is organization-only', () => {
  assert.equal(
    __testing.chooseAffiliation(
      '한국지능정보사회진흥원',
      '(2024.01~ 현재 ) 한국지능정보사회진흥원 인공지능 (AI) 정책실 수석'
    ),
    '한국지능정보사회진흥원 인공지능 (AI) 정책실 수석'
  );
});

test('chooseAffiliation preserves department chair spacing from current career', () => {
  assert.equal(
    __testing.chooseAffiliation('경기과학기술대학교', '현 ) 경기과학기술대학교 전기제어 공학과 학과장 (2020.3~ 현재 )'),
    '경기과학기술대학교 전기제어 공학과 학과장'
  );
});

test('sanitizeAffiliation removes standalone contact label tails', () => {
  assert.equal(
    __testing.sanitizeAffiliation('한국인터넷진흥원 / 디지털위협예방본부 / 디지털보안인증단 단장 핸드폰 )010-3043-9470'),
    '한국인터넷진흥원 / 디지털위협예방본부 / 디지털보안인증단 단장'
  );
});

test('extractBirth supports comma-separated birth dates', () => {
  assert.equal(__testing.extractBirth('생 년 월 일 1979,11.30'), '1979.11.30');
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

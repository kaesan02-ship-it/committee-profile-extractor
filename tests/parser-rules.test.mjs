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
  assert.equal(__testing.extractPhone('연락처 : 010–5567-6216', { requireContext: true }), '010-5567-6216');
  assert.equal(__testing.extractPhone('드림유컨설팅 / 010. 9988. 1981 / email@example.com', { requireContext: true }), '010-9988-1981');
  assert.equal(__testing.extractPhone('(주)태경에스이 소장 / mail@example.com / 010- 4502 - 2985', { requireContext: true }), '010-4502-2985');
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

test('sanitizeAffiliation keeps Jeon-prefixed university names', () => {
  assert.equal(__testing.sanitizeAffiliation('現 전북 대학교 경영학과 교수 (2025 년~현재)'), '전북대학교 경영학과 교수');
  assert.equal(__testing.sanitizeAffiliation('現 전남 대학교 국제처 국제부처장 (2026 년~현재)'), '전남대학교 국제처 국제부처장');
  assert.equal(__testing.sanitizeAffiliation('전 한국대학교 연구원'), '한국대학교연구원');
});

test('splitEducationRecords removes combined mixed-degree fallback records', () => {
  const records = __testing.splitEducationRecords('한국외국어대학교 이란어 / 경영학 (부) 학사 한양사이버대학원 IT MBA 석사');

  assert.deepEqual(records, [
    '한국외국어대학교 이란어 경영학 (부) 학사',
    '한양사이버대학원 IT MBA 석사',
  ]);
  assert.equal(__testing.extractHighestEducation(records), '한양사이버대학원 IT MBA 석사');
});

test('splitEducationRecords stops before inline career section headers', () => {
  assert.deepEqual(
    __testing.splitEducationRecords('충북 대 학교 건축공학과 학사 ➢ 경력사항 및 수행실적'),
    ['충북대학교 건축공학과 학사']
  );
});

test('splitEducationRecords repairs Korean university suffix fragments only', () => {
  assert.deepEqual(
    __testing.splitEducationRecords('충북 대 학교 건축공학과 학사\nYale 대학교 건축대학원 건축학 석사'),
    [
      '충북대학교 건축공학과 학사',
      'Yale 대학교 건축대학원 건축학 석사',
    ]
  );
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

  assert.deepEqual(
    __testing.splitEducationRecords('국가평생교육진흥원 사회복지학 학사\n한성대학교 아동심리학\n학사 서강대학교 교육대학원 평생교육 / 코칭 석사과정 중'),
    [
      '국가평생교육진흥원 사회복지학 학사',
      '한성대학교 아동심리학 학사',
      '서강대학교 교육대학원 평생교육 코칭 석사과정 중',
    ]
  );
});

test('splitEducationRecords expands bracket degree labels with optional counts', () => {
  assert.deepEqual(
    __testing.splitEducationRecords('[학사]\n국민대학교 법학과 [석사]\n서울미디어대학원 AI 소프트웨어학과 (AI 기술경영) 공학석사在'),
    [
      '국민대학교 법학과 학사',
      '서울미디어대학원 AI 소프트웨어학과 (AI 기술경영) 공학석사',
    ]
  );

  assert.deepEqual(
    __testing.splitEducationRecords('[학사 / 5] 경남대 노문학, 한국방송통신대 통계데이터학 / 컴퓨터과학 / 미디어영상학 / 교육학 [석사 / 2] 단국대 경영대학원 경영학 석사, 한국방송통신대 통계대학원 통계데이터학 석사 [박사 / 1] 한성대 컨설팅대학원 매니지먼트 전공 박사'),
    [
      '단국대 경영대학원 경영학 석사',
      '한국방송통신대 통계대학원 통계데이터학 석사',
      '한성대 컨설팅대학원 매니지먼트 전공 박사',
      '경남대 노문학, 한국방송통신대 통계데이터학 컴퓨터과학 미디어영상학 교육학 학사',
    ]
  );
});

test('splitEducationRecords preserves context for colon degree labels', () => {
  assert.deepEqual(
    __testing.splitEducationRecords('박사: 한양대학교 컴퓨터공학 박사수료 (세부전공: 자연어처리, 인공지능) 석사: 한양대학교 컴퓨터공학과 일반대학원 학사: 경주대학교 컴퓨터공학과'),
    [
      '한양대학교 컴퓨터공학 박사수료',
      '한양대학교 컴퓨터공학과 일반대학원 석사',
      '경주대학교 컴퓨터공학과 학사',
    ]
  );
});

test('splitEducationRecords does not carry a previous degree into the next school', () => {
  const records = __testing.splitEducationRecords([
    '성신여자대학교 소비자심리학',
    '박사 연세대학교 교육대학원 인적자원개발',
    '한국외국어대학교 교육대학원 상담심리학석사 한국외국어대학교 서반아어학사',
  ].join('\n'));

  assert.deepEqual(records, [
    '성신여자대학교 소비자심리학 박사',
    '한국외국어대학교 교육대학원 상담심리학석사',
    '한국외국어대학교 서반아어학사',
  ]);
  assert.equal(__testing.extractHighestEducation(records), '성신여자대학교 소비자심리학 박사');
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

test('chooseAffiliation recovers consulting and office names from current career', () => {
  assert.equal(
    __testing.chooseAffiliation('', '現 더이음컨설팅 대표 – HR, 인사, 채용, 컨설팅 – 2021 년 현재 前 커리어넷 기업영업팀 과장'),
    '더이음컨설팅 대표'
  );
  assert.equal(
    __testing.chooseAffiliation('', '現 VAERKSTED 행정 및 브랜딩 컨설팅 (2024. 9~현재) 前 CJ 제일제당 법무팀'),
    'VAERKSTED 행정 및 브랜딩 컨설팅'
  );
  assert.equal(
    __testing.chooseAffiliation('', '(2019. 06. 11~현재) 시야 건축사사무소 대표 (건축계획 · 설계)'),
    '시야 건축사사무소 대표'
  );
  assert.equal(
    __testing.chooseAffiliation('', '現 Ericsson Korea(에릭슨) 인사본부 상무, 인사 총괄 (2024 년 01 월~현재) 前 Hilti Group 글로벌 해양 사업 부문 인사본부 총괄'),
    'Ericsson Korea(에릭슨) 인사본부 상무'
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

test('extractBirth rejects recent career dates as birth dates', () => {
  const recentYear = new Date().getFullYear() - 6;
  assert.equal(__testing.extractBirth(`現 유정노동법률사무소 대표노무사 (${recentYear}. 07. 01)`), '');
  assert.equal(__testing.extractBirth('생 년 월 일 1972.05.19'), '1972.05.19');
});

test('isProfileFormatAnomaly flags embedded profile headers in long career text', () => {
  assert.equal(
    __testing.isProfileFormatAnomaly({
      careerRaw: `${'대표경력 '.repeat(320)} 심사위원 프로필 인적사항 1972. 05. 19 생 년 월 일 중앙대학교 / 교수 연락처 010-7221-6869 성 명`,
    }),
    true
  );
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

test('extractEducationFallbackRecords prefers clean education-only text boxes', () => {
  const records = __testing.extractEducationFallbackRecords([
    '기본 인적사항 이 름 김 세 진 학 력 경력사항 및 수행실적 영남대학교 무역학과 경북대학교 경영학석사 ( 마케팅 ) KAIST 경영학석사 ( 금융공학 )',
    '경북대학교 경영학석사 ( 마케팅 )',
    'KAIST 경영학석사 ( 금융공학 )',
    '대구가톨릭대학교 신학석사 ( 신학 )',
    '영남대학교 경영학박사 ( 인사조직 )',
  ]);

  assert.deepEqual(records, [
    '경북대학교 경영학석사 (마케팅)',
    'KAIST 경영학석사 (금융공학)',
    '대구가톨릭대학교 신학석사 (신학)',
    '영남대학교 경영학박사 (인사조직)',
  ]);
});

test('extractEducationFallbackRecords ignores academic system work descriptions', () => {
  assert.deepEqual(
    __testing.extractEducationFallbackRecords([
      '< 개발업무 > 경주대학교 입시관리, 학사관리시스템, 인사관리시스템 및 수강 신청 등 IT 개발 업무 수행',
    ]),
    []
  );
});

test('fallback extractors recover split contact and evaluation labels', () => {
  assert.equal(
    __testing.extractAffiliationFallbackBody([
      '현소속 / 연락처',
      'iM 뱅크 ( 舊 대구은행 ) 리스크검증팀 팀장 / iM 금융지주 리스크검증팀 팀장 ( 겸직 )',
      'iM Microfinance Myanmar 비상임이사 ( 겸직 ) / dg980210@naver.com / 010-5528-5828',
      '전 문 분 야',
    ]),
    'iM 뱅크 ( 舊 대구은행 ) 리스크검증팀 팀장'
  );

  assert.equal(
    __testing.extractEvaluationFallbackBody([
      '< 면접 위원 >',
      '대구은행 글로벌부문 직원 선발',
      '대구은행 해외유학생 인턴 선발',
      '< 강사 이력 >',
      '금융감독원 청소년 금융경제교육 강사',
    ]),
    '< 면접 위원 > 대구은행 글로벌부문 직원 선발 대구은행 해외유학생 인턴 선발'
  );
});

test('getFixedLayoutProfile extracts the new unlabeled grid without inventing fields', () => {
  const profile = __testing.getFixedLayoutProfile([
    '홍길동',
    '한국전력공사, 한국도로공사',
    '채용면접전문가 교육과정 이수\n(논문) 인재평가 연구',
    '기업은행, 한국산업은행',
    '기타 프로젝트 수행',
    '한국인재연구소 대표',
    'HR, 채용, 교육',
    '1978년 1월 8일',
    'person@example.com',
    '010-1234-5678',
    '인지대학교 영어영문학 학사',
    '숙명여자대학교 인적자원개발 석사',
    '現) 한국인재연구소 대표 (2024년 1월~현재)\n前) 한국기업 팀장 (2016년~2023년)',
  ]);

  assert.equal(profile.affiliation, '한국인재연구소 대표');
  assert.equal(profile.birth, '1978.01.08');
  assert.deepEqual(profile.educationList, [
    '[학사] 인지대학교 영어영문학 학사',
    '[석사] 숙명여자대학교 인적자원개발 석사',
  ]);
  assert.equal(profile.evaluationRaw, '[서류] 한국전력공사, 한국도로공사\n[면접] 기업은행, 한국산업은행');
});

test('getFixedLayoutProfile preserves fixed degree slots instead of shifting nonempty values', () => {
  const profile = __testing.getFixedLayoutProfile([
    { text: '김성중' },
    { text: '現) 한국앙코르커리어 전문위원 (2026년~현재)' },
    { text: '한국해외인프라개발공사, 국가정보원' },
    { text: '평가전문위원 교육 이수' },
    { text: '한국과학기술기획평가원, 국가정보원' },
    { text: '한국앙코르커리어 전문위원' },
    { text: '인사, 행정' },
    { text: '1971년 4월 28일' },
    { text: 'person@example.com' },
    { text: '010-1234-5678' },
    { text: '중앙대 행정학', x: 3240000, y: 1857600 },
    { text: '/ /', x: 7183080, y: 1864800 },
  ]);

  assert.deepEqual(profile.educationList, ['[학사] 중앙대 행정학']);
  assert.equal(profile.educationDegreeConflict, false);
});

test('getFixedLayoutProfile labels placeholder-backed education in bachelor-master-doctor order', () => {
  const profile = __testing.getFixedLayoutProfile([
    { text: '김혜림' },
    { text: '現) 어텀브릿 HR사업본부 이사 (2024년~현재)' },
    { text: '국가과학기술인력개발원, 농협' },
    { text: '채용면접전문가 교육 이수' },
    { text: '소상공인진흥공단, 강서구시설공단' },
    { text: '어텀브릿 HR사업본부 이사' },
    { text: '인사, 채용, 교육' },
    { text: '1984년 5월 19일' },
    { text: 'person@example.com' },
    { text: '010-1234-5678' },
    { text: '인하공업전문대학 항공운항과\n학점은행제 경영학과', placeholderIndex: 30 },
    { text: '연세대학교 호텔외식급식경영학과', placeholderIndex: 31 },
    { text: '한양대학교\n관광학과', placeholderIndex: 32 },
  ]);

  assert.deepEqual(profile.educationList, [
    '[학사] 인하공업전문대학 항공운항과 / 학점은행제 경영학과',
    '[석사] 연세대학교 호텔외식급식경영학과',
    '[박사] 한양대학교 관광학과',
  ]);
});

test('getFixedLayoutProfile flags source text that conflicts with its fixed degree slot', () => {
  const profile = __testing.getFixedLayoutProfile([
    { text: '홍길동' },
    { text: '現) 한국인재연구소 대표 (2024년~현재)' },
    { text: '한국전력공사, 한국도로공사' },
    { text: '전문면접관 교육 이수' },
    { text: '기업은행, 한국산업은행' },
    { text: '한국인재연구소 대표' },
    { text: 'HR, 채용, 교육' },
    { text: '1978년 1월 8일' },
    { text: 'person@example.com' },
    { text: '010-1234-5678' },
    { text: '인지대학교 영어영문학 석사', placeholderIndex: 30 },
  ]);

  assert.equal(profile.educationDegreeConflict, true);
  assert.deepEqual(profile.educationList, ['[학사] 인지대학교 영어영문학 석사']);
});

test('getFixedLayoutProfile does not activate on labeled legacy layouts', () => {
  assert.equal(__testing.getFixedLayoutProfile([
    '기본 인적사항',
    '홍길동',
    '소속 및 연락처',
    '한국인재연구소 대표',
    '전문 분야',
    'HR',
    '1978년 1월 8일',
    'person@example.com',
    '010-1234-5678',
  ]), null);
});

test('getFixedLayoutProfile preserves empty slots and reads strict short-year birth dates', () => {
  const profile = __testing.getFixedLayoutProfile([
    '장은경',
    '現) 가원 대표 (2023년~현재)',
    '농협, 국민은행',
    '전문면접관 교육 이수',
    '한전KPS, 하나은행',
    '',
    '가원 대표',
    'IT, 금융',
    '79.11.26',
    'person@example.com',
    '010-1234-5678',
    '성신여자대학교 컴퓨터정보학부 졸업',
  ]);

  assert.equal(profile.birth, '1979.11.26');
  assert.equal(profile.affiliation, '가원 대표');
  assert.equal(profile.evaluationRaw, '[서류] 농협, 국민은행\n[면접] 한전KPS, 하나은행');
});

test('getFixedLayoutProfile does not duplicate one mixed evaluation paragraph into both categories', () => {
  const profile = __testing.getFixedLayoutProfile([
    '허철',
    '現) 넥스트솔루션 대표 (2017년~현재)',
    '서울시 사업 심사: 참여자 서류 및 면접 평가',
    '넥스트솔루션',
    '채용, 교육',
    '1974.10.29',
    'person@example.com',
    '010-1234-5678',
    '관동대학교 컴퓨터공학',
  ]);

  assert.equal(profile.evaluationRaw, '');
});

test('getFixedLayoutProfile recovers an unmarked dated career block', () => {
  const profile = __testing.getFixedLayoutProfile([
    '김현지',
    '이루다 컨설팅 대표 (2022.01~현재) 에어부산 교육관 (2009.03~2013.07)',
    '기업은행, 산업은행',
    '자격 과정 이수',
    '국민은행, 하나은행',
    '이루다 컨설팅',
    '채용, 교육',
    '1979.03.19',
    'person@example.com',
    '010-1234-5678',
    '영남대학교 심리학과',
  ]);

  assert.match(profile.careerBlock, /이루다 컨설팅 대표/);
});

test('getFixedLayoutProfile recovers two displaced organization evaluation lists only', () => {
  const profile = __testing.getFixedLayoutProfile([
    '박신선',
    '전문면접관 자격과정 이수',
    '내부 실무 인원 채용 참여',
    '피플커리어 대표',
    '',
    '1970년 8월 28일',
    'person@example.com',
    '010-1234-5678',
    '경원대학교 사회체육학과',
    '現) 피플커리어 대표 (2025년~현재)',
    '인사, 채용, 교육',
    '한국석유공사, 한국수출입은행, 코레일공단',
    '한국전력공사, 신용회복위원회, 토지주택공사',
  ]);

  assert.equal(
    profile.evaluationRaw,
    '[서류] 한국석유공사, 한국수출입은행, 코레일공단\n[면접] 한국전력공사, 신용회복위원회, 토지주택공사'
  );
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  formatCareerForTemplate,
  formatEvaluationCareerForTemplate,
} from '../src/lib/profileExcelFormatter.js';

test('formatCareerForTemplate splits spaced current date markers', () => {
  const formatted = formatCareerForTemplate({
    careerDetails: '2021. 10~2022. 03 대신정보통신 부장. 한국교통안전공단 국가자격시스템 개발 / 운영 2024. 01~현 재 SMT 정보기술 부장. 한국교육학술정보원 위탁운영팀 시스템 ADMIN',
  });

  assert.equal(formatted, [
    '現) 2024. 01~현재 SMT 정보기술 부장. 한국교육학술정보원 위탁운영팀 시스템 ADMIN',
    '前) 2021. 10~2022. 03 대신정보통신 부장. 한국교통안전공단 국가자격시스템 개발 / 운영',
  ].join('\n'));
});

test('formatCareerForTemplate keeps trailing date ranges with their career entry', () => {
  assert.equal(
    formatCareerForTemplate({
      careerDetails: '㈜ 앨리오소프트 이사 – 사업관리, 기술영업, 공기업 면접관 및 서류평가 (2023. 08~현재) 마이크로스트레티지 코리아 (주) 이사 – 채용, 조직관리, BI 컨설팅, 세일즈컨설팅, 솔루션 영업 (2003. 06~2023. 07) 시너지 C&C 차장 – SAP 컨설턴트 (BW, CRM)(2002. 01~2003. 05)',
    }),
    [
      '現) ㈜ 앨리오소프트 이사 – 사업관리, 기술영업, 공기업 면접관 및 서류평가 (2023. 08~현재)',
      '前) 마이크로스트레티지 코리아 (주) 이사 – 채용, 조직관리, BI 컨설팅, 세일즈컨설팅, 솔루션 영업 (2003. 06~2023. 07)',
      '前) 시너지 C&C 차장 – SAP 컨설턴트 (BW, CRM)(2002. 01~2003. 05)',
    ].join('\n')
  );
});

test('formatCareerForTemplate keeps leading date ranges attached to following jobs', () => {
  assert.equal(
    formatCareerForTemplate({
      careerDetails: '(1996. 05~2000. 02) 신용보증기금 정보시스템부 팀원 (신용보증시스템 개발) (2000. 02~2004. 02) 신용보증기금 신용정보부 팀원 (신용정보 CRETOP 시스템 기획 및 개발) (2023. 07~현재) 신용보증기금 광화문지점 부지점장',
    }),
    [
      '現) (2023. 07~현재) 신용보증기금 광화문지점 부지점장',
      '前) (1996. 05~2000. 02) 신용보증기금 정보시스템부 팀원 (신용보증시스템 개발)',
      '前) (2000. 02~2004. 02) 신용보증기금 신용정보부 팀원 (신용정보 CRETOP 시스템 기획 및 개발)',
    ].join('\n')
  );
});

test('formatCareerForTemplate handles colon date ranges and ignores activity-only fragments', () => {
  assert.equal(
    formatCareerForTemplate({
      careerDetails: '밸류업컨설팅 (채용평가, 직무교육, 컨설팅, NCS 개발 및 검수): 2013. 07.~현재 한국전문면접평가인증원 사업본부 이사: 2022. 09.~현재 경신 설계팀, 한국오므론전장 연구팀, LS 오토모티브 선행연구팀: 2003. 02.~2013. 04.',
    }),
    [
      '現) 밸류업컨설팅 (채용평가, 직무교육, 컨설팅, NCS 개발 및 검수): 2013. 07.~현재',
      '現) 한국전문면접평가인증원 사업본부 이사: 2022. 09.~현재',
      '前) 경신 설계팀, 한국오므론전장 연구팀, LS 오토모티브 선행연구팀: 2003. 02.~2013. 04.',
    ].join('\n')
  );

  assert.equal(
    formatCareerForTemplate({
      careerDetails: '現 블라썸 컨설팅 / 대표 (2015. 1~현재) (수행실적) - KOICA 선발심사위원 前 앨앤아이컨설팅 ㈜ HR 컨설팅 / 이사 (2008. 5~2014. 12)',
    }),
    [
      '現) 블라썸 컨설팅 / 대표 (2015. 1~현재)',
      '前) 앨앤아이컨설팅 ㈜ HR 컨설팅 / 이사 (2008. 5~2014. 12)',
    ].join('\n')
  );
});

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

test('formatCareerForTemplate keeps hyphen-current dates inside current careers', () => {
  const formatted = formatCareerForTemplate({
    careerDetails: '現 (주)E&C PARTNER 이사 (2019. 01- 현재) 現 고려대학교 평생교육원 상담심리학과 초빙강사 (2014. 03-2023. 08, 2024. 09- 현재) 現 부천대학교 원격평생교육원 강사 (2024. 03- 현재)',
  });

  assert.equal(formatted, [
    '現) (주)E&C PARTNER 이사 (2019. 01- 현재)',
    '現) 고려대학교 평생교육원 상담심리학과 초빙강사 (2014. 03-2023. 08, 2024. 09- 현재)',
    '現) 부천대학교 원격평생교육원 강사 (2024. 03- 현재)',
  ].join('\n'));
});

test('formatCareerForTemplate cleans underscore date ranges and header-only rows', () => {
  const formatted = formatCareerForTemplate({
    careerDetails: '및 수행실적 주요이력 現 커리어엔 대표 (20019 년 _ 현재) _ 조직진단, 역량평가 前 플러스컨설팅 기획 팀장 (2012 년 _ 2015 년) 채용대행',
  });

  assert.equal(formatted, [
    '現) 커리어엔 대표 (2019년~현재) - 조직진단, 역량평가',
    '前) 플러스컨설팅 기획 팀장 (2012 년~2015 년) 채용대행',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate ignores parenthetical interview noise', () => {
  assert.equal(
    formatEvaluationCareerForTemplate({
      careerDetails: '現 세명대학교 교수 前 경희대학교 교수 (면접, 프로젝트 등)',
    }),
    '[미기재] 평가이력 별도 기재 없음'
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

test('formatEvaluationCareerForTemplate handles angle-bracket performance labels', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '<채용면접> 신용보증기금, 한국투자공사, 한국무역보험공사 <영어면접> 한국부동산원, 코트라 <심의위원> 수원지방법원',
  });

  assert.equal(formatted, [
    '[면접] 신용보증기금, 한국투자공사, 한국무역보험공사, 한국부동산원, 코트라',
    '[심사] 수원지방법원',
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

test('formatEvaluationCareerForTemplate uses evaluationRaw from performance sections', () => {
  const formatted = formatEvaluationCareerForTemplate({
    careerRaw: '現 한국 NCS 연구소 이사 (2015. 03~현재)',
    evaluationRaw: '[서류평가] 국민건강보험공단, 기술보증기금, 부산항만공사, 사립학교교직원연금공단, 산업은행 [면접평가] 강원랜드, 광주은행, 건설공제조합, 국민건강보험공단, 국방기술진흥연구소',
  });

  assert.equal(formatted, [
    '[서류] 국민건강보험공단, 기술보증기금, 부산항만공사, 사립학교교직원연금공단 등',
    '[면접] 강원랜드, 광주은행, 건설공제조합, 국민건강보험공단 등',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate groups detailed interview labels', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '[인성면접] 한국중부발전, 한국농수산식품유통공사, 안성시청, 경북 안동의료원, 한국주택금융공사 [토론면접] 한전KPS, 한국수목정원관리원, 한국관광공사 [PT면접] 한국전력기술, 한국에너지공단, 지방공기업평가원 [서류면접] 아동권리보장원, 고용노동교육원, 한국부동산원, 인천공항시설관리, 남북하나재단 [면접모니터링] 부산교육청 교육공무원 채용 [공무원 채용면접] 보건복지부 국립정신건강센터 공무원 채용 업무',
  });

  assert.equal(formatted, [
    '[서류] 아동권리보장원, 고용노동교육원, 한국부동산원, 인천공항시설관리 등',
    '[면접] 한국중부발전, 한국농수산식품유통공사, 안성시청, 경북 안동의료원 등, 한전KPS, 한국수목정원관리원, 한국관광공사, 한국전력기술, 한국에너지공단, 지방공기업평가원, 부산교육청 교육공무원 채용, 보건복지부 국립정신건강센터 공무원 채용 업무',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate handles interview panelist labels', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '[면접관]: 강원랜드, 국립생태원, 농업실용화재단, 국립항공박물관, 한국과학창의재단, 국립낙동강생물자원관 [기업컨설팅] SK 그룹, 롯데백화점, 존슨 & 존슨, KT, 삼성전자',
  });

  assert.equal(formatted, [
    '[면접] 강원랜드, 국립생태원, 농업실용화재단, 국립항공박물관 등',
    '[자문] SK 그룹, 롯데백화점, 존슨 & 존슨, KT 등',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate handles spaced evaluation labels with review blocks', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '서류 평가: 신용보증기금, 대구광역시 동구 등 면접 평가: 신용보증기금, 한국토지주택공사, 한국가스공사 등 정부과제 심사: 신용보증기금 등',
  });

  assert.equal(formatted, [
    '[서류] 신용보증기금, 대구광역시 동구 등',
    '[면접] 신용보증기금, 한국토지주택공사, 한국가스공사 등',
    '[심사] 신용보증기금 등',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate handles hyphen labels and advisory committee labels', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '서류 - NRF 한국연구재단, 한국남동발전 면접 - 한국은행, 금융감독원 <자문위원> 한국철도공사, 서울시 강동구 지속가능위원회',
  });

  assert.equal(formatted, [
    '[서류] NRF 한국연구재단, 한국남동발전',
    '[면접] 한국은행, 금융감독원',
    '[자문] 한국철도공사, 서울시 강동구 지속가능위원회',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate ignores a repeated evaluation project block', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '[HR Project Managing / 채용, 선발 및 평가 컨설팅 운영] 면접: BEI, PT, GD, IB, AI 면접 / 면접관 교육 및 위원장 한국은행, 한국석유공사, 한국도로공사, 한전원자력연료, 한국가스기술공사 서류: 국방기술진흥연구소, 금융감독원, KB 국민은행, 한국예탁결제원 [HR Project Managing / 채용, 선발 및 평가 컨설팅 운영] 면접: 금융감독원, 예금보험공사, 한국부동산신탁, 한국무역보험공사 서류 평가: 금융결제원, KB 국민은행, 한국예탁결제원',
  });

  assert.equal(formatted, [
    '[서류] 국방기술진흥연구소, 금융감독원, KB 국민은행, 한국예탁결제원',
    '[면접] BEI, PT, GD, IB 등',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate keeps distinct interview blocks without repeated project headings', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '[인성면접] 한국은행, 금융감독원 [토론면접] 한국수출입은행, 한국무역보험공사',
  });

  assert.equal(formatted, '[면접] 한국은행, 금융감독원, 한국수출입은행, 한국무역보험공사');
});

test('formatEvaluationCareerForTemplate does not spill a trailing evaluation label into career text', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '[면접] 한국은행, 농협, 서민금융진흥원, 한국벤처투자 [서류] 서울경제진흥원',
    careerRaw: '임팩트그룹코리아 / HR Business Partner: 조직문화 운영',
  });

  assert.equal(formatted, [
    '[서류] 서울경제진흥원',
    '[면접] 한국은행, 농협, 서민금융진흥원, 한국벤처투자',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate ignores bare document words in descriptive phrases', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '채용심사: K-water AI 면접, 농협 AI 개발자 면접, 농협 개발 경력자 서류 평가, 한국전력 IT 직군, 기업은행',
  });

  assert.equal(formatted, '[심사] K-water AI 면접, 농협 AI 개발자 면접, 농협 개발 경력자 서류 평가, 한국전력 IT 직군 등');
});

test('formatEvaluationCareerForTemplate summarizes generic evaluation activities', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '現 코리아에이치알디솔루션 대표 現 강원특별자치도 공무원 면접 평가위원 (2023년~현재) 現 한국 HR 진단평가센터 공공기관 채용 평가위원 (2019년~현재)',
  });

  assert.equal(formatted, '[심사] 現 강원특별자치도 공무원 면접 평가위원 (2023년~현재), 現 한국 HR 진단평가센터 공공기관 채용 평가위원 (2019년~현재)');
});

test('formatEvaluationCareerForTemplate fills rows with no evaluation evidence', () => {
  assert.equal(formatEvaluationCareerForTemplate({}), '[미기재] 평가이력 별도 기재 없음');
});

test('formatEvaluationCareerForTemplate cleans leading colons and profile tails', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '[면접] : 하나은행, 국민은행, 한국연구재단 면접관 Profile',
  });

  assert.equal(formatted, '[면접] 하나은행, 국민은행, 한국연구재단');
});

test('formatEvaluationCareerForTemplate handles public recruitment and panel labels', () => {
  assert.equal(
    formatEvaluationCareerForTemplate({
      evaluationRaw: '[공채면접] 한전 KPS, 사천시청, 한국도로공사 [공채서류전형] 코레일테크, IBK 기업은행',
    }),
    '[서류] 코레일테크, IBK 기업은행\n[면접] 한전 KPS, 사천시청, 한국도로공사'
  );

  assert.equal(
    formatEvaluationCareerForTemplate({
      evaluationRaw: '< 면접 위원 > 대구은행 글로벌부문 직원 선발 대구은행 해외유학생 인턴 선발 < 강사 이력 > 금융감독원 청소년 금융경제교육 강사',
    }),
    '[면접] 대구은행 글로벌부문 직원 선발 대구은행 해외유학생 인턴 선발'
  );
});

test('formatEvaluationCareerForTemplate separates spaced hiring interview labels', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '< 서류 > 기업은행, 한국관세정보원 < 채용 면접 > 한국발명진흥회, 한국과학창의재단 < 자문위원 > 한국소방산업기술원',
  });

  assert.equal(formatted, [
    '[서류] 기업은행, 한국관세정보원',
    '[면접] 한국발명진흥회, 한국과학창의재단',
    '[자문] 한국소방산업기술원',
  ].join('\n'));
});

test('formatEvaluationCareerForTemplate cuts non-evaluation tails', () => {
  assert.equal(
    formatEvaluationCareerForTemplate({
      evaluationRaw: '면접위원: 근로복지공단, 한국원자력환경공단 HR 관련: 기업은행, 서민금융진흥원',
    }),
    '[면접] 근로복지공단, 한국원자력환경공단'
  );

  assert.equal(
    formatEvaluationCareerForTemplate({
      evaluationRaw: '[서류] 표준협회, 충북해양과학관 [면접] KB 국민은행, 산업은행',
    }),
    '[서류] 표준협회, 충북해양과학관\n[면접] KB 국민은행, 산업은행'
  );

  assert.equal(
    formatEvaluationCareerForTemplate({
      evaluationRaw: '< 채용면접 > 신용보증기금 < 심사위원 > 중소상공인희망재단',
    }),
    '[면접] 신용보증기금\n[심사] 중소상공인희망재단'
  );
});

test('formatEvaluationCareerForTemplate does not mix career text when evaluationRaw exists', () => {
  const formatted = formatEvaluationCareerForTemplate({
    evaluationRaw: '면접: IBK 기업은행, 산업은행 서류: 국민카드, 코스콤',
    careerRaw: '㈜ 앨리오소프트 이사 – 공기업 면접관 및 서류평가 (2023.08~ 현재)',
  });

  assert.equal(formatted, [
    '[서류] 국민카드, 코스콤',
    '[면접] IBK 기업은행, 산업은행',
  ].join('\n'));
});

test('formatCareerForTemplate excludes explicit evaluation blocks from career summary', () => {
  const formatted = formatCareerForTemplate({
    affiliation: 'HR 임팩트 대표',
    careerDetails: 'HR 임팩트 대표: (2024~현재) ㈜ 임팩트그룹코리아 이사: 조직개발센터 (2024~현재) ㈜선연그룹 이사: 컨설팅 사업본부 (2021~2024) ㈜ 에스티유니타스 전략기획팀장: 전략 및 신사업 기획 (2015~2017) 면접관 경력 면접: KDB 산업은행, 금융감독원 서류: 기술보증기금',
  });

  assert.equal(formatted, [
    '現) HR 임팩트 대표: (2024~현재)',
    '現) ㈜ 임팩트그룹코리아 이사: 조직개발센터 (2024~현재)',
    '前) ㈜선연그룹 이사: 컨설팅 사업본부 (2021~2024)',
    '前) ㈜ 에스티유니타스 전략기획팀장: 전략 및 신사업 기획 (2015~2017)',
  ].join('\n'));
});

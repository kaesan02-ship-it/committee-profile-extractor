const AFFILIATION_FALSE_EDUCATION_PATTERN = /(학사|석사|박사|전문학사|최종학력|졸업|수료|재학|학위)/i;
const AFFILIATION_ORG_PATTERN = /(대학교|대학원|대학|연구원|연구소|센터|병원|재단|법인|회사|기업|공사|공단|기관|캠퍼스)/i;
const AFFILIATION_DEPARTMENT_PATTERN = /(학과|학부|전공|연구소|센터|본부|실|팀|과|처|부)/i;
const AFFILIATION_MEMBERSHIP_PATTERN =
  /((?:한국|대한|국제|미국|세계|IEEE|ICROS|KSME|KICS|ACS|AIChE)[^()]{0,30}?(?:학회|협회|위원회|연구회)|정회원|종신회원|상임이사|부회장|사업이사|자문위원|평가위원|심사위원|선정위원|운영위원)/i;

const AFFILIATION_POSITION_SOURCE = POSITION_HINTS.slice()
  .sort((a, b) => b.length - a.length)
  .map(escapeRegex)
  .join('|');

const AFFILIATION_POSITION_REGEX = new RegExp(
  `^(.+?(?:${AFFILIATION_POSITION_SOURCE}))(?=\\s|$|[,)])`,
  'i'
);

const normalizeAffiliationSpacing = (text = '') => {
  return cleanInline(text)
    .replace(/([가-힣A-Za-z]{2,})\s+대학교/g, '$1대학교')
    .replace(/([가-힣A-Za-z]{2,})\s+대학원/g, '$1대학원')
    .replace(/([가-힣A-Za-z]{2,})\s+캠퍼스/g, '$1캠퍼스')
    .replace(/([가-힣A-Za-z]{2,})\s+연구소/g, '$1연구소')
    .replace(/([가-힣A-Za-z]{2,})\s+연구원/g, '$1연구원')
    .replace(/([가-힣A-Za-z]{2,})\s+센터/g, '$1센터')
    .replace(/([가-힣A-Za-z]{2,})\s+학과/g, '$1학과')
    .replace(/([가-힣A-Za-z]{2,})\s+학부/g, '$1학부')
    .replace(/([가-힣A-Za-z]{2,})\s+전공/g, '$1전공')
    .replace(/\s{2,}/g, ' ')
    .trim();
};

const stripAffiliationTail = (text = '') => {
  let value = cleanInline(text);
  if (!value) return '';

  // 첫 번째 현재 소속 뒤에 이어지는 두 번째 現/현재/前/전 경력은 잘라냄
  value = value.split(/\s+(?=(?:現|현재|현|前|전)\s+)/)[0].trim();

  // 학회/정회원/자문위원 등 현재 소속이 아닌 활동 정보는 잘라냄
  const membershipMatch = value.match(AFFILIATION_MEMBERSHIP_PATTERN);
  if (membershipMatch && membershipMatch.index != null && membershipMatch.index > 0) {
    value = value.slice(0, membershipMatch.index).trim();
  }

  value = value
    // 재직 기간 제거
    .replace(/\(\s*\d{4}[^)]*?\)/g, ' ')
    .replace(
      /\d{4}(?:\s*년)?(?:[.]\d{1,2}|\s*년\s*\d{1,2}\s*월)?\s*[~∼〜-]\s*(?:\d{4}(?:\s*년)?(?:[.]\d{1,2}|\s*년\s*\d{1,2}\s*월)?|현재)\s*/g,
      ' '
    )
    // 중복 괄호 직함 제거
    .replace(/\(\s*(?:교수|정교수|부교수|조교수|대표|원장|센터장|소장|부장|팀장|실장)\s*\)$/i, ' ')
    .replace(/[|]+/g, ' ')
    .replace(/\s*[,/·]+\s*$/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

  // "대학교 + 학과/학부/전공 + 교수" 같은 핵심 구간만 우선 확보
  const positionMatch = value.match(AFFILIATION_POSITION_REGEX);
  if (positionMatch?.[1] && hasAffiliationHint(positionMatch[1])) {
    value = positionMatch[1].trim();
  }

  return normalizeAffiliationSpacing(value)
    .replace(/^[\])}\s]+/, '')
    .replace(/[([{]\s*$/, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
};

const splitCurrentCareerSegments = (text = '') => {
  const source = sanitizeBracketArtifacts(text)
    .replace(/\s*[•■□▷▶*]+\s*/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (!source) return [];

  const segments = [];
  const regex = /(?:^|\s)(現|현재|현)\s+([\s\S]*?)(?=(?:\s+(?:現|현재|현|前|전)\s+)|$)/gi;

  let match;
  while ((match = regex.exec(source))) {
    const raw = `${match[1]} ${match[2]}`.trim();
    if (raw) segments.push(raw);
  }

  return uniq(segments);
};

const sanitizeAffiliation = (text = '') => {
  let value = normalizeAffiliationSpacing(text);
  if (!value) return '';

  value = value
    .replace(AFFILIATION_LABEL_PATTERN, '')
    .replace(CURRENT_PREFIX_PATTERN, '')
    .replace(PREVIOUS_PREFIX_PATTERN, '')
    .replace(DATE_RANGE_PREFIX_PATTERN, '')
    .replace(EMAIL_PATTERN_GLOBAL, ' ')
    .replace(PHONE_PATTERN, ' ')
    .replace(/\b01[016789]\d{7,8}\b/g, ' ')
    .replace(/\(\s*이메일[^)]*\)/gi, ' ')
    .replace(/\(\s*연락처[^)]*\)/gi, ' ')
    .replace(/(?:핸드폰번호|메일주소|이메일|연락처|휴대폰|휴대전화)\s*[:：]?\s*.*$/i, ' ')
    .replace(/\\/g, ' ')
    .replace(/\s+-\s+.*$/, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

  value = stripAffiliationTail(value);

  value = truncateAtLooseKeyword(value, AFFILIATION_CUTOFF_KEYWORDS)
    .replace(/^[\])}\s]+/, '')
    .replace(/[([{]\s*$/, '')
    .replace(/\s*[/,-]+\s*$/, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (!hasPosition(value) && value.includes('(')) {
    value = value.split('(')[0].trim();
  }

  return normalizeAffiliationSpacing(value);
};

const scoreAffiliationCandidate = (line = '', source = 'affiliation') => {
  const value = sanitizeAffiliation(line);
  if (!value) return -999;

  let score = 0;

  if (source === 'career-current') score += 6;
  if (hasCurrentMarker(line)) score += 4;
  if (hasAffiliationHint(value)) score += 3;
  if (AFFILIATION_ORG_PATTERN.test(value)) score += 3;
  if (AFFILIATION_DEPARTMENT_PATTERN.test(value)) score += 3;
  if (hasPosition(value)) score += 5;
  if (value.length >= 8 && value.length <= 40) score += 3;

  if (value.length > 55) score -= 5;
  if (CONTACT_NOISE_PATTERN.test(value) || extractEmail(value) || extractPhone(value)) score -= 5;
  if (AFFILIATION_FALSE_EDUCATION_PATTERN.test(value)) score -= 4;
  if (AFFILIATION_MEMBERSHIP_PATTERN.test(value)) score -= 8;
  if (CAREER_SECTION_LINE_PATTERN.test(value) && source !== 'career-current') score -= 1;
  if (/\d{4}/.test(value)) score -= 2;

  return score;
};

const extractCurrentCareerLine = (text = '') => {
  const directSegments = splitCurrentCareerSegments(text);

  const fallbackEntries = splitCareerEntries(text)
    .map((line) => line.split(/\s+(?=(?:前|전|\(전\))\s*[)\-:：]?)/)[0])
    .map((line) => safeText(line))
    .filter(Boolean);

  const entries = uniq([...directSegments, ...fallbackEntries]);

  const currentCandidates = entries
    .map((line, index) => ({ line, index }))
    .filter(({ line }) => hasCurrentMarker(line) || /현재/.test(line))
    .map(({ line, index }) => ({ value: sanitizeAffiliation(line), raw: line, index }))
    .filter(({ value }) => Boolean(value))
    .sort(
      (a, b) =>
        scoreAffiliationCandidate(b.raw, 'career-current') - scoreAffiliationCandidate(a.raw, 'career-current') ||
        b.value.length - a.value.length ||
        a.index - b.index
    );

  return currentCandidates[0]?.value || '';
};

const chooseAffiliation = (affiliationText = '', careerText = '') => {
  const candidates = [];

  splitBullets(affiliationText).forEach((line) => {
    const cleaned = sanitizeAffiliation(line);
    if (cleaned) candidates.push({ value: cleaned, source: 'affiliation', raw: line });
  });

  const currentCareer = extractCurrentCareerLine(careerText);
  if (currentCareer) {
    candidates.push({ value: currentCareer, source: 'career-current', raw: currentCareer });
  }

  const scored = uniq(candidates.map((item) => `${item.source}::${item.value}`))
    .map((key) => {
      const [source, value] = key.split('::');
      return { value, source, score: scoreAffiliationCandidate(value, source) };
    })
    .sort(
      (a, b) =>
        b.score - a.score ||
        Number(b.source === 'career-current') - Number(a.source === 'career-current') ||
        b.value.length - a.value.length
    );

  return scored[0]?.value || '';
};

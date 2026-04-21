import JSZip from 'jszip';
import {
  AFFILIATION_CUTOFF_KEYWORDS,
  AFFILIATION_HINTS,
  CAREER_START_PATTERN,
  CURRENT_MARKERS,
  DEGREE_PATTERNS,
  EDUCATION_COMPACT_PATTERNS,
  EDUCATION_DEGREE_KEYWORDS,
  EDUCATION_STATUS_PATTERNS,
  EMPTY_VALUE,
  EXPERTISE_PROTECTED_TERMS,
  EXPERTISE_SPLIT_PATTERN,
  FIELD_LABELS,
  POSITION_HINTS,
  SECTION_STOP_HEADERS,
} from './extractionRules.js';

const PHONE_PATTERN = /0\s*1\s*[016789][\s./-]*\d[\s\d]{2,3}[\s./-]*\d[\s\d]{3}/;
const PHONE_DIGIT_PATTERN = /01[016789]\d{7,8}/;
const EMAIL_PATTERN = /[a-zA-Z0-9._%+-]+\s*@\s*(?:[a-zA-Z0-9-]+\s*\.\s*)+[a-zA-Z]{2,}/;
const EMAIL_PATTERN_GLOBAL = new RegExp(EMAIL_PATTERN.source, 'g');
const DEGREE_PATTERN_GLOBAL = /(?:박사과정수료|박사수료|석박사\s*통합|석박사통합|박사|석사|학사|전문학사|Ph\.?D|Doctor|MBA|M\.?A|M\.?S|B\.?A|B\.?S)/i;
const DEGREE_ONLY_PATTERN = /^(?:박사과정수료|박사수료|석박사\s*통합|석박사통합|박사|석사|학사|전문학사)$/i;
const EDUCATION_NOISE_PATTERN = /^(?:학력|학\s*력|최종학력)\s*[:：]?/;
const CAREER_NOISE_PATTERN = /^(?:주요경력|주요\s*경력|주요이력|주요\s*이력|경력사항|경력)\s*[:：]?/;
const AFFILIATION_LABEL_PATTERN = /^(?:소속|현소속|현재소속|현직|근무처|소속및연락처|소속\s*및\s*연락처|소속\s*\/\s*연락처|소속\s*및\s*직위|소속\s*\/\s*직위)\s*[:：]?/;
const CURRENT_PREFIX_PATTERN = /^(?:현재|현|現|現\)|現\s*\)|\(현\)|\[현\]|\{현\})\s*[-)\]:：.]*/;
const PREVIOUS_PREFIX_PATTERN = /^(?:전|前|\(전\))\s*[-)\]:：.]*/;
const DATE_RANGE_PREFIX_PATTERN = /^(?:\d{4}(?:[.]\d{1,2})?\s*[~∼〜-]\s*(?:\d{4}(?:[.]\d{1,2})?|현재)|\d{4}\s*년\s*~\s*(?:\d{4}\s*년|현재)|\d{4}[.]\d{1,2}\s*~\s*현재)\s*/;
const CONTACT_NOISE_PATTERN = /(연락처|휴대전화|휴대폰|핸드폰번호|전화번호|이메일|E-mail|Email|메일|메일주소)/i;
const EDUCATION_KEYWORD_PATTERN = /(학사|석사|박사|전문학사|대학교|대학원|학위|졸업|수료|재학|학과|학부|전공|University|Department|School|College)/i;
const CAREER_SECTION_LINE_PATTERN = /(경력|이력|재직|근무|수행|프로젝트|담당|위원|심사|평가)/;
const EXPERTISE_LABEL_PATTERN = /^(?:전문분야|전\s*문\s*분야|전문\s*분야|전문\s*산업분야|전문\s*직무분야|주요분야|전공|핵심역량)\s*[:：]?/;
const TOKEN_PLACEHOLDER = '__SLASH__';

const decodeHTML = (text = '') => {
  return String(text)
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
};

const normalize = (text = '') => String(text).replace(/[\s\t\n\r:：·•ㆍ,()[\]{}<>/]/g, '');
const cleanInline = (text = '') => decodeHTML(String(text)).replace(/[“”"`]/g, ' ').replace(/[\u00A0\u3000]/g, ' ').replace(/\s+/g, ' ').trim();
const uniq = (arr = []) => [...new Set(arr.filter(Boolean))];
const safeText = (text = '') => cleanInline(text).replace(/^[-•■□▷▶*]+\s*/, '');
const firstNonEmpty = (...values) => values.map((value) => cleanInline(value)).find(Boolean) || '';
const firstNonEmptyPreserveLines = (...values) => values.map((value) => String(value ?? '').trim()).find(Boolean) || '';
const escapeRegex = (text = '') => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const normalizeDegreeLabel = (text = '') => {
  const value = cleanInline(text).replace(/\s+/g, '');
  if (/석박사통합|통합석박사/i.test(value)) return '석박사 통합';
  if (/박사과정수료/i.test(value)) return '박사과정수료';
  if (/박사수료/i.test(value)) return '박사수료';
  if (/박사/i.test(value)) return '박사';
  if (/석사/i.test(value)) return '석사';
  if (/학사/i.test(value)) return '학사';
  if (/전문학사/i.test(value)) return '전문학사';
  return cleanInline(text);
};

const sanitizeBracketArtifacts = (text = '') => {
  let value = decodeHTML(text)
    .replace(/[［【]/g, '[')
    .replace(/[］】]/g, ']')
    .replace(/[（]/g, '(')
    .replace(/[）]/g, ')')
    .replace(/[“”"`]/g, ' ')
    .replace(/[\u00A0\u3000]/g, ' ');

  let round = 0;
  let square = 0;
  let output = '';

  for (const ch of value) {
    if (ch === '(') {
      round += 1;
      output += ch;
      continue;
    }
    if (ch === ')') {
      if (round > 0) {
        round -= 1;
        output += ch;
      }
      continue;
    }
    if (ch === '[') {
      square += 1;
      output += ch;
      continue;
    }
    if (ch === ']') {
      if (square > 0) {
        square -= 1;
        output += ch;
      }
      continue;
    }
    output += ch;
  }

  output = output.replace(/[[(]+\s*$/g, ' ').replace(/^[\])}\s]+/g, ' ');
  return output.replace(/\s+/g, ' ').trim();
};

const compactEducationSpacing = (text = '') => {
  let value = sanitizeBracketArtifacts(text);

  EDUCATION_COMPACT_PATTERNS.forEach(({ regex, replace }) => {
    value = value.replace(regex, replace);
  });

  return value
    .replace(/([가-힣A-Za-z]{2,})\s+정비\s+공\s+학/g, '$1정비공학')
    .replace(/([가-힣A-Za-z]{2,})\s+항공\s+정비\s+공\s+학/g, '$1항공정비공학')
    .replace(/항공정비공\s+학/g, '항공정비공학')
    .replace(/항공산업\s+학/g, '항공산업학')
    .replace(/([가-힣A-Za-z]{2,})\s+생산\s+공\s+학/g, '$1생산공학')
    .replace(/([가-힣A-Za-z]{2,})\s+산업\s+학/g, '$1산업학')
    .replace(/설계\s+학/g, '설계학')
    .replace(/([가-힣A-Za-z]{2,})\s+([가-힣A-Za-z]{1,12})\s+학(?=\s*(?:학사|석사|박사|과|부|전공|$|,|\)|\n))/g, '$1$2학')
    .replace(/(대학교|대학원|대학|교육원|University|College|School)(?=[가-힣A-Za-z]{2,}(?:학과|학부|전공|공학|산업학|설계학|기술교육원|학))/g, '$1 ')
    .replace(/([가-힣A-Za-z]{2,})\s*-\s*([가-힣A-Za-z]{1,12})\s*학/g, '$1$2학')
    .replace(/([가-힣A-Za-z])-(?=(?:공학|과|부|전공|학))/g, '$1')
    .replace(/-\s*$/g, '')
    .replace(/\s*·\s*/g, ' · ')
    .replace(/\s*,\s*/g, ', ')
    .replace(/\s*\/\s*/g, ' / ')
    .replace(/\[\s*/g, '[')
    .replace(/\s*\]/g, ']')
    .replace(/\(\s*/g, ' (')
    .replace(/\s*\)/g, ') ')
    .replace(/\s{2,}/g, ' ')
    .trim();
};

const prepareEducationSource = (text = '') => {
  return compactEducationSpacing(text)
    .replace(/(?<=[가-힣A-Za-z])(?=\[(?:학사|석사|박사|전문학사|박사수료|박사과정수료)|\((?:학사|석사|박사|전문학사|박사수료|박사과정수료))/g, ' ')
    .replace(/(학사|석사|박사|전문학사|박사수료|박사과정수료|석박사\s*통합)\s*(?=(?:[가-힣A-Za-z]{2,}(?:대학교|대학원|대학|교육원|University|College|School)))/g, '$1\n')
    .replace(/([\])}])\s*(?=(?:[가-힣A-Za-z]{2,}(?:대학교|대학원|대학|교육원|University|College|School)))/g, '$1\n')
    .replace(/\s*[,;/]\s*(?=(?:[가-힣A-Za-z]{2,}(?:대학교|대학원|대학|교육원|University|College|School)|\[|\())/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/\n{2,}/g, '\n')
    .trim();
};

const splitBullets = (text = '') => {
  return uniq(
    sanitizeBracketArtifacts(text)
      .replace(/\s*[•■□▷▶*]+\s*/g, '\n')
      .replace(/\s*[－─]\s*/g, '\n')
      .replace(/\s+(?=(?:\d{2,4}[.]\d{1,2}\s*[~∼〜-]|\d{4}\s*년\s*~))/g, '\n')
      .split(/\n+/)
      .map((line) => safeText(line))
      .filter((line) => line.length > 1)
  );
};

const isLikelyName = (value = '') => /^[가-힣]{2,5}$/.test(cleanInline(value).replace(/\(.*?\)/g, '').replace(/\s/g, ''));

const refineName = (value = '') => {
  let name = cleanInline(value).replace(/\(.*?\)/g, '').replace(/\s/g, '');
  const labels = [
    ...FIELD_LABELS.birth,
    ...FIELD_LABELS.phone,
    ...FIELD_LABELS.email,
    ...FIELD_LABELS.affiliation,
    ...FIELD_LABELS.gender,
    ...FIELD_LABELS.education,
    ...FIELD_LABELS.expertise,
    ...FIELD_LABELS.career,
    '위원',
  ];

  for (const label of labels) {
    if (name.includes(label)) name = name.split(label)[0];
  }

  return isLikelyName(name) ? name : '';
};

const normalizeEmail = (value = '') => cleanInline(value).replace(/\s+/g, '');

const extractEmail = (text = '') => {
  const match = text.match(EMAIL_PATTERN);
  return match ? normalizeEmail(match[0]) : '';
};

const extractPhone = (text = '') => {
  const match = text.match(PHONE_PATTERN);
  if (match) {
    const digits = match[0].replace(/\D/g, '');
    if (digits.length >= 10 && digits.length <= 11) {
      return digits.replace(/(\d{3})(\d{3,4})(\d{4})/, '$1-$2-$3');
    }
  }

  const fallback = text.replace(/\D/g, '').match(PHONE_DIGIT_PATTERN);
  if (fallback) return fallback[0].replace(/(\d{3})(\d{3,4})(\d{4})/, '$1-$2-$3');
  return '';
};

const isValidBirthParts = (year, month, day) => {
  const y = Number(year);
  const m = Number(month);
  const d = Number(day);
  const currentYear = new Date().getFullYear();

  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return false;
  if (y < 1900 || y > currentYear) return false;
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;

  const date = new Date(y, m - 1, d);
  return date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d;
};

const extractBirth = (text = '') => {
  const match = text.match(/(\d{4})[\s./-년]*?(\d{1,2})[\s./-월]*?(\d{1,2})\s*일?/);
  if (match && isValidBirthParts(match[1], match[2], match[3])) {
    return `${match[1]}.${String(match[2]).padStart(2, '0')}.${String(match[3]).padStart(2, '0')}`;
  }

  const digits = text.replace(/\D/g, '');
  if (digits.length >= 8) {
    const value = digits.slice(0, 8);
    if (/^(19|20)\d{6}$/.test(value)) {
      const year = value.slice(0, 4);
      const month = value.slice(4, 6);
      const day = value.slice(6, 8);
      if (isValidBirthParts(year, month, day)) {
        return `${year}.${month}.${day}`;
      }
    }
  }

  return '';
};

const extractGender = (text = '') => {
  const normalized = cleanInline(text).replace(/\s/g, '');
  const labelMatch = normalized.match(/성별[:：]?(남성|여성|남자|여자|남|여)/);
  if (labelMatch) return labelMatch[1].startsWith('남') ? '남' : '여';

  const bareMatch = normalized.match(/(^|[^가-힣])(남성|여성|남자|여자|남|여)([^가-힣]|$)/);
  if (bareMatch) return bareMatch[2].startsWith('남') ? '남' : '여';
  return '';
};

const degreeRank = (line = '') => DEGREE_PATTERNS.find((item) => item.regex.test(line))?.rank || 0;
const educationStatusWeight = (line = '') => EDUCATION_STATUS_PATTERNS.find((item) => item.regex.test(line))?.weight || 0;
const hasPosition = (line = '') => POSITION_HINTS.some((hint) => cleanInline(line).includes(hint));
const hasAffiliationHint = (line = '') => AFFILIATION_HINTS.some((hint) => cleanInline(line).includes(hint));
const hasCurrentMarker = (line = '') => CURRENT_MARKERS.some((marker) => cleanInline(line).startsWith(marker));

const truncateAtLooseKeyword = (text = '', keywords = []) => {
  const value = cleanInline(text);
  let cutIndex = value.length;

  keywords.forEach((keyword) => {
    const regex = new RegExp(keyword.split(/\s+/).map(escapeRegex).join('\\s*'), 'i');
    const match = regex.exec(value);
    if (match && match.index > 0 && match.index < cutIndex) {
      cutIndex = match.index;
    }
  });

  return value.slice(0, cutIndex).trim();
};

const normalizeEducationRecord = (line = '') => {
  let value = compactEducationSpacing(line)
    .replace(EDUCATION_NOISE_PATTERN, '')
    .replace(/^[-•■□▷▶*]+\s*/, '')
    .replace(/^[\])}\s]+/, '')
    .replace(/[([{]\s*$/, '')
    .replace(/[,:;]\s*$/g, '')
    .replace(/-\s*$/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

  value = value.replace(/^,\s*/, '').replace(/,\s*$/, '').trim();
  return value;
};

const splitDegreeList = (text = '') => {
  return uniq(
    sanitizeBracketArtifacts(text)
      .split(/\s*,\s*|\s*·\s*|\s*\/\s*/)
      .map((token) => normalizeDegreeLabel(token))
      .filter((token) => EDUCATION_DEGREE_KEYWORDS.includes(token) || token === '석박사 통합')
  );
};

const expandIntegratedDegree = (prefix = '', degree = '') => {
  const normalizedPrefix = normalizeEducationRecord(prefix).replace(/[,:;/-]\s*$/g, '').trim();
  const normalizedDegree = normalizeDegreeLabel(degree);

  if (!normalizedPrefix) {
    if (normalizedDegree === '석박사 통합') {
      return ['석사', '박사'];
    }
    return normalizedDegree ? [normalizedDegree] : [];
  }

  if (normalizedDegree === '석박사 통합') {
    return [
      normalizeEducationRecord(`${normalizedPrefix} 석사`),
      normalizeEducationRecord(`${normalizedPrefix} 박사`),
    ];
  }

  return [normalizeEducationRecord(`${normalizedPrefix} ${normalizedDegree}`)];
};

const extractEducationContext = (line = '') => {
  const normalizedLine = normalizeEducationRecord(line);
  const degreeMatch = normalizedLine.match(DEGREE_PATTERN_GLOBAL);
  if (!degreeMatch || degreeMatch.index == null || degreeMatch.index <= 0) return '';

  const prefix = normalizedLine
    .slice(0, degreeMatch.index)
    .replace(/\[[^\]]*$/g, '')
    .replace(/\([^)]*$/g, '')
    .replace(/[,:;/-]\s*$/g, '')
    .trim();

  return EDUCATION_KEYWORD_PATTERN.test(prefix) ? prefix : '';
};

const isStandaloneDegreeRecord = (record = '') => {
  const normalized = cleanInline(record)
    .replace(/^[\s[(]+/g, '')
    .replace(/[\s)\]]+$/g, '')
    .replace(/\s+/g, '');

  return DEGREE_ONLY_PATTERN.test(normalized);
};

const isMeaningfulEducationContext = (text = '') => {
  const value = normalizeEducationRecord(text)
    .replace(EDUCATION_NOISE_PATTERN, '')
    .replace(/[,:;/-]\s*$/g, '')
    .trim();

  if (!value) return false;
  if (isStandaloneDegreeRecord(value)) return false;

  return (
    EDUCATION_KEYWORD_PATTERN.test(value) ||
    /(대학교|대학원|대학|교육원|University|College|School|Institute|학과|학부|전공)/i.test(value)
  );
};

const attachContextToDegreeOnly = (record = '', context = '') => {
  if (!record) return '';

  const normalizedRecord = normalizeEducationRecord(record);
  if (!isStandaloneDegreeRecord(normalizedRecord)) {
    return normalizedRecord;
  }

  const degreeOnly = normalizeDegreeLabel(normalizedRecord);
  const normalizedContext = normalizeEducationRecord(context).replace(/[,:;/-]\s*$/g, '').trim();

  return normalizeEducationRecord(normalizedContext ? `${normalizedContext} ${degreeOnly}` : degreeOnly);
};

const extractBracketTaggedEducation = (line = '') => {
  const degreePattern = [...EDUCATION_DEGREE_KEYWORDS, '석박사 통합'].map(escapeRegex).join('|');
  const regex = new RegExp(
    `(?:^|[\\s,])(?:\\[|\\()\\s*(${degreePattern})\\s*(?:\\]|\\))\\s*([^\\[\\(]+?)(?=(?:\\s*(?:\\[|\\()\\s*(?:${degreePattern})\\s*(?:\\]|\\)))|$)`,
    'gi'
  );

  const matches = Array.from(normalizeEducationRecord(line).matchAll(regex));

  return uniq(
    matches
      .flatMap((match) => expandIntegratedDegree(match[2], normalizeDegreeLabel(match[1])))
      .filter(Boolean)
  );
};

const extractTaggedEducationSegments = (text = '') => {
  const degreePattern = [...EDUCATION_DEGREE_KEYWORDS, '석박사 통합'].map(escapeRegex).join('|');
  const regex = new RegExp(`(?:\\[|\\()\\s*(${degreePattern})\\s*(?:\\]|\\))\\s*([\\s\\S]*?)(?=(?:\\s*(?:\\[|\\()\\s*(?:${degreePattern})\\s*(?:\\]|\\)))|$)`, 'gi');
  const source = compactEducationSpacing(text);
  const matches = Array.from(source.matchAll(regex));

  return uniq(
    matches
      .flatMap((match) => {
        const body = normalizeEducationRecord(match[2]).replace(/[,:;/-]\s*$/g, '').trim();
        const degree = normalizeDegreeLabel(match[1]);
        return expandIntegratedDegree(body, degree);
      })
      .filter(Boolean)
  );
};

const cleanupEducationFragment = (segment = '') => {
  const normalizedSegment = normalizeEducationRecord(segment)
    .replace(/^(?:학점은행제?|학점은행)\s*(?:[,/·]|$)\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();

  if (!normalizedSegment) return '';

  const fragments = normalizedSegment
    .split(/\s*\/\s*/)
    .map((part) => normalizeEducationRecord(part))
    .filter(Boolean);

  if (!fragments.length) return normalizedSegment;

  const last = fragments[fragments.length - 1];
  const institutionPattern = /(대학교|대학원|대학|교육원|University|College|School|Institute)/i;

  if (institutionPattern.test(last)) {
    return last;
  }

  if (fragments.length >= 2 && DEGREE_PATTERN_GLOBAL.test(last) && !institutionPattern.test(last)) {
    const priorInstitution = [...fragments.slice(0, -1)].reverse().find((part) => institutionPattern.test(part));
    if (priorInstitution) {
      return normalizeEducationRecord(`${priorInstitution} ${last}`);
    }
  }

  return normalizeEducationRecord(fragments.slice(-2).join(' '));
};

const extractDegreeAnchoredSegments = (text = '') => {
  const degreePattern = [...EDUCATION_DEGREE_KEYWORDS, '석박사 통합'].map(escapeRegex).join('|');
  const degreeAnchoredRegex = new RegExp(`[^\\n,;]+?(?:${degreePattern})(?:\\s*(?:졸업|수료|재학|취득|예정))?`, 'gi');
  const source = prepareEducationSource(text);
  const seeds = source
    .split(/\n+/)
    .flatMap((line) => line.split(/\s*;\s*|\s*\|\s*/))
    .map((line) => normalizeEducationRecord(line))
    .filter(Boolean);

  const segments = [];

  seeds.forEach((seed) => {
    const matches = seed.match(degreeAnchoredRegex);
    if (matches?.length) {
      matches.forEach((match) => {
        const cleaned = cleanupEducationFragment(match);
        if (cleaned) segments.push(cleaned);
      });
      return;
    }

    const cleaned = cleanupEducationFragment(seed);
    if (cleaned && DEGREE_PATTERN_GLOBAL.test(cleaned)) segments.push(cleaned);
  });

  return uniq(segments);
};

const expandParentheticalDegrees = (line = '') => {
  const normalizedLine = normalizeEducationRecord(line);
  const degreePattern = [...EDUCATION_DEGREE_KEYWORDS, '석박사 통합'].map(escapeRegex).join('|');
  const match = normalizedLine.match(
    new RegExp(
      `^(.*?)(?:\\(|\\[)\\s*((?:${degreePattern})(?:\\s*,\\s*(?:${degreePattern}))*)\\s*(?:\\)|\\])\\s*$`,
      'i'
    )
  );

  if (!match) return [];

  const prefix = normalizeEducationRecord(match[1]).replace(/[,:;/-]\s*$/g, '').trim();
  const degrees = splitDegreeList(match[2]);
  if (!prefix || !degrees.length) return [];

  return uniq(
    degrees
      .flatMap((degree) => expandIntegratedDegree(prefix, degree))
      .filter(Boolean)
  );
};

const splitEducationParts = (line = '') => {
  const source = normalizeEducationRecord(line);
  if (!source) return [];

  const parts = [];
  let buffer = '';
  let roundDepth = 0;
  let squareDepth = 0;

  const flush = () => {
    const value = normalizeEducationRecord(buffer);
    if (value) parts.push(value);
    buffer = '';
  };

  const shouldSplitAfter = (index) => {
    const next = source.slice(index + 1).trim();
    if (!next) return false;

    return /^(?:[가-힣A-Za-z]{2,}(?:대학교|대학원|대학|교육원|University|College|School|Institute)|(?:\[|\()?\s*(?:학사|석사|박사|전문학사|박사수료|박사과정수료|석박사\s*통합|석박사통합))/.test(next);
  };

  for (let i = 0; i < source.length; i += 1) {
    const ch = source[i];

    if (ch === '(') {
      roundDepth += 1;
      buffer += ch;
      continue;
    }

    if (ch === ')') {
      roundDepth = Math.max(0, roundDepth - 1);
      buffer += ch;
      continue;
    }

    if (ch === '[') {
      squareDepth += 1;
      buffer += ch;
      continue;
    }

    if (ch === ']') {
      squareDepth = Math.max(0, squareDepth - 1);
      buffer += ch;
      continue;
    }

    if ((ch === '/' || ch === ',') && roundDepth === 0 && squareDepth === 0 && shouldSplitAfter(i)) {
      flush();
      continue;
    }

    buffer += ch;
  }

  flush();
  return parts.length ? parts : [source];
};

const splitEducationRecords = (text = '') => {
  const taggedSegments = extractTaggedEducationSegments(text);
  const anchoredSegments = extractDegreeAnchoredSegments(text);
  const fallbackSegments = splitBullets(prepareEducationSource(text))
    .flatMap((line) => line.split(/\n+/))
    .flatMap((line) => sanitizeBracketArtifacts(line).split(/\s*;\s*|\s*\|\s*/))
    .map((line) => normalizeEducationRecord(line))
    .filter(Boolean);

  const seedLines = uniq([...taggedSegments, ...anchoredSegments, ...fallbackSegments]).filter(Boolean);
  const records = [];

  seedLines.forEach((line) => {
    const normalizedLine = normalizeEducationRecord(line);
    if (!normalizedLine) return;

    if (!DEGREE_PATTERN_GLOBAL.test(normalizedLine) && !isMeaningfulEducationContext(normalizedLine)) {
      return;
    }

    const baseContext = extractEducationContext(normalizedLine);
    let runningContext = baseContext;

    const bracketTagged = extractBracketTaggedEducation(normalizedLine);
    if (bracketTagged.length) {
      bracketTagged.forEach((record) => {
        const normalizedRecord = attachContextToDegreeOnly(record, runningContext || baseContext);
        if (normalizedRecord && DEGREE_PATTERN_GLOBAL.test(normalizedRecord)) {
          records.push(normalizedRecord);
          const nextContext = extractEducationContext(normalizedRecord);
          if (nextContext) runningContext = nextContext;
        }
      });
      return;
    }

    const expanded = expandParentheticalDegrees(normalizedLine);
    if (expanded.length) {
      expanded.forEach((record) => {
        const normalizedRecord = attachContextToDegreeOnly(record, runningContext || baseContext);
        if (normalizedRecord && DEGREE_PATTERN_GLOBAL.test(normalizedRecord)) {
          records.push(normalizedRecord);
          const nextContext = extractEducationContext(normalizedRecord);
          if (nextContext) runningContext = nextContext;
        }
      });
      return;
    }

    const parts = splitEducationParts(normalizedLine);
    if (!parts.length) return;

    parts.forEach((part) => {
      const normalizedPart = normalizeEducationRecord(part);
      if (!normalizedPart) return;

      const nestedBracketTagged = extractBracketTaggedEducation(normalizedPart);
      if (nestedBracketTagged.length) {
        nestedBracketTagged.forEach((record) => {
          const normalizedRecord = attachContextToDegreeOnly(record, runningContext || baseContext);
          if (normalizedRecord && DEGREE_PATTERN_GLOBAL.test(normalizedRecord)) {
            records.push(normalizedRecord);
            const nextContext = extractEducationContext(normalizedRecord);
            if (nextContext) runningContext = nextContext;
          }
        });
        return;
      }

      const nestedExpanded = expandParentheticalDegrees(normalizedPart);
      if (nestedExpanded.length) {
        nestedExpanded.forEach((record) => {
          const normalizedRecord = attachContextToDegreeOnly(record, runningContext || baseContext);
          if (normalizedRecord && DEGREE_PATTERN_GLOBAL.test(normalizedRecord)) {
            records.push(normalizedRecord);
            const nextContext = extractEducationContext(normalizedRecord);
            if (nextContext) runningContext = nextContext;
          }
        });
        return;
      }

      if (isMeaningfulEducationContext(normalizedPart) && !DEGREE_PATTERN_GLOBAL.test(normalizedPart)) {
        runningContext = normalizedPart;
        return;
      }

      const normalizedRecord = attachContextToDegreeOnly(normalizedPart, runningContext || baseContext);

      if (normalizedRecord && DEGREE_PATTERN_GLOBAL.test(normalizedRecord)) {
        records.push(normalizedRecord);
        const nextContext = extractEducationContext(normalizedRecord);
        if (nextContext) runningContext = nextContext;
        return;
      }

      if (records.length && isMeaningfulEducationContext(normalizedPart)) {
        records[records.length - 1] = normalizeEducationRecord(`${records[records.length - 1]} ${normalizedPart}`);
        const nextContext = extractEducationContext(records[records.length - 1]);
        if (nextContext) runningContext = nextContext;
      }
    });
  });

  return uniq(
    records
      .map((record) => normalizeEducationRecord(record))
      .filter(Boolean)
      .filter((record) => DEGREE_PATTERN_GLOBAL.test(record))
  );
};

const extractHighestEducation = (records = []) => {
  if (!records.length) return '';

  let best = null;

  records.forEach((line, index) => {
    const candidate = {
      line: normalizeEducationRecord(line),
      rank: degreeRank(line),
      status: educationStatusWeight(line),
      contextScore: extractEducationContext(line) ? 2 : isStandaloneDegreeRecord(line) ? 0 : 1,
      index,
    };

    if (!best) {
      best = candidate;
      return;
    }

    if (
      candidate.rank > best.rank ||
      (candidate.rank === best.rank && candidate.status > best.status) ||
      (candidate.rank === best.rank && candidate.status === best.status && candidate.contextScore > best.contextScore) ||
      (candidate.rank === best.rank &&
        candidate.status === best.status &&
        candidate.contextScore === best.contextScore &&
        candidate.line.length > best.line.length) ||
      (candidate.rank === best.rank &&
        candidate.status === best.status &&
        candidate.contextScore === best.contextScore &&
        candidate.line.length === best.line.length &&
        candidate.index > best.index)
    ) {
      best = candidate;
    }
  });

  return best?.line || '';
};

const protectExpertiseTerms = (text = '') => {
  let value = decodeHTML(text);
  EXPERTISE_PROTECTED_TERMS.forEach((term) => {
    const pattern = new RegExp(escapeRegex(term), 'gi');
    value = value.replace(pattern, term.replaceAll('/', TOKEN_PLACEHOLDER));
  });
  return value;
};

const restoreProtectedToken = (text = '') => text.replaceAll(TOKEN_PLACEHOLDER, '/');

const normalizeExpertiseToken = (token = '') => {
  return restoreProtectedToken(
    safeText(token)
      .replace(EXPERTISE_LABEL_PATTERN, '')
      .replace(/^[-•■□▷▶*]+\s*/, '')
      .replace(/^\d+[.)]\s*/, '')
  )
    .replace(/\s+/g, ' ')
    .trim();
};

const extractExpertise = (text = '') => {
  const protectedText = protectExpertiseTerms(text);
  const lines = splitBullets(protectedText);
  if (!lines.length) return '';

  const tokens = uniq(
    lines
      .flatMap((line) => line.split(EXPERTISE_SPLIT_PATTERN))
      .map((token) => normalizeExpertiseToken(token))
      .filter((token) => token.length >= 2)
      .filter((token) => !CONTACT_NOISE_PATTERN.test(token))
  );

  return tokens.slice(0, 12).join(', ');
};

const AFFILIATION_FALSE_EDUCATION_PATTERN = /(학사|석사|박사|전문학사|최종학력|졸업|수료|재학|학위)/i;
const AFFILIATION_ORG_PATTERN = /(대학교|대학원|대학|연구원|연구소|센터|병원|재단|법인|회사|기업|공사|공단|기관|캠퍼스)/i;
const AFFILIATION_DEPARTMENT_PATTERN = /(학과|학부|전공|연구소|센터|본부|실|팀|과|처|부)/i;
const AFFILIATION_MEMBERSHIP_PATTERN = /((?:한국|대한|국제|미국|세계|IEEE|ICROS|KSME|KICS|ACS|AIChE)[^()]{0,30}?(?:학회|협회|위원회|연구회)|정회원|종신회원|상임이사|부회장|사업이사|자문위원|평가위원|심사위원|선정위원|운영위원)/i;

const AFFILIATION_POSITION_SOURCE = POSITION_HINTS.slice()
  .sort((a, b) => b.length - a.length)
  .map(escapeRegex)
  .join('|');

const AFFILIATION_POSITION_REGEX = new RegExp(`^(.+?(?:${AFFILIATION_POSITION_SOURCE}))(?=\\s|$|[,)])`, 'i');

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

  value = value.split(/\s+(?=(?:現|현재|현|前|전)\s+)/)[0].trim();

  const membershipMatch = value.match(AFFILIATION_MEMBERSHIP_PATTERN);
  if (membershipMatch && membershipMatch.index != null && membershipMatch.index > 0) {
    value = value.slice(0, membershipMatch.index).trim();
  }

  value = value
    .replace(/\(\s*\d{4}[^)]*?\)/g, ' ')
    .replace(
      /\d{4}(?:\s*년)?(?:[.]\d{1,2}|\s*년\s*\d{1,2}\s*월)?\s*[~∼〜-]\s*(?:\d{4}(?:\s*년)?(?:[.]\d{1,2}|\s*년\s*\d{1,2}\s*월)?|현재)\s*/g,
      ' '
    )
    .replace(/\(\s*(?:교수|정교수|부교수|조교수|대표|원장|센터장|소장|부장|팀장|실장)\s*\)$/i, ' ')
    .replace(/[|]+/g, ' ')
    .replace(/\s*[,/·]+\s*$/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

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

const normalizeCareerLine = (line = '') => {
  return cleanInline(line)
    .replace(CAREER_NOISE_PATTERN, '')
    .replace(/^[-•■□▷▶*]+\s*/, '')
    .replace(/\s+/g, ' ')
    .trim();
};

const isCareerStart = (line = '') => {
  const value = normalizeCareerLine(line);
  return CAREER_START_PATTERN.test(value) || (hasAffiliationHint(value) && hasPosition(value));
};

const splitCareerEntries = (text = '') => {
  const rawLines = sanitizeBracketArtifacts(text)
    .replace(/\s*[•■□▷▶*]+\s*/g, '\n')
    .replace(/\s*[－─]\s*/g, '\n')
    .split(/\n+/)
    .map((line) => normalizeCareerLine(line))
    .filter(Boolean);

  const entries = [];

  rawLines.forEach((line) => {
    if (!entries.length) {
      entries.push(line);
      return;
    }

    if (isCareerStart(line)) {
      entries.push(line);
      return;
    }

    entries[entries.length - 1] = `${entries[entries.length - 1]} ${line}`.replace(/\s+/g, ' ').trim();
  });

  return uniq(
    entries
      .map((entry) => entry.replace(/\s*;\s*/g, ' · ').trim())
      .filter((entry) => entry.length >= 3)
  );
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

const formatEducationDetails = (records = []) => {
  return records
    .map((line, index) => ({ line: normalizeEducationRecord(line), index }))
    .filter(({ line }) => Boolean(line))
    .sort((a, b) => degreeRank(a.line) - degreeRank(b.line) || a.index - b.index)
    .map(({ line }) => line)
    .join('\n');
};

const formatCareerSummary = (entries = []) => entries.slice(0, 6).join('\n');
const formatCareerDetails = (entries = []) => entries.slice(0, 15).join('\n');

const calcAge = (birth = '') => {
  const normalizedBirth = extractBirth(birth);
  const year = parseInt(String(normalizedBirth).slice(0, 4), 10);
  if (!year) return EMPTY_VALUE;
  return `${new Date().getFullYear() - year + 1}세`;
};

const findLabeledValue = (nodes, labels, options = {}) => {
  const { lookAhead = 8, validator = null } = options;
  const labelSet = labels.map((label) => normalize(label));

  for (let index = 0; index < nodes.length; index += 1) {
    const current = normalize(nodes[index]);
    if (!labelSet.includes(current)) continue;

    const joined = nodes.slice(index + 1, index + 1 + lookAhead).join(' ');
    if (validator) {
      const value = validator(joined);
      if (value) return value;
    } else if (joined.trim()) {
      return cleanInline(joined);
    }
  }

  return '';
};

const buildHeaderPattern = (headers = []) => {
  return headers
    .map((header) => header.replace(/\s+/g, '').split('').join('\\s*'))
    .join('|');
};

const findSectionBody = (allText, headers = [], stopHeaders = []) => {
  const stopPattern = buildHeaderPattern(stopHeaders);

  for (const header of headers) {
    const headerPattern = header.replace(/\s+/g, '').split('').join('\\s*');
    const regex = new RegExp(
      `${headerPattern}\\s*

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
const EDUCATION_KEYWORD_PATTERN = /(학사|석사|박사|전문학사|대학교|대학원|학위|졸업|수료|재학|학과|학부|전공|University|Department|School)/i;
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
const escapeRegex = (text = '') => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const normalizeDegreeLabel = (text = '') => {
  const value = cleanInline(text).replace(/\s+/g, '');
  if (/석박사통합/i.test(value)) return '석박사 통합';
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
      .filter((token) => EDUCATION_DEGREE_KEYWORDS.includes(token))
  );
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

const attachContextToDegreeOnly = (record = '', context = '') => {
  if (!record) return '';
  if (!DEGREE_ONLY_PATTERN.test(cleanInline(record).replace(/\s+/g, ''))) return normalizeEducationRecord(record);
  return normalizeEducationRecord(context ? `${context} ${normalizeDegreeLabel(record)}` : normalizeDegreeLabel(record));
};

const extractBracketTaggedEducation = (line = '') => {
  const degreePattern = EDUCATION_DEGREE_KEYWORDS.map(escapeRegex).join('|');
  const regex = new RegExp(`(?:^|[\\s,])(?:\\[|\\()\\s*(${degreePattern})\\s*(?:\\]|\\))\\s*([^\\[\\(]+?)(?=(?:\\s*(?:\\[|\\()\\s*(?:${degreePattern})\\s*(?:\\]|\\)))|$)`, 'gi');
  const matches = Array.from(normalizeEducationRecord(line).matchAll(regex));

  return uniq(
    matches
      .map((match) => {
        const degree = normalizeDegreeLabel(match[1]);
        const body = normalizeEducationRecord(match[2]);
        if (!body) return '';
        return normalizeEducationRecord(`${body} ${degree}`);
      })
      .filter(Boolean)
  );
};

const expandParentheticalDegrees = (line = '') => {
  const normalizedLine = normalizeEducationRecord(line);
  const degreePattern = EDUCATION_DEGREE_KEYWORDS.map(escapeRegex).join('|');
  const match = normalizedLine.match(new RegExp(`^(.*?)(?:\\(|\\[)\\s*((?:${degreePattern})(?:\\s*,\\s*(?:${degreePattern}))*)\\s*(?:\\)|\\])\\s*$`, 'i'));
  if (!match) return [];

  const prefix = normalizeEducationRecord(match[1]).replace(/[,:;/-]\s*$/g, '').trim();
  const degrees = splitDegreeList(match[2]);
  if (!prefix || !degrees.length) return [];

  return uniq(degrees.map((degree) => normalizeEducationRecord(`${prefix} ${degree}`)));
};

const splitEducationRecords = (text = '') => {
  const rawLines = splitBullets(text)
    .flatMap((line) => sanitizeBracketArtifacts(line).split(/\s*;\s*|\s*\|\s*/))
    .map((line) => normalizeEducationRecord(line))
    .filter(Boolean);

  const records = [];

  rawLines.forEach((line) => {
    const context = extractEducationContext(line);
    const bracketTagged = extractBracketTaggedEducation(line);
    if (bracketTagged.length) {
      records.push(...bracketTagged.map((record) => attachContextToDegreeOnly(record, context)));
      return;
    }

    const expanded = expandParentheticalDegrees(line);
    if (expanded.length) {
      records.push(...expanded.map((record) => attachContextToDegreeOnly(record, context)));
      return;
    }

    const parts = line
      .split(/\s*\/\s*(?=[^/]*(?:대학교|대학원|학사|석사|박사|전문학사|석박사))|\s*,\s*(?=[^,]*(?:대학교|대학원|학사|석사|박사|전문학사|석박사))/)
      .map((part) => normalizeEducationRecord(part))
      .filter(Boolean);

    if (!parts.length) return;

    parts.forEach((part) => {
      const expandedPart = expandParentheticalDegrees(part);
      if (expandedPart.length) {
        records.push(...expandedPart.map((record) => attachContextToDegreeOnly(record, context)));
        return;
      }

      const normalizedPart = attachContextToDegreeOnly(part, context);
      if (DEGREE_PATTERN_GLOBAL.test(normalizedPart)) {
        records.push(normalizedPart);
      } else if (records.length && EDUCATION_KEYWORD_PATTERN.test(normalizedPart)) {
        records[records.length - 1] = normalizeEducationRecord(`${records[records.length - 1]} ${normalizedPart}`);
      }
    });
  });

  return uniq(
    records
      .map((record) => normalizeEducationRecord(record))
      .filter(Boolean)
      .filter((record) => DEGREE_PATTERN_GLOBAL.test(record) || EDUCATION_KEYWORD_PATTERN.test(record))
  );
};

const extractHighestEducation = (text = '') => {
  const lines = splitEducationRecords(text);
  if (!lines.length) return '';

  let best = { line: lines[0], rank: degreeRank(lines[0]), status: educationStatusWeight(lines[0]), index: 0 };

  lines.forEach((line, index) => {
    const candidate = {
      line,
      rank: degreeRank(line),
      status: educationStatusWeight(line),
      index,
    };

    if (
      candidate.rank > best.rank ||
      (candidate.rank === best.rank && candidate.status > best.status) ||
      (candidate.rank === best.rank && candidate.status === best.status && candidate.line.length > best.line.length) ||
      (candidate.rank === best.rank && candidate.status === best.status && candidate.line.length === best.line.length && candidate.index > best.index)
    ) {
      best = candidate;
    }
  });

  return best.line;
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

const sanitizeAffiliation = (text = '') => {
  let value = cleanInline(text);
  if (!value) return '';

  value = truncateAtLooseKeyword(value, AFFILIATION_CUTOFF_KEYWORDS)
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
    .replace(/[|]+/g, ' ')
    .replace(/\s+-\s+.*$/, '')
    .replace(/\(\s*\d{4}[^)]*\)\s*$/g, '')
    .replace(/\d{4}(?:[.]\d{1,2})?\s*[~∼〜-]\s*(?:\d{4}(?:[.]\d{1,2})?|현재)\s*$/g, '')
    .replace(/\s*\/\s*(?:[^/]{0,20})?$/, (match) => (/\s+\/\s+/.test(match) ? '' : match))
    .replace(/\s{2,}/g, ' ')
    .trim();

  value = truncateAtLooseKeyword(value, AFFILIATION_CUTOFF_KEYWORDS)
    .replace(/^[\])}\s]+/, '')
    .replace(/[([{]\s*$/, '')
    .replace(/\\/g, ' ')
    .replace(/\s*[/,-]+\s*$/, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (!hasPosition(value) && value.includes('(')) {
    value = value.split('(')[0].trim();
  }

  return value;
};

const scoreAffiliationCandidate = (line = '', source = 'affiliation') => {
  const value = sanitizeAffiliation(line);
  if (!value) return -999;

  let score = 0;
  if (source === 'career-current') score += 5;
  if (hasCurrentMarker(line)) score += 4;
  if (hasAffiliationHint(value)) score += 3;
  if (hasPosition(value)) score += 4;
  if (value.length >= 6 && value.length <= 45) score += 2;
  if (value.length > 60) score -= 6;
  if (CONTACT_NOISE_PATTERN.test(value) || extractEmail(value) || extractPhone(value)) score -= 5;
  if (EDUCATION_KEYWORD_PATTERN.test(value)) score -= 3;
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
  const entries = splitCareerEntries(text)
    .map((line) => line.split(/\s+(?=(?:前|전|\(전\))\s*[)\-:：]?)/)[0])
    .map((line) => safeText(line))
    .filter(Boolean);

  const currentCandidates = entries
    .map((line, index) => ({ line, index }))
    .filter(({ line }) => hasCurrentMarker(line) || /현재/.test(line))
    .map(({ line, index }) => ({ value: sanitizeAffiliation(line), raw: line, index }))
    .filter(({ value }) => Boolean(value))
    .sort((a, b) => scoreAffiliationCandidate(b.raw, 'career-current') - scoreAffiliationCandidate(a.raw, 'career-current') || a.value.length - b.value.length || a.index - b.index);

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
    .sort((a, b) => b.score - a.score || a.value.length - b.value.length);

  return scored[0]?.value || '';
};

const formatCareer = (text = '') => splitCareerEntries(text).slice(0, 15).join('\n');

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
      `${headerPattern}\\s*[:：]?([\\s\\S]*?)(?=\\n\\s*(?:[●○•■□▷▶*-]\\s*)?(?:${stopPattern || '$^'})\\s*[:：]?|$)`,
      'i'
    );
    const match = allText.match(regex);
    if (match?.[1] && cleanInline(match[1]).length > 2) return decodeHTML(match[1]).trim();
  }

  return '';
};

const extractNodesAndText = async (zip) => {
  const slidePaths = Object.keys(zip.files)
    .filter((name) => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const allNodes = [];
  let allText = '';

  for (const path of slidePaths) {
    const file = zip.file(path);
    if (!file) continue;

    const xml = await file.async('string');
    const paragraphs = xml.split(/<a:p>/);

    for (const paragraph of paragraphs) {
      const matches = paragraph.match(/<a:t>([^<]*)<\/a:t>/g) || [];
      const text = matches
        .map((item) => item.replace(/<\/?a:t>/g, '').trim())
        .filter(Boolean)
        .join(' ')
        .trim();

      if (text) {
        allNodes.push(text);
        allText += `${text}\n`;
      }
    }
  }

  return { allNodes, allText };
};

export const createEmptyProfile = (fileName = '') => ({
  fileName,
  name: '',
  gender: '',
  birth: '',
  age: '',
  affiliation: '',
  phone: '',
  email: '',
  education: '',
  educationRaw: '',
  expertise: '',
  career: '',
  error: false,
});

export const parsePptxProfileInput = async (input, fileName = '') => {
  const fallback = createEmptyProfile(fileName);

  try {
    const zip = await JSZip.loadAsync(input);
    const { allNodes, allText } = await extractNodesAndText(zip);
    const joinedText = allNodes.join(' ');
    const row = createEmptyProfile(fileName);

    row.name = firstNonEmpty(
      findLabeledValue(allNodes, FIELD_LABELS.name, { lookAhead: 5, validator: refineName }),
      refineName(allText.match(/(?:위원\s*성명|성명|이름)\s*[:：]?\s*([가-힣\s]{2,8})/)?.[1]),
      refineName(fileName.split(/[_\-.\s]/)[0]),
      EMPTY_VALUE
    );

    row.gender = firstNonEmpty(
      findLabeledValue(allNodes, FIELD_LABELS.gender, { lookAhead: 4, validator: extractGender }),
      extractGender(allText),
      EMPTY_VALUE
    );

    const extractedBirth = firstNonEmpty(
      findLabeledValue(allNodes, FIELD_LABELS.birth, { lookAhead: 10, validator: extractBirth }),
      extractBirth(allText)
    );
    row.birth = firstNonEmpty(extractedBirth, EMPTY_VALUE);

    row.phone = firstNonEmpty(
      findLabeledValue(allNodes, FIELD_LABELS.phone, { lookAhead: 12, validator: extractPhone }),
      extractPhone(allText),
      EMPTY_VALUE
    );

    row.email = firstNonEmpty(
      findLabeledValue(allNodes, FIELD_LABELS.email, { lookAhead: 10, validator: extractEmail }),
      extractEmail(allText),
      EMPTY_VALUE
    );

    const careerBody = firstNonEmpty(
      findSectionBody(allText, FIELD_LABELS.career, SECTION_STOP_HEADERS.career),
      findLabeledValue(allNodes, FIELD_LABELS.career, { lookAhead: 24, validator: cleanInline })
    );

    const affiliationBody = firstNonEmpty(
      findSectionBody(allText, FIELD_LABELS.affiliation, SECTION_STOP_HEADERS.affiliation),
      findLabeledValue(allNodes, FIELD_LABELS.affiliation, { lookAhead: 10, validator: sanitizeAffiliation }),
      joinedText
    );
    row.affiliation = firstNonEmpty(chooseAffiliation(affiliationBody, careerBody), sanitizeAffiliation(affiliationBody), EMPTY_VALUE);

    const educationBody = firstNonEmpty(
      findSectionBody(allText, FIELD_LABELS.education, SECTION_STOP_HEADERS.education),
      findLabeledValue(allNodes, FIELD_LABELS.education, { lookAhead: 12, validator: cleanInline })
    );
    const educationRecords = splitEducationRecords(educationBody);
    row.educationRaw = educationRecords.join('\n');
    row.education = firstNonEmpty(extractHighestEducation(educationBody), educationRecords[0], EMPTY_VALUE);

    const expertiseBody = firstNonEmpty(
      findSectionBody(allText, FIELD_LABELS.expertise, SECTION_STOP_HEADERS.expertise),
      findLabeledValue(allNodes, FIELD_LABELS.expertise, { lookAhead: 10, validator: cleanInline })
    );
    row.expertise = firstNonEmpty(extractExpertise(expertiseBody), EMPTY_VALUE);

    row.career = firstNonEmpty(formatCareer(careerBody), EMPTY_VALUE);
    row.age = row.birth !== EMPTY_VALUE ? calcAge(row.birth) : EMPTY_VALUE;

    return row;
  } catch {
    return {
      ...fallback,
      name: '오류',
      error: true,
    };
  }
};

export const parsePptxProfile = async (file) => {
  return parsePptxProfileInput(file, file?.name || '');
};

export { EMPTY_VALUE };

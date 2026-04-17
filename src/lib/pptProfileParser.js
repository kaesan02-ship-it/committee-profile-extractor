import JSZip from 'jszip';
import {
  AFFILIATION_HINTS,
  CAREER_START_PATTERN,
  CURRENT_MARKERS,
  DEGREE_PATTERNS,
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
const DEGREE_PATTERN_GLOBAL = /(?:박사과정수료|박사수료|박사|석사|학사|전문학사|Ph\.?D|Doctor|MBA|M\.?A|M\.?S|B\.?A|B\.?S)/i;
const EDUCATION_NOISE_PATTERN = /^(?:학력|최종학력)\s*[:：]?/;
const CAREER_NOISE_PATTERN = /^(?:주요경력|주요 경력|주요이력|주요 이력|경력사항|경력)\s*[:：]?/;
const AFFILIATION_LABEL_PATTERN = /^(?:소속|현소속|현재소속|현직|근무처|소속및연락처|소속 및 연락처|소속 \/ 연락처)\s*[:：]?/;
const CURRENT_PREFIX_PATTERN = /^(?:현재|현|現|\(현\)|\[현\]|\{현\})\s*[-)\]:：.]*/;
const PREVIOUS_PREFIX_PATTERN = /^(?:전|前|\(전\))\s*[-)\]:：.]*/;
const DATE_RANGE_PREFIX_PATTERN = /^(?:\d{4}(?:[.]\d{1,2})?\s*[~∼〜-]\s*(?:\d{4}(?:[.]\d{1,2})?|현재)|\d{4}\s*년\s*~\s*(?:\d{4}\s*년|현재))\s*/;
const CONTACT_NOISE_PATTERN = /(연락처|휴대전화|휴대폰|이메일|E-mail|Email|메일)/i;
const EDUCATION_KEYWORD_PATTERN = /(학사|석사|박사|전문학사|대학교|대학원|학위|졸업|수료|재학)/;
const CAREER_SECTION_LINE_PATTERN = /(경력|이력|재직|근무|수행|프로젝트|담당|위원|심사|평가)/;
const EXPERTISE_LABEL_PATTERN = /^(?:전문분야|전 문 분야|전문 분야|주요분야|전공|핵심역량)\s*[:：]?/;
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
const cleanInline = (text = '') => decodeHTML(String(text)).replace(/\s+/g, ' ').trim();
const uniq = (arr = []) => [...new Set(arr.filter(Boolean))];
const safeText = (text = '') => cleanInline(text).replace(/^[-•■□▷▶*]+\s*/, '');
const firstNonEmpty = (...values) => values.map((value) => cleanInline(value)).find(Boolean) || '';
const escapeRegex = (text = '') => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const splitBullets = (text = '') => {
  return uniq(
    decodeHTML(text)
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

const extractBirth = (text = '') => {
  const match = text.match(/(\d{4})[\s./-년]*?(\d{1,2})[\s./-월]*?(\d{1,2})\s*일?/);
  if (match) {
    return `${match[1]}.${String(match[2]).padStart(2, '0')}.${String(match[3]).padStart(2, '0')}`;
  }

  const digits = text.replace(/\D/g, '');
  if (digits.length >= 8) {
    const value = digits.slice(0, 8);
    if (/^(19|20)\d{6}$/.test(value)) {
      return `${value.slice(0, 4)}.${value.slice(4, 6)}.${value.slice(6, 8)}`;
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

const normalizeEducationRecord = (line = '') => {
  return cleanInline(line)
    .replace(EDUCATION_NOISE_PATTERN, '')
    .replace(/^[-•■□▷▶*]+\s*/, '')
    .replace(/\s*,\s*/g, ', ')
    .replace(/\s*\(\s*/g, ' (')
    .replace(/\s*\)\s*/g, ') ')
    .replace(/\s+/g, ' ')
    .trim();
};

const splitEducationRecords = (text = '') => {
  const rawLines = splitBullets(text)
    .flatMap((line) => line.split(/\s*;\s*|\s*\|\s*/))
    .map((line) => normalizeEducationRecord(line))
    .filter(Boolean);

  const records = [];

  rawLines.forEach((line) => {
    const normalizedLine = line.replace(/\s{2,}/g, ' ');
    const matchedRecords = uniq(
      Array.from(normalizedLine.matchAll(/[^,\n/]*?(?:박사과정수료|박사수료|박사|석사|학사|전문학사)(?:\s*\([^)]*\))?/g))
        .map((match) => normalizeEducationRecord(match[0]))
        .filter(Boolean)
    );

    const parts = (matchedRecords.length ? matchedRecords : normalizedLine
      .split(/\s*\/\s*(?=[^/]*(?:대학교|대학원|학사|석사|박사|전문학사))/)
      .flatMap((item) => item.split(/\s*,\s*(?=[^,]*(?:대학교|대학원|학사|석사|박사|전문학사))/))
      .map((item) => normalizeEducationRecord(item))
      .filter(Boolean));

    parts.forEach((part) => {
      if (!part) return;

      if (DEGREE_PATTERN_GLOBAL.test(part)) {
        records.push(part);
      } else if (records.length) {
        records[records.length - 1] = normalizeEducationRecord(`${records[records.length - 1]} ${part}`);
      }
    });
  });

  return uniq(records);
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
      (candidate.rank === best.rank && candidate.status === best.status && candidate.index > best.index) ||
      (candidate.rank === best.rank && candidate.status === best.status && candidate.index === best.index && candidate.line.length > best.line.length)
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
    .replace(/[|]+/g, ' ')
    .replace(/\s+-\s+.*$/, '')
    .replace(/\(\s*(?:\d{4}|현재)[^)]*$/g, '')
    .replace(/\s*\/\s*(?:[^/]{0,20})?$/, (match) => (/\s+\/\s+/.test(match) ? '' : match))
    .replace(/\s{2,}/g, ' ')
    .trim();

  value = value.replace(/\s*\((?:[^)]*?(?:현재|\d{4}\s*년|\d{4}[.]\d{1,2})[^)]*)\)\s*$/g, '').trim();
  if (!hasPosition(value) && value.includes('(')) {
    value = value.split('(')[0].trim();
  }
  value = value.replace(/\s*[/,-]+\s*$/, '').trim();
  return value;
};

const scoreAffiliationCandidate = (line = '', source = 'affiliation') => {
  const value = sanitizeAffiliation(line);
  if (!value) return -999;

  let score = 0;
  if (source === 'career-current') score += 4;
  if (hasCurrentMarker(line)) score += 4;
  if (hasAffiliationHint(value)) score += 3;
  if (hasPosition(value)) score += 3;
  if (/소속|근무|재직/.test(value)) score += 2;
  if (value.length >= 6 && value.length <= 40) score += 1;
  if (value.length > 60) score -= 5;

  if (CONTACT_NOISE_PATTERN.test(value) || extractEmail(value) || extractPhone(value)) score -= 4;
  if (EDUCATION_KEYWORD_PATTERN.test(value)) score -= 3;
  if (CAREER_SECTION_LINE_PATTERN.test(value) && source !== 'career-current') score -= 1;
  if (/\d{4}/.test(value)) score -= 2;

  return score;
};

const extractCurrentCareerLine = (text = '') => {
  const entries = splitCareerEntries(text)
    .map((line) => line.split(/\s+(?=(?:前|전|\(전\))\s*[)\-:：]?)/)[0])
    .map((line) => safeText(line))
    .filter(Boolean);

  const currentCandidates = entries
    .map((line, index) => ({ line, index }))
    .filter(({ line }) => hasCurrentMarker(line) || /현재/.test(line))
    .map(({ line, index }) => ({ value: sanitizeAffiliation(line), index }))
    .filter(({ value }) => Boolean(value))
    .sort((a, b) => scoreAffiliationCandidate(b.value, 'career-current') - scoreAffiliationCandidate(a.value, 'career-current') || b.index - a.index);

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

  const scored = uniq(candidates.map((item) => `${item.source}::${item.value}`)).map((key) => {
    const [source, value] = key.split('::');
    return { value, source, score: scoreAffiliationCandidate(value, source) };
  }).sort((a, b) => b.score - a.score || a.value.length - b.value.length);

  return scored[0]?.value || '';
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
  return CAREER_START_PATTERN.test(value);
};

const splitCareerEntries = (text = '') => {
  const rawLines = decodeHTML(text)
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

const formatCareer = (text = '') => splitCareerEntries(text).slice(0, 15).join('\n');

const calcAge = (birth = '') => {
  const year = parseInt(String(birth).slice(0, 4), 10);
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

    row.birth = firstNonEmpty(
      findLabeledValue(allNodes, FIELD_LABELS.birth, { lookAhead: 10, validator: extractBirth }),
      extractBirth(allText),
      EMPTY_VALUE
    );

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
    row.educationRaw = splitEducationRecords(educationBody).join('\n');
    row.education = firstNonEmpty(extractHighestEducation(educationBody), row.educationRaw.split('\n')[0], EMPTY_VALUE);

    const expertiseBody = firstNonEmpty(
      findSectionBody(allText, FIELD_LABELS.expertise, SECTION_STOP_HEADERS.expertise),
      findLabeledValue(allNodes, FIELD_LABELS.expertise, { lookAhead: 10, validator: cleanInline })
    );
    row.expertise = firstNonEmpty(extractExpertise(expertiseBody), EMPTY_VALUE);

    row.career = firstNonEmpty(formatCareer(careerBody), EMPTY_VALUE);
    row.age = calcAge(row.birth);

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

import { EMPTY_VALUE } from './extractionRules.js';

const normalizeText = (value = '') => {
  const text = String(value ?? '').trim();
  return text && text !== EMPTY_VALUE ? text : EMPTY_VALUE;
};

const isEmptyValue = (value = '') => {
  const text = String(value ?? '').trim();
  return !text || text === EMPTY_VALUE;
};

const normalizeMultiline = (value = '') => {
  const text = normalizeText(value);
  if (text === EMPTY_VALUE) return EMPTY_VALUE;
  return text
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .join('\n');
};

const toLines = (value = '') => String(value ?? '')
  .split(/\n+/)
  .map((line) => line.trim())
  .filter((line) => line && line !== EMPTY_VALUE);

const unique = (arr = []) => [...new Set(arr.map((item) => item.trim()).filter(Boolean))];

const bulletizeMultiline = (value = '') => {
  const text = normalizeMultiline(value);
  if (text === EMPTY_VALUE) return EMPTY_VALUE;
  return text
    .split(/\n+/)
    .map((line) => `• ${line}`)
    .join('\n');
};

const cleanupEntry = (value = '') => String(value ?? '')
  .replace(/\s+/g, ' ')
  .replace(/\s*([~∼〜-])\s*/g, '$1')
  .replace(/\(\s+/g, '(')
  .replace(/\s+\)/g, ')')
  .replace(/\s+([,.;:：])/g, '$1')
  .trim();

const stripEvaluationNoise = (value = '') => cleanupEntry(value)
  .replace(/\s*\((?:서류|면접|프로젝트|평가|심사|자문|채용|,|\s)+등?\)\s*$/gi, '')
  .replace(/\s*(?:면접관\s*)?경력\s*(?:서류|면접)?\s*[:：]?\s*$/gi, '')
  .trim();

const stripEvaluationSections = (value = '') => String(value ?? '')
  .replace(/\s*(?:면접관\s*)?경력\s*(?:서류|면접)\s*[:：][\s\S]*$/i, '')
  .replace(/\s+(?:서류평가|서류전형|서류|채용면접|면접전형|면접)\s*[:：][\s\S]*$/i, '')
  .trim();

const CAREER_MARKER_LOOKAHEAD = /(?=(?:現|前)\s*[).:：-]?\s*|(?:현재|현|전)\s*(?:[).:：-]\s*|\s+))/g;

const splitCareerEntries = (value = '') => {
  const text = normalizeMultiline(stripEvaluationSections(value));
  if (text === EMPTY_VALUE) return [];

  return unique(
    text
      .replace(new RegExp(`\\s+${CAREER_MARKER_LOOKAHEAD.source}`, 'g'), '\n')
      .replace(/\s+(?=\(?\s*\d{4}\s*[.]\s*\d{1,2}(?:\s*[.]\s*\d{1,2})?\s*[~∼〜-]\s*(?:\d{4}|현재|現))/g, '\n')
      .replace(/\s+(?=\(?\s*\d{4}\s*[.]\s*\d{1,2}(?:\s*[.]\s*\d{1,2})?\s*[~∼〜-]\s*\d{4})/g, '\n')
      .replace(/\s+(?=\d{4}\s*년\s*\d{1,2}\s*월\s*[~∼〜-])/g, '\n')
      .split(/\n+/)
      .map(stripEvaluationNoise)
      .filter((line) => line.length >= 3)
      .filter((line) => !isEvaluationOnlyLine(line))
  );
};

const hasCurrentMarker = (line = '') => /^(?:現\s*[).:：-]?|현재\s*(?:[).:：-]|\s)|현\s*(?:[).:：-]|\s))|[~∼〜-]\s*(?:현재|現)(?:\)|\s|$)/.test(line);

const normalizeCareerLine = (line = '', fallbackPrefix = '前') => {
  const cleaned = stripEvaluationNoise(line);
  if (!cleaned) return '';

  if (/^(?:現\s*[).:：-]?|현재\s*(?:[).:：-]|\s)|현\s*(?:[).:：-]|\s))/.test(cleaned)) {
    return cleaned.replace(/^(?:現|현재|현)\s*[).:：-]?\s*/i, '現) ');
  }

  if (/^(?:前\s*[).:：-]?|전\s*(?:[).:：-]|\s))/.test(cleaned)) {
    return cleaned.replace(/^(?:前|전)\s*[).:：-]?\s*/i, '前) ');
  }

  if (hasCurrentMarker(cleaned)) return `現) ${cleaned}`;
  return `${fallbackPrefix}) ${cleaned}`;
};

const isEvaluationOnlyLine = (line = '') => {
  if (/^\[?\s*(서류|면접|채용면접|서류평가|면접전형|심사|자문)\s*\]?[:：]/.test(line)) return true;
  if (/^(?:면접관\s*)?경력\s*(?:서류|면접)\s*[:：]/.test(line)) return true;
  return false;
};

export const formatEducationForTemplate = (educationDetails = '', education = '') => {
  const source = toLines(educationDetails).length ? educationDetails : education;
  if (isEmptyValue(source)) return '';
  return bulletizeMultiline(source).replaceAll('•', '●');
};

export const formatCareerForTemplate = (row = {}) => {
  const source = row.careerDetails || row.careerRaw || row.career;
  const careerLines = splitCareerEntries(source);
  const lines = [];
  const currentLine = careerLines.find((line) => hasCurrentMarker(line));
  const previousLines = careerLines.filter((line) => line !== currentLine);

  if (currentLine) {
    lines.push(normalizeCareerLine(currentLine, '現'));
  } else if (row.affiliation && row.affiliation !== EMPTY_VALUE) {
    lines.push(`現) ${cleanupEntry(row.affiliation)}`);
  }

  previousLines.slice(0, 3).forEach((line) => {
    lines.push(normalizeCareerLine(line, '前'));
  });

  return unique(lines).slice(0, 4).join('\n');
};

const extractEvaluationBlocks = (text = '') => {
  const source = normalizeMultiline(text).replace(/\n+/g, ' ');
  if (source === EMPTY_VALUE) return [];

  const labelRegex = /(?:\[?\s*(서류평가|서류전형|서류|채용면접|면접전형|면접|심사|자문)\s*\]?|(?:면접관\s*)?경력\s*(서류|면접))\s*[:：]/g;
  const matches = [...source.matchAll(labelRegex)];

  return matches.map((match, index) => {
    const rawLabel = match[1] || match[2] || '';
    const nextIndex = matches[index + 1]?.index ?? source.length;
    const body = source.slice((match.index ?? 0) + match[0].length, nextIndex);
    const label = /서류/.test(rawLabel) ? '서류' : /면접/.test(rawLabel) ? '면접' : rawLabel;
    return { label, body: cleanupEvaluationBody(body) };
  }).filter((item) => item.body);
};

const cleanupEvaluationBody = (value = '') => cleanupEntry(value)
  .replace(/^(?:외\s*)?다수[,.]?\s*/i, '')
  .replace(/\s*(?:기타|자격|논문|주요이력|주요실적|수행실적)\s*.*$/i, '')
  .replace(/\s{2,}/g, ' ')
  .replace(/\s*[,/]+\s*$/g, '')
  .trim();

const compactEvaluationList = (value = '') => {
  const cleaned = cleanupEvaluationBody(value);
  if (!cleaned) return '';

  const items = cleaned
    .split(/\s*,\s*/)
    .map((item) => item.trim())
    .filter(Boolean);

  if (items.length >= 5) return `${items.slice(0, 4).join(', ')} 등`;
  if (cleaned.length <= 180) return cleaned;
  return `${cleaned.slice(0, 177).trim()}...`;
};

export const formatEvaluationCareerForTemplate = (row = {}) => {
  const baseText = String(row.careerRaw || row.careerDetails || row.career || '');
  if (isEmptyValue(baseText)) return '';

  const groups = new Map();
  extractEvaluationBlocks(baseText).forEach(({ label, body }) => {
    const key = label === '서류' ? '서류' : label === '면접' ? '면접' : label;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(compactEvaluationList(body));
  });

  const lines = [];
  ['서류', '면접', '심사', '자문'].forEach((label) => {
    const values = unique(groups.get(label) || []);
    if (values.length) lines.push(`[${label}] ${values.join(', ')}`);
  });

  return lines.join('\n');
};

export const __testing = {
  compactEvaluationList,
  extractEvaluationBlocks,
  splitCareerEntries,
};

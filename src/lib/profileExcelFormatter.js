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

  const labelRegex = /(?:\[\s*(서류평가|서류전형|서류|채용면접|면접평가|면접전형|면접|채용심사|심사|자문)\s*\]\s*|(?:서류평가|서류전형|서류|채용면접|면접평가|면접전형|면접|채용심사|심사|자문)\s*[:：]\s*|(?:면접관\s*)?경력\s*(서류|면접)\s*[:：]?\s*|(?:면접경력|서류평가\s*경력)\s*[:：]?\s*|(?:서류전형|면접전형|채용면접)\s*[-–]\s*|(?:^|\s)(서류|면접)\s+(?!평가|전형|심사|관련|과제|교육)(?=[가-힣A-Z0-9][^:：]{0,50}(?:은행|공단|공사|재단|진흥원|관리원|보험공사|보증공사|보증기금|연구원|병원|청|부|시청|구청|KDB|KB|IBK|LH|KOTRA|KOICA)))/g;
  const matches = [...source.matchAll(labelRegex)];

  return matches.map((match, index) => {
    const rawLabel = match[1] || match[2] || match[3] || match[0] || '';
    const nextIndex = matches[index + 1]?.index ?? source.length;
    const body = source.slice((match.index ?? 0) + match[0].length, nextIndex);
    const label = /서류/.test(rawLabel) ? '서류' : /면접/.test(rawLabel) ? '면접' : /심사/.test(rawLabel) ? '심사' : rawLabel;
    return { label, body: cleanupEvaluationBody(body) };
  }).filter((item) => item.body);
};

const cleanupEvaluationBody = (value = '') => cleanupEntry(value)
  .replace(/^\s*[-–]\s*/, '')
  .replace(/^\[\s*(?:금융|공공기관|공공|공기업|사기업|민간|대학교|대학|기타)\s*[-:：]?\s*/, '')
  .replace(/^(?:외\s*)?다수[,.]?\s*/i, '')
  .replace(/(KDB\s*산업은행)산업은행/g, '$1')
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

const extractGenericEvaluationActivities = (texts = []) => {
  const genericPattern = /(채용\s*평가\s*위원|채용평가위원|채용\s*심사|채용심사|채용\s*전문\s*면접관|채용전문면접관|전문\s*면접관|전문면접관|공공기관[^,]{0,20}면접관|채용컨설팅\s*및\s*면접관|면접\s*평가\s*위원|면접\s*심사|서류\s*평가|평가위원)/;

  const candidates = texts
    .flatMap((text) => String(text ?? '')
        .replace(/\n+/g, ' ')
        .split(/\s+(?=(?:現|前|현|전)\s)|\s*[•■□▷▶*]\s*/))
    .map((line) => cleanupEntry(line).replace(/^(?:및\s*)?(?:주요이력|주요경력|경력사항)\s*/i, ''))
    .filter((line) => line && genericPattern.test(line))
    .filter((line) => !/(?:면접관련|면접관\s*Profile|채용전문가\s*\d*\s*급?\s*과정\s*수료|과정\s*수료)/i.test(line))
    .map((line) => {
      if (line.length <= 120) return line;
      const match = line.match(genericPattern);
      const index = match?.index ?? 0;
      const start = Math.max(0, index - 45);
      const end = Math.min(line.length, index + 90);
      return `${start > 0 ? '...' : ''}${line.slice(start, end).trim()}${end < line.length ? '...' : ''}`;
    });

  const seen = [];
  return candidates.sort((a, b) => a.length - b.length).filter((candidate) => {
    const key = candidate.replace(/^\.\.\./, '').replace(/\.\.\.$/, '').replace(/[^\p{L}\p{N}]+/gu, '');
    if (seen.some((seenKey) => seenKey.includes(key) || key.includes(seenKey))) return false;
    seen.push(key);
    return true;
  }).slice(0, 4);
};

export const formatEvaluationCareerForTemplate = (row = {}) => {
  const sourceTexts = unique([
    row.evaluationRaw,
    row.careerRaw,
    row.careerDetails,
    row.career,
  ].map((value) => String(value ?? '')).filter((value) => !isEmptyValue(value)));
  if (!sourceTexts.length) return '';

  const groups = new Map();
  sourceTexts.forEach((sourceText) => {
    extractEvaluationBlocks(sourceText).forEach(({ label, body }) => {
      const key = label === '서류' ? '서류' : label === '면접' ? '면접' : label;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(compactEvaluationList(body));
    });
  });

  if (!groups.size) {
    const genericActivities = extractGenericEvaluationActivities(sourceTexts);
    if (genericActivities.length) groups.set('심사', genericActivities);
  }

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

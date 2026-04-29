import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { parsePptxProfileInput } from '../src/lib/pptProfileParser.js';

const EMPTY_VALUE = '미상';
const DEFAULT_OUT_DIR = 'reports/latest';
const COMPARE_FIELDS = ['affiliation', 'phone', 'education', 'gender'];

const OUTPUT_COLUMNS = [
  'fileName',
  'sourceRelativePath',
  'reviewTags',
  'name',
  'gender',
  'birth',
  'age',
  'affiliation',
  'phone',
  'email',
  'education',
  'educationDetails',
  'expertise',
  'career',
  'careerDetails',
  'error',
];

const FIELD_ALIASES = {
  fileName: ['fileName', '파일명', '원본파일', '원본 파일', 'sourceRelativePath', '원본상대경로'],
  sourceRelativePath: ['sourceRelativePath', '원본상대경로', '상대경로', '경로'],
  reviewTags: ['reviewTags', '검수태그', '태그'],
  name: ['name', '성명', '위원 성명', '위원성명', '이름'],
  gender: ['gender', '성별'],
  birth: ['birth', '출생년월일', '생년월일'],
  affiliation: ['affiliation', '현소속', '소속', '현재소속'],
  phone: ['phone', '연락처', '전화번호', '휴대폰'],
  email: ['email', '이메일', '이메일주소', '메일주소'],
  education: ['education', '최종학력', '학력'],
  educationDetails: ['educationDetails', '학력상세', '학력 정리'],
  expertise: ['expertise', '전문분야', '분야'],
  career: ['career', '경력요약', '주요경력'],
  careerDetails: ['careerDetails', '경력상세', '경력 정리'],
};

const isUrl = (value = '') => /^https?:\/\//i.test(value);

const parseArgs = (argv) => {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith('--')) continue;
    const key = token.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith('--')) {
      args[key] = true;
      continue;
    }
    args[key] = next;
    i += 1;
  }
  return args;
};

const usage = () => `
Usage:
  npm run batch:parse -- --input sample-ppts.zip --out reports/run
  npm run batch:parse -- --input sample-ppts.zip --ground-truth ground_truth.csv --before current_results.xlsx

Options:
  --input          PPTX, ZIP, directory, or public download URL
  --out           Output directory (default: ${DEFAULT_OUT_DIR})
  --ground-truth  Optional CSV/XLSX/JSON with verified values
  --before        Optional CSV/XLSX/JSON from the previous extraction run
  --limit         Optional max number of PPTX files to parse
`;

const readBuffer = async (target) => {
  if (isUrl(target)) {
    const response = await fetch(target);
    if (!response.ok) throw new Error(`Download failed ${response.status}: ${target}`);
    return Buffer.from(await response.arrayBuffer());
  }
  return fs.readFile(target);
};

const readText = async (target) => {
  if (isUrl(target)) {
    const response = await fetch(target);
    if (!response.ok) throw new Error(`Download failed ${response.status}: ${target}`);
    return response.text();
  }
  return fs.readFile(target, 'utf8');
};

const getExt = (target = '') => {
  const clean = String(target).split('?')[0].split('#')[0];
  return path.extname(clean).toLowerCase();
};

const walkDir = async (dir) => {
  const entries = await fs.readdir(dir, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...await walkDir(fullPath));
    } else {
      files.push(fullPath);
    }
  }
  return files;
};

const collectPptxInputs = async (input) => {
  if (isUrl(input)) {
    const buffer = await readBuffer(input);
    return collectPptxFromBuffer(buffer, input);
  }

  const stat = await fs.stat(input);
  if (stat.isDirectory()) {
    const files = (await walkDir(input))
      .filter((name) => name.toLowerCase().endsWith('.pptx'))
      .filter((name) => !path.basename(name).startsWith('~$'))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    return Promise.all(files.map(async (filePath) => ({
      fileName: path.basename(filePath),
      sourceRelativePath: path.relative(input, filePath).replaceAll('\\', '/'),
      buffer: await fs.readFile(filePath),
    })));
  }

  const buffer = await fs.readFile(input);
  return collectPptxFromBuffer(buffer, input);
};

const collectPptxFromBuffer = async (buffer, sourceName) => {
  const ext = getExt(sourceName);
  if (ext === '.pptx') {
    return [{
      fileName: path.basename(sourceName),
      sourceRelativePath: path.basename(sourceName),
      buffer,
    }];
  }

  if (ext !== '.zip') {
    throw new Error(`Unsupported input type: ${sourceName}`);
  }

  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files)
    .filter((name) => name.toLowerCase().endsWith('.pptx'))
    .filter((name) => !name.split('/').pop().startsWith('~$'))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const sourceRoot = path.basename(String(sourceName).split('?')[0], '.zip');
  return Promise.all(names.map(async (name) => ({
    fileName: name.split('/').pop(),
    sourceRelativePath: `${sourceRoot}/${name}`,
    buffer: await zip.file(name).async('nodebuffer'),
  })));
};

const csvEscape = (value = '') => {
  const text = Array.isArray(value) ? value.join(', ') : String(value ?? '');
  if (/[",\n\r]/.test(text)) return `"${text.replaceAll('"', '""')}"`;
  return text;
};

const writeCsv = async (filePath, rows, columns = OUTPUT_COLUMNS) => {
  const lines = [
    columns.join(','),
    ...rows.map((row) => columns.map((column) => csvEscape(row[column])).join(',')),
  ];
  await fs.writeFile(filePath, `${lines.join('\n')}\n`, 'utf8');
};

const parseCsv = (text = '') => {
  const rows = [];
  let row = [];
  let cell = '';
  let quoted = false;

  const pushCell = () => {
    row.push(cell);
    cell = '';
  };

  const pushRow = () => {
    if (row.length || cell) {
      pushCell();
      rows.push(row);
      row = [];
    }
  };

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];

    if (quoted) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') {
        quoted = false;
      } else {
        cell += ch;
      }
      continue;
    }

    if (ch === '"') {
      quoted = true;
    } else if (ch === ',') {
      pushCell();
    } else if (ch === '\n') {
      pushRow();
    } else if (ch !== '\r') {
      cell += ch;
    }
  }

  pushRow();
  if (!rows.length) return [];

  const headers = rows[0].map((header) => header.trim());
  return rows.slice(1).map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ''])));
};

const pick = (source, aliases = []) => {
  for (const key of aliases) {
    if (source[key] != null && String(source[key]).trim() !== '') return source[key];
  }
  return '';
};

const normalizeRow = (row = {}) => {
  const normalized = {};
  for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
    normalized[field] = pick(row, aliases);
  }
  return {
    ...row,
    ...normalized,
    reviewTags: Array.isArray(normalized.reviewTags)
      ? normalized.reviewTags
      : String(normalized.reviewTags || '').split(/\s*,\s*/).filter(Boolean),
  };
};

const loadRows = async (target) => {
  if (!target) return [];
  const ext = getExt(target);

  if (ext === '.json') {
    const rows = JSON.parse(await readText(target));
    return (Array.isArray(rows) ? rows : []).map(normalizeRow);
  }

  if (ext === '.xlsx' || ext === '.xls') {
    const workbook = XLSX.read(await readBuffer(target), { type: 'buffer' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(firstSheet, { defval: '' }).map(normalizeRow);
  }

  return parseCsv(await readText(target)).map(normalizeRow);
};

const rowKey = (row = {}) => String(row.sourceRelativePath || row.fileName || '').trim();

const indexRows = (rows = []) => {
  const index = new Map();
  rows.forEach((row) => {
    const primary = rowKey(row);
    if (primary) index.set(primary, row);
    if (row.fileName) index.set(String(row.fileName).trim(), row);
  });
  return index;
};

const isEmpty = (value = '') => {
  const text = String(value ?? '').trim();
  return !text || text === EMPTY_VALUE;
};

const normalizeForCompare = (value = '', field = '') => {
  if (isEmpty(value)) return '';
  const text = String(value ?? '').trim();
  if (field === 'phone') return text.replace(/\D/g, '');
  if (field === 'gender') return text.replace(/\s/g, '');
  return text.toLowerCase().replace(/[\s()[\]{}.,·:：;/-]/g, '');
};

const isMatch = (actual = '', expected = '', field = '') => {
  const a = normalizeForCompare(actual, field);
  const e = normalizeForCompare(expected, field);
  return Boolean(a && e && a === e);
};

const precisionByField = (rows, truthIndex) => {
  const stats = Object.fromEntries(COMPARE_FIELDS.map((field) => [field, { correct: 0, predicted: 0, precision: null }]));

  rows.forEach((row) => {
    const truth = truthIndex.get(rowKey(row)) || truthIndex.get(row.fileName);
    if (!truth) return;

    COMPARE_FIELDS.forEach((field) => {
      const predicted = row[field];
      const expected = truth[field];
      if (isEmpty(predicted)) return;

      stats[field].predicted += 1;
      if (isMatch(predicted, expected, field)) stats[field].correct += 1;
    });
  });

  Object.values(stats).forEach((item) => {
    item.precision = item.predicted ? item.correct / item.predicted : null;
  });

  return stats;
};

const pct = (value) => (value == null ? 'n/a' : `${(value * 100).toFixed(1)}%`);

const tagCounts = (rows = []) => {
  const counts = new Map();
  rows.flatMap((row) => row.reviewTags || []).forEach((tag) => {
    counts.set(tag, (counts.get(tag) || 0) + 1);
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1]);
};

const buildRepresentativeRows = (afterRows, beforeIndex, truthIndex) => {
  const rows = [];
  const hasBefore = beforeIndex.size > 0;

  afterRows.forEach((after) => {
    const before = beforeIndex.get(rowKey(after)) || beforeIndex.get(after.fileName) || {};
    const truth = truthIndex.get(rowKey(after)) || truthIndex.get(after.fileName) || {};

    COMPARE_FIELDS.forEach((field) => {
      const changed = hasBefore && normalizeForCompare(before[field], field) !== normalizeForCompare(after[field], field);
      const failed = truth[field] && !isMatch(after[field], truth[field], field);
      if (!changed && !failed) return;

      rows.push({
        fileName: after.fileName,
        field,
        expected: truth[field] || '',
        before: before[field] || '',
        after: after[field] || '',
        reviewTags: (after.reviewTags || []).join(', '),
      });
    });
  });

  return rows.slice(0, 20);
};

const markdownTable = (rows, columns) => {
  if (!rows.length) return '_No representative cases available._';
  const header = `| ${columns.join(' |')} |`;
  const divider = `| ${columns.map(() => '---').join(' |')} |`;
  const body = rows.map((row) => `| ${columns.map((column) => String(row[column] ?? '').replace(/\n/g, '<br>').replace(/\|/g, '\\|')).join(' |')} |`);
  return [header, divider, ...body].join('\n');
};

const buildReport = ({ input, outDir, rows, beforeRows, truthRows }) => {
  const beforeIndex = indexRows(beforeRows);
  const truthIndex = indexRows(truthRows);
  const hasTruth = truthRows.length > 0;
  const hasBefore = beforeRows.length > 0;
  const afterPrecision = hasTruth ? precisionByField(rows, truthIndex) : null;
  const beforePrecision = hasTruth && hasBefore ? precisionByField(beforeRows, truthIndex) : null;

  const lines = [
    '# PPT 추출 비교 리포트',
    '',
    `- Input: ${input}`,
    `- Parsed files: ${rows.length}`,
    `- Output directory: ${outDir}`,
    `- Generated at: ${new Date().toISOString()}`,
    '',
    '## 검수 태그',
    '',
    markdownTable(tagCounts(rows).map(([tag, count]) => ({ tag, count })), ['tag', 'count']),
    '',
    '## Precision',
    '',
  ];

  if (!hasTruth) {
    lines.push('ground_truth 파일이 없어 precision은 계산하지 않았습니다.');
  } else {
    const precisionRows = COMPARE_FIELDS.map((field) => {
      const before = beforePrecision?.[field];
      const after = afterPrecision[field];
      const delta = before?.precision == null || after.precision == null
        ? 'n/a'
        : `${((after.precision - before.precision) * 100).toFixed(1)}pp`;
      return {
        field,
        before: before ? `${before.correct}/${before.predicted} (${pct(before.precision)})` : 'n/a',
        after: `${after.correct}/${after.predicted} (${pct(after.precision)})`,
        delta,
      };
    });
    lines.push(markdownTable(precisionRows, ['field', 'before', 'after', 'delta']));
  }

  lines.push('', '## 대표 실패/변경 케이스 20건', '');
  if (!hasBefore && !hasTruth) {
    lines.push('before 결과나 ground_truth 파일이 없어 대표 실패/변경 케이스를 만들 수 없습니다.');
  } else {
    lines.push(markdownTable(buildRepresentativeRows(rows, beforeIndex, truthIndex), ['fileName', 'field', 'expected', 'before', 'after', 'reviewTags']));
  }

  return `${lines.join('\n')}\n`;
};

const flattenRow = (row) => ({
  ...row,
  reviewTags: (row.reviewTags || []).join(', '),
  educationDetails: row.educationDetails || '',
  careerDetails: row.careerDetails || '',
  error: row.error ? 'true' : 'false',
});

const main = async () => {
  const args = parseArgs(process.argv.slice(2));
  if (!args.input || args.help) {
    console.log(usage());
    process.exit(args.help ? 0 : 1);
  }

  const outDir = args.out || DEFAULT_OUT_DIR;
  await fs.mkdir(outDir, { recursive: true });

  const inputs = await collectPptxInputs(args.input);
  const selectedInputs = args.limit ? inputs.slice(0, Number(args.limit)) : inputs;
  const rows = [];

  for (const item of selectedInputs) {
    const parsed = await parsePptxProfileInput(item.buffer, item.fileName);
    rows.push({
      ...parsed,
      sourceRelativePath: item.sourceRelativePath,
    });
  }

  const flattened = rows.map(flattenRow);
  await fs.writeFile(path.join(outDir, 'parsed_results.json'), `${JSON.stringify(rows, null, 2)}\n`, 'utf8');
  await writeCsv(path.join(outDir, 'parsed_results.csv'), flattened);
  await writeCsv(
    path.join(outDir, 'qa_flags.csv'),
    rows.filter((row) => row.reviewTags?.length).map(flattenRow),
  );

  const beforeRows = await loadRows(args.before);
  const truthRows = await loadRows(args['ground-truth']);
  const report = buildReport({ input: args.input, outDir, rows, beforeRows, truthRows });
  await fs.writeFile(path.join(outDir, 'comparison_report.md'), report, 'utf8');

  console.log(`Parsed ${rows.length} PPTX files`);
  console.log(`Wrote ${path.join(outDir, 'parsed_results.csv')}`);
  console.log(`Wrote ${path.join(outDir, 'comparison_report.md')}`);
};

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

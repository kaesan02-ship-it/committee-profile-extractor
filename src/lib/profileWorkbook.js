import { EMPTY_VALUE } from './extractionRules.js';
import {
  formatCareerForTemplate,
  formatEducationForTemplate,
  formatEvaluationCareerForTemplate,
  formatIbkCurrentAffiliation,
  formatIbkPreviousCareer,
} from './profileExcelFormatter.js';

const valueOrBlank = (value) => value === EMPTY_VALUE ? '' : value ?? '';

const formatBirth = (birth = '') => {
  const match = /^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/.exec(birth);
  return match ? `${match[1]}년 ${Number(match[2])}월 ${Number(match[3])}일` : '';
};

const profileColumns = [
  { header: '원본폴더', key: 'sourceFolder', width: 34 },
  { header: '원본상대경로', key: 'sourceRelativePath', width: 46, hidden: true },
  { header: 'No.', key: 'no', width: 8.125 },
  { header: '이름', key: 'name', width: 12 },
  { header: '생년월일', key: 'birth', width: 20 },
  { header: '연령', key: 'age', width: 10 },
  { header: '학력', key: 'education', width: 42 },
  { header: '주요경력(일부)', key: 'career', width: 62.875 },
  { header: '평가경력', key: 'evaluationCareer', width: 36 },
  { header: '분야', key: 'field', width: 24 },
  { header: '연락처', key: 'phone', width: 18 },
  { header: '이메일주소', key: 'email', width: 28 },
];

const ibkColumns = [
  { header: '연번', key: 'no', width: 8 },
  { header: '분야', key: 'field', width: 14 },
  { header: '성명', key: 'name', width: 14 },
  { header: '선정', key: 'selected', width: 10 },
  { header: '성별', key: 'gender', width: 10 },
  { header: '최종학력(전공)', key: 'education', width: 36 },
  { header: '현재 소속', key: 'currentAffiliation', width: 34 },
  { header: '이전 경력', key: 'previousCareer', width: 42 },
  { header: '채용 관련 활동 경험', key: 'recruitmentExperience', width: 54 },
  { header: '생년월일', key: 'birth', width: 16 },
  { header: 'IBK참여이력', key: 'ibkHistory', width: 18 },
  { header: '연락처', key: 'phone', width: 18 },
];

const applySheetStyle = (sheet, { headerColor, bodyHeight, leftAligned }) => {
  const borderEdge = { style: 'thin', color: { indexed: 64 } };
  const border = { top: borderEdge, bottom: borderEdge, left: borderEdge, right: borderEdge };
  const headerFont = { name: '맑은 고딕', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const bodyFont = { name: '맑은 고딕', size: 10, color: { argb: 'FF000000' } };
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerColor } };
  const bodyFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
  const center = { vertical: 'middle', horizontal: 'center', wrapText: true };
  const left = { ...center, horizontal: 'left' };

  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: sheet.columnCount },
  };

  sheet.eachRow((row, rowNumber) => {
    const isHeader = rowNumber === 1;
    let lineCount = 1;

    sheet.columns.forEach((column, index) => {
      const cell = row.getCell(index + 1);
      cell.font = isHeader ? headerFont : bodyFont;
      cell.fill = isHeader ? headerFill : bodyFill;
      cell.border = border;
      cell.alignment = !isHeader && leftAligned.includes(column.key) ? left : center;

      if (!column.hidden && typeof cell.value === 'string') {
        lineCount = Math.max(lineCount, cell.value.split('\n').length);
      }
    });

    row.height = isHeader ? 26.1 : Math.min(409, Math.max(bodyHeight, lineCount * 15 + 8));
  });
};

export const buildProfileWorkbook = async (results) => {
  const { default: ExcelJS } = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'committee-profile-extractor';

  const profileSheet = workbook.addWorksheet('프로필 요약');
  profileSheet.columns = profileColumns;
  const ibkSheet = workbook.addWorksheet('IBK기업은행 양식');
  ibkSheet.columns = ibkColumns;

  results.forEach((row, index) => {
    const no = index + 1;
    const name = valueOrBlank(row.name);
    const birth = formatBirth(row.birth);
    const phone = valueOrBlank(row.phone);
    const evaluationCareer = formatEvaluationCareerForTemplate(row);

    profileSheet.addRow({
      sourceFolder: valueOrBlank(row.sourceFolder),
      sourceRelativePath: valueOrBlank(row.sourceRelativePath),
      no,
      name,
      birth,
      age: String(row.age ?? '').replace(/\D/g, ''),
      education: formatEducationForTemplate(row.educationDetails, row.education),
      career: formatCareerForTemplate(row),
      evaluationCareer,
      field: valueOrBlank(row.expertise),
      phone,
      email: valueOrBlank(row.email),
    });

    ibkSheet.addRow({
      no,
      field: '',
      name,
      selected: '',
      gender: valueOrBlank(row.gender),
      education: valueOrBlank(row.education),
      currentAffiliation: formatIbkCurrentAffiliation(row),
      previousCareer: formatIbkPreviousCareer(row),
      recruitmentExperience: evaluationCareer,
      birth,
      ibkHistory: '',
      phone,
    });
  });

  applySheetStyle(profileSheet, {
    headerColor: 'FF4C4F57',
    bodyHeight: 72,
    leftAligned: ['education', 'career', 'evaluationCareer'],
  });
  applySheetStyle(ibkSheet, {
    headerColor: 'FF2E3440',
    bodyHeight: 48,
    leftAligned: ['education', 'previousCareer', 'recruitmentExperience'],
  });

  return workbook;
};

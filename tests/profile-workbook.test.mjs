import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import { buildProfileWorkbook } from '../src/lib/profileWorkbook.js';

test('the downloaded workbook follows the two-sheet profile template', async () => {
  const source = {
    name: '테스트위원',
    gender: '여',
    birth: '1984.09.09',
    age: '42세',
    affiliation: '테스트연구원 책임연구원',
    education: '테스트대학교 석사',
    careerDetails: '前) 테스트기업 팀장 (2018~2020)',
    evaluationRaw: '[서류] 테스트공단\n[면접] 테스트공사',
    expertise: '인사, 채용',
    phone: '010-2345-6789',
    email: 'test@example.com',
    sourceFolder: 'test',
    sourceRelativePath: 'test/test.pptx',
  };
  const results = Array.from({ length: 40 }, () => source);
  const original = await buildProfileWorkbook(results);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await original.xlsx.writeBuffer());

  const [profile, ibk] = workbook.worksheets;
  assert.deepEqual(workbook.worksheets.map(({ name }) => name), ['프로필 요약', 'IBK기업은행 양식']);
  assert.deepEqual(profile.getRow(1).values.slice(1), [
    '원본폴더', '원본상대경로', 'No.', '이름', '생년월일', '연령', '학력',
    '주요경력(일부)', '평가경력', '분야', '연락처', '이메일주소',
  ]);
  assert.equal(profile.getCell('D2').value, source.name);
  assert.equal(profile.getCell('C41').value, 40);
  assert.equal(profile.getColumn('B').hidden, true);
  assert.equal(profile.getCell('B2').value, source.sourceRelativePath);
  assert.deepEqual(['J2', 'K2', 'L2'].map((address) => profile.getCell(address).value), [
    source.expertise, source.phone, source.email,
  ]);
  assert.equal(profile.getCell('G2').alignment.horizontal, 'left');
  assert.match(profile.getCell('I2').value, /\[서류\].*\n\[면접\]/);
  assert.equal(ibk.getCell('G1').value, '현재 소속');
  assert.equal(ibk.getCell('G2').value, source.affiliation);
  assert.equal(ibk.getCell('G2').alignment.horizontal, 'center');
  assert.equal(ibk.getCell('H2').value, '테스트기업 팀장');

  const headerColors = ['FF4C4F57', 'FF2E3440'];
  workbook.worksheets.forEach((sheet, index) => {
    assert.equal(sheet.rowCount, 41);
    assert.equal(sheet.autoFilter, 'A1:L1');
    assert.equal(sheet.getCell('A1').fill.fgColor.argb, headerColors[index]);
    assert.equal(sheet.getCell('A2').fill.fgColor.argb, 'FFFFFFFF');
    assert.ok((sheet.views ?? []).every(({ state }) => state !== 'frozen'));
  });
  for (const address of ['J2', 'K2', 'L2']) {
    const { horizontal, vertical, wrapText } = profile.getCell(address).alignment;
    assert.deepEqual({ horizontal, vertical, wrapText }, {
      horizontal: 'center', vertical: 'middle', wrapText: true,
    });
  }
});

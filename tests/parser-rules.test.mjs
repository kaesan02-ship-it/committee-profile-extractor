import test from 'node:test';
import assert from 'node:assert/strict';
import { EMPTY_VALUE, __testing, tagSuspiciousProfile } from '../src/lib/pptProfileParser.js';

test('extractPhone keeps labeled compact phone numbers', () => {
  assert.equal(__testing.extractPhone('01012345678'), '010-1234-5678');
});

test('extractPhone rejects unlabeled digit-only matches in full-document fallback', () => {
  const text = '생년월일 1973.07.07 경력 2010.01~2020.02 평가번호 01012345678';
  assert.equal(__testing.extractPhone(text, { requireContext: true }), '');
});

test('extractPhone requires nearby context for compact full-document matches', () => {
  const text = `연락처\n\n주요경력 ${'평가 '.repeat(12)} 평가번호 01012345678`;
  assert.equal(__testing.extractPhone(text, { requireContext: true }), '');
});

test('extractPhone accepts formatted phones in full-document fallback', () => {
  assert.equal(__testing.extractPhone('비상 연락 010-1234-5678', { requireContext: true }), '010-1234-5678');
});

test('extractGender ignores 남/여 choice templates', () => {
  assert.equal(__testing.extractGender('성별(남/여)'), '');
  assert.equal(__testing.extractGender('성별: 남'), '남');
});

test('chooseAffiliation refuses generic document text without organization signal', () => {
  assert.equal(__testing.chooseAffiliation('전문분야 경영전략 주요경력 평가위원', ''), '');
});

test('tagSuspiciousProfile flags missing fields for review', () => {
  const tags = tagSuspiciousProfile({
    phone: EMPTY_VALUE,
    affiliation: EMPTY_VALUE,
    education: EMPTY_VALUE,
    gender: EMPTY_VALUE,
    educationList: [],
    error: false,
  });

  assert.deepEqual(tags, ['phone_missing', 'affiliation_missing', 'education_review', 'gender_missing']);
});

# committee-profile-extractor

PPTX 파일에 들어 있는 위원 프로필 정보를 읽어 Excel 파일로 정리해주는 React + Vite 기반 웹앱입니다.

원본 프로젝트인 `ppt-excel-converter`는 브라우저에서 PPTX 슬라이드 XML을 직접 읽어 성명, 생년월일, 연락처, 이메일, 학력, 주요경력을 추출하는 구조입니다. 이 리포지토리는 그 구조를 바탕으로 다음 항목까지 확장한 업데이트 버전입니다.

- 위원 성명
- 성별
- 출생년월일
- 현소속
- 연락처
- 이메일주소
- 최종학력
- 전문분야
- 경력

## 핵심 개선점

### 1) 별도 리포지토리 운영용 구조 정리
- `vite.config.js`에서 저장소명을 한 곳에서 관리
- GitHub Pages 배포용 workflow 포함
- README, lint, build, preview 동작 정리

### 2) 추출 로직 모듈화
- `src/lib/pptProfileParser.js`: PPTX 파싱 및 필드 추출 로직
- `src/lib/extractionRules.js`: 샘플 PPT에 맞춰 튜닝할 라벨/섹션 규칙

### 3) 추출 필드 확장
- 기존: 성명, 생년월일, 연락처, 이메일, 학력, 주요경력
- 확장: 성별, 현소속, 최종학력, 전문분야, 경력 통합 정리

## 프로젝트 구조

```bash
committee-profile-extractor/
├─ .github/
│  └─ workflows/
│     └─ deploy.yml
├─ src/
│  ├─ lib/
│  │  ├─ extractionRules.js
│  │  └─ pptProfileParser.js
│  ├─ App.jsx
│  ├─ index.css
│  └─ main.jsx
├─ .gitignore
├─ eslint.config.js
├─ index.html
├─ package.json
├─ README.md
└─ vite.config.js
```

## 설치 및 실행

```bash
npm install
npm run dev
```

## 빌드

```bash
npm run build
npm run preview
```

## GitHub Pages 배포

### 1. 저장소명 확인
기본 설정은 아래처럼 되어 있습니다.

```js
const repoName = 'committee-profile-extractor';
```

만약 GitHub 저장소명이 다르면 `vite.config.js`의 `repoName`을 바꿔주세요.

### 2. Pages 설정
GitHub 저장소 Settings → Pages → Source 를 **GitHub Actions**로 두면 됩니다.

### 3. main 브랜치에 push
workflow가 자동으로 빌드 후 Pages 배포를 수행합니다.

## 샘플 PPT 기준 튜닝 방법

정확도 튜닝은 `src/lib/extractionRules.js` 중심으로 진행합니다.

### 자주 손보는 항목
- `FIELD_LABELS`: 문서마다 다른 표기 보강
  - 예: `현소속`, `현재소속`, `근무처`, `소속기관`
- `SECTION_STOP_HEADERS`: 섹션 경계 잘못 잡힐 때 종료 헤더 추가
- `AFFILIATION_HINTS`: 현소속 인식 키워드 추가

### 예시
어떤 샘플에서 전문분야 헤더가 `주요전문영역`으로 들어오면 아래에 추가합니다.

```js
expertise: ['전문분야', '주요분야', '전공', '핵심역량', '주요전문영역']
```

## 출력 결과

Excel 파일은 아래 3개 시트로 저장됩니다.

1. `1.위원DB`  
   최종 정리된 마스터 데이터
2. `2.원문정리`  
   학력/경력 원문 확인용
3. `3.요약`  
   총 파일 수, 정상 추출 수, 오류 수

## 권장 작업 순서

1. 샘플 PPT 3~10개로 먼저 테스트
2. 현소속 / 전문분야 / 경력에서 누락 패턴 확인
3. `src/lib/extractionRules.js` 보정
4. 다시 테스트 후 배포

## 참고 원본
- 원본 저장소: https://github.com/kaesan02-ship-it/ppt-excel-converter
- 원본 App 로직 참고: https://raw.githubusercontent.com/kaesan02-ship-it/ppt-excel-converter/main/src/App.jsx
- 원본 Vite 설정 참고: https://raw.githubusercontent.com/kaesan02-ship-it/ppt-excel-converter/main/vite.config.js
- 원본 패키지 설정 참고: https://raw.githubusercontent.com/kaesan02-ship-it/ppt-excel-converter/main/package.json

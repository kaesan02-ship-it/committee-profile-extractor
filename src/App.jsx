import React, { useCallback, useMemo, useRef, useState } from 'react';
import JSZip from 'jszip';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import { EMPTY_VALUE, parsePptxProfile } from './lib/pptProfileParser';

const normalizeRelativePath = (relativePath = '', fallbackName = '') => {
  const normalized = String(relativePath || '').replaceAll('\\', '/').replace(/^\/+|\/+$/g, '');
  return normalized || fallbackName;
};

const getSourceFolder = (relativePath = '', fallbackLabel = '개별 업로드') => {
  const segments = normalizeRelativePath(relativePath).split('/').filter(Boolean);
  return segments.length > 1 ? segments[segments.length - 2] : fallbackLabel;
};

const createUploadItem = (file, relativePath = '') => {
  const sourceRelativePath = normalizeRelativePath(relativePath || file.webkitRelativePath, file.name);

  return {
    file,
    sourceFolder: getSourceFolder(sourceRelativePath),
    sourceRelativePath,
  };
};

const buildWorkbook = (files, results) => {
  const masterRows = results.map((row) => ({
    파일명: row.fileName,
    원본폴더: row.sourceFolder,
    원본상대경로: row.sourceRelativePath,
    '위원 성명': row.name,
    성별: row.gender,
    출생년월일: row.birth,
    연령: row.age,
    현소속: row.affiliation,
    연락처: row.phone,
    이메일주소: row.email,
    최종학력: row.education,
    전문분야: row.expertise,
    경력: row.career,
  }));

  const rawRows = results.map((row) => ({
    파일명: row.fileName,
    원본폴더: row.sourceFolder,
    원본상대경로: row.sourceRelativePath,
    성명: row.name,
    학력원문: row.educationRaw,
    경력원문: row.career,
  }));

  const summaryRows = [
    { 항목: '총 파일 수', 값: files.length },
    { 항목: '정상 추출 수', 값: results.filter((row) => !row.error).length },
    { 항목: '오류 수', 값: results.filter((row) => row.error).length },
    { 항목: '생성일시', 값: new Date().toLocaleString('ko-KR') },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(masterRows), '1.위원DB');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rawRows), '2.원문정리');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryRows), '3.요약');
  return workbook;
};

function App() {
  const [files, setFiles] = useState([]);
  const [results, setResults] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [currentIdx, setCurrentIdx] = useState(-1);
  const [progress, setProgress] = useState(0);

  const fileInputRef = useRef(null);
  const folderInputRef = useRef(null);

  const onDrop = useCallback(async (acceptedFiles) => {
    const collected = [];

    for (const file of acceptedFiles) {
      if (file.name.endsWith('.pptx') && !file.name.startsWith('~$')) {
        collected.push(createUploadItem(file, file.webkitRelativePath || file.name));
        continue;
      }

      if (file.name.endsWith('.zip')) {
        const zip = await JSZip.loadAsync(file);
        const pptxNames = Object.keys(zip.files).filter(
          (name) => name.endsWith('.pptx') && !name.split('/').pop().startsWith('~$')
        );

        for (const name of pptxNames) {
          const blob = await zip.file(name).async('blob');
          const extractedFile = new File([blob], name.split('/').pop(), {
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          });
          collected.push(createUploadItem(extractedFile, `${file.name.replace(/\.zip$/i, '')}/${name}`));
        }
      }
    }

    setFiles((prev) => [...prev, ...collected]);
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    noClick: true,
  });

  const startProcessing = async () => {
    if (!files.length) return;

    setIsProcessing(true);
    setResults([]);
    setCurrentIdx(-1);
    setProgress(0);

    const nextResults = [];
    for (let i = 0; i < files.length; i += 1) {
      setCurrentIdx(i);
      const upload = files[i];
      const parsed = await parsePptxProfile(upload.file);
      nextResults.push({
        ...parsed,
        sourceFolder: upload.sourceFolder,
        sourceRelativePath: upload.sourceRelativePath,
      });
      setResults([...nextResults]);
      setProgress(Math.round(((i + 1) / files.length) * 100));
    }

    setIsProcessing(false);
  };

  const handleDownload = () => {
    const workbook = buildWorkbook(files, results);
    const stamp = new Date().toLocaleDateString('ko-KR').replace(/\. /g, '').replace(/\./g, '-');
    XLSX.writeFile(workbook, `위원_프로필_DB_${stamp}.xlsx`);
  };

  const removeFile = (targetIndex) => {
    setFiles((prev) => prev.filter((_, index) => index !== targetIndex));
    setResults((prev) => prev.filter((_, index) => index !== targetIndex));
  };

  const clearAll = () => {
    setFiles([]);
    setResults([]);
    setProgress(0);
    setCurrentIdx(-1);
    setIsProcessing(false);
  };

  const stats = useMemo(() => ({
    total: files.length,
    completed: results.filter((row) => !row.error).length,
    errors: results.filter((row) => row.error).length,
  }), [files.length, results]);

  return (
    <div className="card">
      <div className="header">
        <h1>위원 프로필 PPTX → Excel 추출기</h1>
        <p>
          원본 프로젝트는 React + Vite 기반으로 PPTX 슬라이드 XML을 직접 읽어 성명, 생년월일, 연락처,
          이메일, 학력, 경력을 추출하는 구조입니다. 이번 버전은 별도 리포지토리 운영을 전제로
          현소속·최종학력·전문분야까지 확장하고 튜닝 포인트를 분리했습니다.
        </p>
      </div>

      <div className="guide-box">
        <strong>현재 추출 컬럼</strong><br />
        위원 성명 / 성별 / 출생년월일 / 현소속 / 연락처 / 이메일주소 / 최종학력 / 전문분야 / 경력 / 원본폴더 / 원본상대경로
        <br /><br />
        <strong>샘플 PPT 기준 추가 튜닝 포인트</strong><br />
        src/lib/extractionRules.js 에서 라벨 별칭, 섹션 종료 헤더, 현소속 키워드를 조정하면 됩니다.
      </div>

      <div className="upload-options">
        <button className="upload-option-btn" onClick={() => fileInputRef.current?.click()}>
          📁 파일/ZIP 선택
        </button>
        <button className="upload-option-btn" onClick={() => folderInputRef.current?.click()}>
          📂 폴더 선택
        </button>
        <input
          type="file"
          ref={fileInputRef}
          style={{ display: 'none' }}
          multiple
          onChange={(event) => onDrop(Array.from(event.target.files || []))}
        />
        <input
          type="file"
          ref={folderInputRef}
          style={{ display: 'none' }}
          webkitdirectory="true"
          onChange={(event) => onDrop(Array.from(event.target.files || []))}
        />
      </div>

      <div {...getRootProps()} className={`dropzone ${isDragActive ? 'dropzone-active' : ''}`}>
        <input {...getInputProps()} />
        <div>
          <div style={{ fontSize: '3rem', marginBottom: '1rem' }}>📄</div>
          <h3 style={{ marginBottom: '0.75rem' }}>PPTX 또는 ZIP 파일을 드래그해서 올려주세요</h3>
          <p style={{ color: 'var(--text-muted)', lineHeight: 1.7 }}>
            ZIP 내부의 PPTX도 자동 분리합니다. 폴더 선택 업로드 시 원본폴더와 원본상대경로를 함께 기록합니다.
            임시 파일(~$)은 제외합니다.
          </p>
        </div>
      </div>

      {files.length > 0 && (
        <div className="result-area">
          <div className="info-grid">
            <div className="stat-card">
              <div className="stat-value">{stats.total}</div>
              <div className="stat-label">업로드 파일</div>
            </div>
            <div className="stat-card">
              <div className="stat-value">{stats.completed}</div>
              <div className="stat-label">정상 추출</div>
            </div>
            <div className="stat-card">
              <div className="stat-value">{stats.errors}</div>
              <div className="stat-label">오류</div>
            </div>
          </div>

          <div className="file-list">
            {files.map((upload, index) => {
              const status = index < results.length ? (results[index]?.error ? '오류' : '완료') : index === currentIdx ? '처리중' : '대기';
              const statusClass = index < results.length ? (results[index]?.error ? 'status-error' : 'status-success') : 'status-pending';

              return (
                <div className="file-item" key={`${upload.sourceRelativePath}-${index}`}>
                  <div>
                    <div>{upload.file.name}</div>
                    <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                      {upload.sourceFolder} · {upload.sourceRelativePath}
                    </div>
                  </div>
                  <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
                    <span className={`status-badge ${statusClass}`}>{status}</span>
                    {!isProcessing && (
                      <button
                        className="btn-secondary"
                        style={{ padding: '0.4rem 0.8rem' }}
                        onClick={() => removeFile(index)}
                      >
                        삭제
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>

          <div className="action-row" style={{ marginTop: '1rem' }}>
            <button className="btn-primary" disabled={isProcessing} onClick={startProcessing}>
              {isProcessing ? '추출 중...' : '업데이트 버전 실행'}
            </button>
            <button className="btn-secondary" disabled={!results.length || isProcessing} onClick={handleDownload}>
              Excel 다운로드
            </button>
            <button className="btn-secondary" disabled={isProcessing} onClick={clearAll}>
              전체 초기화
            </button>
          </div>

          {isProcessing && (
            <div className="progress-wrap">
              <div style={{ marginBottom: '0.5rem', color: 'var(--text-muted)' }}>
                {currentIdx + 1} / {files.length} · {files[currentIdx]?.file?.name || '분석 중'}
              </div>
              <div className="progress-bar-container">
                <div className="progress-bar" style={{ width: `${progress}%` }} />
              </div>
            </div>
          )}

          {results.length > 0 && (
            <div className="data-preview">
              <table>
                <thead>
                  <tr>
                    <th>원본폴더</th>
                    <th>원본상대경로</th>
                    <th>성명</th>
                    <th>성별</th>
                    <th>출생년월일</th>
                    <th>현소속</th>
                    <th>최종학력</th>
                    <th>전문분야</th>
                    <th>경력</th>
                  </tr>
                </thead>
                <tbody>
                  {results.map((row, index) => (
                    <tr key={`${row.fileName}-${index}`}>
                      <td>{row.sourceFolder || EMPTY_VALUE}</td>
                      <td style={{ minWidth: '220px' }}>{row.sourceRelativePath || EMPTY_VALUE}</td>
                      <td>{row.name || EMPTY_VALUE}</td>
                      <td>{row.gender || EMPTY_VALUE}</td>
                      <td>{row.birth || EMPTY_VALUE}</td>
                      <td>{row.affiliation || EMPTY_VALUE}</td>
                      <td>{row.education || EMPTY_VALUE}</td>
                      <td>{row.expertise || EMPTY_VALUE}</td>
                      <td style={{ whiteSpace: 'pre-line', minWidth: '260px' }}>{row.career || EMPTY_VALUE}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export default App;

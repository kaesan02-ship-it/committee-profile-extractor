import React, { useCallback, useMemo, useRef, useState } from 'react';
import JSZip from 'jszip';
import { useDropzone } from 'react-dropzone';
import { EMPTY_VALUE, parsePptxProfile } from './lib/pptProfileParser';
import { buildProfileWorkbook } from './lib/profileWorkbook.js';

const normalizeRelativePath = (relativePath = '', fallbackName = '') => {
  const normalized = String(relativePath || '')
    .replaceAll('\\', '/')
    .replace(/^\/+|\/+$/g, '');
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

const normalizeText = (value = '') => {
  const text = String(value ?? '').trim();
  return text && text !== EMPTY_VALUE ? text : EMPTY_VALUE;
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

const normalizeReviewTags = (tags = []) => {
  if (!Array.isArray(tags) || !tags.length) return EMPTY_VALUE;
  return tags.join(', ');
};

const bulletizeMultiline = (value = '') => {
  const text = normalizeMultiline(value);
  if (text === EMPTY_VALUE) return EMPTY_VALUE;
  return text
    .split(/\n+/)
    .map((line) => `• ${line}`)
    .join('\n');
};

const createStamp = () => {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mi = String(now.getMinutes()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}_${hh}${mi}`;
};

const downloadWorkbook = async (results) => {
  const workbook = await buildProfileWorkbook(results);
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob(
    [buffer],
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
  );

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `위원_프로필_DB_${createStamp()}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

function App() {
  const [files, setFiles] = useState([]);
  const [results, setResults] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [currentIdx, setCurrentIdx] = useState(-1);
  const [progress, setProgress] = useState(0);

  const fileInputRef = useRef(null);
  const folderInputRef = useRef(null);

  const onDrop = useCallback(async (acceptedFiles) => {
    const collected = [];

    for (const file of acceptedFiles) {
      const lowerFileName = file.name.toLowerCase();

      if (lowerFileName.endsWith('.pptx') && !file.name.startsWith('~$')) {
        collected.push(createUploadItem(file, file.webkitRelativePath || file.name));
        continue;
      }

      if (lowerFileName.endsWith('.zip')) {
        const zip = await JSZip.loadAsync(file);
        const pptxNames = Object.keys(zip.files).filter(
          (name) => name.toLowerCase().endsWith('.pptx') && !name.split('/').pop().startsWith('~$')
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
      const parsed = await parsePptxProfile(upload.file, upload.sourceRelativePath);

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

  const handleDownload = async () => {
    if (!results.length) return;

    setIsDownloading(true);
    try {
      await downloadWorkbook(results);
    } finally {
      setIsDownloading(false);
    }
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
    setIsDownloading(false);
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
          학력은 학사·석사·박사 단위로 줄바꿈해 정리하고, 경력도 항목별 줄바꿈으로 내려받을 수 있도록
          엑셀 저장 형식을 개선한 버전입니다.
        </p>
      </div>

      <div className="guide-box">
        <strong>Excel 저장 양식</strong><br />
        프로필 요약과 IBK기업은행 양식으로 저장합니다.
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

          <div className="action-row">
            <button className="btn-primary" disabled={isProcessing} onClick={startProcessing}>
              {isProcessing ? '추출 중...' : '업데이트 버전 실행'}
            </button>
            <button
              className="btn-secondary"
              disabled={!results.length || isProcessing || isDownloading}
              onClick={handleDownload}
            >
              {isDownloading ? 'Excel 생성 중...' : '예쁘게 Excel 다운로드'}
            </button>
            <button className="btn-secondary" disabled={isProcessing || isDownloading} onClick={clearAll}>
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

          <div className="file-list">
            {files.map((upload, index) => {
              const status = index < results.length
                ? (results[index]?.error ? '오류' : '완료')
                : index === currentIdx
                  ? '처리중'
                  : '대기';
              const statusClass = index < results.length
                ? (results[index]?.error ? 'status-error' : 'status-success')
                : 'status-pending';

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

          {results.length > 0 && (
            <div className="data-preview">
              <table>
                <thead>
                  <tr>
                    <th>원본폴더</th>
                    <th>원본상대경로</th>
                    <th>검수태그</th>
                    <th>성명</th>
                    <th>현소속</th>
                    <th>최종학력</th>
                    <th>학력상세</th>
                    <th>경력상세</th>
                  </tr>
                </thead>
                <tbody>
                  {results.map((row, index) => (
                    <tr key={`${row.fileName}-${index}`}>
                      <td>{row.sourceFolder || EMPTY_VALUE}</td>
                      <td style={{ minWidth: '220px' }}>{row.sourceRelativePath || EMPTY_VALUE}</td>
                      <td style={{ minWidth: '180px' }}>{normalizeReviewTags(row.reviewTags)}</td>
                      <td>{row.name || EMPTY_VALUE}</td>
                      <td style={{ minWidth: '220px' }}>{row.affiliation || EMPTY_VALUE}</td>
                      <td style={{ minWidth: '220px' }}>{row.education || EMPTY_VALUE}</td>
                      <td style={{ whiteSpace: 'pre-line', minWidth: '280px' }}>
                        {bulletizeMultiline(row.educationDetails)}
                      </td>
                      <td style={{ whiteSpace: 'pre-line', minWidth: '320px' }}>
                        {bulletizeMultiline(row.careerDetails)}
                      </td>
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

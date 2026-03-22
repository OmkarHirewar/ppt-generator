import React, { useState, useRef, useCallback } from 'react';
import axios from 'axios';

// ── Styles ─────────────────────────────────────────────────
const S = {
  // Layout
  page: {
    minHeight: '100vh',
    background: 'var(--gray)',
    fontFamily: "'DM Sans', sans-serif",
  },

  // Navbar
  nav: {
    background: 'var(--navy)',
    padding: '0 2.5rem',
    height: 64,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    position: 'sticky',
    top: 0,
    zIndex: 100,
    boxShadow: '0 2px 20px rgba(0,0,0,0.18)',
  },
  navLogo: {
    fontFamily: "'Syne', sans-serif",
    fontWeight: 800,
    fontSize: '1.35rem',
    color: 'var(--white)',
    letterSpacing: '-0.5px',
  },
  navBadge: {
    background: 'var(--teal)',
    color: 'var(--navy)',
    fontSize: '0.72rem',
    fontWeight: 700,
    padding: '4px 12px',
    borderRadius: 20,
    letterSpacing: '0.5px',
    textTransform: 'uppercase',
  },

  // Hero
  hero: {
    background: 'linear-gradient(135deg, var(--navy) 0%, #0e3d7a 60%, #1abebc15 100%)',
    padding: '5rem 2rem 4.5rem',
    textAlign: 'center',
  },
  heroTag: {
    display: 'inline-block',
    background: 'rgba(26,190,188,0.15)',
    border: '1px solid rgba(26,190,188,0.35)',
    color: 'var(--teal)',
    fontSize: '0.78rem',
    fontWeight: 600,
    letterSpacing: '1.5px',
    textTransform: 'uppercase',
    padding: '5px 16px',
    borderRadius: 30,
    marginBottom: '1.4rem',
    animation: 'fadeUp 0.5s ease both',
  },
  heroH1: {
    fontFamily: "'Syne', sans-serif",
    fontSize: 'clamp(2rem, 5vw, 3.4rem)',
    fontWeight: 800,
    color: 'var(--white)',
    lineHeight: 1.15,
    marginBottom: '1.1rem',
    animation: 'fadeUp 0.6s ease both',
    animationDelay: '0.1s',
  },
  heroP: {
    fontSize: '1.05rem',
    color: 'rgba(255,255,255,0.65)',
    maxWidth: 540,
    margin: '0 auto',
    lineHeight: 1.7,
    fontWeight: 300,
    animation: 'fadeUp 0.7s ease both',
    animationDelay: '0.2s',
  },

  // Main card
  wrapper: {
    maxWidth: 820,
    margin: '-2rem auto 4rem',
    padding: '0 1.5rem',
    position: 'relative',
    zIndex: 10,
  },
  card: {
    background: 'var(--white)',
    borderRadius: 'var(--radius)',
    boxShadow: 'var(--shadow)',
    padding: '2.5rem 3rem',
    animation: 'fadeUp 0.8s ease both',
    animationDelay: '0.3s',
  },
  cardTitle: {
    fontFamily: "'Syne', sans-serif",
    fontSize: '1.35rem',
    fontWeight: 700,
    color: 'var(--navy)',
    marginBottom: '0.3rem',
  },
  cardSub: {
    fontSize: '0.88rem',
    color: 'var(--muted)',
    marginBottom: '2rem',
  },

  // Upload grid
  uploadGrid: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '1.4rem',
    marginBottom: '1.8rem',
  },

  // Drop zone
  dropzone: (active, hasFile) => ({
    border: `2px dashed ${hasFile ? 'var(--teal)' : active ? 'var(--teal)' : '#d0daea'}`,
    borderRadius: 12,
    padding: '1.8rem 1.2rem',
    textAlign: 'center',
    cursor: 'pointer',
    transition: 'all 0.2s',
    background: hasFile ? '#e8faf8' : active ? '#f0fafa' : 'var(--gray)',
    position: 'relative',
    minHeight: 180,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '0.6rem',
  }),
  dropIcon: (color) => ({
    width: 48, height: 48,
    background: color || 'linear-gradient(135deg, var(--navy), var(--teal))',
    borderRadius: 12,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '1.5rem',
    margin: '0 auto',
  }),
  dropTitle: {
    fontFamily: "'Syne', sans-serif",
    fontSize: '0.95rem',
    fontWeight: 700,
    color: 'var(--navy)',
  },
  dropSub: {
    fontSize: '0.78rem',
    color: 'var(--muted)',
  },
  dropAccept: {
    fontSize: '0.72rem',
    color: 'var(--teal)',
    fontWeight: 600,
    background: 'rgba(26,190,188,0.1)',
    padding: '2px 8px',
    borderRadius: 6,
  },
  fileName: {
    fontSize: '0.82rem',
    color: '#0a5c4a',
    fontWeight: 600,
    background: 'rgba(26,190,188,0.15)',
    padding: '4px 10px',
    borderRadius: 8,
    maxWidth: '90%',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },

  // Progress
  progressWrap: {
    background: '#e8f0fe',
    borderRadius: 8,
    height: 6,
    overflow: 'hidden',
    marginBottom: '0.5rem',
  },
  progressBar: {
    height: '100%',
    background: 'linear-gradient(90deg, var(--teal), var(--navy))',
    borderRadius: 8,
    animation: 'progress 8s ease forwards',
  },

  // Buttons
  btnGenerate: (disabled) => ({
    width: '100%',
    padding: '1rem',
    background: disabled
      ? '#b0bcd4'
      : 'linear-gradient(135deg, var(--navy) 0%, #1565c0 100%)',
    color: 'var(--white)',
    border: 'none',
    borderRadius: 12,
    fontSize: '1rem',
    fontFamily: "'Syne', sans-serif",
    fontWeight: 700,
    letterSpacing: '0.5px',
    cursor: disabled ? 'not-allowed' : 'pointer',
    transition: 'all 0.2s',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '0.6rem',
  }),
  btnDownload: {
    width: '100%',
    padding: '1rem',
    background: 'linear-gradient(135deg, #0a7c5c, var(--teal))',
    color: 'var(--white)',
    border: 'none',
    borderRadius: 12,
    fontSize: '1rem',
    fontFamily: "'Syne', sans-serif",
    fontWeight: 700,
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '0.6rem',
    marginTop: '1rem',
    transition: 'all 0.2s',
  },

  // Status boxes
  statusBox: (type) => {
    const map = {
      loading: { bg: '#eef4ff', border: '#b3d0f5', color: '#1a3a6e' },
      success: { bg: '#e6faf5', border: 'var(--teal)', color: '#0a5c4a' },
      error:   { bg: '#fff0f0', border: '#f5a0a0', color: '#8b0000' },
    };
    const t = map[type] || map.loading;
    return {
      display: 'flex',
      alignItems: 'flex-start',
      gap: '0.9rem',
      background: t.bg,
      border: `1.5px solid ${t.border}`,
      borderRadius: 12,
      padding: '1.2rem 1.4rem',
      marginTop: '1.4rem',
      color: t.color,
      fontSize: '0.9rem',
    };
  },

  // Spinner
  spinner: {
    width: 22, height: 22,
    border: '3px solid rgba(26,62,110,0.15)',
    borderTopColor: 'var(--navy)',
    borderRadius: '50%',
    animation: 'spin 0.8s linear infinite',
    flexShrink: 0,
  },

  // Steps
  stepsSection: {
    maxWidth: 820,
    margin: '0 auto 4rem',
    padding: '0 1.5rem',
  },
  stepsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
    gap: '1.1rem',
    marginTop: '1.5rem',
  },
  stepCard: {
    background: 'var(--white)',
    borderRadius: 14,
    padding: '1.5rem 1.3rem',
    boxShadow: '0 4px 20px rgba(10,45,94,0.06)',
    borderTop: '3px solid var(--teal)',
  },
  stepNum: {
    fontFamily: "'Syne', sans-serif",
    fontSize: '2rem',
    fontWeight: 800,
    color: '#d0e8f5',
    lineHeight: 1,
    marginBottom: '0.3rem',
  },
  stepTitle: {
    fontFamily: "'Syne', sans-serif",
    fontSize: '0.9rem',
    fontWeight: 700,
    color: 'var(--navy)',
    marginBottom: '0.3rem',
  },
  stepDesc: {
    fontSize: '0.8rem',
    color: 'var(--muted)',
    lineHeight: 1.6,
  },

  // Footer
  footer: {
    background: 'var(--navy)',
    color: 'rgba(255,255,255,0.4)',
    textAlign: 'center',
    padding: '1.8rem',
    fontSize: '0.82rem',
  },

  label: {
    fontSize: '0.83rem',
    fontWeight: 600,
    color: 'var(--navy)',
    marginBottom: '0.5rem',
    display: 'block',
    letterSpacing: '0.2px',
  },
};

// ── DropZone Component ─────────────────────────────────────
function DropZone({ label, accept, icon, acceptLabel, file, onFile }) {
  const [dragActive, setDragActive] = useState(false);
  const inputRef = useRef();

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragActive(false);
    const f = e.dataTransfer.files[0];
    if (f) onFile(f);
  }, [onFile]);

  const handleChange = (e) => {
    if (e.target.files[0]) onFile(e.target.files[0]);
  };

  return (
    <div>
      <span style={S.label}>{label}</span>
      <div
        style={S.dropzone(dragActive, !!file)}
        onClick={() => inputRef.current.click()}
        onDragOver={(e) => { e.preventDefault(); setDragActive(true); }}
        onDragLeave={() => setDragActive(false)}
        onDrop={handleDrop}
      >
        <input
          ref={inputRef}
          type="file"
          accept={accept}
          style={{ display: 'none' }}
          onChange={handleChange}
        />
        <div style={S.dropIcon()}>{icon}</div>
        {file ? (
          <>
            <div style={S.dropTitle}>✅ File Selected</div>
            <div style={S.fileName}>{file.name}</div>
            <div style={S.dropSub}>Click to change</div>
          </>
        ) : (
          <>
            <div style={S.dropTitle}>Drop file here</div>
            <div style={S.dropSub}>or click to browse</div>
            <div style={S.dropAccept}>{acceptLabel}</div>
          </>
        )}
      </div>
    </div>
  );
}

// ── Main App ───────────────────────────────────────────────
export default function App() {
  const [docFile,      setDocFile]      = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [status,       setStatus]       = useState(null); // null | 'loading' | 'success' | 'error'
  const [errorMsg,     setErrorMsg]     = useState('');
  const [downloadUrl,  setDownloadUrl]  = useState(null);
  const [downloadName, setDownloadName] = useState('presentation.pptx');
  const [progress,     setProgress]     = useState(0);

  const canGenerate = docFile && templateFile && status !== 'loading';

  const handleGenerate = async () => {
    if (!canGenerate) return;

    setStatus('loading');
    setErrorMsg('');
    setDownloadUrl(null);
    setProgress(0);

    const formData = new FormData();
    formData.append('document', docFile);
    formData.append('template', templateFile);

    try {
      const res = await axios.post('/generate-ppt', formData, {
        responseType: 'blob',
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (e) => {
          setProgress(Math.round((e.loaded / e.total) * 30));
        },
        timeout: 120000, // 2 min timeout
      });

      // Get filename from header
      const disposition = res.headers['content-disposition'];
      let fname = 'presentation.pptx';
      if (disposition) {
        const match = disposition.match(/filename="?([^"]+)"?/);
        if (match) fname = match[1];
      }

      const url = URL.createObjectURL(new Blob([res.data]));
      setDownloadUrl(url);
      setDownloadName(fname);
      setStatus('success');
      setProgress(100);

    } catch (err) {
      setStatus('error');
      if (err.response) {
        try {
          const text = await err.response.data.text();
          const json = JSON.parse(text);
          setErrorMsg(json.detail || json.error || 'Server error');
        } catch {
          setErrorMsg(`Server error: ${err.response.status}`);
        }
      } else {
        setErrorMsg(err.message || 'Network error — is the backend running?');
      }
    }
  };

  const handleDownload = () => {
    if (!downloadUrl) return;
    const a = document.createElement('a');
    a.href = downloadUrl;
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const handleReset = () => {
    setDocFile(null);
    setTemplateFile(null);
    setStatus(null);
    setErrorMsg('');
    setDownloadUrl(null);
    setProgress(0);
  };

  return (
    <div style={S.page}>

      {/* ── Navbar ── */}
      <nav style={S.nav}>
        <div style={S.navLogo}>
          my<span style={{ color: 'var(--teal)' }}>robo</span>.in
        </div>
        <div style={S.navBadge}>⚡ PPT Generator</div>
      </nav>

      {/* ── Hero ── */}
      <section style={S.hero}>
        <div style={S.heroTag}>Template-Aware AI Generation</div>
        <h1 style={S.heroH1}>
          Document → PowerPoint<br/>
          <span style={{ color: 'var(--teal)' }}>In Your Own Template</span>
        </h1>
        <p style={S.heroP}>
          Upload any document and your own PPT template.
          Get a perfectly formatted presentation that matches
          your design — fonts, colors, layout and all.
        </p>
      </section>

      {/* ── Main Card ── */}
      <div style={S.wrapper}>
        <div style={S.card}>
          <div style={S.cardTitle}>Generate Your Presentation</div>
          <div style={S.cardSub}>
            Upload your document and reference template to get started
          </div>

          {/* Upload Grid */}
          <div style={S.uploadGrid}>
            <DropZone
              label="📄 Document"
              accept=".pdf,.docx,.txt,.rtf"
              icon="📄"
              acceptLabel=".pdf  .docx  .txt  .rtf"
              file={docFile}
              onFile={setDocFile}
            />
            <DropZone
              label="🎨 PPT Template"
              accept=".pptx"
              icon="🎨"
              acceptLabel=".pptx only"
              file={templateFile}
              onFile={setTemplateFile}
            />
          </div>

          {/* Generate Button */}
          <button
            style={S.btnGenerate(!canGenerate)}
            onClick={handleGenerate}
            disabled={!canGenerate}
          >
            {status === 'loading' ? (
              <>
                <div style={S.spinner} />
                <span>Generating...</span>
              </>
            ) : (
              <>
                <span>✨</span>
                <span>Generate PowerPoint</span>
              </>
            )}
          </button>

          {/* Progress Bar */}
          {status === 'loading' && (
            <div style={{ marginTop: '1rem' }}>
              <div style={S.progressWrap}>
                <div style={S.progressBar} />
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--muted)', textAlign: 'center' }}>
                Reading document → Analyzing template → Building slides...
              </div>
            </div>
          )}

          {/* Loading Status */}
          {status === 'loading' && (
            <div style={S.statusBox('loading')}>
              <div style={S.spinner} />
              <div>
                <strong>Processing your files...</strong><br />
                <small>
                  Extracting content from your document, analyzing the template design,
                  and generating matching slides. This takes 10–30 seconds.
                </small>
              </div>
            </div>
          )}

          {/* Success Status */}
          {status === 'success' && (
            <>
              <div style={S.statusBox('success')}>
                <span style={{ fontSize: '1.5rem' }}>🎉</span>
                <div>
                  <strong>Your PPT is ready!</strong><br />
                  <small>
                    Generated with your template's exact colors, fonts, and layout.
                    Click download below.
                  </small>
                </div>
              </div>
              <button style={S.btnDownload} onClick={handleDownload}>
                <span>⬇️</span>
                <span>Download {downloadName}</span>
              </button>
              <button
                onClick={handleReset}
                style={{
                  width: '100%', marginTop: '0.7rem', padding: '0.7rem',
                  background: 'none', border: '1.5px solid var(--border)',
                  borderRadius: 10, cursor: 'pointer', color: 'var(--muted)',
                  fontSize: '0.88rem', fontFamily: "'DM Sans', sans-serif",
                }}
              >
                ↺ Generate Another
              </button>
            </>
          )}

          {/* Error Status */}
          {status === 'error' && (
            <>
              <div style={S.statusBox('error')}>
                <span style={{ fontSize: '1.5rem' }}>❌</span>
                <div>
                  <strong>Generation failed</strong><br />
                  <small>{errorMsg}</small>
                </div>
              </div>
              <button
                onClick={handleReset}
                style={{
                  width: '100%', marginTop: '0.8rem', padding: '0.8rem',
                  background: 'var(--navy)', border: 'none',
                  borderRadius: 10, cursor: 'pointer', color: 'white',
                  fontSize: '0.9rem', fontFamily: "'Syne', sans-serif",
                  fontWeight: 700,
                }}
              >
                ↺ Try Again
              </button>
            </>
          )}
        </div>
      </div>

      {/* ── How It Works ── */}
      <div style={S.stepsSection}>
        <div style={{
          fontFamily: "'Syne', sans-serif",
          fontSize: '1.4rem', fontWeight: 800, color: 'var(--navy)'
        }}>
          How it works
        </div>
        <div style={{ fontSize: '0.88rem', color: 'var(--muted)', marginTop: '0.3rem' }}>
          Four steps from document to finished presentation
        </div>
        <div style={S.stepsGrid}>
          {[
            { n: '01', t: 'Upload Files',        d: 'Drop your document (PDF/DOCX/TXT) and your PPT template.' },
            { n: '02', t: 'Extract Content',     d: 'Backend parses your document into title, sections, and bullet points.' },
            { n: '03', t: 'Analyze Template',    d: 'Colors, fonts, layouts and placeholder positions are extracted from your PPT.' },
            { n: '04', t: 'Generate & Download', d: 'A new PPT is built matching your template design with the document content.' },
          ].map(s => (
            <div key={s.n} style={S.stepCard}>
              <div style={S.stepNum}>{s.n}</div>
              <div style={S.stepTitle}>{s.t}</div>
              <div style={S.stepDesc}>{s.d}</div>
            </div>
          ))}
        </div>
      </div>

      {/* ── Supported Formats ── */}
      <div style={{ ...S.stepsSection, marginTop: '-2rem' }}>
        <div style={{
          background: 'var(--white)',
          borderRadius: 14,
          padding: '1.5rem 2rem',
          boxShadow: '0 4px 20px rgba(10,45,94,0.06)',
          display: 'grid',
          gridTemplateColumns: '1fr 1fr',
          gap: '1.5rem',
        }}>
          <div>
            <div style={{ fontFamily: "'Syne', sans-serif", fontWeight: 700, color: 'var(--navy)', marginBottom: '0.6rem' }}>
              📄 Document Formats
            </div>
            {[
              { fmt: '.pdf',  label: 'PDF',  note: 'Full support' },
              { fmt: '.docx', label: 'DOCX', note: 'Full support' },
              { fmt: '.txt',  label: 'TXT',  note: 'Full support' },
              { fmt: '.rtf',  label: 'RTF',  note: 'Full support' },
            ].map(f => (
              <div key={f.fmt} style={{ display: 'flex', gap: '0.6rem', marginBottom: '0.4rem', alignItems: 'center' }}>
                <span style={{ background: 'var(--teal)', color: 'white', borderRadius: 4, padding: '1px 7px', fontSize: '0.75rem', fontWeight: 700 }}>{f.fmt}</span>
                <span style={{ fontSize: '0.85rem', color: 'var(--muted)' }}>{f.label} — {f.note}</span>
              </div>
            ))}
          </div>
          <div>
            <div style={{ fontFamily: "'Syne', sans-serif", fontWeight: 700, color: 'var(--navy)', marginBottom: '0.6rem' }}>
              🎨 Template Format
            </div>
            <div style={{ display: 'flex', gap: '0.6rem', marginBottom: '0.4rem', alignItems: 'center' }}>
              <span style={{ background: 'var(--navy)', color: 'white', borderRadius: 4, padding: '1px 7px', fontSize: '0.75rem', fontWeight: 700 }}>.pptx</span>
              <span style={{ fontSize: '0.85rem', color: 'var(--muted)' }}>PowerPoint — Required</span>
            </div>
            <div style={{ fontSize: '0.8rem', color: 'var(--muted)', marginTop: '0.8rem', lineHeight: 1.6 }}>
              The template's colors, fonts, header, footer, and layout will be applied to every generated slide.
            </div>
          </div>
        </div>
      </div>

      {/* ── Footer ── */}
      <footer style={S.footer}>
        Built for <span style={{ color: 'var(--teal)', fontWeight: 600 }}>myrobo.in</span>
        &nbsp;·&nbsp; Future Skills for Future Schools
        &nbsp;·&nbsp; Powered by python-pptx
      </footer>
    </div>
  );
}

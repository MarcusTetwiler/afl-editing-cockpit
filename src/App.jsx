import { useState, useEffect, useRef, useCallback } from "react";
import mammoth from "mammoth";
import * as XLSX from "xlsx";

// ── Palette ───────────────────────────────────────────────────────────────────
const C = {
  charcoal:   "#1C1C1C",
  parchment:  "#F2ECD8",
  amber:      "#D4820A",
  redOrange:  "#C0392B",
  dimParch:   "#E8DFC4",
  smokeDark:  "#2A2A2A",
  smokeLight: "#3A3938",
  textMuted:  "#8A8070",
  textLight:  "#F2ECD8",
  textDim:    "#B8AC96",
  pass:       "#4A7C59",
  warn:       "#D4820A",
  regress:    "#C0392B",
};

// ── Signal Definitions ────────────────────────────────────────────────────────
const SIGNALS = [
  {
    id: "seven_word",
    name: "7-Word Narration Rule",
    short: "7-WORD",
    description: "Non-dialogue sentences under 7 words",
  },
  {
    id: "interiority",
    name: "Interiority Leak",
    short: "INTERIORITY",
    description: "Stated thoughts, feelings, motives",
  },
  {
    id: "pronoun_ambiguity",
    name: "Pronoun Ambiguity",
    short: "PRONOUN",
    description: 'Vague "it," "this," "that" with unclear referent',
  },
  {
    id: "repeated_openers",
    name: "Repeated Openers",
    short: "OPENERS",
    description: "3+ consecutive sentences starting with same word",
  },
  {
    id: "flat_wave",
    name: "Flat Sentence Wave",
    short: "FLAT WAVE",
    description: "Paragraphs with monotonous sentence-length clustering",
  },
  {
    id: "filter_verb",
    name: "Filter Verb / Perception Leak",
    short: "FILTER VERB",
    description: '"saw," "looked," "watched," "noticed," "felt," "realized"',
  },
];

// ── Text Extraction ───────────────────────────────────────────────────────────
async function extractTextFromDocx(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value;
}

// ── Signal Analysis Engine ────────────────────────────────────────────────────
function runSignals(rawText) {
  const findings = [];
  const paragraphs = rawText.split(/\n+/).filter(p => p.trim().length > 10);
  const chapterPattern = /^(chapter|prologue|epilogue|part\s+\d|andrew|elena|reese|logan|jack|anika|ryba|rupert)\s*[IVXivx\d]*/i;

  let currentChapter = "Opening";
  let paraIndex = 0;

  const interiorityPatterns = [
    /\b(he|she|they|andrew|elena|reese|logan|jack|anika|ryba|rupert)\s+(knew|felt|realized|decided|thought|wondered|understood|believed|hoped|feared|wanted|needed|wished|sensed|remembered|recognized|considered|assumed|supposed|imagined|expected)\b/gi,
    /\b(his|her|their)\s+(mind|thoughts?|feelings?|emotions?|instinct|gut)\b/gi,
  ];

  const pronounPatterns = [
    /\b(it|this|that)\s+(was|is|had|has|did|does|would|could|should|seemed|looked|felt|appeared)\b/gi,
    /^(It|This|That)\s+(was|is|had|seemed|would)/,
  ];

  const filterPatterns = [
    /\b(saw|looked|watched|noticed|realized|felt|heard|smelled|tasted|sensed|observed|spotted|glimpsed|perceived|detected)\b/gi,
  ];

  for (const para of paragraphs) {
    paraIndex++;
    const trimmed = para.trim();

    if (chapterPattern.test(trimmed) && trimmed.length < 60) {
      currentChapter = trimmed.replace(/\s+/g, " ").toUpperCase();
      continue;
    }

    const isDialogue = (trimmed.match(/"/g) || []).length >= 2;
    const sentences = trimmed.match(/[^.!?]+[.!?]+/g) || [trimmed];

    // 7-Word Rule
    if (!isDialogue) {
      for (const sent of sentences) {
        const s = sent.trim();
        if (!s) continue;
        const wordCount = s.split(/\s+/).filter(Boolean).length;
        if (wordCount > 0 && wordCount < 7) {
          findings.push({ signal_id: "seven_word", chapter: currentChapter, sentence: s, issue: "7-WORD NARRATION RULE" });
        }
      }
    }

    // Interiority Leak
    for (const sent of sentences) {
      const s = sent.trim();
      for (const pat of interiorityPatterns) {
        pat.lastIndex = 0;
        if (pat.test(s)) {
          findings.push({ signal_id: "interiority", chapter: currentChapter, sentence: s, issue: "INTERIORITY LEAK" });
          break;
        }
      }
    }

    // Pronoun Ambiguity
    for (const sent of sentences) {
      const s = sent.trim();
      for (const pat of pronounPatterns) {
        pat.lastIndex = 0;
        if (pat.test(s)) {
          findings.push({ signal_id: "pronoun_ambiguity", chapter: currentChapter, sentence: s, issue: "PRONOUN AMBIGUITY" });
          break;
        }
      }
    }

    // Repeated Openers
    if (sentences.length >= 3) {
      const openers = sentences.map(s => s.trim().split(/\s+/)[0]?.toLowerCase().replace(/[^a-z]/g, "") || "");
      let streak = 1;
      for (let i = 1; i < openers.length; i++) {
        if (openers[i] && openers[i] === openers[i - 1]) {
          streak++;
          if (streak >= 3) {
            findings.push({
              signal_id: "repeated_openers",
              chapter: currentChapter,
              sentence: sentences.slice(i - 2, i + 1).join(" ").trim(),
              issue: "REPEATED OPENERS",
            });
            streak = 1;
          }
        } else {
          streak = 1;
        }
      }
    }

    // Flat Sentence Wave
    if (sentences.length >= 4) {
      const lengths = sentences.map(s => s.trim().split(/\s+/).length);
      const mean = lengths.reduce((a, b) => a + b, 0) / lengths.length;
      const variance = lengths.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / lengths.length;
      const stdDev = Math.sqrt(variance);
      if (stdDev < 3.5 && mean > 8) {
        findings.push({
          signal_id: "flat_wave",
          chapter: currentChapter,
          sentence: `[Para ${paraIndex}] ${trimmed.slice(0, 120)}…`,
          issue: "FLAT SENTENCE WAVE",
        });
      }
    }

    // Filter Verb
    for (const sent of sentences) {
      const s = sent.trim();
      for (const pat of filterPatterns) {
        pat.lastIndex = 0;
        if (pat.test(s)) {
          findings.push({ signal_id: "filter_verb", chapter: currentChapter, sentence: s, issue: "FILTER VERB" });
          break;
        }
      }
    }
  }

  return findings;
}

// ── Export ────────────────────────────────────────────────────────────────────
function exportToXlsx(findings, draftName) {
  const rows = [["Chapter / Location", "Flagged Sentence", "Issue Type"]];
  for (const f of findings) rows.push([f.chapter, f.sentence, f.issue]);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Findings");
  XLSX.writeFile(wb, `${draftName}_findings.xlsx`);
}

// ── Persistence ───────────────────────────────────────────────────────────────
const STORAGE_KEY = "afl_cockpit_sessions";

function loadSessions() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
}

function saveSessions(sessions) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(
      sessions.map(s => ({ ...s, findings: s.findings?.slice(0, 2000) }))
    ));
  } catch (e) { console.warn("localStorage quota hit"); }
}

// ── Health Score ──────────────────────────────────────────────────────────────
function computeHealth(findings, wordCount) {
  if (!wordCount) return 0;
  const rate = (findings.length / wordCount) * 1000;
  return Math.round(Math.max(40, Math.min(100, 100 - rate * 0.6)) * 10) / 10;
}

// ── Sub-components ────────────────────────────────────────────────────────────
function SignalCard({ signal, count, prevCount }) {
  const delta = prevCount !== undefined ? count - prevCount : null;
  const isPass = count === 0;
  const isRegress = delta !== null && delta > 0;

  return (
    <div style={{
      background: C.smokeDark,
      border: `1px solid ${isRegress ? C.regress : isPass ? C.pass : C.smokeLight}`,
      padding: "16px 20px",
      position: "relative",
    }}>
      {isRegress && <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "2px", background: C.regress }} />}
      <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.textMuted, fontFamily: "Barlow Condensed, sans-serif", marginBottom: "8px" }}>
        {signal.short}
      </div>
      <div style={{ display: "flex", alignItems: "flex-end", gap: "12px" }}>
        <div style={{ fontSize: "36px", fontFamily: "Barlow Condensed, sans-serif", fontWeight: 700, color: isPass ? C.pass : isRegress ? C.regress : C.parchment, lineHeight: 1 }}>
          {count.toLocaleString()}
        </div>
        {delta !== null && (
          <div style={{ fontSize: "13px", fontFamily: "Barlow Condensed, sans-serif", color: isRegress ? C.regress : C.pass, paddingBottom: "4px" }}>
            {isRegress ? `▲ +${delta}` : delta < 0 ? `▼ ${delta}` : "—"}
          </div>
        )}
        {isPass && <div style={{ fontSize: "12px", color: C.pass, paddingBottom: "4px", fontFamily: "Barlow Condensed, sans-serif" }}>✓ PASS</div>}
      </div>
      <div style={{ fontSize: "11px", color: C.textDim, marginTop: "8px", fontFamily: "IM Fell English, serif", fontStyle: "italic" }}>
        {signal.description}
      </div>
    </div>
  );
}

function HealthBar({ score, prevScore }) {
  const delta = prevScore !== undefined ? (score - prevScore).toFixed(1) : null;
  const pct = ((score - 40) / 60) * 100;
  return (
    <div style={{ marginBottom: "32px" }}>
      <div style={{ display: "flex", alignItems: "baseline", gap: "16px", marginBottom: "12px" }}>
        <div style={{ fontSize: "64px", fontFamily: "Barlow Condensed, sans-serif", fontWeight: 700, color: C.amber, lineHeight: 1 }}>
          {score}
        </div>
        <div style={{ fontFamily: "Barlow Condensed, sans-serif" }}>
          <div style={{ fontSize: "14px", color: C.textMuted, letterSpacing: "0.12em" }}>/ 100</div>
          {delta !== null && (
            <div style={{ fontSize: "16px", color: parseFloat(delta) >= 0 ? C.pass : C.regress }}>
              {parseFloat(delta) >= 0 ? `▲ +${delta}` : `▼ ${delta}`} since previous
            </div>
          )}
        </div>
      </div>
      <div style={{ height: "4px", background: C.smokeLight, position: "relative" }}>
        <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: `${pct}%`, background: `linear-gradient(90deg, ${C.amber}, ${C.redOrange})` }} />
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [sessions, setSessions] = useState([]);
  const [activeSession, setActiveSession] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [uploadStep, setUploadStep] = useState("");
  const [uploadTime, setUploadTime] = useState(null);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [filterSignal, setFilterSignal] = useState("all");
  const [findingsSearch, setFindingsSearch] = useState("");
  const fileInputRef = useRef();

  useEffect(() => {
    const saved = loadSessions();
    if (saved.length > 0) {
      setSessions(saved);
      setActiveSession(saved[saved.length - 1]);
    }
  }, []);

  const handleUpload = useCallback(async (file) => {
    if (!file?.name.endsWith(".docx")) { alert("Please upload a .docx file."); return; }
    const draftName = file.name.replace(".docx", "");
    setUploading(true);
    const startTime = Date.now();
    try {
      setUploadStep("Parsing manuscript…");
      const rawText = await extractTextFromDocx(file);
      setUploadStep("Running 6 signal passes…");
      await new Promise(r => setTimeout(r, 100));
      const findings = runSignals(rawText);
      setUploadStep("Computing health score…");
      const wordCount = rawText.split(/\s+/).filter(Boolean).length;
      const health = computeHealth(findings, wordCount);
      const counts = {};
      for (const sig of SIGNALS) counts[sig.id] = 0;
      for (const f of findings) counts[f.signal_id] = (counts[f.signal_id] || 0) + 1;
      const newSession = { id: Date.now(), draftName, uploadedAt: new Date().toISOString(), wordCount, health, counts, findings };
      setSessions(prev => {
        const next = [...prev, newSession];
        saveSessions(next);
        return next;
      });
      setActiveSession(newSession);
      setUploadTime(((Date.now() - startTime) / 1000).toFixed(1));
    } catch (err) {
      alert("Error parsing manuscript: " + err.message);
    } finally {
      setUploading(false);
      setUploadStep("");
    }
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleUpload(file);
  }, [handleUpload]);

  const activeIdx = sessions.findIndex(s => s.id === activeSession?.id);
  const prevSession = activeIdx > 0 ? sessions[activeIdx - 1] : null;
  const totalFindings = activeSession ? Object.values(activeSession.counts).reduce((a, b) => a + b, 0) : 0;
  const totalPrev = prevSession ? Object.values(prevSession.counts).reduce((a, b) => a + b, 0) : null;

  const filteredFindings = (activeSession?.findings || []).filter(f => {
    if (filterSignal !== "all" && f.signal_id !== filterSignal) return false;
    if (findingsSearch && !f.sentence.toLowerCase().includes(findingsSearch.toLowerCase()) && !f.chapter.toLowerCase().includes(findingsSearch.toLowerCase())) return false;
    return true;
  });

  return (
    <div style={{ minHeight: "100vh", background: C.charcoal, color: C.parchment, fontFamily: "Barlow Condensed, sans-serif" }}>
      {/* Header */}
      <div style={{ borderBottom: `1px solid ${C.smokeLight}`, padding: "0 32px", display: "flex", alignItems: "center", justifyContent: "space-between", height: "56px" }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: "16px" }}>
          <span style={{ fontSize: "18px", fontWeight: 700, letterSpacing: "0.15em", color: C.amber }}>AFL</span>
          <span style={{ fontSize: "14px", letterSpacing: "0.2em", color: C.textMuted }}>EDITING COCKPIT</span>
        </div>
        <div style={{ display: "flex", gap: "4px" }}>
          {["dashboard", "findings"].map(tab => (
            <button key={tab} onClick={() => setActiveTab(tab)} style={{
              background: activeTab === tab ? C.smokeLight : "transparent",
              border: "none", color: activeTab === tab ? C.parchment : C.textMuted,
              padding: "8px 16px", fontSize: "12px", letterSpacing: "0.15em",
              cursor: "pointer", fontFamily: "Barlow Condensed, sans-serif", textTransform: "uppercase",
            }}>
              {tab}
              {tab === "findings" && activeSession && (
                <span style={{ marginLeft: "8px", background: C.amber, color: C.charcoal, padding: "1px 6px", fontSize: "10px", fontWeight: 700 }}>
                  {totalFindings.toLocaleString()}
                </span>
              )}
            </button>
          ))}
        </div>
      </div>

      <div style={{ padding: "32px", maxWidth: "1100px", margin: "0 auto" }}>
        {/* Version Rail */}
        {sessions.length > 0 && (
          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap", marginBottom: "24px" }}>
            {sessions.map((s, i) => (
              <div key={s.id} onClick={() => setActiveSession(s)} style={{
                padding: "4px 12px", cursor: "pointer",
                background: s.id === activeSession?.id ? C.amber : C.smokeLight,
                color: s.id === activeSession?.id ? C.charcoal : C.textDim,
                fontFamily: "Barlow Condensed, sans-serif", fontSize: "12px", letterSpacing: "0.1em",
                fontWeight: s.id === activeSession?.id ? 700 : 400,
              }}>
                {s.draftName} <span style={{ opacity: 0.7 }}>{s.health}</span>
              </div>
            ))}
          </div>
        )}

        {/* Upload Zone */}
        <div
          onDrop={handleDrop} onDragOver={e => e.preventDefault()}
          onClick={() => !uploading && fileInputRef.current?.click()}
          style={{
            border: `1px dashed ${C.smokeLight}`, padding: "20px 28px", marginBottom: "32px",
            cursor: uploading ? "default" : "pointer",
            display: "flex", alignItems: "center", justifyContent: "space-between",
          }}
          onMouseEnter={e => { if (!uploading) e.currentTarget.style.borderColor = C.amber; }}
          onMouseLeave={e => { e.currentTarget.style.borderColor = C.smokeLight; }}
        >
          <input ref={fileInputRef} type="file" accept=".docx" style={{ display: "none" }}
            onChange={e => e.target.files[0] && handleUpload(e.target.files[0])} />
          {uploading ? (
            <div>
              <div style={{ fontSize: "13px", color: C.amber, letterSpacing: "0.1em" }}>{uploadStep}</div>
              <div style={{ width: "220px", height: "2px", background: C.smokeLight, marginTop: "10px", overflow: "hidden" }}>
                <div style={{ height: "100%", background: C.amber, width: "40%", animation: "pulse 1.4s ease-in-out infinite" }} />
              </div>
            </div>
          ) : (
            <>
              <div>
                <div style={{ fontSize: "14px", letterSpacing: "0.12em", color: C.textDim }}>DROP DRAFT — .docx</div>
                <div style={{ fontSize: "11px", color: C.textMuted, marginTop: "4px", fontStyle: "italic", fontFamily: "IM Fell English, serif" }}>
                  Each upload starts a new session. History is preserved.
                </div>
              </div>
              <div style={{ fontSize: "11px", color: C.textMuted }}>
                {uploadTime ? `Last run: ${uploadTime}s` : "Click or drag"}
              </div>
            </>
          )}
        </div>

        {/* Empty state */}
        {!activeSession && !uploading && (
          <div style={{ textAlign: "center", padding: "80px 32px", color: C.textMuted }}>
            <div style={{ fontSize: "40px", marginBottom: "16px", opacity: 0.2 }}>◉</div>
            <div style={{ fontSize: "14px", letterSpacing: "0.2em" }}>NO DRAFT LOADED</div>
            <div style={{ fontSize: "12px", marginTop: "8px", fontFamily: "IM Fell English, serif", fontStyle: "italic" }}>
              Upload a manuscript to initialize the cockpit
            </div>
          </div>
        )}

        {/* Dashboard */}
        {activeSession && activeTab === "dashboard" && (
          <div>
            <div style={{ fontSize: "10px", letterSpacing: "0.2em", color: C.textMuted, marginBottom: "16px" }}>
              DRAFT HEALTH — {activeSession.draftName} — {activeSession.wordCount?.toLocaleString()} words
            </div>
            <HealthBar score={activeSession.health} prevScore={prevSession?.health} />

            <div style={{ display: "flex", gap: "40px", marginBottom: "32px", paddingBottom: "24px", borderBottom: `1px solid ${C.smokeLight}` }}>
              <div>
                <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.textMuted, marginBottom: "4px" }}>OPEN FINDINGS</div>
                <div style={{ fontSize: "32px", fontWeight: 700 }}>{totalFindings.toLocaleString()}</div>
              </div>
              {prevSession && totalPrev > totalFindings && (
                <div>
                  <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.textMuted, marginBottom: "4px" }}>RESOLVED SINCE {prevSession.draftName}</div>
                  <div style={{ fontSize: "32px", fontWeight: 700, color: C.pass }}>{(totalPrev - totalFindings).toLocaleString()}</div>
                </div>
              )}
              {prevSession && totalFindings > totalPrev && (
                <div>
                  <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.regress, marginBottom: "4px" }}>⚠ REGRESSION</div>
                  <div style={{ fontSize: "32px", fontWeight: 700, color: C.regress }}>+{(totalFindings - totalPrev).toLocaleString()}</div>
                </div>
              )}
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: "12px", marginBottom: "32px" }}>
              {SIGNALS.map(sig => (
                <SignalCard key={sig.id} signal={sig} count={activeSession.counts[sig.id] || 0} prevCount={prevSession?.counts[sig.id]} />
              ))}
            </div>

            <button onClick={() => exportToXlsx(activeSession.findings || [], activeSession.draftName)} style={{
              background: "transparent", border: `1px solid ${C.amber}`, color: C.amber,
              padding: "10px 24px", fontSize: "12px", letterSpacing: "0.15em",
              cursor: "pointer", fontFamily: "Barlow Condensed, sans-serif", textTransform: "uppercase",
            }}>
              EXPORT FINDINGS
            </button>
          </div>
        )}

        {/* Findings */}
        {activeSession && activeTab === "findings" && (
          <div>
            <div style={{ display: "flex", gap: "8px", marginBottom: "20px", flexWrap: "wrap", alignItems: "center" }}>
              <select value={filterSignal} onChange={e => setFilterSignal(e.target.value)} style={{
                background: C.smokeDark, border: `1px solid ${C.smokeLight}`, color: C.parchment,
                padding: "8px 12px", fontSize: "12px", fontFamily: "Barlow Condensed, sans-serif", letterSpacing: "0.1em", cursor: "pointer",
              }}>
                <option value="all">ALL SIGNALS ({totalFindings})</option>
                {SIGNALS.map(sig => (
                  <option key={sig.id} value={sig.id}>{sig.short} ({activeSession.counts[sig.id] || 0})</option>
                ))}
              </select>
              <input value={findingsSearch} onChange={e => setFindingsSearch(e.target.value)}
                placeholder="Search findings…"
                style={{
                  background: C.smokeDark, border: `1px solid ${C.smokeLight}`, color: C.parchment,
                  padding: "8px 16px", fontSize: "12px", fontFamily: "Barlow Condensed, sans-serif",
                  flex: 1, minWidth: "200px", outline: "none",
                }} />
              <span style={{ fontSize: "11px", color: C.textMuted }}>{filteredFindings.length.toLocaleString()} shown</span>
              <button onClick={() => exportToXlsx(filteredFindings, `${activeSession.draftName}_${filterSignal}`)} style={{
                background: "transparent", border: `1px solid ${C.smokeLight}`, color: C.textDim,
                padding: "8px 16px", fontSize: "11px", letterSpacing: "0.12em",
                cursor: "pointer", fontFamily: "Barlow Condensed, sans-serif",
              }}>
                EXPORT VIEW
              </button>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
              {filteredFindings.slice(0, 500).map((f, i) => (
                <div key={i} style={{
                  display: "grid", gridTemplateColumns: "160px 100px 1fr", gap: "0",
                  background: i % 2 === 0 ? C.smokeDark : "transparent",
                  padding: "8px 12px", fontSize: "12px", alignItems: "start",
                }}>
                  <div style={{ color: C.textMuted, fontSize: "10px", letterSpacing: "0.08em", paddingTop: "2px" }}>
                    {f.chapter.slice(0, 24)}
                  </div>
                  <div style={{
                    background: C.smokeLight, color: C.amber, fontSize: "9px",
                    letterSpacing: "0.1em", padding: "2px 6px", alignSelf: "start",
                    display: "inline-block", whiteSpace: "nowrap",
                  }}>
                    {f.issue}
                  </div>
                  <div style={{ color: C.dimParch, fontFamily: "IM Fell English, serif", fontSize: "13px", lineHeight: 1.5 }}>
                    {f.sentence.slice(0, 300)}
                  </div>
                </div>
              ))}
              {filteredFindings.length > 500 && (
                <div style={{ textAlign: "center", padding: "16px", color: C.textMuted, fontSize: "11px" }}>
                  Showing 500 of {filteredFindings.length.toLocaleString()} — export xlsx for full list
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Health Timeline */}
      {sessions.length >= 2 && activeSession && (
        <div style={{ borderTop: `1px solid ${C.smokeLight}`, padding: "16px 32px", display: "flex", gap: "32px", alignItems: "center" }}>
          <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.textMuted, flexShrink: 0 }}>HEALTH TIMELINE</div>
          {sessions.map(s => (
            <div key={s.id} onClick={() => setActiveSession(s)} style={{
              cursor: "pointer", textAlign: "center",
              opacity: s.id === activeSession.id ? 1 : 0.45,
            }}>
              <div style={{ fontSize: "20px", fontWeight: 700, color: C.amber, fontFamily: "Barlow Condensed, sans-serif" }}>{s.health}</div>
              <div style={{ fontSize: "9px", color: C.textMuted, letterSpacing: "0.08em" }}>{s.draftName}</div>
            </div>
          ))}
        </div>
      )}

      <style>{`
        @keyframes pulse {
          0%, 100% { opacity: 0.4; transform: translateX(-100%); }
          50% { opacity: 1; transform: translateX(250%); }
        }
        input::placeholder { color: #8A8070; }
        select option { background: #2A2A2A; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #1C1C1C; }
        ::-webkit-scrollbar-thumb { background: #3A3938; }
      `}</style>
    </div>
  );
}

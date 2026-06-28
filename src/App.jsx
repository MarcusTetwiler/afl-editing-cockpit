import { useState, useEffect, useRef, useCallback } from "react";
import mammoth from "mammoth";
import * as XLSX from "xlsx";

// ── Palette ───────────────────────────────────────────────────────────────────
const C = {
  charcoal:   "#1C1C1C", parchment:  "#F2ECD8", amber:      "#D4820A",
  redOrange:  "#C0392B", dimParch:   "#E8DFC4", smokeDark:  "#2A2A2A",
  smokeLight: "#3A3938", textMuted:  "#8A8070", textDim:    "#B8AC96",
  pass:       "#4A7C59", regress:    "#C0392B",
};

// ═══════════════════════════════════════════════════════════════════════════════
// LAYER 1 — SENTENCE CLASSIFIER
// ═══════════════════════════════════════════════════════════════════════════════

const ATTRIBUTION_VERBS = new Set([
  "said","called","asked","replied","added","continued","murmured","answered",
  "shouted","whispered","snapped","growled","began","finished","offered","noted",
  "admitted","insisted","suggested","repeated","demanded","confirmed","reported",
  "stated","announced","declared","explained","responded","returned","countered",
  "cut","ordered","urged","warned","reminded","prompted","pressed","pushed",
  "breathed","managed","allowed","agreed","disagreed","objected","conceded",
  "laughed","sighed","grunted","barked","hissed","spat","called","echoed",
]);

const HEADER_PATTERNS = [
  /^[A-Z][A-Z\s\d,.'"\-–—]{8,}$/, // ALL CAPS line
  /^(chapter|prologue|epilogue|part|interlude)\s/i,
  /^(andrew|elena|reese|logan|jack|anika|ryba|rupert)\s+(i{1,3}|iv|v|vi{0,3}|ix|x|\d+)$/i,
  /^(andrew|elena|reese|logan|jack|anika|ryba|rupert)$/i,
];

const DISPLAY_TEXT_PATTERNS = [
  /^\[.*\]$/, // [SYSTEM TEXT]
  /^[A-Z\s\d:\/\-–—]{10,}$/, // TIER 3 REFERRAL FLAG
  /REASSESSMENT|REFERRAL FLAG|ADVISORY|NOTICE:|ALERT:|STATUS:/,
  /^\d{1,2}:\d{2}/, // timestamps
];

const FRONT_MATTER_WORDS = new Set([
  "acknowledgments","acknowledgements","dedication","epigraph","preface",
  "foreword","author","copyright","isbn","reserved","rights","published",
]);

// Strip display text markers from a sentence
function stripDisplayMarkers(s) {
  return s.replace(/\[.*?\]/g, "").replace(/^[A-Z\s]{10,}:/,"").trim();
}

// Extract action beat from a dialogue sentence
// Returns { quote, attribution_verb, action_beat, action_word_count }
function parseDialogueSentence(s) {
  // Find last closing quote
  const quoteEnd = Math.max(s.lastIndexOf('"'), s.lastIndexOf('\u201d'));
  if (quoteEnd === -1) return null;

  const afterQuote = s.slice(quoteEnd + 1).trim().replace(/^[,.]/, "").trim();
  if (!afterQuote) return { quote: s, attribution_verb: null, action_beat: "", action_word_count: 0 };

  const words = afterQuote.split(/\s+/).filter(Boolean);
  // Find attribution verb (usually first or second word: "he said" / "said Andrew")
  let attrIdx = -1;
  for (let i = 0; i < Math.min(words.length, 3); i++) {
    const w = words[i].toLowerCase().replace(/[^a-z]/g, "");
    if (ATTRIBUTION_VERBS.has(w)) { attrIdx = i; break; }
  }

  // Remove subject + attribution verb from action beat
  let actionStart = attrIdx === -1 ? 0 : attrIdx + 1;
  // Also skip pronoun before verb: "he said" → attrIdx=1 → skip words[0] too
  if (attrIdx > 0) actionStart = attrIdx + 1;
  const actionBeat = words.slice(actionStart).join(" ");
  const action_word_count = actionBeat.split(/\s+/).filter(Boolean).length;

  return {
    quote: s.slice(0, quoteEnd + 1),
    attribution_verb: attrIdx >= 0 ? words[attrIdx] : null,
    action_beat: actionBeat,
    action_word_count,
  };
}

// Classify a single sentence
function classifySentence(s) {
  const trimmed = s.trim();
  if (!trimmed) return "EMPTY";

  // Front matter
  const lower = trimmed.toLowerCase();
  if (FRONT_MATTER_WORDS.has(lower.split(/\s+/)[0])) return "FRONT_MATTER";

  // Header
  for (const pat of HEADER_PATTERNS) {
    if (pat.test(trimmed)) return "HEADER";
  }
  // Short all-caps (chapter labels)
  if (/^[A-Z\s]{3,40}$/.test(trimmed) && trimmed.split(" ").length <= 6) return "HEADER";

  // Display text
  for (const pat of DISPLAY_TEXT_PATTERNS) {
    if (pat.test(trimmed)) return "DISPLAY_TEXT";
  }

  // Does it contain dialogue?
  const hasOpenQuote  = /[""\u201c]/.test(trimmed);
  const hasCloseQuote = /[""\u201d]/.test(trimmed);

  if (!hasOpenQuote && !hasCloseQuote) return "NARRATION";

  // Pure dialogue: starts and ends with quotes, minimal outside content
  const startsWithQuote = /^[""\u201c]/.test(trimmed);
  const endsWithQuote   = /[""\u201d][.!?]?\s*$/.test(trimmed);

  if (startsWithQuote && endsWithQuote) return "PURE_DIALOGUE";

  if (startsWithQuote) {
    const parsed = parseDialogueSentence(trimmed);
    if (!parsed) return "PURE_DIALOGUE";
    if (parsed.action_word_count >= 5) return "DIALOGUE_ACTION";
    return "DIALOGUE_TAG";
  }

  // Narration with embedded quote
  return "NARRATION";
}

// ═══════════════════════════════════════════════════════════════════════════════
// LAYER 2 — ZONE DETECTION
// ═══════════════════════════════════════════════════════════════════════════════

function detectZone(line, currentZone) {
  const t = line.trim();
  if (!t) return currentZone;

  // Chapter boundaries reset to SCENE
  for (const pat of HEADER_PATTERNS) {
    if (pat.test(t)) return "SCENE";
  }
  if (/^(title|byline|by\s+\w)/i.test(t)) return "FRONT_MATTER";

  return currentZone;
}

// ═══════════════════════════════════════════════════════════════════════════════
// LAYER 3 — SIGNALS
// ═══════════════════════════════════════════════════════════════════════════════

// Disposition constants
const D = {
  ACTIONABLE: "actionable",
  REVIEW:     "review_only",
  EXCLUDED:   "excluded",
};

// ── Signal helpers ─────────────────────────────────────────────────────────────

const INTERIORITY_COGNITION_VERBS = new Set([
  "knew","thought","wondered","understood","believed","hoped","feared",
  "wished","sensed","remembered","recognized","considered","assumed",
  "supposed","imagined","realized",
]);
// These only count when subject is human/POV character
const INTERIORITY_CONDITIONAL = new Set(["felt","decided","expected","needed","wanted"]);

const POV_NAMES = new Set(["andrew","elena","reese","logan","jack","anika","ryba","rupert"]);
const POV_PRONOUNS = new Set(["he","she","they","his","her","their"]);
const NON_HUMAN_SUBJECTS = new Set([
  "system","plan","route","amount","road","machine","drone","signal","mission",
  "screen","display","unit","command","protocol","schedule","sequence","process",
  "device","network","camera","sensor","data","feed","broadcast","channel",
]);

function isHumanSubject(subjectWord) {
  const w = (subjectWord || "").toLowerCase().replace(/[^a-z]/g,"");
  if (POV_PRONOUNS.has(w)) return true;
  if (POV_NAMES.has(w)) return true;
  if (NON_HUMAN_SUBJECTS.has(w)) return false;
  return true; // default assume human
}

function detectInteriority(sentence) {
  const words = sentence.split(/\s+/);
  for (let i = 0; i < words.length - 1; i++) {
    const w = words[i].toLowerCase().replace(/[^a-z]/g,"");
    const subj = i > 0 ? words[i-1].toLowerCase().replace(/[^a-z]/g,"") : "";

    if (INTERIORITY_COGNITION_VERBS.has(w)) {
      return { found: true, confidence: 90 };
    }
    if (INTERIORITY_CONDITIONAL.has(w) && isHumanSubject(subj)) {
      // Extra check: "he needed to", "she decided that" = interiority
      // "the amount needed", "route needed" = not
      const next = (words[i+1] || "").toLowerCase().replace(/[^a-z]/g,"");
      if (["to","that","it","him","her","them","more","less","one"].includes(next)) {
        return { found: true, confidence: 75 };
      }
    }
  }
  // Check possessive + mind/thought
  if (/\b(his|her|their)\s+(mind|thoughts?|feelings?|emotions?|instinct|gut|conscience)\b/i.test(sentence)) {
    return { found: true, confidence: 85 };
  }
  return { found: false };
}

const FILTER_VERBS = new Set([
  "saw","looked","watched","noticed","realized","felt","heard","smelled",
  "tasted","sensed","observed","spotted","glimpsed","perceived","detected",
]);

// GPS-Lite patterns
function detectGPS(sentence, wordCount) {
  const reasons = [];
  const lower = sentence.toLowerCase();

  // 1. Abstract subject + vague verb ("views were", "concerns remained")
  if (/\b(views?|concerns?|issues?|matters?|aspects?|elements?)\s+(were|are|remain|seemed|appeared|became)\b/i.test(sentence)) {
    reasons.push({ type: "ABSTRACT_SUBJECT", confidence: 70, reason: "Abstract subject with vague state verb" });
  }

  // 2. Repeated noun causing reparse
  const nounMatches = sentence.match(/\b(\w{4,})\b.*?\b\1\b/gi) || [];
  for (const m of nounMatches) {
    const word = m.split(/\s+/)[0].toLowerCase();
    if (!["that","this","with","from","they","their","there","when","then","than","have","been","were","will"].includes(word)) {
      reasons.push({ type: "REPEATED_NOUN", confidence: 60, reason: `Repeated noun "${word}" may cause reparse` });
      break;
    }
  }

  // 3. Missing/delayed referent: "the next dropped", "the first showed"
  if (/\bthe\s+(next|first|last|second|third|former|latter)\s+(dropped|showed|hit|fell|came|went|moved|turned|cut)\b/i.test(sentence)) {
    reasons.push({ type: "DELAYED_REFERENT", confidence: 85, reason: "Ordinal/determiner without clear referent" });
  }

  // 4. Overloaded: 3+ clauses (as/while/with/when chains)
  const clauseCount = (sentence.match(/\b(as|while|when|although|though|because|since|after|before|until|unless|despite|who|which|that)\b/gi) || []).length;
  if (clauseCount >= 3 && wordCount > 25) {
    reasons.push({ type: "OVERLOADED_PATH", confidence: 65, reason: `${clauseCount} subordinate clauses in one sentence` });
  }

  // 5. Participial phrase attachment risk (opening -ing phrase, long subject gap)
  if (/^[A-Z][a-z]+ing\b/.test(sentence) && wordCount > 12) {
    reasons.push({ type: "PARTICIPIAL_RISK", confidence: 60, reason: "Opening participial phrase — attachment may be ambiguous" });
  }

  // 6. Action verb + misleading object pairing
  const misleadingPairs = [
    [/\bdropped\s+his\s+(chair|seat|position|post)\b/i, "dropped + body-position noun reads as physical drop"],
    [/\btook\s+the\s+(floor|ground|field|stage)\b/i, "took + spatial noun may mislead first parse"],
    [/\bheld\s+the\s+(line|floor|room|stage|court)\b/i, "held + abstract location may mislead"],
    [/\bkilled\s+the\s+(lights?|sound|feed|signal|line)\b/i, "killed + tech noun — violence read before correction"],
  ];
  for (const [pat, reason] of misleadingPairs) {
    if (pat.test(sentence)) {
      reasons.push({ type: "MISLEADING_PAIR", confidence: 80, reason });
    }
  }

  // 7. as/while/with chain over threshold
  const chainWords = (sentence.match(/\b(as|while|with)\b/gi) || []).length;
  if (chainWords >= 3) {
    reasons.push({ type: "CHAIN_OVERLOAD", confidence: 65, reason: `${chainWords}x as/while/with in one sentence` });
  }

  if (reasons.length === 0) return null;
  // Return highest confidence finding
  reasons.sort((a,b) => b.confidence - a.confidence);
  return reasons[0];
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN ANALYSIS ENGINE
// ═══════════════════════════════════════════════════════════════════════════════

const SIGNAL_IDS = ["seven_word","interiority","pronoun_ambiguity","repeated_openers","flat_wave","filter_verb","gps_lite"];

function runSignals(rawText) {
  const allFindings = [];
  const lines = rawText.split(/\n/);

  let currentChapter = "Opening";
  let currentZone = "FRONT_MATTER";
  let paraIndex = 0;
  let inFrontMatter = true;

  // Paragraph buffer
  const paragraphs = [];
  let currentPara = [];

  for (const line of lines) {
    const t = line.trim();
    if (!t) {
      if (currentPara.length) { paragraphs.push({ lines: currentPara, chapter: currentChapter }); currentPara = []; }
      continue;
    }

    // Zone/chapter update
    for (const pat of HEADER_PATTERNS) {
      if (pat.test(t) && t.length < 80) {
        currentChapter = t.replace(/\s+/g," ");
        inFrontMatter = false;
        currentZone = "SCENE";
        break;
      }
    }
    if (inFrontMatter && t.length < 60 && !/[.!?]$/.test(t)) {
      currentZone = "FRONT_MATTER";
    }

    currentPara.push({ text: t, chapter: currentChapter, zone: currentZone });
  }
  if (currentPara.length) paragraphs.push({ lines: currentPara, chapter: currentChapter });

  for (const para of paragraphs) {
    paraIndex++;
    const chapter = para.lines[0]?.chapter || currentChapter;
    const zone = para.lines[0]?.zone || "SCENE";
    const fullText = para.lines.map(l => l.text).join(" ");

    if (zone === "FRONT_MATTER") continue;
    if (fullText.trim().length < 15) continue;

    // Split into sentences
    const rawSentences = fullText.match(/[^.!?]+[.!?]+(?:\s|$)|[^.!?]+$/g) || [fullText];
    const sentences = rawSentences.map(s => s.trim()).filter(Boolean);

    // ── Flat Sentence Wave (paragraph-level) ──────────────────────────────────
    if (sentences.length >= 4) {
      const narrationSents = sentences.filter(s => {
        const cls = classifySentence(s);
        return cls === "NARRATION" || cls === "DIALOGUE_ACTION";
      });
      if (narrationSents.length >= 4) {
        const lengths = narrationSents.map(s => s.split(/\s+/).length);
        const mean = lengths.reduce((a,b)=>a+b,0)/lengths.length;
        const variance = lengths.reduce((a,b)=>a+Math.pow(b-mean,2),0)/lengths.length;
        const stdDev = Math.sqrt(variance);
        if (stdDev < 3.5 && mean > 8) {
          allFindings.push({
            signal_id: "flat_wave",
            chapter,
            sentence: `[Para ${paraIndex}] ${fullText.slice(0,130)}…`,
            issue: "FLAT SENTENCE WAVE",
            disposition: D.REVIEW,
            confidence: 70,
            reason: `StdDev ${stdDev.toFixed(1)} across ${narrationSents.length} narration sentences (mean ${mean.toFixed(1)} words)`,
          });
        }
      }
    }

    // ── Repeated Openers (paragraph-level) ───────────────────────────────────
    if (sentences.length >= 3) {
      const openers = sentences.map(s => s.trim().split(/\s+/)[0]?.toLowerCase().replace(/[^a-z]/g,"") || "");
      let streak = 1;
      for (let i = 1; i < openers.length; i++) {
        if (openers[i] && openers[i] === openers[i-1] && openers[i].length > 1) {
          streak++;
          if (streak >= 3) {
            allFindings.push({
              signal_id: "repeated_openers",
              chapter,
              sentence: sentences.slice(i-2, i+1).join("  "),
              issue: "REPEATED OPENERS",
              disposition: D.ACTIONABLE,
              confidence: 90,
              reason: `"${openers[i]}" opens ${streak} consecutive sentences`,
            });
            streak = 1;
          }
        } else { streak = 1; }
      }
    }

    // ── Sentence-level signals ────────────────────────────────────────────────
    for (const rawSent of sentences) {
      const s = rawSent.trim();
      if (!s || s.length < 4) continue;

      const cls = classifySentence(s);
      const wordCount = s.split(/\s+/).filter(Boolean).length;

      // Exclude headers, front matter, display text from all signals
      if (["HEADER","FRONT_MATTER","DISPLAY_TEXT","EMPTY"].includes(cls)) continue;

      // ── 7-Word Narration Rule ─────────────────────────────────────────────
      if (cls === "NARRATION" && wordCount < 7 && wordCount > 0) {
        allFindings.push({
          signal_id: "seven_word",
          chapter,
          sentence: s,
          issue: "7-WORD NARRATION RULE",
          disposition: D.ACTIONABLE,
          confidence: 85,
          reason: `${wordCount}-word narration sentence`,
        });
      } else if (cls === "DIALOGUE_ACTION") {
        const parsed = parseDialogueSentence(s);
        if (parsed && parsed.action_word_count > 0 && parsed.action_word_count < 7) {
          allFindings.push({
            signal_id: "seven_word",
            chapter,
            sentence: s,
            issue: "7-WORD NARRATION RULE",
            disposition: D.REVIEW,
            confidence: 60,
            reason: `Action beat is ${parsed.action_word_count} words: "${parsed.action_beat}"`,
          });
        }
      }

      // ── Interiority Leak ──────────────────────────────────────────────────
      if (cls === "NARRATION") {
        const result = detectInteriority(s);
        if (result.found) {
          allFindings.push({
            signal_id: "interiority",
            chapter,
            sentence: s,
            issue: "INTERIORITY LEAK",
            disposition: result.confidence >= 85 ? D.ACTIONABLE : D.REVIEW,
            confidence: result.confidence,
            reason: `Cognition verb detected in narration`,
          });
        }
      }

      // ── Pronoun Ambiguity ─────────────────────────────────────────────────
      if (cls === "NARRATION") {
        const pronounPats = [
          /\b(it|this|that)\s+(was|is|had|has|did|does|would|could|should|seemed|looked|felt|appeared)\b/gi,
          /^(It|This|That)\s+(was|is|had|seemed|would|felt|appeared)/,
        ];
        for (const pat of pronounPats) {
          pat.lastIndex = 0;
          if (pat.test(s)) {
            allFindings.push({
              signal_id: "pronoun_ambiguity",
              chapter,
              sentence: s,
              issue: "PRONOUN AMBIGUITY",
              disposition: D.REVIEW,
              confidence: 70,
              reason: "Vague pronoun with unclear referent",
            });
            break;
          }
        }
      }

      // ── Filter Verb ───────────────────────────────────────────────────────
      if (cls === "NARRATION" || cls === "DIALOGUE_ACTION") {
        const target = cls === "DIALOGUE_ACTION" ? (parseDialogueSentence(s)?.action_beat || s) : s;
        const filterPat = /\b(saw|looked|watched|noticed|realized|felt|heard|smelled|tasted|sensed|observed|spotted|glimpsed|perceived|detected)\b/gi;
        filterPat.lastIndex = 0;
        if (filterPat.test(target)) {
          allFindings.push({
            signal_id: "filter_verb",
            chapter,
            sentence: s,
            issue: "FILTER VERB",
            disposition: D.REVIEW,
            confidence: 75,
            reason: "Perception/filter verb in narration",
          });
        }
      }

      // ── GPS-Lite ──────────────────────────────────────────────────────────
      if (cls === "NARRATION" && wordCount >= 8) {
        const gps = detectGPS(s, wordCount);
        if (gps && gps.confidence >= 60) {
          allFindings.push({
            signal_id: "gps_lite",
            chapter,
            sentence: s,
            issue: `GPS: ${gps.type.replace(/_/g," ")}`,
            disposition: gps.confidence >= 75 ? D.ACTIONABLE : D.REVIEW,
            confidence: gps.confidence,
            reason: gps.reason,
          });
        }
      }
    }
  }

  return allFindings;
}

// ═══════════════════════════════════════════════════════════════════════════════
// SIGNAL DEFINITIONS (for UI)
// ═══════════════════════════════════════════════════════════════════════════════

const SIGNALS = [
  { id: "seven_word",        name: "7-Word Narration Rule",         short: "7-WORD",       description: "Narration sentences under 7 words (excludes dialogue, headers, display text)" },
  { id: "interiority",       name: "Interiority Leak",              short: "INTERIORITY",  description: "Stated thoughts, feelings, cognition — human subject + cognition verb" },
  { id: "pronoun_ambiguity", name: "Pronoun Ambiguity",             short: "PRONOUN",      description: 'Vague "it," "this," "that" with unclear referent in narration' },
  { id: "repeated_openers",  name: "Repeated Openers",              short: "OPENERS",      description: "3+ consecutive sentences opening with the same word" },
  { id: "flat_wave",         name: "Flat Sentence Wave",            short: "FLAT WAVE",    description: "Paragraphs with monotonous sentence-length clustering" },
  { id: "filter_verb",       name: "Filter Verb / Perception Leak", short: "FILTER VERB",  description: '"saw," "looked," "watched," "noticed" in narration' },
  { id: "gps_lite",          name: "GPS-Lite: Garden-Path Scanner", short: "GPS-LITE",     description: "Heuristic first-parse friction — misleading verbs, delayed referents, clause overload" },
];

// ═══════════════════════════════════════════════════════════════════════════════
// EXPORT
// ═══════════════════════════════════════════════════════════════════════════════

function exportToXlsx(findings, draftName) {
  const rows = [["Chapter / Location", "Flagged Sentence", "Issue Type", "Disposition", "Confidence", "Reason"]];
  for (const f of findings) {
    rows.push([f.chapter, f.sentence, f.issue, f.disposition || "", f.confidence || "", f.reason || ""]);
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [20,18,60,16,12,12,40].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Findings");
  XLSX.writeFile(wb, `${draftName}_findings.xlsx`);
}

// ═══════════════════════════════════════════════════════════════════════════════
// PERSISTENCE
// ═══════════════════════════════════════════════════════════════════════════════

const STORAGE_KEY = "afl_cockpit_v2";

function loadSessions() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]"); } catch { return []; }
}
function saveSessions(sessions) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(
      sessions.map(s => ({ ...s, findings: s.findings?.slice(0,2000) }))
    ));
  } catch { console.warn("localStorage quota hit"); }
}

function computeHealth(findings, wordCount) {
  if (!wordCount) return 0;
  const actionable = findings.filter(f => f.disposition === D.ACTIONABLE).length;
  const rate = (actionable / wordCount) * 1000;
  return Math.round(Math.max(40, Math.min(100, 100 - rate * 0.55)) * 10) / 10;
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMPONENTS
// ═══════════════════════════════════════════════════════════════════════════════

function SignalCard({ signal, actionable, raw, prevActionable, onClick }) {
  const delta = prevActionable !== undefined ? actionable - prevActionable : null;
  const isPass = actionable === 0;
  const isRegress = delta !== null && delta > 0;
  const excluded = raw - actionable;

  return (
    <div onClick={onClick} style={{
      background: C.smokeDark,
      border: `1px solid ${isRegress ? C.regress : isPass ? C.pass : C.smokeLight}`,
      padding: "16px 20px", position: "relative", cursor: "pointer",
      transition: "border-color 0.15s",
    }}
    onMouseEnter={e => e.currentTarget.style.borderColor = C.amber}
    onMouseLeave={e => e.currentTarget.style.borderColor = isRegress ? C.regress : isPass ? C.pass : C.smokeLight}
    >
      {isRegress && <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "2px", background: C.regress }} />}
      <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.textMuted, fontFamily: "Barlow Condensed, sans-serif", marginBottom: "8px" }}>
        {signal.short}
      </div>
      <div style={{ display: "flex", alignItems: "flex-end", gap: "10px" }}>
        <div style={{ fontSize: "36px", fontFamily: "Barlow Condensed, sans-serif", fontWeight: 700, lineHeight: 1,
          color: isPass ? C.pass : isRegress ? C.regress : C.parchment }}>
          {actionable.toLocaleString()}
        </div>
        {delta !== null && (
          <div style={{ fontSize: "12px", fontFamily: "Barlow Condensed, sans-serif", color: isRegress ? C.regress : C.pass, paddingBottom: "4px" }}>
            {isRegress ? `▲ +${delta}` : delta < 0 ? `▼ ${delta}` : "—"}
          </div>
        )}
        {isPass && <div style={{ fontSize: "11px", color: C.pass, paddingBottom: "4px", fontFamily: "Barlow Condensed, sans-serif" }}>✓ PASS</div>}
      </div>
      {excluded > 0 && (
        <div style={{ fontSize: "10px", color: C.textMuted, marginTop: "4px", fontFamily: "Barlow Condensed, sans-serif" }}>
          {excluded.toLocaleString()} excluded / low-conf
        </div>
      )}
      <div style={{ fontSize: "11px", color: C.textDim, marginTop: "8px", fontFamily: "IM Fell English, serif", fontStyle: "italic", lineHeight: 1.3 }}>
        {signal.description}
      </div>
    </div>
  );
}

function HealthBar({ score, prevScore }) {
  const delta = prevScore !== undefined ? (score - prevScore).toFixed(1) : null;
  const pct = Math.max(0, Math.min(100, ((score - 40) / 60) * 100));
  return (
    <div style={{ marginBottom: "28px" }}>
      <div style={{ display: "flex", alignItems: "baseline", gap: "16px", marginBottom: "10px" }}>
        <div style={{ fontSize: "64px", fontFamily: "Barlow Condensed, sans-serif", fontWeight: 700, color: C.amber, lineHeight: 1 }}>{score}</div>
        <div>
          <div style={{ fontSize: "14px", color: C.textMuted, letterSpacing: "0.12em", fontFamily: "Barlow Condensed, sans-serif" }}>/ 100</div>
          {delta !== null && (
            <div style={{ fontSize: "15px", fontFamily: "Barlow Condensed, sans-serif", color: parseFloat(delta) >= 0 ? C.pass : C.regress }}>
              {parseFloat(delta) >= 0 ? `▲ +${delta}` : `▼ ${delta}`} since previous
            </div>
          )}
        </div>
      </div>
      <div style={{ height: "3px", background: C.smokeLight }}>
        <div style={{ height: "100%", width: `${pct}%`, background: `linear-gradient(90deg, ${C.amber}, ${C.redOrange})` }} />
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════════════════

export default function App() {
  const [sessions, setSessions]           = useState([]);
  const [activeSession, setActiveSession] = useState(null);
  const [uploading, setUploading]         = useState(false);
  const [uploadStep, setUploadStep]       = useState("");
  const [uploadTime, setUploadTime]       = useState(null);
  const [activeTab, setActiveTab]         = useState("dashboard");
  const [filterSignal, setFilterSignal]   = useState("all");
  const [filterDisp, setFilterDisp]       = useState("actionable");
  const [search, setSearch]               = useState("");
  const fileInputRef = useRef();

  useEffect(() => {
    const saved = loadSessions();
    if (saved.length) { setSessions(saved); setActiveSession(saved[saved.length - 1]); }
  }, []);

  const handleUpload = useCallback(async (file) => {
    if (!file?.name.endsWith(".docx")) { alert("Please upload a .docx file."); return; }
    const draftName = file.name.replace(".docx","");
    setUploading(true);
    const t0 = Date.now();
    try {
      setUploadStep("Parsing manuscript…");
      const ab = await file.arrayBuffer();
      const { value: rawText } = await mammoth.extractRawText({ arrayBuffer: ab });

      setUploadStep("Classifying sentences…");
      await new Promise(r => setTimeout(r, 80));

      setUploadStep("Running 7 signal passes…");
      await new Promise(r => setTimeout(r, 80));
      const findings = runSignals(rawText);

      setUploadStep("Computing health score…");
      const wordCount = rawText.split(/\s+/).filter(Boolean).length;
      const health = computeHealth(findings, wordCount);

      // Counts: actionable only for dashboard; raw for signal cards
      const counts = {}, rawCounts = {};
      for (const sig of SIGNALS) { counts[sig.id] = 0; rawCounts[sig.id] = 0; }
      for (const f of findings) {
        rawCounts[f.signal_id] = (rawCounts[f.signal_id]||0) + 1;
        if (f.disposition !== D.EXCLUDED) counts[f.signal_id] = (counts[f.signal_id]||0) + 1;
      }

      const session = { id: Date.now(), draftName, uploadedAt: new Date().toISOString(), wordCount, health, counts, rawCounts, findings };

      setSessions(prev => {
        const next = [...prev, session];
        saveSessions(next);
        return next;
      });
      setActiveSession(session);
      setUploadTime(((Date.now()-t0)/1000).toFixed(1));
      setActiveTab("dashboard");
    } catch (err) {
      alert("Error: " + err.message);
    } finally {
      setUploading(false); setUploadStep("");
    }
  }, []);

  const handleDrop = useCallback(e => {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f) handleUpload(f);
  }, [handleUpload]);

  const activeIdx  = sessions.findIndex(s => s.id === activeSession?.id);
  const prevSess   = activeIdx > 0 ? sessions[activeIdx - 1] : null;

  const totalActionable = activeSession ? Object.values(activeSession.counts).reduce((a,b)=>a+b,0) : 0;
  const totalRaw        = activeSession ? Object.values(activeSession.rawCounts||{}).reduce((a,b)=>a+b,0) : 0;
  const prevTotal       = prevSess ? Object.values(prevSess.counts).reduce((a,b)=>a+b,0) : null;

  const filteredFindings = (activeSession?.findings || []).filter(f => {
    if (filterSignal !== "all" && f.signal_id !== filterSignal) return false;
    if (filterDisp !== "all" && f.disposition !== filterDisp) return false;
    if (search && !f.sentence.toLowerCase().includes(search.toLowerCase()) && !f.chapter.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  const dispColor = { actionable: C.amber, review_only: C.textDim, excluded: C.textMuted };

  return (
    <div style={{ minHeight: "100vh", background: C.charcoal, color: C.parchment, fontFamily: "Barlow Condensed, sans-serif" }}>
      {/* Header */}
      <div style={{ borderBottom: `1px solid ${C.smokeLight}`, padding: "0 32px", display: "flex", alignItems: "center", justifyContent: "space-between", height: "56px" }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: "16px" }}>
          <span style={{ fontSize: "18px", fontWeight: 700, letterSpacing: "0.15em", color: C.amber }}>AFL</span>
          <span style={{ fontSize: "14px", letterSpacing: "0.2em", color: C.textMuted }}>EDITING COCKPIT</span>
        </div>
        <div style={{ display: "flex", gap: "4px" }}>
          {["dashboard","findings"].map(tab => (
            <button key={tab} onClick={() => setActiveTab(tab)} style={{
              background: activeTab === tab ? C.smokeLight : "transparent", border: "none",
              color: activeTab === tab ? C.parchment : C.textMuted, padding: "8px 16px",
              fontSize: "12px", letterSpacing: "0.15em", cursor: "pointer",
              fontFamily: "Barlow Condensed, sans-serif", textTransform: "uppercase",
            }}>
              {tab}
              {tab === "findings" && activeSession && (
                <span style={{ marginLeft: "8px", background: C.amber, color: C.charcoal, padding: "1px 6px", fontSize: "10px", fontWeight: 700 }}>
                  {totalActionable.toLocaleString()}
                </span>
              )}
            </button>
          ))}
        </div>
      </div>

      <div style={{ padding: "32px", maxWidth: "1140px", margin: "0 auto" }}>
        {/* Version Rail */}
        {sessions.length > 0 && (
          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap", marginBottom: "24px" }}>
            {sessions.map(s => (
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
        <div onDrop={handleDrop} onDragOver={e => e.preventDefault()}
          onClick={() => !uploading && fileInputRef.current?.click()}
          style={{ border: `1px dashed ${C.smokeLight}`, padding: "18px 28px", marginBottom: "32px",
            cursor: uploading ? "default" : "pointer", display: "flex", alignItems: "center", justifyContent: "space-between" }}
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
                  Each upload starts a new session. History preserved.
                </div>
              </div>
              <div style={{ fontSize: "11px", color: C.textMuted }}>{uploadTime ? `Last run: ${uploadTime}s` : "Click or drag"}</div>
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

        {/* ── DASHBOARD ── */}
        {activeSession && activeTab === "dashboard" && (
          <div>
            <div style={{ fontSize: "10px", letterSpacing: "0.2em", color: C.textMuted, marginBottom: "16px" }}>
              DRAFT HEALTH — {activeSession.draftName} — {activeSession.wordCount?.toLocaleString()} words
            </div>
            <HealthBar score={activeSession.health} prevScore={prevSess?.health} />

            <div style={{ display: "flex", gap: "40px", marginBottom: "32px", paddingBottom: "24px", borderBottom: `1px solid ${C.smokeLight}`, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.textMuted, marginBottom: "4px" }}>ACTIONABLE FINDINGS</div>
                <div style={{ fontSize: "32px", fontWeight: 700 }}>{totalActionable.toLocaleString()}</div>
                <div style={{ fontSize: "10px", color: C.textMuted, marginTop: "2px" }}>{totalRaw.toLocaleString()} raw total</div>
              </div>
              {prevSess && prevTotal > totalActionable && (
                <div>
                  <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.textMuted, marginBottom: "4px" }}>RESOLVED SINCE {prevSess.draftName}</div>
                  <div style={{ fontSize: "32px", fontWeight: 700, color: C.pass }}>{(prevTotal - totalActionable).toLocaleString()}</div>
                </div>
              )}
              {prevSess && totalActionable > prevTotal && (
                <div>
                  <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.regress, marginBottom: "4px" }}>⚠ REGRESSION</div>
                  <div style={{ fontSize: "32px", fontWeight: 700, color: C.regress }}>+{(totalActionable - prevTotal).toLocaleString()}</div>
                </div>
              )}
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: "12px", marginBottom: "28px" }}>
              {SIGNALS.map(sig => (
                <SignalCard key={sig.id} signal={sig}
                  actionable={activeSession.counts[sig.id] || 0}
                  raw={activeSession.rawCounts?.[sig.id] || 0}
                  prevActionable={prevSess?.counts[sig.id]}
                  onClick={() => { setFilterSignal(sig.id); setFilterDisp("all"); setActiveTab("findings"); }}
                />
              ))}
            </div>

            <button onClick={() => exportToXlsx(activeSession.findings || [], activeSession.draftName)} style={{
              background: "transparent", border: `1px solid ${C.amber}`, color: C.amber,
              padding: "10px 24px", fontSize: "12px", letterSpacing: "0.15em",
              cursor: "pointer", fontFamily: "Barlow Condensed, sans-serif",
            }}>
              EXPORT ALL FINDINGS
            </button>
          </div>
        )}

        {/* ── FINDINGS ── */}
        {activeSession && activeTab === "findings" && (
          <div>
            <div style={{ display: "flex", gap: "8px", marginBottom: "20px", flexWrap: "wrap", alignItems: "center" }}>
              <select value={filterSignal} onChange={e => setFilterSignal(e.target.value)} style={{
                background: C.smokeDark, border: `1px solid ${C.smokeLight}`, color: C.parchment,
                padding: "8px 12px", fontSize: "12px", fontFamily: "Barlow Condensed, sans-serif", cursor: "pointer",
              }}>
                <option value="all">ALL SIGNALS</option>
                {SIGNALS.map(sig => (
                  <option key={sig.id} value={sig.id}>{sig.short} ({activeSession.counts[sig.id]||0})</option>
                ))}
              </select>

              <select value={filterDisp} onChange={e => setFilterDisp(e.target.value)} style={{
                background: C.smokeDark, border: `1px solid ${C.smokeLight}`, color: C.parchment,
                padding: "8px 12px", fontSize: "12px", fontFamily: "Barlow Condensed, sans-serif", cursor: "pointer",
              }}>
                <option value="actionable">ACTIONABLE</option>
                <option value="review_only">REVIEW ONLY</option>
                <option value="all">ALL DISPOSITIONS</option>
              </select>

              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search…"
                style={{ background: C.smokeDark, border: `1px solid ${C.smokeLight}`, color: C.parchment,
                  padding: "8px 14px", fontSize: "12px", fontFamily: "Barlow Condensed, sans-serif", flex: 1, minWidth: "160px", outline: "none" }} />

              <span style={{ fontSize: "11px", color: C.textMuted }}>{filteredFindings.length.toLocaleString()} shown</span>

              <button onClick={() => exportToXlsx(filteredFindings, `${activeSession.draftName}_${filterSignal}_${filterDisp}`)} style={{
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
                  display: "grid", gridTemplateColumns: "140px 90px 90px 1fr",
                  background: i % 2 === 0 ? C.smokeDark : "transparent",
                  padding: "8px 12px", alignItems: "start", gap: "0",
                }}>
                  <div style={{ color: C.textMuted, fontSize: "10px", letterSpacing: "0.07em", paddingTop: "2px" }}>
                    {f.chapter.slice(0,22)}
                  </div>
                  <div style={{ background: C.smokeLight, color: C.amber, fontSize: "9px", letterSpacing: "0.08em", padding: "2px 5px", alignSelf: "start", whiteSpace: "nowrap" }}>
                    {f.issue.slice(0,18)}
                  </div>
                  <div style={{ fontSize: "9px", color: dispColor[f.disposition] || C.textMuted, paddingTop: "2px", letterSpacing: "0.06em" }}>
                    {(f.disposition||"").toUpperCase().replace("_"," ")} {f.confidence ? `·${f.confidence}%` : ""}
                  </div>
                  <div style={{ color: C.dimParch, fontFamily: "IM Fell English, serif", fontSize: "13px", lineHeight: 1.5 }}>
                    {f.sentence.slice(0,300)}
                    {f.reason && <span style={{ display: "block", fontSize: "10px", color: C.textMuted, fontFamily: "Barlow Condensed, sans-serif", marginTop: "2px", fontStyle: "normal" }}>{f.reason}</span>}
                  </div>
                </div>
              ))}
              {filteredFindings.length > 500 && (
                <div style={{ textAlign: "center", padding: "16px", color: C.textMuted, fontSize: "11px" }}>
                  Showing 500 of {filteredFindings.length.toLocaleString()} — export xlsx for full list
                </div>
              )}
              {filteredFindings.length === 0 && (
                <div style={{ textAlign: "center", padding: "40px", color: C.textMuted, fontSize: "12px", letterSpacing: "0.1em" }}>
                  NO FINDINGS MATCH CURRENT FILTERS
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Health Timeline */}
      {sessions.length >= 2 && activeSession && (
        <div style={{ borderTop: `1px solid ${C.smokeLight}`, padding: "14px 32px", display: "flex", gap: "32px", alignItems: "center" }}>
          <div style={{ fontSize: "10px", letterSpacing: "0.15em", color: C.textMuted, flexShrink: 0 }}>HEALTH TIMELINE</div>
          {sessions.map(s => (
            <div key={s.id} onClick={() => setActiveSession(s)} style={{ cursor: "pointer", textAlign: "center", opacity: s.id === activeSession.id ? 1 : 0.45 }}>
              <div style={{ fontSize: "20px", fontWeight: 700, color: C.amber, fontFamily: "Barlow Condensed, sans-serif" }}>{s.health}</div>
              <div style={{ fontSize: "9px", color: C.textMuted, letterSpacing: "0.08em" }}>{s.draftName}</div>
            </div>
          ))}
        </div>
      )}

      <style>{`
        @keyframes pulse { 0%,100%{opacity:.3;transform:translateX(-100%)} 50%{opacity:1;transform:translateX(250%)} }
        input::placeholder{color:#8A8070;}
        select option{background:#2A2A2A;}
        ::-webkit-scrollbar{width:6px;}
        ::-webkit-scrollbar-track{background:#1C1C1C;}
        ::-webkit-scrollbar-thumb{background:#3A3938;}
      `}</style>
    </div>
  );
}

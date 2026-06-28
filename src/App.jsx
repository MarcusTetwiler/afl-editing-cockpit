import { useState, useEffect, useRef, useCallback } from "react";
import mammoth from "mammoth";
import * as XLSX from "xlsx";

// ── Palette ───────────────────────────────────────────────────────────────────
const C = {
  charcoal:   "#1C1C1C", parchment:  "#F2ECD8", amber:      "#D4820A",
  redOrange:  "#C0392B", dimParch:   "#E8DFC4", smokeDark:  "#2A2A2A",
  smokeLight: "#3A3938", textMuted:  "#8A8070", textDim:    "#B8AC96",
  pass:       "#4A7C59", regress:    "#C0392B", overlay:    "rgba(0,0,0,0.72)",
};

// ═══════════════════════════════════════════════════════════════════════════════
// SENTENCE CLASSIFIER
// ═══════════════════════════════════════════════════════════════════════════════

const ATTRIBUTION_VERBS = new Set([
  "said","called","asked","replied","added","continued","murmured","answered",
  "shouted","whispered","snapped","growled","began","finished","offered","noted",
  "admitted","insisted","suggested","repeated","demanded","confirmed","reported",
  "stated","announced","declared","explained","responded","returned","countered",
  "cut","ordered","urged","warned","reminded","prompted","pressed","pushed",
  "breathed","managed","allowed","agreed","disagreed","objected","conceded",
  "laughed","sighed","grunted","barked","hissed","spat","echoed",
]);

const HEADER_PATTERNS = [
  /^(chapter|prologue|epilogue|part|interlude)\s/i,
  /^(andrew|elena|reese|logan|jack|anika|ryba|rupert)\s+(i{1,3}|iv|v|vi{0,3}|ix|x|\d+)$/i,
  /^(andrew|elena|reese|logan|jack|anika|ryba|rupert)$/i,
];

const DISPLAY_TEXT_PATTERNS = [
  /^\[.*\]$/,
  /REASSESSMENT|REFERRAL FLAG|ADVISORY|NOTICE:|ALERT:|STATUS:/,
];

const FRONT_MATTER_WORDS = new Set([
  "acknowledgments","acknowledgements","dedication","epigraph","preface",
  "foreword","author","copyright","isbn","reserved","rights","published",
]);

function parseDialogueSentence(s) {
  const quoteEnd = Math.max(s.lastIndexOf('"'), s.lastIndexOf('\u201d'));
  if (quoteEnd === -1) return null;
  const afterQuote = s.slice(quoteEnd + 1).trim().replace(/^[,.]/, "").trim();
  if (!afterQuote) return { action_beat: "", action_word_count: 0 };
  const words = afterQuote.split(/\s+/).filter(Boolean);
  let attrIdx = -1;
  for (let i = 0; i < Math.min(words.length, 3); i++) {
    if (ATTRIBUTION_VERBS.has(words[i].toLowerCase().replace(/[^a-z]/g, ""))) { attrIdx = i; break; }
  }
  const actionStart = attrIdx >= 0 ? attrIdx + 1 : 0;
  const actionBeat = words.slice(actionStart).join(" ");
  return { action_beat: actionBeat, action_word_count: actionBeat.split(/\s+/).filter(Boolean).length };
}

function classifySentence(s) {
  const t = s.trim();
  if (!t) return "EMPTY";
  const lower = t.toLowerCase();
  if (FRONT_MATTER_WORDS.has(lower.split(/\s+/)[0])) return "FRONT_MATTER";
  for (const pat of HEADER_PATTERNS) if (pat.test(t)) return "HEADER";
  if (/^[A-Z\s]{3,40}$/.test(t) && t.split(" ").length <= 6) return "HEADER";
  for (const pat of DISPLAY_TEXT_PATTERNS) if (pat.test(t)) return "DISPLAY_TEXT";
  const hasOpen  = /[""\u201c]/.test(t);
  const hasClose = /[""\u201d]/.test(t);
  if (!hasOpen && !hasClose) return "NARRATION";
  const startsQ = /^[""\u201c]/.test(t);
  const endsQ   = /[""\u201d][.!?]?\s*$/.test(t);
  if (startsQ && endsQ) return "PURE_DIALOGUE";
  if (startsQ) {
    const parsed = parseDialogueSentence(t);
    if (!parsed || parsed.action_word_count < 5) return "DIALOGUE_TAG";
    return "DIALOGUE_ACTION";
  }
  return "NARRATION";
}

// ═══════════════════════════════════════════════════════════════════════════════
// BUILT-IN SIGNAL DETECTORS
// ═══════════════════════════════════════════════════════════════════════════════

const D = { ACTIONABLE: "actionable", REVIEW: "review_only", EXCLUDED: "excluded" };

const INTERIORITY_COGNITION = new Set([
  "knew","thought","wondered","understood","believed","hoped","feared",
  "wished","sensed","remembered","recognized","considered","assumed",
  "supposed","imagined","realized",
]);
const INTERIORITY_CONDITIONAL = new Set(["felt","decided","expected","needed","wanted"]);
const POV_NAMES    = new Set(["andrew","elena","reese","logan","jack","anika","ryba","rupert"]);
const POV_PRONOUNS = new Set(["he","she","they","his","her","their"]);
const NON_HUMAN    = new Set(["system","plan","route","amount","road","machine","drone","signal",
  "mission","screen","display","unit","command","protocol","schedule","sequence",
  "process","device","network","camera","sensor","data","feed","broadcast","channel"]);

function isHumanSubject(w) {
  const c = (w||"").toLowerCase().replace(/[^a-z]/g,"");
  if (POV_PRONOUNS.has(c) || POV_NAMES.has(c)) return true;
  if (NON_HUMAN.has(c)) return false;
  return true;
}

function detectInteriority(s) {
  const words = s.split(/\s+/);
  for (let i = 0; i < words.length; i++) {
    const w = words[i].toLowerCase().replace(/[^a-z]/g,"");
    const subj = i > 0 ? words[i-1].toLowerCase().replace(/[^a-z]/g,"") : "";
    if (INTERIORITY_COGNITION.has(w)) return { found:true, confidence:90 };
    if (INTERIORITY_CONDITIONAL.has(w) && isHumanSubject(subj)) {
      const next = (words[i+1]||"").toLowerCase().replace(/[^a-z]/g,"");
      if (["to","that","it","him","her","them","more","less","one"].includes(next))
        return { found:true, confidence:75 };
    }
  }
  if (/\b(his|her|their)\s+(mind|thoughts?|feelings?|emotions?|instinct|gut|conscience)\b/i.test(s))
    return { found:true, confidence:85 };
  return { found:false };
}

function detectGPS(s, wordCount) {
  const checks = [
    [/\b(views?|concerns?|issues?|matters?|aspects?|elements?)\s+(were|are|remain|seemed|appeared|became)\b/i, "ABSTRACT_SUBJECT", 70, "Abstract subject with vague state verb"],
    [/\bthe\s+(next|first|last|second|third|former|latter)\s+(dropped|showed|hit|fell|came|went|moved|turned|cut)\b/i, "DELAYED_REFERENT", 85, "Ordinal without clear referent"],
    [/\bdropped\s+his\s+(chair|seat|position|post)\b/i, "MISLEADING_PAIR", 80, "dropped + position noun reads as physical drop"],
    [/\btook\s+the\s+(floor|ground|field|stage)\b/i, "MISLEADING_PAIR", 80, "took + spatial noun may mislead first parse"],
    [/\bheld\s+the\s+(line|floor|room|stage|court)\b/i, "MISLEADING_PAIR", 75, "held + abstract location may mislead"],
    [/\bkilled\s+the\s+(lights?|sound|feed|signal|line)\b/i, "MISLEADING_PAIR", 80, "killed + tech noun — violence read before correction"],
  ];
  for (const [pat, type, conf, reason] of checks) {
    if (pat.test(s)) return { type, confidence: conf, reason };
  }
  const clauseCount = (s.match(/\b(as|while|when|although|though|because|since|after|before|until|unless|despite|who|which|that)\b/gi)||[]).length;
  if (clauseCount >= 3 && wordCount > 25)
    return { type:"OVERLOADED_PATH", confidence:65, reason:`${clauseCount} subordinate clauses` };
  if (/^[A-Z][a-z]+ing\b/.test(s) && wordCount > 12)
    return { type:"PARTICIPIAL_RISK", confidence:60, reason:"Opening participial phrase — attachment may be ambiguous" };
  const chain = (s.match(/\b(as|while|with)\b/gi)||[]).length;
  if (chain >= 3) return { type:"CHAIN_OVERLOAD", confidence:65, reason:`${chain}x as/while/with chain` };
  const nounMatch = s.match(/\b(\w{5,})\b.*?\b\1\b/i);
  if (nounMatch) {
    const w = nounMatch[1].toLowerCase();
    if (!["that","this","with","from","they","their","there","when","then","than","have","been","were","will"].includes(w))
      return { type:"REPEATED_NOUN", confidence:60, reason:`Repeated noun "${w}" may cause reparse` };
  }
  return null;
}

// ── Run all signals on parsed text ─────────────────────────────────────────────
function runBuiltinSignals(sentences, customSignals = []) {
  const findings = [];

  // Paragraph grouping for para-level signals
  // sentences = [{text, cls, chapter, paraIdx, wordCount}]
  const byPara = {};
  for (const s of sentences) {
    if (!byPara[s.paraIdx]) byPara[s.paraIdx] = [];
    byPara[s.paraIdx].push(s);
  }

  // Para-level: Flat Wave + Repeated Openers
  for (const [, sents] of Object.entries(byPara)) {
    const chapter = sents[0].chapter;
    const narr = sents.filter(s => s.cls === "NARRATION" || s.cls === "DIALOGUE_ACTION");

    // Flat Wave
    if (narr.length >= 4) {
      const lens = narr.map(s => s.wordCount);
      const mean = lens.reduce((a,b)=>a+b,0)/lens.length;
      const stdDev = Math.sqrt(lens.reduce((a,b)=>a+Math.pow(b-mean,2),0)/lens.length);
      if (stdDev < 3.5 && mean > 8) {
        findings.push({ signal_id:"flat_wave", signalType:"defect", chapter, sentenceId:null, hash:null, sentence:`[Para] ${sents.map(s=>s.text).join(" ").slice(0,130)}…`,
          issue:"FLAT SENTENCE WAVE", disposition:D.REVIEW, confidence:70,
          reason:`StdDev ${stdDev.toFixed(1)}, mean ${mean.toFixed(1)} words across ${narr.length} sentences` });
      }
    }

    // Repeated Openers
    if (sents.length >= 3) {
      const openers = sents.map(s => s.text.trim().split(/\s+/)[0]?.toLowerCase().replace(/[^a-z]/g,"") || "");
      let streak = 1;
      for (let i = 1; i < openers.length; i++) {
        if (openers[i] && openers[i].length > 1 && openers[i] === openers[i-1]) {
          streak++;
          if (streak >= 3) {
            findings.push({ signal_id:"repeated_openers", signalType:"defect", chapter, sentenceId:null, hash:null,
              sentence: sents.slice(i-2,i+1).map(s=>s.text).join("  "),
              issue:"REPEATED OPENERS", disposition:D.ACTIONABLE, confidence:90,
              reason:`"${openers[i]}" opens ${streak} consecutive sentences` });
            streak = 1;
          }
        } else { streak = 1; }
      }
    }
  }

  // Sentence-level signals
  for (const s of sentences) {
    const { text, cls, chapter, wordCount, sentenceId, hash } = s;
    if (["HEADER","FRONT_MATTER","DISPLAY_TEXT","EMPTY"].includes(cls)) continue;

    // 7-Word
    if (cls === "NARRATION" && wordCount > 0 && wordCount < 7) {
      findings.push({ signal_id:"seven_word", signalType:"defect", chapter, sentenceId, hash, sentence:text, issue:"7-WORD NARRATION RULE",
        disposition:D.ACTIONABLE, confidence:85, reason:`${wordCount}-word narration sentence` });
    } else if (cls === "DIALOGUE_ACTION") {
      const p = parseDialogueSentence(text);
      if (p && p.action_word_count > 0 && p.action_word_count < 7) {
        findings.push({ signal_id:"seven_word", signalType:"defect", chapter, sentenceId, hash, sentence:text, issue:"7-WORD NARRATION RULE",
          disposition:D.REVIEW, confidence:60, reason:`Action beat is ${p.action_word_count} words: "${p.action_beat}"` });
      }
    }

    // Interiority
    if (cls === "NARRATION") {
      const r = detectInteriority(text);
      if (r.found) findings.push({ signal_id:"interiority", signalType:"defect", chapter, sentenceId, hash, sentence:text, issue:"INTERIORITY LEAK",
        disposition: r.confidence >= 85 ? D.ACTIONABLE : D.REVIEW, confidence:r.confidence,
        reason:"Cognition verb in narration" });
    }

    // Pronoun Ambiguity
    if (cls === "NARRATION") {
      const pats = [/\b(it|this|that)\s+(was|is|had|has|did|does|would|could|should|seemed|looked|felt|appeared)\b/gi, /^(It|This|That)\s+(was|is|had|seemed|would|felt|appeared)/];
      for (const pat of pats) { pat.lastIndex=0; if (pat.test(text)) { findings.push({ signal_id:"pronoun_ambiguity", signalType:"defect", chapter, sentenceId, hash, sentence:text, issue:"PRONOUN AMBIGUITY", disposition:D.REVIEW, confidence:70, reason:"Vague pronoun with unclear referent" }); break; } }
    }

    // Filter Verb
    if (cls === "NARRATION" || cls === "DIALOGUE_ACTION") {
      const target = cls === "DIALOGUE_ACTION" ? (parseDialogueSentence(text)?.action_beat || text) : text;
      if (/\b(saw|looked|watched|noticed|realized|felt|heard|smelled|tasted|sensed|observed|spotted|glimpsed|perceived|detected)\b/gi.test(target))
        findings.push({ signal_id:"filter_verb", signalType:"defect", chapter, sentenceId, hash, sentence:text, issue:"FILTER VERB", disposition:D.REVIEW, confidence:75, reason:"Perception verb in narration" });
    }

    // GPS-Lite
    if (cls === "NARRATION" && wordCount >= 8) {
      const g = detectGPS(text, wordCount);
      if (g && g.confidence >= 60) findings.push({ signal_id:"gps_lite", signalType:"defect", chapter, sentenceId, hash, sentence:text,
        issue:`GPS: ${g.type.replace(/_/g," ")}`, disposition: g.confidence >= 75 ? D.ACTIONABLE : D.REVIEW,
        confidence:g.confidence, reason:g.reason });
    }

    // Custom signals
    for (const cs of customSignals) {
      if (!cs.enabled) continue;
      const eligibleTypes = cs.scope === "narration" ? ["NARRATION"] : cs.scope === "dialogue" ? ["PURE_DIALOGUE","DIALOGUE_TAG","DIALOGUE_ACTION"] : ["NARRATION","PURE_DIALOGUE","DIALOGUE_TAG","DIALOGUE_ACTION"];
      if (!eligibleTypes.includes(cls)) continue;
      const target = text;
      let matched = false;
      for (const kw of cs.keywords) {
        const kwtrimmed = kw.trim();
        if (!kwtrimmed) continue;
        try {
          const pat = new RegExp(`\\b${kwtrimmed.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}\\b`, "gi");
          if (pat.test(target)) { matched = true; break; }
        } catch {}
      }
      if (matched) findings.push({ signal_id: cs.id, signalType: cs.signalType || "defect", chapter, sentenceId, hash, sentence: text, issue: cs.name.toUpperCase(),
        disposition: D.REVIEW, confidence: 75, reason: `Custom signal: ${cs.keywords.slice(0,3).join(", ")}` });
    }
  }

  return findings;
}

// ── Slug helpers ──────────────────────────────────────────────────────────────
function slugify(s) {
  return (s || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40);
}

// Lightweight 6-char hash of a string (FNV-1a variant)
function textHash(s) {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = (h * 16777619) >>> 0; }
  return h.toString(16).slice(0, 6);
}

// ── Parse raw text into sentence index ────────────────────────────────────────
function buildSentenceIndex(rawText) {
  const lines = rawText.split(/\n/);
  let currentChapter = "Opening";
  let paraIdx = 0;
  let inFrontMatter = true;
  const sentences = [];

  const parasRaw = [];
  let cur = [];
  for (const line of lines) {
    const t = line.trim();
    if (!t) { if (cur.length) { parasRaw.push(cur); cur = []; } continue; }
    cur.push(t);
  }
  if (cur.length) parasRaw.push(cur);

  for (const para of parasRaw) {
    paraIdx++;
    const fullText = para.join(" ");
    const t = para[0];

    // Chapter detection
    for (const pat of HEADER_PATTERNS) {
      if (pat.test(t) && t.length < 80) { currentChapter = t; inFrontMatter = false; break; }
    }
    if (/^[A-Z\s]{4,}$/.test(t) && t.length < 80 && t.split(" ").length <= 8) { currentChapter = t; inFrontMatter = false; }

    if (inFrontMatter && fullText.length < 120 && !/[.!?]{1}\s*$/.test(fullText)) continue;

    const rawSents = fullText.match(/[^.!?]+[.!?]+(?:\s|$)|[^.!?]+$/g) || [fullText];
    const chapterSlug = slugify(currentChapter);
    let sentIdx = 0;
    for (const rs of rawSents) {
      const s = rs.trim();
      if (!s || s.length < 4) continue;
      sentIdx++;
      const cls = classifySentence(s);
      const pStr = String(paraIdx).padStart(4, "0");
      const sStr = String(sentIdx).padStart(3, "0");
      const sentenceId = `${chapterSlug}:p${pStr}:s${sStr}`;
      sentences.push({
        text: s,
        cls,
        chapter: currentChapter,
        chapterSlug,
        paraIdx,
        sentIdx,
        sentenceId,
        hash: textHash(s),
        wordCount: s.split(/\s+/).filter(Boolean).length,
      });
    }
  }
  return sentences;
}

// ═══════════════════════════════════════════════════════════════════════════════
// SIGNAL TYPE SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

// signalType drives dashboard language, delta direction, and card colour logic
// defect      → count should go DOWN   — regression = bad, improvement = good
// motif       → count should PERSIST   — no regression logic; presence is health
// continuity  → count should PERSIST   — track anchors, flag disappearance
// reader_note → resolve manually       — imported human/AI concerns

const SIGNAL_TYPES = {
  defect:      { label:"Defect",       color:"#C0392B", desiredDir:"down",    deltaLabel: (d) => d > 0 ? `▲ +${d} regression` : d < 0 ? `▼ ${Math.abs(d)} resolved` : "—" },
  motif:       { label:"Motif",        color:"#D4820A", desiredDir:"persist", deltaLabel: (d) => d > 0 ? `▲ +${d}` : d < 0 ? `▼ ${Math.abs(d)}` : "—" },
  continuity:  { label:"Continuity",   color:"#6B8F9E", desiredDir:"persist", deltaLabel: (d) => d > 0 ? `▲ +${d}` : d < 0 ? `▼ ${Math.abs(d)} missing` : "—" },
  reader_note: { label:"Reader Note",  color:"#8A8070", desiredDir:"down",    deltaLabel: (d) => d > 0 ? `▲ +${d}` : d < 0 ? `▼ ${Math.abs(d)} resolved` : "—" },
};

// ═══════════════════════════════════════════════════════════════════════════════
// BUILT-IN SIGNAL DEFINITIONS
// ═══════════════════════════════════════════════════════════════════════════════

const BUILTIN_SIGNALS = [
  { id:"seven_word",        name:"7-Word Narration Rule",         short:"7-WORD",      signalType:"defect",      description:"Narration sentences under 7 words — excludes dialogue, headers, display text", builtin:true },
  { id:"interiority",       name:"Interiority Leak",              short:"INTERIORITY", signalType:"defect",      description:"Stated thoughts, feelings, cognition — human subject + cognition verb", builtin:true },
  { id:"pronoun_ambiguity", name:"Pronoun Ambiguity",             short:"PRONOUN",     signalType:"defect",      description:'Vague "it," "this," "that" with unclear referent in narration', builtin:true },
  { id:"repeated_openers",  name:"Repeated Openers",              short:"OPENERS",     signalType:"defect",      description:"3+ consecutive sentences opening with the same word", builtin:true },
  { id:"flat_wave",         name:"Flat Sentence Wave",            short:"FLAT WAVE",   signalType:"defect",      description:"Paragraphs with monotonous sentence-length clustering", builtin:true },
  { id:"filter_verb",       name:"Filter Verb / Perception Leak", short:"FILTER VERB", signalType:"defect",      description:'"saw," "looked," "watched," "noticed" in narration', builtin:true },
  { id:"gps_lite",          name:"GPS-Lite: Garden-Path Scanner", short:"GPS-LITE",    signalType:"defect",      description:"Heuristic first-parse friction — misleading verbs, delayed referents, clause overload", builtin:true },
];

// ═══════════════════════════════════════════════════════════════════════════════
// PERSISTENCE
// ═══════════════════════════════════════════════════════════════════════════════

const STORAGE_KEY    = "afl_cockpit_v3";
const CUSTOM_SIG_KEY = "afl_cockpit_custom_signals";

function loadSessions() { try { return JSON.parse(localStorage.getItem(STORAGE_KEY)||"[]"); } catch { return []; } }
function saveSessions(s) { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(s.map(x=>({...x,findings:x.findings?.slice(0,2000)})))); } catch {} }
function loadCustomSignals() { try { return JSON.parse(localStorage.getItem(CUSTOM_SIG_KEY)||"[]"); } catch { return []; } }
function saveCustomSignals(s) { try { localStorage.setItem(CUSTOM_SIG_KEY, JSON.stringify(s)); } catch {} }

// ═══════════════════════════════════════════════════════════════════════════════
// IMPORT FINDINGS LIST (xlsx → imported signal)
// ═══════════════════════════════════════════════════════════════════════════════

// Normalise a sentence for dedup comparison
function normSentence(s) {
  return (s || "").toLowerCase().replace(/[^a-z0-9\s]/g, "").replace(/\s+/g, " ").trim();
}

// Returns overlap ratio between two normalised strings
function overlapRatio(a, b) {
  if (!a || !b) return 0;
  const shorter = a.length < b.length ? a : b;
  const longer  = a.length < b.length ? b : a;
  if (longer.includes(shorter)) return 1;
  // Word-level Jaccard
  const setA = new Set(a.split(" "));
  const setB = new Set(b.split(" "));
  const inter = [...setA].filter(w => setB.has(w)).length;
  return inter / Math.max(setA.size, setB.size);
}

async function parseImportedXlsx(file) {
  // Returns [{signalName, sentences:[]}]
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  const sheets = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    // Column A = index 0; skip header row if it looks like metadata (short, no period)
    const sentences = [];
    for (const row of rows) {
      const cell = String(row[0] || "").trim();
      if (!cell) continue;
      if (cell.length < 15) continue; // skip short labels/headers
      if (!/[a-z]/i.test(cell)) continue; // skip all-symbol rows
      sentences.push(cell);
    }
    if (sentences.length > 0) {
      sheets.push({ signalName: sheetName.trim(), sentences });
    }
  }
  return sheets;
}

function deduplicateImported(importedSentences, existingFindings) {
  const existingNorms = existingFindings.map(f => normSentence(f.sentence));
  return importedSentences.map(s => {
    const norm = normSentence(s);
    const isDupe = existingNorms.some(en => overlapRatio(norm, en) >= 0.85);
    return { sentence: s, isDupe };
  });
}

function computeHealth(findings, wordCount) {
  if (!wordCount) return 0;
  const actionable = findings.filter(f => f.disposition === D.ACTIONABLE).length;
  const rate = (actionable / wordCount) * 1000;
  return Math.round(Math.max(40, Math.min(100, 100 - rate * 0.55)) * 10) / 10;
}

function exportToXlsx(findings, name) {
  const rows = [["Sentence ID","Hash","Chapter / Location","Flagged Sentence","Issue Type","Signal Type","Disposition","Confidence","Reason"]];
  for (const f of findings) rows.push([f.sentenceId||"", f.hash||"", f.chapter, f.sentence, f.issue, f.signalType||"defect", f.disposition||"", f.confidence||"", f.reason||""]);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [22,8,20,60,18,12,12,10,40].map(w=>({wch:w}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Findings");
  XLSX.writeFile(wb, `${name}_findings.xlsx`);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADD SIGNAL MODAL
// ═══════════════════════════════════════════════════════════════════════════════

function trafficLight(count) {
  if (count === 0)   return { color:"#666", label:"No matches" };
  if (count <= 10)   return { color:C.pass,      label:"Surgical — looks good" };
  if (count <= 50)   return { color:C.amber,     label:"Review — may need tuning" };
  return               { color:C.regress,    label:"Noisy — consider narrowing keywords" };
}

function AddSignalModal({ sentenceIndex, onSave, onClose }) {
  const [name,       setName]       = useState("");
  const [keywords,   setKeywords]   = useState("");
  const [scope,      setScope]      = useState("narration");
  const [signalType, setSignalType] = useState("defect");
  const [preview,    setPreview]    = useState(null);
  const [running,    setRunning]    = useState(false);

  const eligible = {
    narration: ["NARRATION"],
    dialogue:  ["PURE_DIALOGUE","DIALOGUE_TAG","DIALOGUE_ACTION"],
    all:       ["NARRATION","PURE_DIALOGUE","DIALOGUE_TAG","DIALOGUE_ACTION"],
  }[scope];

  const runPreview = useCallback(() => {
    const kws = keywords.split(",").map(k=>k.trim()).filter(Boolean);
    if (!kws.length || !sentenceIndex?.length) { setPreview(null); return; }
    setRunning(true);
    setTimeout(() => {
      const hits = [];
      for (const s of sentenceIndex) {
        if (!eligible.includes(s.cls)) continue;
        for (const kw of kws) {
          try {
            const pat = new RegExp(`\\b${kw.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}\\b`, "gi");
            if (pat.test(s.text)) { hits.push(s); break; }
          } catch {}
        }
      }
      const examples = hits.slice(0,3).map(s=>({ chapter:s.chapter, text:s.text }));
      setPreview({ count:hits.length, examples });
      setRunning(false);
    }, 80);
  }, [keywords, scope, sentenceIndex]);

  // Auto-preview on keyword/scope change
  useEffect(() => {
    const t = setTimeout(runPreview, 400);
    return () => clearTimeout(t);
  }, [keywords, scope, runPreview]);

  const light = preview ? trafficLight(preview.count) : null;

  const handleSave = () => {
    const kws = keywords.split(",").map(k=>k.trim()).filter(Boolean);
    if (!name.trim() || !kws.length) return;
    const id = "custom_" + Date.now();
    onSave({ id, name: name.trim(), short: name.trim().toUpperCase().slice(0,10), description:`Custom: ${kws.slice(0,4).join(", ")}`, keywords:kws, scope, signalType, enabled:true, builtin:false });
  };

  const inputStyle = {
    background: C.smokeDark, border:`1px solid ${C.smokeLight}`, color:C.parchment,
    padding:"10px 14px", fontSize:"13px", fontFamily:"Barlow Condensed, sans-serif",
    width:"100%", outline:"none", letterSpacing:"0.05em",
  };

  const labelStyle = { fontSize:"10px", letterSpacing:"0.18em", color:C.textMuted, display:"block", marginBottom:"6px" };

  return (
    <div style={{ position:"fixed", inset:0, background:C.overlay, zIndex:1000, display:"flex", alignItems:"center", justifyContent:"center" }}
      onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{ background:C.smokeDark, border:`1px solid ${C.smokeLight}`, width:"560px", maxWidth:"94vw", maxHeight:"90vh", overflowY:"auto" }}>
        {/* Modal header */}
        <div style={{ padding:"20px 24px 16px", borderBottom:`1px solid ${C.smokeLight}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div>
            <div style={{ fontSize:"16px", fontWeight:700, letterSpacing:"0.15em", color:C.amber }}>ADD SIGNAL</div>
            <div style={{ fontSize:"11px", color:C.textMuted, marginTop:"2px", fontFamily:"IM Fell English, serif", fontStyle:"italic" }}>
              Keyword-based detector — matches against classified sentence index
            </div>
          </div>
          <button onClick={onClose} style={{ background:"none", border:"none", color:C.textMuted, fontSize:"20px", cursor:"pointer", lineHeight:1 }}>×</button>
        </div>

        <div style={{ padding:"24px" }}>
          {/* Signal name */}
          <div style={{ marginBottom:"20px" }}>
            <label style={labelStyle}>SIGNAL NAME</label>
            <input value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. Motif Tracker — Upper Lane" style={inputStyle} />
          </div>

          {/* Keywords */}
          <div style={{ marginBottom:"20px" }}>
            <label style={labelStyle}>KEYWORDS — comma separated, plain text or simple phrase</label>
            <input value={keywords} onChange={e=>setKeywords(e.target.value)}
              placeholder="e.g. upper lane, lighter, whisper pin, Moscowvite" style={inputStyle} />
            <div style={{ fontSize:"10px", color:C.textMuted, marginTop:"6px" }}>
              Each keyword matched as whole word. Case-insensitive.
            </div>
          </div>

          {/* Scope */}
          <div style={{ marginBottom:"20px" }}>
            <label style={labelStyle}>RUN ON</label>
            <div style={{ display:"flex", gap:"8px" }}>
              {[["narration","Narration only"],["dialogue","Dialogue only"],["all","All sentences"]].map(([val,label])=>(
                <button key={val} onClick={()=>setScope(val)} style={{
                  background: scope===val ? C.amber : C.smokeLight,
                  color: scope===val ? C.charcoal : C.textDim,
                  border:"none", padding:"8px 16px", fontSize:"11px", letterSpacing:"0.1em",
                  cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif",
                }}>
                  {label.toUpperCase()}
                </button>
              ))}
            </div>
          </div>

          {/* Signal type */}
          <div style={{ marginBottom:"24px" }}>
            <label style={labelStyle}>SIGNAL TYPE</label>
            <div style={{ display:"flex", gap:"8px", flexWrap:"wrap" }}>
              {Object.entries(SIGNAL_TYPES).map(([val, meta])=>(
                <button key={val} onClick={()=>setSignalType(val)} style={{
                  background: signalType===val ? meta.color : C.smokeLight,
                  color: signalType===val ? "#fff" : C.textDim,
                  border:"none", padding:"8px 14px", fontSize:"11px", letterSpacing:"0.1em",
                  cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif",
                }}>
                  {meta.label.toUpperCase()}
                </button>
              ))}
            </div>
            <div style={{ fontSize:"10px", color:C.textMuted, marginTop:"6px" }}>
              {signalType === "defect"      && "Count should go down. Regression flagged if it rises."}
              {signalType === "motif"       && "Count should persist or rise. No regression logic. Tracks healthy structure."}
              {signalType === "continuity"  && "Tracks anchors across manuscript. Flags disappearance."}
              {signalType === "reader_note" && "Imported human or AI concern. Resolve manually."}
            </div>
          </div>

          {/* Preview */}
          <div style={{ borderTop:`1px solid ${C.smokeLight}`, paddingTop:"20px", marginBottom:"24px" }}>
            <div style={{ fontSize:"10px", letterSpacing:"0.18em", color:C.textMuted, marginBottom:"12px" }}>
              LIVE PREVIEW {running && "— scanning…"}
            </div>

            {!sentenceIndex?.length && (
              <div style={{ fontSize:"12px", color:C.textMuted, fontStyle:"italic", fontFamily:"IM Fell English, serif" }}>
                Upload a draft first to preview signal results.
              </div>
            )}

            {preview && sentenceIndex?.length > 0 && (
              <>
                {/* Traffic light */}
                <div style={{ display:"flex", alignItems:"center", gap:"16px", marginBottom:"16px" }}>
                  <div style={{ width:"14px", height:"14px", borderRadius:"50%", background:light.color, flexShrink:0 }} />
                  <div>
                    <span style={{ fontSize:"28px", fontWeight:700, fontFamily:"Barlow Condensed, sans-serif", color:light.color }}>
                      {preview.count.toLocaleString()}
                    </span>
                    <span style={{ fontSize:"11px", color:C.textMuted, marginLeft:"10px", letterSpacing:"0.1em" }}>
                      {light.label.toUpperCase()}
                    </span>
                  </div>
                </div>

                {/* Examples */}
                {preview.examples.length > 0 && (
                  <div style={{ display:"flex", flexDirection:"column", gap:"8px" }}>
                    {preview.examples.map((ex,i)=>(
                      <div key={i} style={{ background:C.charcoal, padding:"10px 14px", borderLeft:`2px solid ${C.smokeLight}` }}>
                        <div style={{ fontSize:"9px", color:C.textMuted, letterSpacing:"0.1em", marginBottom:"4px" }}>{ex.chapter.slice(0,30)}</div>
                        <div style={{ fontSize:"12px", color:C.dimParch, fontFamily:"IM Fell English, serif", lineHeight:1.5 }}>{ex.text.slice(0,200)}</div>
                      </div>
                    ))}
                  </div>
                )}

                {preview.count === 0 && (
                  <div style={{ fontSize:"12px", color:C.textMuted, fontStyle:"italic", fontFamily:"IM Fell English, serif" }}>
                    No matches found. Try broader keywords or a different scope.
                  </div>
                )}
              </>
            )}

            {!preview && keywords.trim() && sentenceIndex?.length > 0 && (
              <div style={{ fontSize:"12px", color:C.textMuted }}>Scanning…</div>
            )}
          </div>

          {/* Actions */}
          <div style={{ display:"flex", gap:"12px", justifyContent:"flex-end" }}>
            <button onClick={onClose} style={{ background:"transparent", border:`1px solid ${C.smokeLight}`, color:C.textDim, padding:"10px 20px", fontSize:"12px", letterSpacing:"0.12em", cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif" }}>
              CANCEL
            </button>
            <button onClick={handleSave} disabled={!name.trim() || !keywords.trim()}
              style={{ background: (!name.trim()||!keywords.trim()) ? C.smokeLight : C.amber,
                color: (!name.trim()||!keywords.trim()) ? C.textMuted : C.charcoal,
                border:"none", padding:"10px 24px", fontSize:"12px", letterSpacing:"0.15em",
                cursor: (!name.trim()||!keywords.trim()) ? "default" : "pointer",
                fontFamily:"Barlow Condensed, sans-serif", fontWeight:700 }}>
              ADD TO DASHBOARD
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// SIGNAL MANAGEMENT PANEL (hamburger)
// ═══════════════════════════════════════════════════════════════════════════════

function SignalPanel({ builtinSignals, customSignals, onToggleCustom, onDeleteCustom, onAddSignal, onImport, onClose }) {
  const importRef = useRef();
  return (
    <div style={{ position:"fixed", inset:0, background:C.overlay, zIndex:999, display:"flex", justifyContent:"flex-end" }}
      onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{ background:C.smokeDark, borderLeft:`1px solid ${C.smokeLight}`, width:"340px", maxWidth:"90vw", height:"100%", overflowY:"auto", display:"flex", flexDirection:"column" }}>
        <div style={{ padding:"20px 24px 16px", borderBottom:`1px solid ${C.smokeLight}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div style={{ fontSize:"14px", fontWeight:700, letterSpacing:"0.18em", color:C.amber }}>SIGNALS</div>
          <button onClick={onClose} style={{ background:"none", border:"none", color:C.textMuted, fontSize:"20px", cursor:"pointer" }}>×</button>
        </div>

        <div style={{ padding:"16px 24px", flex:1 }}>
          {/* Built-in */}
          <div style={{ fontSize:"9px", letterSpacing:"0.18em", color:C.textMuted, marginBottom:"10px" }}>BUILT-IN — always active</div>
          {builtinSignals.map(sig=>(
            <div key={sig.id} style={{ padding:"10px 0", borderBottom:`1px solid ${C.smokeLight}`, display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
              <div>
                <div style={{ fontSize:"12px", color:C.parchment, letterSpacing:"0.08em" }}>{sig.name}</div>
                <div style={{ fontSize:"10px", color:C.textMuted, marginTop:"3px", fontFamily:"IM Fell English, serif", fontStyle:"italic", lineHeight:1.3 }}>{sig.description}</div>
              </div>
              <div style={{ fontSize:"9px", color:C.pass, letterSpacing:"0.1em", flexShrink:0, marginLeft:"12px", marginTop:"2px" }}>ON</div>
            </div>
          ))}

          {/* Custom */}
          <div style={{ fontSize:"9px", letterSpacing:"0.18em", color:C.textMuted, marginTop:"20px", marginBottom:"10px" }}>CUSTOM SIGNALS</div>
          {customSignals.length === 0 && (
            <div style={{ fontSize:"11px", color:C.textMuted, fontStyle:"italic", fontFamily:"IM Fell English, serif", marginBottom:"12px" }}>
              No custom signals yet.
            </div>
          )}
          {customSignals.map(sig=>(
            <div key={sig.id} style={{ padding:"10px 0", borderBottom:`1px solid ${C.smokeLight}` }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <div style={{ display:"flex", alignItems:"center", gap:"8px" }}>
                  <div style={{ fontSize:"12px", color: sig.enabled ? C.parchment : C.textMuted, letterSpacing:"0.08em" }}>{sig.name}</div>
                  {sig.imported && <div style={{ fontSize:"8px", background:C.smokeLight, color:C.textMuted, padding:"1px 5px", letterSpacing:"0.1em" }}>IMPORTED</div>}
                </div>
                <div style={{ display:"flex", gap:"10px", alignItems:"center", flexShrink:0, marginLeft:"10px" }}>
                  <button onClick={()=>onToggleCustom(sig.id)} style={{ background:"none", border:"none", color: sig.enabled ? C.pass : C.textMuted, fontSize:"9px", letterSpacing:"0.1em", cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif" }}>
                    {sig.enabled ? "ON" : "OFF"}
                  </button>
                  <button onClick={()=>onDeleteCustom(sig.id)} style={{ background:"none", border:"none", color:C.textMuted, fontSize:"13px", cursor:"pointer", lineHeight:1 }}>×</button>
                </div>
              </div>
              <div style={{ fontSize:"10px", color:C.textMuted, marginTop:"3px", fontFamily:"IM Fell English, serif", fontStyle:"italic" }}>
                {sig.imported ? sig.description : `${sig.keywords?.slice(0,4).join(", ")} · ${sig.scope}`}
              </div>
            </div>
          ))}

          <button onClick={onAddSignal} style={{ marginTop:"16px", background:"transparent", border:`1px solid ${C.amber}`, color:C.amber, padding:"10px 20px", fontSize:"11px", letterSpacing:"0.15em", cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif", width:"100%" }}>
            + ADD SIGNAL
          </button>

          {/* Divider */}
          <div style={{ borderTop:`1px solid ${C.smokeLight}`, margin:"20px 0" }} />

          {/* Import section */}
          <div style={{ fontSize:"9px", letterSpacing:"0.18em", color:C.textMuted, marginBottom:"8px" }}>IMPORT FINDINGS LIST</div>
          <div style={{ fontSize:"11px", color:C.textDim, fontFamily:"IM Fell English, serif", fontStyle:"italic", marginBottom:"12px", lineHeight:1.4 }}>
            Upload an xlsx where the sheet name is the signal name and column A contains flagged sentences.
          </div>
          <input ref={importRef} type="file" accept=".xlsx" style={{ display:"none" }}
            onChange={e => { if (e.target.files[0]) { onImport(e.target.files[0]); e.target.value=""; } }} />
          <button onClick={()=>importRef.current?.click()} style={{ background:"transparent", border:`1px solid ${C.smokeLight}`, color:C.textDim, padding:"10px 20px", fontSize:"11px", letterSpacing:"0.15em", cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif", width:"100%" }}>
            ↑ IMPORT .XLSX LIST
          </button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMPONENTS
// ═══════════════════════════════════════════════════════════════════════════════

function SignalCard({ signal, actionable, raw, prevActionable, onClick }) {
  const delta    = prevActionable !== undefined ? actionable - prevActionable : null;
  const excluded = Math.max(0, (raw||0) - actionable);
  const sigType  = SIGNAL_TYPES[signal.signalType] || SIGNAL_TYPES.defect;

  // For defect/reader_note: down is good. For motif/continuity: direction is neutral.
  const isDefect   = signal.signalType === "defect" || signal.signalType === "reader_note";
  const isPass     = isDefect && actionable === 0;
  const isRegress  = isDefect && delta !== null && delta > 0;
  const borderColor = isRegress ? C.regress : isPass ? C.pass : C.smokeLight;

  // Type accent strip at top
  const typeColor = sigType.color;

  return (
    <div onClick={onClick} style={{ background:C.smokeDark, border:`1px solid ${borderColor}`, padding:"16px 20px", position:"relative", cursor:"pointer" }}
      onMouseEnter={e=>e.currentTarget.style.borderColor=C.amber}
      onMouseLeave={e=>e.currentTarget.style.borderColor=borderColor}>
      {/* Signal type accent strip */}
      <div style={{ position:"absolute", top:0, left:0, right:0, height:"2px", background: isRegress ? C.regress : typeColor, opacity: isRegress ? 1 : 0.6 }} />
      {signal.builtin === false && (
        <div style={{ position:"absolute", top:"8px", right:"10px", fontSize:"8px", color:C.textMuted, letterSpacing:"0.1em" }}>
          {signal.imported ? "IMPORTED" : "CUSTOM"}
        </div>
      )}
      {/* Signal type label */}
      <div style={{ display:"flex", alignItems:"center", gap:"6px", marginBottom:"8px" }}>
        <div style={{ fontSize:"10px", letterSpacing:"0.15em", color:C.textMuted, fontFamily:"Barlow Condensed, sans-serif" }}>{signal.short}</div>
        <div style={{ fontSize:"8px", letterSpacing:"0.12em", color:typeColor, fontFamily:"Barlow Condensed, sans-serif", opacity:0.85 }}>
          {sigType.label.toUpperCase()}
        </div>
      </div>
      <div style={{ display:"flex", alignItems:"flex-end", gap:"10px" }}>
        <div style={{ fontSize:"36px", fontFamily:"Barlow Condensed, sans-serif", fontWeight:700, lineHeight:1,
          color: isPass ? C.pass : isRegress ? C.regress : C.parchment }}>
          {actionable.toLocaleString()}
        </div>
        {delta !== null && delta !== 0 && (
          <div style={{ fontSize:"12px", fontFamily:"Barlow Condensed, sans-serif",
            color: isDefect ? (delta > 0 ? C.regress : C.pass) : C.textDim,
            paddingBottom:"4px" }}>
            {sigType.deltaLabel(delta)}
          </div>
        )}
        {isPass && <div style={{ fontSize:"11px", color:C.pass, paddingBottom:"4px", fontFamily:"Barlow Condensed, sans-serif" }}>✓ PASS</div>}
      </div>
      {excluded > 0 && <div style={{ fontSize:"10px", color:C.textMuted, marginTop:"3px", fontFamily:"Barlow Condensed, sans-serif" }}>{excluded.toLocaleString()} excluded</div>}
      <div style={{ fontSize:"11px", color:C.textDim, marginTop:"8px", fontFamily:"IM Fell English, serif", fontStyle:"italic", lineHeight:1.3 }}>{signal.description}</div>
    </div>
  );
}

function HealthBar({ score, prevScore }) {
  const delta = prevScore !== undefined ? (score-prevScore).toFixed(1) : null;
  const pct = Math.max(0,Math.min(100,((score-40)/60)*100));
  return (
    <div style={{ marginBottom:"28px" }}>
      <div style={{ display:"flex", alignItems:"baseline", gap:"16px", marginBottom:"10px" }}>
        <div style={{ fontSize:"64px", fontFamily:"Barlow Condensed, sans-serif", fontWeight:700, color:C.amber, lineHeight:1 }}>{score}</div>
        <div>
          <div style={{ fontSize:"14px", color:C.textMuted, letterSpacing:"0.12em", fontFamily:"Barlow Condensed, sans-serif" }}>/ 100</div>
          {delta!==null && <div style={{ fontSize:"15px", fontFamily:"Barlow Condensed, sans-serif", color:parseFloat(delta)>=0?C.pass:C.regress }}>{parseFloat(delta)>=0?`▲ +${delta}`:`▼ ${delta}`} since previous</div>}
        </div>
      </div>
      <div style={{ height:"3px", background:C.smokeLight }}>
        <div style={{ height:"100%", width:`${pct}%`, background:`linear-gradient(90deg,${C.amber},${C.redOrange})` }} />
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════════════════

export default function App() {
  const [sessions,       setSessions]       = useState([]);
  const [activeSession,  setActiveSession]  = useState(null);
  const [sentenceIndex,  setSentenceIndex]  = useState(null); // persists for preview
  const [customSignals,  setCustomSignals]  = useState([]);
  const [uploading,      setUploading]      = useState(false);
  const [uploadStep,     setUploadStep]     = useState("");
  const [uploadTime,     setUploadTime]     = useState(null);
  const [activeTab,      setActiveTab]      = useState("dashboard");
  const [filterSignal,   setFilterSignal]   = useState("all");
  const [filterDisp,     setFilterDisp]     = useState("actionable");
  const [search,         setSearch]         = useState("");
  const [showPanel,      setShowPanel]      = useState(false);
  const [showAddModal,   setShowAddModal]   = useState(false);
  const fileInputRef = useRef();

  useEffect(() => {
    const saved = loadSessions();
    if (saved.length) { setSessions(saved); setActiveSession(saved[saved.length-1]); }
    setCustomSignals(loadCustomSignals());
  }, []);

  const allSignals = [...BUILTIN_SIGNALS, ...customSignals];

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
      await new Promise(r=>setTimeout(r,60));
      const index = buildSentenceIndex(rawText);
      setSentenceIndex(index);

      setUploadStep("Running signals…");
      await new Promise(r=>setTimeout(r,60));
      const findings = runBuiltinSignals(index, customSignals);

      setUploadStep("Computing health…");
      const wordCount = rawText.split(/\s+/).filter(Boolean).length;
      const health = computeHealth(findings, wordCount);

      const counts = {}, rawCounts = {};
      for (const sig of allSignals) { counts[sig.id]=0; rawCounts[sig.id]=0; }
      for (const f of findings) {
        rawCounts[f.signal_id] = (rawCounts[f.signal_id]||0)+1;
        if (f.disposition !== D.EXCLUDED) counts[f.signal_id] = (counts[f.signal_id]||0)+1;
      }

      const session = { id:Date.now(), draftName, uploadedAt:new Date().toISOString(), wordCount, health, counts, rawCounts, findings };
      setSessions(prev => { const next=[...prev,session]; saveSessions(next); return next; });
      setActiveSession(session);
      setUploadTime(((Date.now()-t0)/1000).toFixed(1));
      setActiveTab("dashboard");
    } catch(err) { alert("Error: "+err.message); }
    finally { setUploading(false); setUploadStep(""); }
  }, [customSignals, allSignals]);

  const handleDrop = useCallback(e => { e.preventDefault(); const f=e.dataTransfer.files[0]; if(f) handleUpload(f); }, [handleUpload]);

  const handleAddCustomSignal = useCallback((sig) => {
    const updated = [...customSignals, sig];
    setCustomSignals(updated);
    saveCustomSignals(updated);
    setShowAddModal(false);
    setShowPanel(false);
  }, [customSignals]);

  const handleToggleCustom = useCallback((id) => {
    const updated = customSignals.map(s=>s.id===id?{...s,enabled:!s.enabled}:s);
    setCustomSignals(updated); saveCustomSignals(updated);
  }, [customSignals]);

  const handleDeleteCustom = useCallback((id) => {
    const updated = customSignals.filter(s=>s.id!==id);
    setCustomSignals(updated); saveCustomSignals(updated);
  }, [customSignals]);

  const handleImport = useCallback(async (file) => {
    try {
      const sheets = await parseImportedXlsx(file);
      if (!sheets.length) { alert("No valid sheets found. Check format: sheet name = signal name, column A = sentences."); return; }

      const existingFindings = activeSession?.findings || [];
      const newSignals = [];

      for (const { signalName, sentences } of sheets) {
        const deduped = deduplicateImported(sentences, existingFindings);
        const id = "imported_" + Date.now() + "_" + Math.random().toString(36).slice(2,6);
        const findings = deduped.map(({ sentence, isDupe }) => ({
          signal_id: id,
          chapter: "IMPORTED",
          sentence,
          issue: signalName.toUpperCase(),
          disposition: isDupe ? "excluded" : "review_only",
          confidence: 80,
          reason: isDupe ? "Duplicate — already caught by existing signal" : `Imported from ${file.name}`,
        }));

        const actionableCount = findings.filter(f => f.disposition !== "excluded").length;
        const dupCount = findings.length - actionableCount;

        newSignals.push({
          id,
          name: signalName,
          short: signalName.toUpperCase().slice(0, 10),
          description: `Imported from ${file.name} — ${actionableCount} unique, ${dupCount} dupes excluded`,
          keywords: [],
          scope: "all",
          enabled: true,
          builtin: false,
          imported: true,
          importedFindings: findings,
        });
      }

      // Merge into customSignals
      const updated = [...customSignals, ...newSignals];
      setCustomSignals(updated);
      saveCustomSignals(updated);

      // Merge imported findings into active session
      if (activeSession) {
        const allNewFindings = newSignals.flatMap(s => s.importedFindings);
        const updatedSession = {
          ...activeSession,
          findings: [...existingFindings, ...allNewFindings],
          counts: { ...activeSession.counts },
          rawCounts: { ...activeSession.rawCounts },
        };
        for (const sig of newSignals) {
          updatedSession.counts[sig.id] = sig.importedFindings.filter(f => f.disposition !== "excluded").length;
          updatedSession.rawCounts[sig.id] = sig.importedFindings.length;
        }
        setActiveSession(updatedSession);
        setSessions(prev => {
          const next = prev.map(s => s.id === updatedSession.id ? updatedSession : s);
          saveSessions(next);
          return next;
        });
      }

      setShowPanel(false);
      const totalImported = newSignals.reduce((a, s) => a + s.importedFindings.length, 0);
      const totalDupes    = newSignals.reduce((a, s) => a + s.importedFindings.filter(f=>f.disposition==="excluded").length, 0);
      alert(`Imported ${newSignals.length} signal${newSignals.length>1?"s":""}.\n${totalImported} sentences — ${totalDupes} duplicates excluded.`);
    } catch (err) {
      alert("Import error: " + err.message);
    }
  }, [activeSession, customSignals]);

  const activeIdx      = sessions.findIndex(s=>s.id===activeSession?.id);
  const prevSess       = activeIdx > 0 ? sessions[activeIdx-1] : null;
  const totalActionable= activeSession ? Object.values(activeSession.counts).reduce((a,b)=>a+b,0) : 0;
  const totalRaw       = activeSession ? Object.values(activeSession.rawCounts||{}).reduce((a,b)=>a+b,0) : 0;
  const prevTotal      = prevSess ? Object.values(prevSess.counts).reduce((a,b)=>a+b,0) : null;
  const dispColor      = { actionable:C.amber, review_only:C.textDim };

  const filteredFindings = (activeSession?.findings||[]).filter(f => {
    if (filterSignal!=="all" && f.signal_id!==filterSignal) return false;
    if (filterDisp!=="all" && f.disposition!==filterDisp) return false;
    if (search && !f.sentence.toLowerCase().includes(search.toLowerCase()) && !f.chapter.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  return (
    <div style={{ minHeight:"100vh", background:C.charcoal, color:C.parchment, fontFamily:"Barlow Condensed, sans-serif" }}>
      {/* Header */}
      <div style={{ borderBottom:`1px solid ${C.smokeLight}`, padding:"0 24px", display:"flex", alignItems:"center", justifyContent:"space-between", height:"56px" }}>
        <div style={{ display:"flex", alignItems:"baseline", gap:"16px" }}>
          <span style={{ fontSize:"18px", fontWeight:700, letterSpacing:"0.15em", color:C.amber }}>AFL</span>
          <span style={{ fontSize:"14px", letterSpacing:"0.2em", color:C.textMuted }}>EDITING COCKPIT</span>
        </div>
        <div style={{ display:"flex", gap:"4px", alignItems:"center" }}>
          {["dashboard","findings"].map(tab=>(
            <button key={tab} onClick={()=>setActiveTab(tab)} style={{ background:activeTab===tab?C.smokeLight:"transparent", border:"none", color:activeTab===tab?C.parchment:C.textMuted, padding:"8px 16px", fontSize:"12px", letterSpacing:"0.15em", cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif", textTransform:"uppercase" }}>
              {tab}
              {tab==="findings"&&activeSession&&<span style={{ marginLeft:"8px", background:C.amber, color:C.charcoal, padding:"1px 6px", fontSize:"10px", fontWeight:700 }}>{totalActionable.toLocaleString()}</span>}
            </button>
          ))}
          {/* Hamburger */}
          <button onClick={()=>setShowPanel(true)} style={{ background:"none", border:`1px solid ${C.smokeLight}`, color:C.textDim, padding:"7px 12px", cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif", fontSize:"16px", marginLeft:"8px", lineHeight:1 }} title="Signal management">
            ☰
          </button>
        </div>
      </div>

      <div style={{ padding:"28px 32px", maxWidth:"1160px", margin:"0 auto" }}>
        {/* Version Rail */}
        {sessions.length > 0 && (
          <div style={{ display:"flex", gap:"8px", flexWrap:"wrap", marginBottom:"20px" }}>
            {sessions.map(s=>(
              <div key={s.id} onClick={()=>setActiveSession(s)} style={{ padding:"4px 12px", cursor:"pointer", background:s.id===activeSession?.id?C.amber:C.smokeLight, color:s.id===activeSession?.id?C.charcoal:C.textDim, fontFamily:"Barlow Condensed, sans-serif", fontSize:"12px", letterSpacing:"0.1em", fontWeight:s.id===activeSession?.id?700:400 }}>
                {s.draftName} <span style={{ opacity:0.7 }}>{s.health}</span>
              </div>
            ))}
          </div>
        )}

        {/* Upload Zone */}
        <div onDrop={handleDrop} onDragOver={e=>e.preventDefault()} onClick={()=>!uploading&&fileInputRef.current?.click()}
          style={{ border:`1px dashed ${C.smokeLight}`, padding:"18px 28px", marginBottom:"28px", cursor:uploading?"default":"pointer", display:"flex", alignItems:"center", justifyContent:"space-between" }}
          onMouseEnter={e=>{if(!uploading)e.currentTarget.style.borderColor=C.amber;}}
          onMouseLeave={e=>{e.currentTarget.style.borderColor=C.smokeLight;}}>
          <input ref={fileInputRef} type="file" accept=".docx" style={{ display:"none" }} onChange={e=>e.target.files[0]&&handleUpload(e.target.files[0])} />
          {uploading ? (
            <div>
              <div style={{ fontSize:"13px", color:C.amber, letterSpacing:"0.1em" }}>{uploadStep}</div>
              <div style={{ width:"220px", height:"2px", background:C.smokeLight, marginTop:"10px", overflow:"hidden" }}>
                <div style={{ height:"100%", background:C.amber, width:"40%", animation:"pulse 1.4s ease-in-out infinite" }} />
              </div>
            </div>
          ) : (
            <>
              <div>
                <div style={{ fontSize:"14px", letterSpacing:"0.12em", color:C.textDim }}>DROP DRAFT — .docx</div>
                <div style={{ fontSize:"11px", color:C.textMuted, marginTop:"4px", fontStyle:"italic", fontFamily:"IM Fell English, serif" }}>Each upload starts a new session. History preserved.</div>
              </div>
              <div style={{ fontSize:"11px", color:C.textMuted }}>{uploadTime?`Last run: ${uploadTime}s`:"Click or drag"}</div>
            </>
          )}
        </div>

        {!activeSession&&!uploading&&(
          <div style={{ textAlign:"center", padding:"80px 32px", color:C.textMuted }}>
            <div style={{ fontSize:"40px", marginBottom:"16px", opacity:0.2 }}>◉</div>
            <div style={{ fontSize:"14px", letterSpacing:"0.2em" }}>NO DRAFT LOADED</div>
            <div style={{ fontSize:"12px", marginTop:"8px", fontFamily:"IM Fell English, serif", fontStyle:"italic" }}>Upload a manuscript to initialize the cockpit</div>
          </div>
        )}

        {/* ── DASHBOARD ── */}
        {activeSession&&activeTab==="dashboard"&&(
          <div>
            <div style={{ fontSize:"10px", letterSpacing:"0.2em", color:C.textMuted, marginBottom:"16px" }}>
              DRAFT HEALTH — {activeSession.draftName} — {activeSession.wordCount?.toLocaleString()} words
            </div>
            <HealthBar score={activeSession.health} prevScore={prevSess?.health} />

            <div style={{ display:"flex", gap:"40px", marginBottom:"28px", paddingBottom:"20px", borderBottom:`1px solid ${C.smokeLight}`, flexWrap:"wrap" }}>
              <div>
                <div style={{ fontSize:"10px", letterSpacing:"0.15em", color:C.textMuted, marginBottom:"4px" }}>ACTIONABLE FINDINGS</div>
                <div style={{ fontSize:"32px", fontWeight:700 }}>{totalActionable.toLocaleString()}</div>
                <div style={{ fontSize:"10px", color:C.textMuted, marginTop:"2px" }}>{totalRaw.toLocaleString()} raw total</div>
              </div>
              {prevSess&&prevTotal>totalActionable&&(
                <div>
                  <div style={{ fontSize:"10px", letterSpacing:"0.15em", color:C.textMuted, marginBottom:"4px" }}>RESOLVED SINCE {prevSess.draftName}</div>
                  <div style={{ fontSize:"32px", fontWeight:700, color:C.pass }}>{(prevTotal-totalActionable).toLocaleString()}</div>
                </div>
              )}
              {prevSess&&totalActionable>prevTotal&&(
                <div>
                  <div style={{ fontSize:"10px", letterSpacing:"0.15em", color:C.regress, marginBottom:"4px" }}>⚠ REGRESSION</div>
                  <div style={{ fontSize:"32px", fontWeight:700, color:C.regress }}>+{(totalActionable-prevTotal).toLocaleString()}</div>
                </div>
              )}
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(240px,1fr))", gap:"12px", marginBottom:"24px" }}>
              {allSignals.filter(sig=>!sig.builtin||true).map(sig=>(
                (!sig.enabled && sig.builtin===false) ? null :
                <SignalCard key={sig.id} signal={sig}
                  actionable={activeSession.counts[sig.id]||0}
                  raw={activeSession.rawCounts?.[sig.id]||0}
                  prevActionable={prevSess?.counts[sig.id]}
                  onClick={()=>{ setFilterSignal(sig.id); setFilterDisp("all"); setActiveTab("findings"); }}
                />
              ))}
            </div>

            <button onClick={()=>exportToXlsx(activeSession.findings||[],activeSession.draftName)} style={{ background:"transparent", border:`1px solid ${C.amber}`, color:C.amber, padding:"10px 24px", fontSize:"12px", letterSpacing:"0.15em", cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif" }}>
              EXPORT ALL FINDINGS
            </button>
          </div>
        )}

        {/* ── FINDINGS ── */}
        {activeSession&&activeTab==="findings"&&(
          <div>
            <div style={{ display:"flex", gap:"8px", marginBottom:"20px", flexWrap:"wrap", alignItems:"center" }}>
              <select value={filterSignal} onChange={e=>setFilterSignal(e.target.value)} style={{ background:C.smokeDark, border:`1px solid ${C.smokeLight}`, color:C.parchment, padding:"8px 12px", fontSize:"12px", fontFamily:"Barlow Condensed, sans-serif", cursor:"pointer" }}>
                <option value="all">ALL SIGNALS</option>
                {allSignals.map(sig=><option key={sig.id} value={sig.id}>{sig.short} ({activeSession.counts[sig.id]||0})</option>)}
              </select>
              <select value={filterDisp} onChange={e=>setFilterDisp(e.target.value)} style={{ background:C.smokeDark, border:`1px solid ${C.smokeLight}`, color:C.parchment, padding:"8px 12px", fontSize:"12px", fontFamily:"Barlow Condensed, sans-serif", cursor:"pointer" }}>
                <option value="actionable">ACTIONABLE</option>
                <option value="review_only">REVIEW ONLY</option>
                <option value="all">ALL DISPOSITIONS</option>
              </select>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search…" style={{ background:C.smokeDark, border:`1px solid ${C.smokeLight}`, color:C.parchment, padding:"8px 14px", fontSize:"12px", fontFamily:"Barlow Condensed, sans-serif", flex:1, minWidth:"140px", outline:"none" }} />
              <span style={{ fontSize:"11px", color:C.textMuted }}>{filteredFindings.length.toLocaleString()} shown</span>
              <button onClick={()=>exportToXlsx(filteredFindings,`${activeSession.draftName}_${filterSignal}_${filterDisp}`)} style={{ background:"transparent", border:`1px solid ${C.smokeLight}`, color:C.textDim, padding:"8px 16px", fontSize:"11px", letterSpacing:"0.12em", cursor:"pointer", fontFamily:"Barlow Condensed, sans-serif" }}>EXPORT VIEW</button>
            </div>

            <div style={{ display:"flex", flexDirection:"column", gap:"2px" }}>
              {filteredFindings.slice(0,500).map((f,i)=>(
                <div key={i} style={{ display:"grid", gridTemplateColumns:"130px 90px 80px 1fr", background:i%2===0?C.smokeDark:"transparent", padding:"8px 12px", alignItems:"start" }}>
                  <div style={{ color:C.textMuted, fontSize:"10px", letterSpacing:"0.07em", paddingTop:"2px" }}>{f.chapter.slice(0,22)}</div>
                  <div style={{ background:C.smokeLight, color:C.amber, fontSize:"9px", letterSpacing:"0.08em", padding:"2px 5px", alignSelf:"start", whiteSpace:"nowrap", overflow:"hidden" }}>{f.issue.slice(0,16)}</div>
                  <div style={{ fontSize:"9px", color:dispColor[f.disposition]||C.textMuted, paddingTop:"2px", letterSpacing:"0.06em" }}>{(f.disposition||"").toUpperCase().replace("_"," ")}{f.confidence?` ·${f.confidence}%`:""}</div>
                  <div style={{ color:C.dimParch, fontFamily:"IM Fell English, serif", fontSize:"13px", lineHeight:1.5 }}>
                    {f.sentence.slice(0,300)}
                    {f.reason&&<span style={{ display:"block", fontSize:"10px", color:C.textMuted, fontFamily:"Barlow Condensed, sans-serif", marginTop:"2px", fontStyle:"normal" }}>{f.reason}</span>}
                  </div>
                </div>
              ))}
              {filteredFindings.length>500&&<div style={{ textAlign:"center", padding:"16px", color:C.textMuted, fontSize:"11px" }}>Showing 500 of {filteredFindings.length.toLocaleString()} — export xlsx for full list</div>}
              {filteredFindings.length===0&&<div style={{ textAlign:"center", padding:"40px", color:C.textMuted, fontSize:"12px", letterSpacing:"0.1em" }}>NO FINDINGS MATCH CURRENT FILTERS</div>}
            </div>
          </div>
        )}
      </div>

      {/* Health Timeline */}
      {sessions.length>=2&&activeSession&&(
        <div style={{ borderTop:`1px solid ${C.smokeLight}`, padding:"14px 32px", display:"flex", gap:"32px", alignItems:"center" }}>
          <div style={{ fontSize:"10px", letterSpacing:"0.15em", color:C.textMuted, flexShrink:0 }}>HEALTH TIMELINE</div>
          {sessions.map(s=>(
            <div key={s.id} onClick={()=>setActiveSession(s)} style={{ cursor:"pointer", textAlign:"center", opacity:s.id===activeSession.id?1:0.45 }}>
              <div style={{ fontSize:"20px", fontWeight:700, color:C.amber, fontFamily:"Barlow Condensed, sans-serif" }}>{s.health}</div>
              <div style={{ fontSize:"9px", color:C.textMuted, letterSpacing:"0.08em" }}>{s.draftName}</div>
            </div>
          ))}
        </div>
      )}

      {/* Signal Panel */}
      {showPanel&&(
        <SignalPanel builtinSignals={BUILTIN_SIGNALS} customSignals={customSignals}
          onToggleCustom={handleToggleCustom} onDeleteCustom={handleDeleteCustom}
          onAddSignal={()=>{setShowAddModal(true); setShowPanel(false);}}
          onImport={handleImport}
          onClose={()=>setShowPanel(false)} />
      )}

      {/* Add Signal Modal */}
      {showAddModal&&(
        <AddSignalModal sentenceIndex={sentenceIndex} onSave={handleAddCustomSignal} onClose={()=>setShowAddModal(false)} />
      )}

      <style>{`
        @keyframes pulse{0%,100%{opacity:.3;transform:translateX(-100%)}50%{opacity:1;transform:translateX(250%)}}
        input::placeholder{color:#8A8070;}
        select option{background:#2A2A2A;}
        ::-webkit-scrollbar{width:6px;}
        ::-webkit-scrollbar-track{background:#1C1C1C;}
        ::-webkit-scrollbar-thumb{background:#3A3938;}
      `}</style>
    </div>
  );
}

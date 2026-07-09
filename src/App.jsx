import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import mammoth from "mammoth";
import * as XLSX from "xlsx";

const C = {
  charcoal:"#1C1C1C",parchment:"#F2ECD8",amber:"#D4820A",redOrange:"#C0392B",
  dimParch:"#E8DFC4",smokeDark:"#2A2A2A",smokeLight:"#3A3938",
  textMuted:"#8A8070",textDim:"#B8AC96",pass:"#4A7C59",regress:"#C0392B",overlay:"rgba(0,0,0,0.72)",
};
const SIGNAL_TYPES={
  defect:{label:"Defect",color:"#C0392B",desiredDir:"down"},
  motif:{label:"Motif",color:"#D4820A",desiredDir:"persist"},
  continuity:{label:"Continuity",color:"#6B8F9E",desiredDir:"persist"},
  reader_note:{label:"Reader Note",color:"#8A8070",desiredDir:"down"},
};
const ATTRIBUTION_VERBS=new Set(["said","called","asked","replied","added","continued","murmured","answered","shouted","whispered","snapped","growled","began","finished","offered","noted","admitted","insisted","suggested","repeated","demanded","confirmed","reported","stated","announced","declared","explained","responded","returned","countered","cut","ordered","urged","warned","reminded","prompted","pressed","pushed","breathed","managed","allowed","agreed","disagreed","objected","conceded","laughed","sighed","grunted","barked","hissed","spat","echoed"]);
const HEADER_PATTERNS=[/^(chapter|prologue|epilogue|part|interlude)\s/i,/^(andrew|elena|reese|logan|jack|anika|ryba|rupert)\s+(i{1,3}|iv|v|vi{0,3}|ix|x|\d+)$/i,/^(andrew|elena|reese|logan|jack|anika|ryba|rupert)$/i];
const DISPLAY_TEXT_PATTERNS=[/^\[.*\]$/,/REASSESSMENT|REFERRAL FLAG|ADVISORY|NOTICE:|ALERT:|STATUS:/];
function parseDialogueSentence(s){const qe=Math.max(s.lastIndexOf('"'),s.lastIndexOf('\u201d'));if(qe===-1)return null;const aq=s.slice(qe+1).trim().replace(/^[,.]/, "").trim();if(!aq)return{action_beat:"",action_word_count:0};const words=aq.split(/\s+/).filter(Boolean);let ai=-1;for(let i=0;i<Math.min(words.length,3);i++){if(ATTRIBUTION_VERBS.has(words[i].toLowerCase().replace(/[^a-z]/g,""))){ai=i;break;}}const ab=words.slice(ai>=0?ai+1:0).join(" ");return{action_beat:ab,action_word_count:ab.split(/\s+/).filter(Boolean).length};}
function classifySentence(s){const t=s.trim();if(!t)return"EMPTY";for(const p of HEADER_PATTERNS)if(p.test(t))return"HEADER";if(/^[A-Z\s]{3,40}$/.test(t)&&t.split(" ").length<=6)return"HEADER";for(const p of DISPLAY_TEXT_PATTERNS)if(p.test(t))return"DISPLAY_TEXT";const ho=/[""\u201c]/.test(t),hc=/[""\u201d]/.test(t);if(!ho&&!hc)return"NARRATION";const sq=/^[""\u201c]/.test(t);if(sq&&/[""\u201d][.!?]?\s*$/.test(t))return"PURE_DIALOGUE";if(sq){const p=parseDialogueSentence(t);return(!p||p.action_word_count<5)?"DIALOGUE_TAG":"DIALOGUE_ACTION";}return"NARRATION";}
function slugify(s){return(s||"").toLowerCase().replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"").slice(0,40);}
function textHash(s){let h=2166136261;for(let i=0;i<s.length;i++){h^=s.charCodeAt(i);h=(h*16777619)>>>0;}return h.toString(16).slice(0,6);}
function buildSentenceIndex(rawText){const lines=rawText.split(/\n/);let ch="Opening",pi=0,ln=0,fm=true;const sents=[],pr=[];let cur=[];for(const line of lines){const t=line.trim();if(!t){if(cur.length){pr.push(cur);cur=[];}continue;}cur.push(t);}if(cur.length)pr.push(cur);for(const para of pr){pi++;const ft=para.join(" "),t=para[0];for(const pat of HEADER_PATTERNS){if(pat.test(t)&&t.length<80){ch=t;fm=false;break;}}if(/^[A-Z\s]{4,}$/.test(t)&&t.length<80&&t.split(" ").length<=8){ch=t;fm=false;}if(fm&&ft.length<120&&!/[.!?]\s*$/.test(ft))continue;const rs=ft.match(/[^.!?]+[.!?]+(?:\s|$)|[^.!?]+$/g)||[ft];const cs=slugify(ch);let si=0;for(const r of rs){const s=r.trim();if(!s||s.length<4)continue;si++;ln++;const cls=classifySentence(s);const sid=`${cs}:p${String(pi).padStart(4,"0")}:s${String(si).padStart(3,"0")}`;sents.push({text:s,cls,chapter:ch,chapterSlug:cs,paraIdx:pi,sentIdx:si,lineNum:ln,sentenceId:sid,paragraphId:`${cs}:p${String(pi).padStart(4,"0")}`,hash:textHash(s),wordCount:s.split(/\s+/).filter(Boolean).length});}}return sents;}
function buildIndexFromTaxonomy(rows){return rows.map((r,i)=>({text:String(r.sentence_text||""),cls:(r.is_dialogue===true||r.is_dialogue==="true"||r.is_dialogue===1)?"PURE_DIALOGUE":"NARRATION",chapter:String(r.chapter_id||""),chapterSlug:String(r.chapter_id||"").toLowerCase().replace(/\s+/g,"-"),paraIdx:Number(r.paragraph_index_in_chapter)||i,sentIdx:Number(r.sentence_index_in_paragraph)||1,lineNum:i+1,sentenceId:String(r.sentence_id||""),paragraphId:String(r.paragraph_id||""),hash:String(r.sentence_hash||""),wordCount:Number(r.word_count)||String(r.sentence_text||"").split(/\s+/).filter(Boolean).length}));}
const D={ACTIONABLE:"actionable",REVIEW:"review_only",EXCLUDED:"excluded"};
const IC=new Set(["knew","thought","wondered","understood","believed","hoped","feared","wished","sensed","remembered","recognized","considered","assumed","supposed","imagined","realized"]);
const ICD=new Set(["felt","decided","expected","needed","wanted"]);
const PN=new Set(["andrew","elena","reese","logan","jack","anika","ryba","rupert"]);
const PP=new Set(["he","she","they","his","her","their"]);
const NH=new Set(["system","plan","route","amount","road","machine","drone","signal","mission","screen","display","unit","command","protocol","schedule","sequence","process","device","network","camera","sensor","data","feed","broadcast","channel"]);
function isHuman(w){const c=(w||"").toLowerCase().replace(/[^a-z]/g,"");if(PP.has(c)||PN.has(c))return true;if(NH.has(c))return false;return true;}
function detectInteriority(s){const w=s.split(/\s+/);for(let i=0;i<w.length;i++){const ww=w[i].toLowerCase().replace(/[^a-z]/g,""),sub=i>0?w[i-1].toLowerCase().replace(/[^a-z]/g,""):"";if(IC.has(ww))return{found:true,confidence:90};if(ICD.has(ww)&&isHuman(sub)){const n=(w[i+1]||"").toLowerCase().replace(/[^a-z]/g,"");if(["to","that","it","him","her","them","more","less","one"].includes(n))return{found:true,confidence:75};}}if(/\b(his|her|their)\s+(mind|thoughts?|feelings?|emotions?|instinct|gut|conscience)\b/i.test(s))return{found:true,confidence:85};return{found:false};}
function detectGPS(s,wc){const chk=[[/\b(views?|concerns?|issues?|matters?)\s+(were|are|remain|seemed)\b/i,"ABSTRACT_SUBJECT",70,"Abstract subject with vague verb"],[/\bthe\s+(next|first|last|second)\s+(dropped|showed|hit|fell|came)\b/i,"DELAYED_REFERENT",85,"Ordinal without clear referent"],[/\bdropped\s+his\s+(chair|seat|position)\b/i,"MISLEADING_PAIR",80,"position noun reads as drop"],[/\btook\s+the\s+(floor|ground|stage)\b/i,"MISLEADING_PAIR",80,"spatial noun may mislead"],[/\bkilled\s+the\s+(lights?|sound|feed|signal)\b/i,"MISLEADING_PAIR",80,"tech noun — violence read first"]];for(const[p,t,c,r]of chk)if(p.test(s))return{type:t,confidence:c,reason:r};const cc=(s.match(/\b(as|while|when|although|though|because|since|after|before|until|unless|who|which|that)\b/gi)||[]).length;if(cc>=3&&wc>25)return{type:"OVERLOADED_PATH",confidence:65,reason:`${cc} subordinate clauses`};if(/^[A-Z][a-z]+ing\b/.test(s)&&wc>12)return{type:"PARTICIPIAL_RISK",confidence:60,reason:"Opening participial — attachment ambiguous"};return null;}
function runBuiltinSignals(sentences,customSignals=[]){const findings=[];const byPara={};for(const s of sentences){if(!byPara[s.paraIdx])byPara[s.paraIdx]=[];byPara[s.paraIdx].push(s);}
for(const[,sents]of Object.entries(byPara)){const ch=sents[0].chapter;const narr=sents.filter(s=>s.cls==="NARRATION"||s.cls==="DIALOGUE_ACTION");if(narr.length>=4){const lens=narr.map(s=>s.wordCount);const mean=lens.reduce((a,b)=>a+b,0)/lens.length;const sd=Math.sqrt(lens.reduce((a,b)=>a+Math.pow(b-mean,2),0)/lens.length);if(sd<3.5&&mean>8)findings.push({signal_id:"flat_wave",signalType:"defect",chapter:ch,sentenceId:null,paragraphId:null,hash:null,lineNum:sents[0]?.lineNum||null,sentence:`[Para] ${sents.map(s=>s.text).join(" ").slice(0,130)}…`,issue:"FLAT SENTENCE WAVE",disposition:D.REVIEW,confidence:70,reason:`StdDev ${sd.toFixed(1)}, mean ${mean.toFixed(1)}`});}
if(sents.length>=3){const op=sents.map(s=>s.text.trim().split(/\s+/)[0]?.toLowerCase().replace(/[^a-z]/g,"")||"");let str=1;for(let i=1;i<op.length;i++){if(op[i]&&op[i].length>1&&op[i]===op[i-1]){str++;if(str>=3){findings.push({signal_id:"repeated_openers",signalType:"defect",chapter:ch,sentenceId:null,paragraphId:null,hash:null,lineNum:sents[Math.max(0,i-2)]?.lineNum||null,sentence:sents.slice(i-2,i+1).map(s=>s.text).join("  "),issue:"REPEATED OPENERS",disposition:D.ACTIONABLE,confidence:90,reason:`"${op[i]}" opens ${str} sentences`});str=1;}}else str=1;}}}
for(const s of sentences){const{text,cls,chapter,wordCount,sentenceId,hash,lineNum}=s;const paragraphId=s.paragraphId||(sentenceId?sentenceId.split(':').slice(0,2).join(':'):"");if(["HEADER","FRONT_MATTER","DISPLAY_TEXT","EMPTY"].includes(cls))continue;
if(cls==="NARRATION"&&wordCount>0&&wordCount<7)findings.push({signal_id:"seven_word",signalType:"defect",chapter,sentenceId,paragraphId,hash,lineNum,sentence:text,issue:"7-WORD NARRATION RULE",disposition:D.ACTIONABLE,confidence:85,reason:`${wordCount} words`});
else if(cls==="DIALOGUE_ACTION"){const p=parseDialogueSentence(text);if(p&&p.action_word_count>0&&p.action_word_count<7)findings.push({signal_id:"seven_word",signalType:"defect",chapter,sentenceId,paragraphId,hash,lineNum,sentence:text,issue:"7-WORD NARRATION RULE",disposition:D.REVIEW,confidence:60,reason:`Action beat ${p.action_word_count} words`});}
if(cls==="NARRATION"){const r=detectInteriority(text);if(r.found)findings.push({signal_id:"interiority",signalType:"defect",chapter,sentenceId,paragraphId,hash,lineNum,sentence:text,issue:"INTERIORITY LEAK",disposition:r.confidence>=85?D.ACTIONABLE:D.REVIEW,confidence:r.confidence,reason:"Cognition verb in narration"});}
if(cls==="NARRATION"){const pts=[/\b(it|this|that)\s+(was|is|had|has|did|does|would|could|should|seemed|looked|felt|appeared)\b/gi,/^(It|This|That)\s+(was|is|had|seemed|would|felt|appeared)/];for(const pt of pts){pt.lastIndex=0;if(pt.test(text)){findings.push({signal_id:"pronoun_ambiguity",signalType:"defect",chapter,sentenceId,paragraphId,hash,lineNum,sentence:text,issue:"PRONOUN AMBIGUITY",disposition:D.REVIEW,confidence:70,reason:"Vague pronoun"});break;}}}
if(cls==="NARRATION"||cls==="DIALOGUE_ACTION"){const tgt=cls==="DIALOGUE_ACTION"?(parseDialogueSentence(text)?.action_beat||text):text;if(/\b(saw|looked|watched|noticed|realized|felt|heard|smelled|tasted|sensed|observed|spotted|glimpsed|perceived|detected)\b/gi.test(tgt))findings.push({signal_id:"filter_verb",signalType:"defect",chapter,sentenceId,paragraphId,hash,lineNum,sentence:text,issue:"FILTER VERB",disposition:D.REVIEW,confidence:75,reason:"Perception verb"});}
if(cls==="NARRATION"&&wordCount>=8){const g=detectGPS(text,wordCount);if(g&&g.confidence>=60)findings.push({signal_id:"gps_lite",signalType:"defect",chapter,sentenceId,paragraphId,hash,lineNum,sentence:text,issue:`GPS: ${g.type.replace(/_/g," ")}`,disposition:g.confidence>=75?D.ACTIONABLE:D.REVIEW,confidence:g.confidence,reason:g.reason});}
for(const cs of customSignals){if(!cs.enabled||cs.imported)continue;const el=cs.scope==="narration"?["NARRATION"]:cs.scope==="dialogue"?["PURE_DIALOGUE","DIALOGUE_TAG","DIALOGUE_ACTION"]:["NARRATION","PURE_DIALOGUE","DIALOGUE_TAG","DIALOGUE_ACTION"];if(!el.includes(cls))continue;let m=false;for(const kw of cs.keywords){const k=kw.trim();if(!k)continue;try{const pt=new RegExp(`\\b${k.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}\\b`,"gi");if(pt.test(text)){m=true;break;}}catch{}}if(m)findings.push({signal_id:cs.id,signalType:cs.signalType||"defect",chapter,sentenceId,paragraphId,hash,lineNum,sentence:text,issue:cs.name.toUpperCase(),disposition:D.REVIEW,confidence:75,reason:`Custom: ${cs.keywords.slice(0,3).join(", ")}`});}}return findings;}
const BUILTIN_SIGNALS=[{id:"seven_word",name:"7-Word Narration Rule",short:"7-WORD",signalType:"defect",description:"Narration sentences under 7 words",builtin:true},{id:"interiority",name:"Interiority Leak",short:"INTERIORITY",signalType:"defect",description:"Stated thoughts, feelings, cognition",builtin:true},{id:"pronoun_ambiguity",name:"Pronoun Ambiguity",short:"PRONOUN",signalType:"defect",description:'Vague "it," "this," "that" in narration',builtin:true},{id:"repeated_openers",name:"Repeated Openers",short:"OPENERS",signalType:"defect",description:"3+ consecutive same-word openers",builtin:true},{id:"flat_wave",name:"Flat Sentence Wave",short:"FLAT WAVE",signalType:"defect",description:"Monotonous sentence-length paragraphs",builtin:true},{id:"filter_verb",name:"Filter Verb / Perception Leak",short:"FILTER VERB",signalType:"defect",description:'"saw," "looked," "noticed" in narration',builtin:true},{id:"gps_lite",name:"GPS-Lite: Garden-Path Scanner",short:"GPS-LITE",signalType:"defect",description:"Heuristic first-parse friction",builtin:true}];
const SK="afl_cockpit_v5",CSK="afl_cockpit_custom_signals",DSK="afl_cockpit_dismissals";
function loadSessions(){try{return JSON.parse(localStorage.getItem(SK)||"[]");}catch{return[];}}
function saveSessions(s){try{localStorage.setItem(SK,JSON.stringify(s.map(x=>({...x,findings:x.findings?.slice(0,2000)}))));}catch{}}
function loadCustomSignals(){try{return JSON.parse(localStorage.getItem(CSK)||"[]");}catch{return[];}}
function saveCustomSignals(s){try{localStorage.setItem(CSK,JSON.stringify(s));}catch{}}
function loadDismissals(){try{return new Set(JSON.parse(localStorage.getItem(DSK)||"[]"));}catch{return new Set();}}
function saveDismissals(s){try{localStorage.setItem(DSK,JSON.stringify([...s]));}catch{}}
function computeHealth(findings,wc){if(!wc)return 0;const a=findings.filter(f=>f.disposition===D.ACTIONABLE).length;return Math.round(Math.max(40,Math.min(100,100-(a/wc)*1000*0.55))*10)/10;}
function exportToXlsx(findings,name){const rows=[["Line #","Sentence ID","Paragraph ID","Chapter ID","Flagged Sentence","Issue Type","Signal Type","Disposition","Confidence","Reason"]];for(const f of findings){const pid=f.paragraphId||(f.sentenceId?f.sentenceId.split(':').slice(0,2).join(':'):"");rows.push([f.lineNum||"",f.sentenceId||"",pid,f.chapter,f.sentence,f.issue,f.signalType||"defect",f.disposition||"",f.confidence||"",f.reason||""]);}const ws=XLSX.utils.aoa_to_sheet(rows);ws["!cols"]=[8,28,22,18,60,18,12,12,10,40].map(w=>({wch:w}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Findings");XLSX.writeFile(wb,`${name}_findings.xlsx`);}
function normSentence(s){return(s||"").toLowerCase().replace(/[^a-z0-9\s]/g,"").replace(/\s+/g," ").trim();}
function overlapRatio(a,b){if(!a||!b)return 0;if(a===b)return 1;const sh=a.length<b.length?a:b,lo=a.length<b.length?b:a;if(lo.includes(sh))return 1;const sa=new Set(a.split(" ")),sb=new Set(b.split(" "));return[...sa].filter(w=>sb.has(w)).length/Math.max(sa.size,sb.size);}
const GS=/^(sheet\d*|data|findings|results|output|list|export|tab\d*)$/i;
function cleanFn(f){return f.replace(/\.(xlsx|xls|csv)$/i,"").replace(/[_\-]+/g," ").replace(/\b(A\d{2,3}|v\d+|\(\d+\)|\d{4,})\b/gi,"").replace(/\s{2,}/g," ").trim();}
async function parseImportedXlsx(file){const ab=await file.arrayBuffer();const wb=XLSX.read(ab,{type:"array"});const fn=cleanFn(file.name);const sheets=[];for(const sn of wb.SheetNames){const ws=wb.Sheets[sn];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});const sents=[];for(const r of rows){const c=String(r[0]||"").trim();if(!c||c.length<15||!/[a-z]/i.test(c))continue;sents.push(c);}if(!sents.length)continue;const ig=GS.test(sn.trim());const sn2=ig?(wb.SheetNames.length>1?`${fn} — ${sn}`:fn):sn.trim();sheets.push({signalName:sn2,sentences:sents});}return sheets;}
function deduplicateImported(sents,existing){const en=existing.map(f=>normSentence(f.sentence));return sents.map(s=>({sentence:s,isDupe:en.some(e=>overlapRatio(normSentence(s),e)>=0.85)}));}

function SignalPanel({builtinSignals,customSignals,onToggleCustom,onDeleteCustom,onRename,onAddSignal,onImport,onClose}){
  const ir=useRef();const[eid,setEid]=useState(null);const[en,setEn]=useState("");
  return(<div style={{position:"fixed",inset:0,background:C.overlay,zIndex:999,display:"flex",justifyContent:"flex-end"}} onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div style={{background:C.smokeDark,borderLeft:`1px solid ${C.smokeLight}`,width:"340px",maxWidth:"90vw",height:"100%",overflowY:"auto"}}>
      <div style={{padding:"20px 24px 16px",borderBottom:`1px solid ${C.smokeLight}`,display:"flex",justifyContent:"space-between"}}>
        <div style={{fontSize:"14px",fontWeight:700,letterSpacing:"0.18em",color:C.amber}}>SIGNALS</div>
        <button onClick={onClose} style={{background:"none",border:"none",color:C.textMuted,fontSize:"20px",cursor:"pointer"}}>×</button>
      </div>
      <div style={{padding:"16px 24px"}}>
        <div style={{fontSize:"9px",letterSpacing:"0.18em",color:C.textMuted,marginBottom:"10px"}}>BUILT-IN</div>
        {builtinSignals.map(s=>(<div key={s.id} style={{padding:"8px 0",borderBottom:`1px solid ${C.smokeLight}`}}><div style={{fontSize:"12px",color:C.parchment}}>{s.name}</div><div style={{fontSize:"10px",color:C.textMuted,fontStyle:"italic",fontFamily:"IM Fell English, serif"}}>{s.description}</div></div>))}
        <div style={{fontSize:"9px",letterSpacing:"0.18em",color:C.textMuted,marginTop:"20px",marginBottom:"10px"}}>CUSTOM</div>
        {!customSignals.length&&<div style={{fontSize:"11px",color:C.textMuted,fontStyle:"italic",fontFamily:"IM Fell English, serif",marginBottom:"12px"}}>No custom signals.</div>}
        {customSignals.map(s=>(<div key={s.id} style={{padding:"10px 0",borderBottom:`1px solid ${C.smokeLight}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div style={{flex:1,minWidth:0}}>
              {eid===s.id?(<input autoFocus value={en} onChange={e=>setEn(e.target.value)} onBlur={()=>{onRename(s.id,en);setEid(null);}} onKeyDown={e=>{if(e.key==="Enter"){onRename(s.id,en);setEid(null);}if(e.key==="Escape")setEid(null);}} style={{background:"transparent",border:"none",borderBottom:`1px solid ${C.amber}`,color:C.parchment,fontSize:"12px",fontFamily:"Barlow Condensed, sans-serif",outline:"none",width:"100%"}}/>
              ):(<div onClick={()=>{setEid(s.id);setEn(s.name);}} style={{fontSize:"12px",color:s.enabled?C.parchment:C.textMuted,cursor:"text",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.name}</div>)}
            </div>
            <div style={{display:"flex",gap:"10px",marginLeft:"10px",flexShrink:0}}>
              <button onClick={()=>onToggleCustom(s.id)} style={{background:"none",border:"none",color:s.enabled?C.pass:C.textMuted,fontSize:"9px",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>{s.enabled?"ON":"OFF"}</button>
              <button onClick={()=>onDeleteCustom(s.id)} style={{background:"none",border:"none",color:C.textMuted,fontSize:"13px",cursor:"pointer"}}>×</button>
            </div>
          </div>
          <div style={{fontSize:"10px",color:C.textMuted,fontStyle:"italic",fontFamily:"IM Fell English, serif",marginTop:"2px"}}>{s.imported?s.description:`${s.keywords?.slice(0,4).join(", ")} · ${s.scope}`}</div>
        </div>))}
        <button onClick={onAddSignal} style={{marginTop:"16px",background:"transparent",border:`1px solid ${C.amber}`,color:C.amber,padding:"10px 20px",fontSize:"11px",letterSpacing:"0.15em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",width:"100%"}}>+ ADD SIGNAL</button>
        <div style={{borderTop:`1px solid ${C.smokeLight}`,margin:"20px 0"}}/>
        <div style={{fontSize:"9px",letterSpacing:"0.18em",color:C.textMuted,marginBottom:"8px"}}>IMPORT FINDINGS LIST</div>
        <div style={{fontSize:"11px",color:C.textDim,fontFamily:"IM Fell English, serif",fontStyle:"italic",marginBottom:"12px",lineHeight:1.4}}>Sheet name = signal name, col A = flagged sentences.</div>
        <input ref={ir} type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>{if(e.target.files[0]){onImport(e.target.files[0]);e.target.value="";}}}/>
        <button onClick={()=>ir.current?.click()} style={{background:"transparent",border:`1px solid ${C.smokeLight}`,color:C.textDim,padding:"10px 20px",fontSize:"11px",letterSpacing:"0.15em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",width:"100%"}}>↑ IMPORT FINDINGS .XLSX</button>
      </div>
    </div>
  </div>);}

function AddSignalModal({sentenceIndex,onSave,onClose}){
  const[name,setName]=useState("");const[kw,setKw]=useState("");const[scope,setScope]=useState("narration");const[st,setSt]=useState("defect");const[preview,setPreview]=useState(null);const[running,setRunning]=useState(false);
  const el=scope==="narration"?["NARRATION"]:scope==="dialogue"?["PURE_DIALOGUE","DIALOGUE_TAG","DIALOGUE_ACTION"]:["NARRATION","PURE_DIALOGUE","DIALOGUE_TAG","DIALOGUE_ACTION"];
  const runPreview=useCallback(()=>{const kws=kw.split(",").map(k=>k.trim()).filter(Boolean);if(!kws.length||!sentenceIndex?.length){setPreview(null);return;}setRunning(true);setTimeout(()=>{const hits=[];for(const s of sentenceIndex){if(!el.includes(s.cls))continue;for(const k of kws){try{const p=new RegExp(`\\b${k.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}\\b`,"gi");if(p.test(s.text)){hits.push(s);break;}}catch{}}}setPreview({count:hits.length,examples:hits.slice(0,3)});setRunning(false);},80);},[kw,scope,sentenceIndex]);
  useEffect(()=>{const t=setTimeout(runPreview,400);return()=>clearTimeout(t);},[kw,scope,runPreview]);
  const light=preview?(preview.count===0?{c:"#666",l:"No matches"}:preview.count<=10?{c:C.pass,l:"Surgical"}:preview.count<=50?{c:C.amber,l:"Review"}:{c:C.regress,l:"Noisy"}):null;
  return(<div style={{position:"fixed",inset:0,background:C.overlay,zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div style={{background:C.smokeDark,border:`1px solid ${C.smokeLight}`,width:"520px",maxWidth:"94vw",maxHeight:"90vh",overflowY:"auto",padding:"24px"}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:"20px"}}><div style={{fontSize:"16px",fontWeight:700,color:C.amber}}>ADD SIGNAL</div><button onClick={onClose} style={{background:"none",border:"none",color:C.textMuted,fontSize:"20px",cursor:"pointer"}}>×</button></div>
      <div style={{marginBottom:"16px"}}><div style={{fontSize:"10px",color:C.textMuted,marginBottom:"4px"}}>NAME</div><input value={name} onChange={e=>setName(e.target.value)} style={{background:C.charcoal,border:`1px solid ${C.smokeLight}`,color:C.parchment,padding:"9px 12px",width:"100%",outline:"none",fontFamily:"Barlow Condensed, sans-serif",fontSize:"13px"}}/></div>
      <div style={{marginBottom:"16px"}}><div style={{fontSize:"10px",color:C.textMuted,marginBottom:"4px"}}>KEYWORDS — comma separated</div><input value={kw} onChange={e=>setKw(e.target.value)} style={{background:C.charcoal,border:`1px solid ${C.smokeLight}`,color:C.parchment,padding:"9px 12px",width:"100%",outline:"none",fontFamily:"Barlow Condensed, sans-serif",fontSize:"13px"}}/></div>
      <div style={{marginBottom:"16px"}}><div style={{fontSize:"10px",color:C.textMuted,marginBottom:"4px"}}>SCOPE</div><div style={{display:"flex",gap:"8px"}}>{[["narration","Narration"],["dialogue","Dialogue"],["all","All"]].map(([v,l])=>(<button key={v} onClick={()=>setScope(v)} style={{background:scope===v?C.amber:C.smokeLight,color:scope===v?C.charcoal:C.textDim,border:"none",padding:"7px 14px",fontSize:"11px",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>{l}</button>))}</div></div>
      <div style={{marginBottom:"20px"}}><div style={{fontSize:"10px",color:C.textMuted,marginBottom:"4px"}}>TYPE</div><div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>{Object.entries(SIGNAL_TYPES).map(([v,m])=>(<button key={v} onClick={()=>setSt(v)} style={{background:st===v?m.color:C.smokeLight,color:st===v?"#fff":C.textDim,border:"none",padding:"7px 12px",fontSize:"11px",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>{m.label}</button>))}</div></div>
      {preview&&sentenceIndex?.length>0&&(<div style={{marginBottom:"20px",padding:"12px",background:C.charcoal}}>
        <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"10px"}}><div style={{width:"12px",height:"12px",borderRadius:"50%",background:light.c}}/><span style={{fontSize:"24px",fontWeight:700,color:light.c,fontFamily:"Barlow Condensed, sans-serif"}}>{preview.count}</span><span style={{fontSize:"10px",color:C.textMuted}}>{light.l}</span></div>
        {preview.examples.map((ex,i)=>(<div key={i} style={{fontSize:"11px",color:C.dimParch,fontFamily:"IM Fell English, serif",borderLeft:`2px solid ${C.smokeLight}`,paddingLeft:"8px",marginBottom:"6px"}}>{ex.text.slice(0,150)}</div>))}
      </div>)}
      <div style={{display:"flex",gap:"10px",justifyContent:"flex-end"}}>
        <button onClick={onClose} style={{background:"transparent",border:`1px solid ${C.smokeLight}`,color:C.textDim,padding:"9px 18px",fontSize:"11px",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>CANCEL</button>
        <button onClick={()=>{const kws=kw.split(",").map(k=>k.trim()).filter(Boolean);if(!name.trim()||!kws.length)return;onSave({id:"custom_"+Date.now(),name:name.trim(),short:name.trim().toUpperCase().slice(0,10),description:`Custom: ${kws.slice(0,4).join(", ")}`,keywords:kws,scope,signalType:st,enabled:true,builtin:false});}} disabled={!name.trim()||!kw.trim()} style={{background:(!name.trim()||!kw.trim())?C.smokeLight:C.amber,color:(!name.trim()||!kw.trim())?C.textMuted:C.charcoal,border:"none",padding:"9px 20px",fontSize:"11px",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",fontWeight:700}}>ADD TO DASHBOARD</button>
      </div>
    </div>
  </div>);}

function SignalCard({signal,actionable,raw,prevActionable,onClick}){
  const delta=prevActionable!==undefined?actionable-prevActionable:null;const excl=Math.max(0,(raw||0)-actionable);const st=SIGNAL_TYPES[signal.signalType]||SIGNAL_TYPES.defect;const isd=signal.signalType==="defect"||signal.signalType==="reader_note";const isp=isd&&actionable===0;const isr=isd&&delta!==null&&delta>0;const bc=isr?C.regress:isp?C.pass:C.smokeLight;
  return(<div onClick={onClick} style={{background:C.smokeDark,border:`1px solid ${bc}`,padding:"16px 20px",position:"relative",cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.borderColor=C.amber} onMouseLeave={e=>e.currentTarget.style.borderColor=bc}>
    <div style={{position:"absolute",top:0,left:0,right:0,height:"2px",background:isr?C.regress:st.color,opacity:isr?1:0.6}}/>
    {signal.builtin===false&&<div style={{position:"absolute",top:"8px",right:"10px",fontSize:"8px",color:C.textMuted,letterSpacing:"0.1em"}}>{signal.imported?"IMPORTED":"CUSTOM"}</div>}
    <div style={{display:"flex",alignItems:"center",gap:"6px",marginBottom:"8px"}}>
      <div style={{fontSize:"10px",letterSpacing:"0.15em",color:C.textMuted,fontFamily:"Barlow Condensed, sans-serif"}}>{signal.short}</div>
      <div style={{fontSize:"8px",color:st.color,letterSpacing:"0.12em",fontFamily:"Barlow Condensed, sans-serif",opacity:0.85}}>{st.label.toUpperCase()}</div>
    </div>
    <div style={{display:"flex",alignItems:"flex-end",gap:"10px"}}>
      <div style={{fontSize:"36px",fontFamily:"Barlow Condensed, sans-serif",fontWeight:700,lineHeight:1,color:isp?C.pass:isr?C.regress:C.parchment}}>{actionable.toLocaleString()}</div>
      {delta!==null&&delta!==0&&<div style={{fontSize:"12px",fontFamily:"Barlow Condensed, sans-serif",color:isd?(delta>0?C.regress:C.pass):C.textDim,paddingBottom:"4px"}}>{delta>0?`▲ +${delta}`:`▼ ${Math.abs(delta)}`}</div>}
      {isp&&<div style={{fontSize:"11px",color:C.pass,paddingBottom:"4px",fontFamily:"Barlow Condensed, sans-serif"}}>✓ PASS</div>}
    </div>
    {excl>0&&<div style={{fontSize:"10px",color:C.textMuted,marginTop:"3px",fontFamily:"Barlow Condensed, sans-serif"}}>{excl.toLocaleString()} excluded</div>}
    <div style={{fontSize:"11px",color:C.textDim,marginTop:"8px",fontFamily:"IM Fell English, serif",fontStyle:"italic",lineHeight:1.3}}>{signal.description}</div>
  </div>);}

function HealthBar({score,prevScore}){const d=prevScore!==undefined?(score-prevScore).toFixed(1):null;const p=Math.max(0,Math.min(100,((score-40)/60)*100));return(<div style={{marginBottom:"28px"}}><div style={{display:"flex",alignItems:"baseline",gap:"16px",marginBottom:"10px"}}><div style={{fontSize:"64px",fontFamily:"Barlow Condensed, sans-serif",fontWeight:700,color:C.amber,lineHeight:1}}>{score}</div><div><div style={{fontSize:"14px",color:C.textMuted,letterSpacing:"0.12em",fontFamily:"Barlow Condensed, sans-serif"}}>/ 100</div>{d!==null&&<div style={{fontSize:"15px",fontFamily:"Barlow Condensed, sans-serif",color:parseFloat(d)>=0?C.pass:C.regress}}>{parseFloat(d)>=0?`▲ +${d}`:`▼ ${d}`} since previous</div>}</div></div><div style={{height:"3px",background:C.smokeLight}}><div style={{height:"100%",width:`${p}%`,background:`linear-gradient(90deg,${C.amber},${C.redOrange})`}}/></div></div>);}

// ── Review Queue ──────────────────────────────────────────────────────────────
const RQ_KEY="afl_review_queue_v1";
const RQ_STATUS={FIX:"Fix",SKIP:"Skip",DESKTOP:"Needs Desktop"};

function ReviewQueue({findings,draftName,onDismiss}){
  const[reviews,setReviews]=useState(()=>{try{return JSON.parse(localStorage.getItem(RQ_KEY)||"{}");}catch{return{};}});
  const[index,setIndex]=useState(0);
  const[draft,setDraft]=useState("");
  const[sessionCount,setSessionCount]=useState(0);
  const[toast,setToast]=useState("");
  const[popKey,setPopKey]=useState(0);
  const toastTimer=useRef(null);
  const total=findings.length;
  const fkey=f=>f.sentenceId||f.sentence?.slice(0,80)||"x";

  useEffect(()=>{
    const first=findings.findIndex(f=>!reviews[fkey(f)]?.status);
    setIndex(first===-1?0:first);
  },[findings.length]);

  useEffect(()=>{
    const f=findings[index];
    if(f) setDraft(reviews[fkey(f)]?.note||"");
  },[index,findings]);

  const reviewedCount=useMemo(()=>findings.filter(f=>reviews[fkey(f)]?.status).length,[reviews,findings]);
  const pct=total?Math.round((reviewedCount/total)*100):0;

  function saveR(r){setReviews(r);try{localStorage.setItem(RQ_KEY,JSON.stringify(r));}catch{}}
  function showToast(msg){setToast(msg);if(toastTimer.current)clearTimeout(toastTimer.current);toastTimer.current=setTimeout(()=>setToast(""),1600);}

  function advanceFrom(i,revs){
    for(let s=1;s<=findings.length;s++){const j=(i+s)%findings.length;if(!revs[fkey(findings[j])]?.status)return j;}
    return -1;
  }

  function commit(status){
    const f=findings[index];if(!f)return;
    const k=fkey(f);
    const updated={...reviews,[k]:{note:draft,status,ts:new Date().toISOString()}};
    saveR(updated);setSessionCount(c=>c+1);setPopKey(p=>p+1);
    const next=advanceFrom(index,updated);
    if(next!==-1)setIndex(next);
    showToast(status===RQ_STATUS.SKIP?"Skipped →":"Saved \u2713");
  }

  function jump(delta){
    const f=findings[index];
    if(f){const k=fkey(f);saveR({...reviews,[k]:{...reviews[k],note:draft}});}
    setIndex(j=>Math.min(Math.max(j+delta,0),total-1));
    setPopKey(p=>p+1);
  }

  function clearReviews(){saveR({});setIndex(0);setSessionCount(0);showToast("Reviews cleared");}

  function exportXlsx(){
    const rows=[["Line #","Sentence ID","Paragraph ID","Chapter ID","Flagged Sentence","Issue Type","Signal Type","Disposition","Confidence","Reason","Author Note","Status","Reviewed At"],...findings.map(f=>{const k=fkey(f);const rev=reviews[k]||{};const pid=f.paragraphId||(f.sentenceId?f.sentenceId.split(":").slice(0,2).join(":"):"");return[f.lineNum||"",f.sentenceId||"",pid,f.chapter,f.sentence,f.issue,f.signalType||"",f.disposition||"",f.confidence||"",f.reason||"",rev.note||"",rev.status||"",rev.ts||""];})];
    const ws=XLSX.utils.aoa_to_sheet(rows);ws["!cols"]=[8,28,22,18,60,18,12,12,10,40,40,14,20].map(w=>({wch:w}));
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Review");
    XLSX.writeFile(wb,`${draftName||"afl"}_reviewed.xlsx`);
    showToast("Export ready");
  }

  const sChip=st=>{const m={[RQ_STATUS.FIX]:{bg:"#1A3A24",c:C.pass},[RQ_STATUS.SKIP]:{bg:C.smokeLight,c:C.textMuted},[RQ_STATUS.DESKTOP]:{bg:"#2A2010",c:C.amber}}[st]||{bg:C.smokeLight,c:C.textMuted};return <span style={{background:m.bg,color:m.c,fontSize:"9px",letterSpacing:"0.15em",padding:"2px 8px",fontFamily:"Barlow Condensed, sans-serif",fontWeight:700}}>{st.toUpperCase()}</span>;};

  if(total===0) return(
    <div style={{textAlign:"center",padding:"60px 32px",color:C.textMuted}}>
      <div style={{fontSize:"32px",marginBottom:"16px",opacity:0.2}}>&#9678;</div>
      <div style={{fontSize:"14px",letterSpacing:"0.2em",marginBottom:"8px"}}>NO FINDINGS IN QUEUE</div>
      <div style={{fontSize:"11px",fontFamily:"IM Fell English, serif",fontStyle:"italic",lineHeight:1.6}}>Set filters in the Findings tab first, then switch here to review.</div>
    </div>
  );

  const f=findings[index];
  const k=f?fkey(f):null;
  const cs=k?reviews[k]?.status:null;
  const allDone=reviewedCount===total;

  return(
    <div style={{maxWidth:"680px"}}>
      <div style={{marginBottom:"20px"}}>
        <div style={{height:"3px",background:C.smokeLight,marginBottom:"8px"}}>
          <div style={{height:"100%",width:pct+"%",background:"linear-gradient(90deg,"+C.amber+","+C.redOrange+")",transition:"width 300ms ease"}}/>
        </div>
        <div style={{display:"flex",gap:"24px",fontSize:"11px",color:C.textMuted,fontFamily:"Barlow Condensed, sans-serif",letterSpacing:"0.08em",flexWrap:"wrap"}}>
          <span><span style={{color:C.parchment,fontWeight:700}}>{reviewedCount}</span> / {total} reviewed &middot; {pct}%</span>
          <span><span style={{color:C.parchment,fontWeight:700}}>{sessionCount}</span> this session</span>
          <span>{draftName}</span>
        </div>
      </div>

      {allDone&&(
        <div style={{background:"#1A3A24",border:"1px solid "+C.pass,padding:"16px 20px",marginBottom:"16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontSize:"12px",fontWeight:700,color:C.pass,letterSpacing:"0.15em",fontFamily:"Barlow Condensed, sans-serif"}}>ALL {total} REVIEWED</div>
            <div style={{fontSize:"11px",color:C.textMuted,marginTop:"2px",fontFamily:"Barlow Condensed, sans-serif"}}>{findings.filter(f=>reviews[fkey(f)]?.status===RQ_STATUS.DESKTOP).length} needs desktop &middot; {findings.filter(f=>reviews[fkey(f)]?.status===RQ_STATUS.SKIP).length} skipped</div>
          </div>
          <button onClick={exportXlsx} style={{background:"transparent",border:"1px solid "+C.amber,color:C.amber,padding:"9px 18px",fontSize:"11px",letterSpacing:"0.12em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>EXPORT .XLSX</button>
        </div>
      )}

      {f&&(
        <div key={popKey} style={{background:C.smokeDark,border:"1px solid "+(cs===RQ_STATUS.FIX?C.pass:cs===RQ_STATUS.DESKTOP?C.amber:C.smokeLight),padding:"20px",marginBottom:"12px",borderLeft:"3px solid "+(cs===RQ_STATUS.FIX?C.pass:cs===RQ_STATUS.DESKTOP?C.amber:C.smokeLight)}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"14px",gap:"10px"}}>
            <div style={{display:"flex",gap:"8px",alignItems:"center",flexWrap:"wrap"}}>
              <div style={{background:C.smokeLight,color:C.amber,fontSize:"9px",letterSpacing:"0.1em",padding:"2px 6px",fontFamily:"Barlow Condensed, sans-serif"}}>{f.issue}</div>
              {f.lineNum&&<div style={{fontSize:"10px",color:C.textMuted,fontFamily:"Barlow Condensed, sans-serif"}}>L{f.lineNum.toLocaleString()}</div>}
              <div style={{fontSize:"10px",color:C.textMuted,fontFamily:"Barlow Condensed, sans-serif"}}>{(f.chapter||"").slice(0,24)}</div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:"10px",flexShrink:0}}>
              {cs&&sChip(cs)}
              <div style={{fontSize:"10px",color:C.textMuted,fontFamily:"Barlow Condensed, sans-serif"}}>{index+1} / {total}</div>
            </div>
          </div>
          <div style={{borderLeft:"2px solid "+C.amber,paddingLeft:"14px",marginBottom:"14px"}}>
            <div style={{fontSize:"15px",color:C.parchment,fontFamily:"IM Fell English, serif",lineHeight:1.7}}>{f.sentence}</div>
          </div>
          {f.reason&&<div style={{fontSize:"10px",color:C.textMuted,letterSpacing:"0.06em",fontFamily:"Barlow Condensed, sans-serif",marginBottom:"14px"}}>{"\u2193"} {f.reason}</div>}
          {f.sentenceId&&<div style={{fontSize:"9px",color:C.smokeLight,letterSpacing:"0.08em",fontFamily:"Barlow Condensed, sans-serif",marginBottom:"14px",borderTop:"1px solid "+C.smokeLight,paddingTop:"10px"}}>{f.sentenceId}</div>}
          <div style={{fontSize:"9px",letterSpacing:"0.18em",color:C.textMuted,fontFamily:"Barlow Condensed, sans-serif",marginBottom:"6px"}}>AUTHOR NOTE</div>
          <textarea value={draft} onChange={e=>setDraft(e.target.value)} onKeyDown={e=>{if((e.metaKey||e.ctrlKey)&&e.key==="Enter")commit(RQ_STATUS.FIX);}} placeholder="Revision or note for the desktop pass..." rows={4} style={{width:"100%",background:C.charcoal,border:"1px solid "+C.smokeLight,color:C.parchment,padding:"10px 12px",fontSize:"14px",fontFamily:"IM Fell English, serif",lineHeight:1.6,resize:"vertical",outline:"none",boxSizing:"border-box"}}/>
          <div style={{fontSize:"9px",color:C.textMuted,marginTop:"4px",fontFamily:"Barlow Condensed, sans-serif"}}>Cmd+Enter to save &amp; next</div>
        </div>
      )}

      {f&&(
        <div style={{display:"grid",gap:"8px",marginBottom:"14px"}}>
          <button onClick={()=>commit(RQ_STATUS.FIX)} style={{background:"#1A3A24",border:"1px solid "+C.pass,color:C.pass,padding:"11px 18px",fontSize:"12px",letterSpacing:"0.12em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",fontWeight:700,width:"100%"}}>SAVE NOTE &amp; NEXT</button>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px"}}>
            <button onClick={()=>commit(RQ_STATUS.SKIP)} style={{background:C.smokeDark,border:"1px solid "+C.smokeLight,color:C.textMuted,padding:"11px 14px",fontSize:"11px",letterSpacing:"0.12em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",width:"100%"}}>SKIP</button>
            <button onClick={()=>commit(RQ_STATUS.DESKTOP)} style={{background:"#2A2010",border:"1px solid "+C.amber,color:C.amber,padding:"11px 14px",fontSize:"11px",letterSpacing:"0.12em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",width:"100%"}}>NEEDS DESKTOP</button>
            <button onClick={()=>{if(!f||!onDismiss)return;onDismiss(f);setPopKey(p=>p+1);setSessionCount(c=>c+1);const next=advanceFrom(index,reviews);if(next!==-1)setIndex(next);showToast("Author skipped — removed from count");}} style={{background:"#1A1A2A",border:"1px solid #5A5A8A",color:"#8A8ACA",padding:"11px 14px",fontSize:"11px",letterSpacing:"0.12em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",width:"100%"}}>AUTHOR SKIP</button>
          </div>
        </div>
      )}

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"10px"}}>
        <button onClick={()=>jump(-1)} disabled={index===0} style={{background:"none",border:"none",color:index===0?C.smokeLight:C.textMuted,cursor:index===0?"default":"pointer",fontSize:"12px",fontFamily:"Barlow Condensed, sans-serif",letterSpacing:"0.1em"}}>&#8592; BACK</button>
        <button onClick={exportXlsx} style={{background:"none",border:"none",color:C.amber,cursor:"pointer",fontSize:"12px",fontFamily:"Barlow Condensed, sans-serif",letterSpacing:"0.1em"}}>EXPORT .XLSX</button>
        <button onClick={()=>jump(1)} disabled={index===total-1} style={{background:"none",border:"none",color:index===total-1?C.smokeLight:C.textMuted,cursor:index===total-1?"default":"pointer",fontSize:"12px",fontFamily:"Barlow Condensed, sans-serif",letterSpacing:"0.1em"}}>FORWARD &#8594;</button>
      </div>
      <div style={{textAlign:"center"}}>
        <button onClick={clearReviews} style={{background:"none",border:"none",color:C.textMuted,cursor:"pointer",fontSize:"10px",fontFamily:"Barlow Condensed, sans-serif"}}>clear all reviews</button>
      </div>

      {toast&&<div style={{position:"fixed",bottom:"24px",left:"50%",transform:"translateX(-50%)",background:C.parchment,color:C.charcoal,fontSize:"12px",fontWeight:700,letterSpacing:"0.1em",padding:"9px 20px",boxShadow:"0 4px 16px rgba(0,0,0,0.4)",fontFamily:"Barlow Condensed, sans-serif",zIndex:100}}>{toast}</div>}
    </div>
  );
}

export default function App(){
  const[sessions,setSessions]=useState([]);const[activeSession,setActiveSession]=useState(null);const[sentenceIndex,setSentenceIndex]=useState(null);const[customSignals,setCustomSignals]=useState([]);const[uploading,setUploading]=useState(false);const[uploadStep,setUploadStep]=useState("");const[uploadTime,setUploadTime]=useState(null);const[activeTab,setActiveTab]=useState("dashboard");const[filterSignal,setFilterSignal]=useState("all");const[filterDisp,setFilterDisp]=useState("actionable");const[search,setSearch]=useState("");const[sortBy,setSortBy]=useState("line");const[showPanel,setShowPanel]=useState(false);const[showAddModal,setShowAddModal]=useState(false);
  const[dismissals,setDismissals]=useState(()=>loadDismissals());
  const fileInputRef=useRef();const sirRef=useRef(null);
  useEffect(()=>{const s=loadSessions();if(s.length){setSessions(s);setActiveSession(s[s.length-1]);}setCustomSignals(loadCustomSignals());},[]);
  useEffect(()=>{sirRef.current=sentenceIndex;},[sentenceIndex]);
  const allSignals=[...BUILTIN_SIGNALS,...customSignals];
  const runAndBuild=useCallback(async(index,draftName)=>{const findings=runBuiltinSignals(index,customSignals);const wc=index.reduce((a,s)=>a+s.wordCount,0);const health=computeHealth(findings,wc);const counts={},rc={};for(const sig of allSignals){counts[sig.id]=0;rc[sig.id]=0;}for(const f of findings){rc[f.signal_id]=(rc[f.signal_id]||0)+1;if(f.disposition!==D.EXCLUDED)counts[f.signal_id]=(counts[f.signal_id]||0)+1;}return{draftName,wordCount:wc,health,counts,rawCounts:rc,findings};},[customSignals,allSignals]);
  const handleTaxonomyImport=useCallback(async(file)=>{setUploading(true);const t0=Date.now();try{setUploadStep("Reading taxonomy…");const ab=await file.arrayBuffer();const wb=XLSX.read(ab,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{defval:""});if(!rows.length)throw new Error("No rows found. Run build_afl_index.py first.");const req=["sentence_id","sentence_text","chapter_id","paragraph_id","version_id"];const miss=req.filter(r=>!Object.keys(rows[0]).includes(r));if(miss.length)throw new Error(`Missing: ${miss.join(", ")}. Run build_afl_index.py first.`);setUploadStep("Building index…");await new Promise(r=>setTimeout(r,60));const idx=buildIndexFromTaxonomy(rows);setSentenceIndex(idx);sirRef.current=idx;setUploadStep("Running signals…");await new Promise(r=>setTimeout(r,60));const vid=String(rows[0].version_id||file.name.replace(".xlsx",""));const result=await runAndBuild(idx,vid);const session={id:Date.now(),...result,uploadedAt:new Date().toISOString(),source:"taxonomy"};setSessions(prev=>{const next=[...prev,session];saveSessions(next);return next;});setActiveSession(session);setUploadTime(((Date.now()-t0)/1000).toFixed(1));setActiveTab("dashboard");}catch(err){alert("Error: "+err.message);}finally{setUploading(false);setUploadStep("");}},[ runAndBuild]);
  const handleUpload=useCallback(async(file)=>{if(!file?.name.endsWith(".docx")){alert("Please upload a .docx file.");return;}setUploading(true);const t0=Date.now();try{setUploadStep("Parsing manuscript…");const ab=await file.arrayBuffer();const{value:rawText}=await mammoth.extractRawText({arrayBuffer:ab});setUploadStep("Classifying…");await new Promise(r=>setTimeout(r,60));const idx=buildSentenceIndex(rawText);setSentenceIndex(idx);sirRef.current=idx;setUploadStep("Running signals…");await new Promise(r=>setTimeout(r,60));const result=await runAndBuild(idx,file.name.replace(".docx",""));const session={id:Date.now(),...result,uploadedAt:new Date().toISOString(),source:"docx"};setSessions(prev=>{const next=[...prev,session];saveSessions(next);return next;});setActiveSession(session);setUploadTime(((Date.now()-t0)/1000).toFixed(1));setActiveTab("dashboard");}catch(err){alert("Error: "+err.message);}finally{setUploading(false);setUploadStep("");}},[ runAndBuild]);
  const handleDrop=useCallback(e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(!f)return;if(f.name.endsWith(".xlsx"))handleTaxonomyImport(f);else handleUpload(f);},[handleTaxonomyImport,handleUpload]);
  const addCustomSignal=useCallback(sig=>{const u=[...customSignals,sig];setCustomSignals(u);saveCustomSignals(u);setShowAddModal(false);setShowPanel(false);},[customSignals]);
  const toggleCustom=useCallback(id=>{const u=customSignals.map(s=>s.id===id?{...s,enabled:!s.enabled}:s);setCustomSignals(u);saveCustomSignals(u);},[customSignals]);
  const deleteCustom=useCallback(id=>{const u=customSignals.filter(s=>s.id!==id);setCustomSignals(u);saveCustomSignals(u);},[customSignals]);
  const renameCustom=useCallback((id,n)=>{if(!n.trim())return;const u=customSignals.map(s=>s.id===id?{...s,name:n.trim(),short:n.trim().toUpperCase().slice(0,10)}:s);setCustomSignals(u);saveCustomSignals(u);},[customSignals]);
  const importFindings=useCallback(async(file)=>{try{const sheets=await parseImportedXlsx(file);if(!sheets.length){alert("No valid sheets.");return;}const ef=activeSession?.findings||[];const ci=sirRef.current||[];const nm=ci.map(s=>({norm:normSentence(s.text),lineNum:s.lineNum,chapter:s.chapter,sentenceId:s.sentenceId,paragraphId:s.paragraphId,hash:s.hash}));const ns=[];for(const{signalName,sentences}of sheets){const dd=deduplicateImported(sentences,ef);const id="imported_"+Date.now()+"_"+Math.random().toString(36).slice(2,6);const fi=dd.map(({sentence,isDupe})=>{const norm=normSentence(sentence);let ln=null,mc=null,mid=null,mh=null,mp=null,bs=0;for(const e of nm){const sc=overlapRatio(norm,e.norm);if(sc>bs&&sc>=0.6){bs=sc;ln=e.lineNum;mc=e.chapter;mid=e.sentenceId;mh=e.hash;mp=e.paragraphId;}}return{signal_id:id,signalType:"reader_note",chapter:mc||"UNMATCHED",sentenceId:mid,paragraphId:mp,hash:mh,lineNum:ln,sentence,issue:signalName.toUpperCase(),disposition:isDupe?"excluded":"review_only",confidence:Math.round(bs*100)||80,reason:isDupe?"Duplicate":`Imported · ${ln?`L${ln}`:"no match"}`};});ns.push({id,name:signalName,short:signalName.toUpperCase().slice(0,10),description:`Imported ${fi.filter(f=>f.disposition!=="excluded").length} unique`,keywords:[],scope:"all",enabled:true,builtin:false,imported:true,importedFindings:fi});}const u=[...customSignals,...ns];setCustomSignals(u);saveCustomSignals(u);if(activeSession){const af=ns.flatMap(s=>s.importedFindings);const us={...activeSession,findings:[...ef,...af],counts:{...activeSession.counts},rawCounts:{...activeSession.rawCounts}};for(const sig of ns){us.counts[sig.id]=sig.importedFindings.filter(f=>f.disposition!=="excluded").length;us.rawCounts[sig.id]=sig.importedFindings.length;}setActiveSession(us);setSessions(prev=>{const next=prev.map(s=>s.id===us.id?us:s);saveSessions(next);return next;});}setShowPanel(false);alert(`Imported ${ns.length} signal(s).`);}catch(err){alert("Import error: "+err.message);}},[activeSession,customSignals]);

  const handleDismiss=useCallback(f=>{
    const k=f.sentenceId||f.sentence?.slice(0,80)||"x";
    const updated=new Set([...dismissals,k]);
    setDismissals(updated);
    saveDismissals(updated);
  },[dismissals]);
  const fkey=f=>f.sentenceId||f.sentence?.slice(0,80)||"x";
  const ai=sessions.findIndex(s=>s.id===activeSession?.id);const ps=ai>0?sessions[ai-1]:null;
  // Exclude dismissed findings from all counts
  const activeFindings=(activeSession?.findings||[]).filter(f=>!dismissals.has(fkey(f)));
  const activeCounts=useMemo(()=>{const c={};for(const sig of allSignals) c[sig.id]=0;for(const f of activeFindings){if(f.disposition!==D.EXCLUDED) c[f.signal_id]=(c[f.signal_id]||0)+1;}return c;},[activeFindings,allSignals]);
  const ta=Object.values(activeCounts).reduce((a,b)=>a+b,0);
  const tr=activeFindings.length;
  const pt=ps?Object.values(ps.counts).reduce((a,b)=>a+b,0):null;const dc={actionable:C.amber,review_only:C.textDim};
  const ff=(()=>{let r=activeFindings.filter(f=>{if(filterSignal!=="all"&&f.signal_id!==filterSignal)return false;if(filterDisp!=="all"&&f.disposition!==filterDisp)return false;if(search&&!f.sentence.toLowerCase().includes(search.toLowerCase())&&!f.chapter.toLowerCase().includes(search.toLowerCase()))return false;return true;});if(sortBy==="line")r=[...r].sort((a,b)=>(a.lineNum||999999)-(b.lineNum||999999));if(sortBy==="chapter")r=[...r].sort((a,b)=>a.chapter.localeCompare(b.chapter)||(a.lineNum||0)-(b.lineNum||0));if(sortBy==="signal")r=[...r].sort((a,b)=>a.signal_id.localeCompare(b.signal_id)||(a.lineNum||0)-(b.lineNum||0));return r;})();
  return(<div style={{minHeight:"100vh",background:C.charcoal,color:C.parchment,fontFamily:"Barlow Condensed, sans-serif"}}>
    <div style={{borderBottom:`1px solid ${C.smokeLight}`,padding:"0 24px",display:"flex",alignItems:"center",justifyContent:"space-between",height:"56px"}}>
      <div style={{display:"flex",alignItems:"baseline",gap:"16px"}}><span style={{fontSize:"18px",fontWeight:700,letterSpacing:"0.15em",color:C.amber}}>AFL</span><span style={{fontSize:"14px",letterSpacing:"0.2em",color:C.textMuted}}>EDITING COCKPIT</span></div>
      <div style={{display:"flex",gap:"4px",alignItems:"center"}}>
        {["dashboard","findings","review"].map(tab=>(<button key={tab} onClick={()=>setActiveTab(tab)} style={{background:activeTab===tab?C.smokeLight:"transparent",border:"none",color:activeTab===tab?C.parchment:C.textMuted,padding:"8px 16px",fontSize:"12px",letterSpacing:"0.15em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",textTransform:"uppercase"}}>{tab==="review"?"REVIEW QUEUE":tab}{tab==="findings"&&activeSession&&<span style={{marginLeft:"8px",background:C.amber,color:C.charcoal,padding:"1px 6px",fontSize:"10px",fontWeight:700}}>{ta.toLocaleString()}</span>}</button>))}
        <button onClick={()=>setShowPanel(true)} style={{background:"none",border:`1px solid ${C.smokeLight}`,color:C.textDim,padding:"7px 12px",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",fontSize:"16px",marginLeft:"8px",lineHeight:1}}>☰</button>
      </div>
    </div>
    <div style={{padding:"28px 32px",maxWidth:"1160px",margin:"0 auto"}}>
      {/* Version progression strip — replaces both chip rail and bottom timeline */}
      {sessions.length>0&&(
        <div style={{marginBottom:"24px",overflowX:"auto",paddingBottom:"4px"}}>
          <div style={{display:"flex",alignItems:"center",minWidth:"min-content",gap:"0"}}>
            {sessions.map((s,i)=>{
              const isActive=s.id===activeSession?.id;
              const prev=sessions[i-1];
              const delta=prev?s.health-prev.health:null;
              const scoreColor=delta===null?C.textDim:delta>0?C.pass:delta<0?C.regress:C.textDim;
              return(
                <div key={s.id} style={{display:"flex",alignItems:"center"}}>
                  {i>0&&(
                    <div style={{display:"flex",alignItems:"center",margin:"0 2px"}}>
                      <div style={{width:"24px",height:"1px",background:delta>0?C.pass:delta<0?C.regress:C.smokeLight}}/>
                      <div style={{fontSize:"9px",color:delta>0?C.pass:delta<0?C.regress:C.smokeLight,marginLeft:"2px",marginRight:"2px",lineHeight:1}}>{delta>0?"↑":delta<0?"↓":"→"}</div>
                      <div style={{width:"4px",height:"1px",background:delta>0?C.pass:delta<0?C.regress:C.smokeLight}}/>
                    </div>
                  )}
                  <div onClick={()=>setActiveSession(s)} style={{cursor:"pointer",padding:"6px 10px",borderBottom:isActive?`2px solid ${C.amber}`:"2px solid transparent",transition:"all 0.15s",minWidth:"72px",textAlign:"center"}}
                    onMouseEnter={e=>e.currentTarget.style.background=C.smokeDark}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <div style={{fontSize:"18px",fontWeight:700,fontFamily:"Barlow Condensed, sans-serif",color:isActive?C.amber:scoreColor,lineHeight:1,marginBottom:"3px"}}>{s.health}</div>
                    <div style={{fontSize:"9px",color:isActive?C.parchment:C.textMuted,letterSpacing:"0.06em",fontFamily:"Barlow Condensed, sans-serif",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:"80px"}}>{s.draftName}</div>
                    {s.source==="taxonomy"&&<div style={{width:"4px",height:"4px",borderRadius:"50%",background:C.pass,margin:"2px auto 0"}}/>}
                  </div>
                </div>
              );
            })}
            {/* Add new button */}
            <div style={{marginLeft:"8px",display:"flex",alignItems:"center"}}>
              <div style={{width:"20px",height:"1px",background:C.smokeLight}}/>
              <div onClick={()=>fileInputRef.current?.click()} style={{cursor:"pointer",padding:"6px 10px",color:C.textMuted,fontSize:"10px",letterSpacing:"0.1em",fontFamily:"Barlow Condensed, sans-serif",whiteSpace:"nowrap",border:`1px dashed ${C.smokeLight}`,marginLeft:"4px"}}
                onMouseEnter={e=>{e.currentTarget.style.color=C.amber;e.currentTarget.style.borderColor=C.amber;}}
                onMouseLeave={e=>{e.currentTarget.style.color=C.textMuted;e.currentTarget.style.borderColor=C.smokeLight;}}>
                + ADD DRAFT
              </div>
            </div>
          </div>
        </div>
      )}
      <div onDrop={handleDrop} onDragOver={e=>e.preventDefault()} style={{border:`1px solid ${C.smokeLight}`,padding:"18px 28px",marginBottom:"28px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <input ref={fileInputRef} type="file" accept=".xlsx,.docx" style={{display:"none"}} onChange={e=>{const f=e.target.files[0];if(!f)return;if(f.name.endsWith(".xlsx"))handleTaxonomyImport(f);else handleUpload(f);}}/>
        {uploading?(<div><div style={{fontSize:"13px",color:C.amber,letterSpacing:"0.1em"}}>{uploadStep}</div><div style={{width:"220px",height:"2px",background:C.smokeLight,marginTop:"10px",overflow:"hidden"}}><div style={{height:"100%",background:C.amber,width:"40%",animation:"pulse 1.4s ease-in-out infinite"}}/></div></div>):(<><div><div style={{display:"flex",gap:"10px",alignItems:"center",marginBottom:"8px"}}><button onClick={()=>fileInputRef.current?.click()} style={{background:C.amber,border:"none",color:C.charcoal,padding:"9px 20px",fontSize:"12px",letterSpacing:"0.12em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif",fontWeight:700}}>IMPORT TAXONOMY .XLSX</button><button onClick={()=>fileInputRef.current?.click()} style={{background:"transparent",border:`1px solid ${C.smokeLight}`,color:C.textMuted,padding:"9px 14px",fontSize:"11px",letterSpacing:"0.1em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>or .docx fallback</button></div><div style={{fontSize:"10px",color:C.textMuted,fontFamily:"IM Fell English, serif",fontStyle:"italic"}}>Run <span style={{color:C.textDim,fontStyle:"normal"}}>python build_afl_index.py manuscript.docx</span> first for stable IDs</div></div><div style={{fontSize:"11px",color:C.textMuted}}>{uploadTime?`Last run: ${uploadTime}s`:""}</div></>)}
      </div>
      {!activeSession&&!uploading&&(<div style={{textAlign:"center",padding:"80px 32px",color:C.textMuted}}><div style={{fontSize:"40px",marginBottom:"16px",opacity:0.2}}>◉</div><div style={{fontSize:"14px",letterSpacing:"0.2em"}}>NO DRAFT LOADED</div><div style={{fontSize:"12px",marginTop:"8px",fontFamily:"IM Fell English, serif",fontStyle:"italic"}}>Import taxonomy index or drop a DOCX</div></div>)}
      {activeSession&&activeTab==="dashboard"&&(<div>
        <div style={{fontSize:"10px",letterSpacing:"0.2em",color:C.textMuted,marginBottom:"16px"}}>DRAFT HEALTH — {activeSession.draftName} — {activeSession.wordCount?.toLocaleString()} words{activeSession.source==="taxonomy"&&<span style={{marginLeft:"12px",color:C.pass,fontSize:"9px"}}>● TAXONOMY INDEX</span>}</div>
        <HealthBar score={activeSession.health} prevScore={ps?.health}/>
        <div style={{display:"flex",gap:"40px",marginBottom:"28px",paddingBottom:"20px",borderBottom:`1px solid ${C.smokeLight}`,flexWrap:"wrap"}}>
          <div><div style={{fontSize:"10px",letterSpacing:"0.15em",color:C.textMuted,marginBottom:"4px"}}>ACTIONABLE FINDINGS</div><div style={{fontSize:"32px",fontWeight:700}}>{ta.toLocaleString()}</div><div style={{fontSize:"10px",color:C.textMuted,marginTop:"2px"}}>{tr.toLocaleString()} raw</div></div>
          {ps&&pt>ta&&<div><div style={{fontSize:"10px",letterSpacing:"0.15em",color:C.textMuted,marginBottom:"4px"}}>RESOLVED</div><div style={{fontSize:"32px",fontWeight:700,color:C.pass}}>{(pt-ta).toLocaleString()}</div></div>}
          {ps&&ta>pt&&<div><div style={{fontSize:"10px",letterSpacing:"0.15em",color:C.regress,marginBottom:"4px"}}>⚠ REGRESSION</div><div style={{fontSize:"32px",fontWeight:700,color:C.regress}}>+{(ta-pt).toLocaleString()}</div></div>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(240px,1fr))",gap:"12px",marginBottom:"24px"}}>
          {allSignals.filter(s=>s.enabled!==false).map(sig=>(<SignalCard key={sig.id} signal={sig} actionable={activeCounts[sig.id]||0} raw={activeSession?.rawCounts?.[sig.id]||0} prevActionable={ps?.counts[sig.id]} onClick={()=>{setFilterSignal(sig.id);setFilterDisp("all");setActiveTab("findings");}}/>))}
        </div>
        <div style={{display:"flex",gap:"16px",alignItems:"center",flexWrap:"wrap"}}>
          <button onClick={()=>exportToXlsx(activeFindings,activeSession.draftName)} style={{background:"transparent",border:"1px solid "+C.amber,color:C.amber,padding:"10px 24px",fontSize:"12px",letterSpacing:"0.15em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>EXPORT ALL FINDINGS</button>
          {dismissals.size>0&&<div style={{fontSize:"11px",color:C.textMuted,fontFamily:"Barlow Condensed, sans-serif"}}>{dismissals.size} author-skipped &middot; <button onClick={()=>{const e=new Set();setDismissals(e);saveDismissals(e);}} style={{background:"none",border:"none",color:C.textMuted,cursor:"pointer",fontSize:"11px",fontFamily:"Barlow Condensed, sans-serif",textDecoration:"underline"}}>restore all</button></div>}
        </div>
      </div>)}
      {activeSession&&activeTab==="findings"&&(<div>
        <div style={{display:"flex",gap:"8px",marginBottom:"20px",flexWrap:"wrap",alignItems:"center"}}>
          <select value={filterSignal} onChange={e=>setFilterSignal(e.target.value)} style={{background:C.smokeDark,border:`1px solid ${C.smokeLight}`,color:C.parchment,padding:"8px 12px",fontSize:"12px",fontFamily:"Barlow Condensed, sans-serif",cursor:"pointer"}}><option value="all">ALL SIGNALS</option>{allSignals.map(s=><option key={s.id} value={s.id}>{s.short} ({activeSession.counts[s.id]||0})</option>)}</select>
          <select value={filterDisp} onChange={e=>setFilterDisp(e.target.value)} style={{background:C.smokeDark,border:`1px solid ${C.smokeLight}`,color:C.parchment,padding:"8px 12px",fontSize:"12px",fontFamily:"Barlow Condensed, sans-serif",cursor:"pointer"}}><option value="actionable">ACTIONABLE</option><option value="review_only">REVIEW ONLY</option><option value="all">ALL</option></select>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search…" style={{background:C.smokeDark,border:`1px solid ${C.smokeLight}`,color:C.parchment,padding:"8px 14px",fontSize:"12px",fontFamily:"Barlow Condensed, sans-serif",flex:1,minWidth:"140px",outline:"none"}}/>
          <div style={{display:"flex",gap:"2px"}}>{[["line","LINE #"],["chapter","CHAPTER"],["signal","SIGNAL"]].map(([v,l])=>(<button key={v} onClick={()=>setSortBy(v)} style={{background:sortBy===v?C.smokeLight:"transparent",border:`1px solid ${sortBy===v?C.amber:C.smokeLight}`,color:sortBy===v?C.parchment:C.textMuted,padding:"6px 10px",fontSize:"10px",letterSpacing:"0.1em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>{l}</button>))}</div>
          <span style={{fontSize:"11px",color:C.textMuted}}>{ff.length.toLocaleString()} shown</span>
          <button onClick={()=>exportToXlsx(ff,`${activeSession.draftName}_${filterSignal}_${filterDisp}`)} style={{background:"transparent",border:`1px solid ${C.smokeLight}`,color:C.textDim,padding:"8px 16px",fontSize:"11px",letterSpacing:"0.12em",cursor:"pointer",fontFamily:"Barlow Condensed, sans-serif"}}>EXPORT VIEW</button>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:"2px"}}>
          {ff.slice(0,500).map((f,i)=>(<div key={i} style={{display:"grid",gridTemplateColumns:"52px 130px 90px 80px 1fr",background:i%2===0?C.smokeDark:"transparent",padding:"8px 12px",alignItems:"start"}}>
            <div style={{color:C.textMuted,fontSize:"10px",paddingTop:"2px",fontVariantNumeric:"tabular-nums"}}>{f.lineNum?f.lineNum.toLocaleString():"—"}</div>
            <div style={{color:C.textMuted,fontSize:"10px",paddingTop:"2px"}}>{f.chapter.slice(0,22)}</div>
            <div style={{background:C.smokeLight,color:C.amber,fontSize:"9px",padding:"2px 5px",alignSelf:"start",whiteSpace:"nowrap",overflow:"hidden"}}>{f.issue.slice(0,16)}</div>
            <div style={{fontSize:"9px",color:dc[f.disposition]||C.textMuted,paddingTop:"2px"}}>{(f.disposition||"").toUpperCase().replace("_"," ")}{f.confidence?` ·${f.confidence}%`:""}</div>
            <div style={{color:C.dimParch,fontFamily:"IM Fell English, serif",fontSize:"13px",lineHeight:1.5}}>{f.sentence.slice(0,300)}{f.reason&&<span style={{display:"block",fontSize:"10px",color:C.textMuted,fontFamily:"Barlow Condensed, sans-serif",marginTop:"2px",fontStyle:"normal"}}>{f.reason}</span>}</div>
          </div>))}
          {ff.length>500&&<div style={{textAlign:"center",padding:"16px",color:C.textMuted,fontSize:"11px"}}>Showing 500 of {ff.length.toLocaleString()} — export for full list</div>}
          {ff.length===0&&<div style={{textAlign:"center",padding:"40px",color:C.textMuted,fontSize:"12px",letterSpacing:"0.1em"}}>NO FINDINGS MATCH</div>}
        </div>
      </div>)}
      {activeTab==="review"&&<div style={{padding:"0 32px 28px",maxWidth:"1160px",margin:"0 auto"}}><ReviewQueue findings={ff} draftName={activeSession?.draftName||""} onDismiss={handleDismiss}/></div>}
    </div>
    {showPanel&&<SignalPanel builtinSignals={BUILTIN_SIGNALS} customSignals={customSignals} onToggleCustom={toggleCustom} onDeleteCustom={deleteCustom} onRename={renameCustom} onAddSignal={()=>{setShowAddModal(true);setShowPanel(false);}} onImport={importFindings} onClose={()=>setShowPanel(false)}/>}
    {showAddModal&&<AddSignalModal sentenceIndex={sentenceIndex} onSave={addCustomSignal} onClose={()=>setShowAddModal(false)}/>}
    <style>{`@keyframes pulse{0%,100%{opacity:.3;transform:translateX(-100%)}50%{opacity:1;transform:translateX(250%)}}@keyframes rqpop{0%{transform:scale(0.97);opacity:0.5}100%{transform:scale(1);opacity:1}}input::placeholder{color:#8A8070;}select option{background:#2A2A2A;}::-webkit-scrollbar{width:6px;}::-webkit-scrollbar-track{background:#1C1C1C;}::-webkit-scrollbar-thumb{background:#3A3938;}`}</style>
  </div>);}

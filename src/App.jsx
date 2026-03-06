import { useState, useRef, useEffect, useCallback } from "react";
import * as mammoth from "mammoth";

// ── Baked-in SOPs ─────────────────────────────────────────────────────────
const PRELOADED_SOPS = [];

// ── Mermaid loader ────────────────────────────────────────────────────────
const loadMermaid = () => new Promise((res) => {
  if (window.mermaid) return res(window.mermaid);
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/mermaid/10.6.1/mermaid.min.js";
  s.onload = () => {
    window.mermaid.initialize({ startOnLoad: false, theme: "base", themeVariables: { primaryColor: "#005587", primaryTextColor: "#fff", lineColor: "#005587", secondaryColor: "#8DB3E2" } });
    res(window.mermaid);
  };
  document.head.appendChild(s);
});

// ── docx loader ───────────────────────────────────────────────────────────
const loadDocx = () => new Promise((res) => {
  if (window.docx) return res(window.docx);
  const s = document.createElement("script");
  s.src = "https://unpkg.com/docx@9.5.3/build/index.umd.js";
  s.onload = () => res(window.docx);
  document.head.appendChild(s);
});

// ── API call ──────────────────────────────────────────────────────────────
async function callClaude(system, user) {
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 1000, system, messages: [{ role: "user", content: user }] }),
  });
  const d = await r.json();
  if (d.error) throw new Error(d.error.message);
  return d.content?.map(b => b.text || "").join("") || "";
}

// ── Helpers ───────────────────────────────────────────────────────────────
function chunkText(text, size = 1500) {
  const c = [];
  for (let i = 0; i < text.length; i += size) c.push(text.slice(i, i + size));
  return c;
}

function buildContext(docs, query) {
  const words = query.toLowerCase().split(/\W+/);
  const scored = docs.flatMap(d => d.chunks.map(c => ({ name: d.name, chunk: c, score: words.reduce((s, w) => s + (c.toLowerCase().includes(w) ? 1 : 0), 0) })));
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, 5).map(s => `[${s.name}]\n${s.chunk}`).join("\n\n---\n\n");
}

function today() {
  return new Date().toLocaleDateString("en-US", { month: "2-digit", day: "2-digit", year: "numeric" });
}

function cleanName(name) {
  return name.replace(/\.docx$/i, "").replace(/[_-]/g, " ").replace(/SWP \d+/i, "").replace(/\s+/g, " ").trim();
}

// ── SOP Template Reference (from Recruit_-_Upgrading_Job_Reqs.docx, v2.0) ─
// This is the canonical Denver Health HR SWP template as of 02/08/2026.
// Key specs confirmed from source file:
//   - All header cells: fill #8DB3E2 (light blue), dark text, bold
//   - Data cells: no fill (white), normal weight
//   - No alternating row shading in steps table
//   - Border color: #8DB3E2 single 4pt
//   - Column widths (DXA): Step=434, What=1151, Who=761, How=1679, Why=975  (sum=5000, scaled to full width)
//   - Metadata table: label col ~1800, value col ~7560
//   - Change log: 3 cols, header #8DB3E2, dark text
//   - Screenshots/Notes: single full-width cell, bold "Screenshots / Notes" header inline (no separate header row)
//   - Font: Calibri 10pt (size=20 in half-points)
//   - Page: US Letter, ~1" margins
// Writing style from example SOP:
//   - What: short noun phrase (e.g. "Request upgrade", "Evaluate upgrade criteria")
//   - Who: role title (e.g. "TA Partner", "Hiring Manager")
//   - How: 1-2 sentences with specific system actions (e.g. "Email the TA Partner requesting...")
//   - Why: 1 sentence justification (e.g. "Initiates the evaluation and decision.")
const SOP_TEMPLATE_REFERENCE = `
DENVER HEALTH SWP TEMPLATE (canonical as of 02/08/2026, converted by Brian Caputo)
Source: Recruit_-_Upgrading_Job_Reqs.docx v2.0

METADATA TABLE (Table 0):
  Row labels (col 0, fill #8DB3E2, bold): Department | Job / Role | Process Name | Date Created | Author | Version
  Values (col 1, white, normal): free text

PROCESS STEPS TABLE (Table 1):
  Header row (fill #8DB3E2, bold): Step | What | Who | How | Why
  Data rows (white, no shading): step number | short noun phrase | role title | 1-2 sentence instruction | 1 sentence rationale
  Example step:
    Step 1 | Request upgrade | Hiring Manager | Email the TA Partner requesting the job requisition be upgraded to a higher paid job profile. | Initiates the evaluation and decision.
    Step 2 | Evaluate upgrade criteria | TA Partner | Compare the Job Level on the current job profile vs the requested job profile. Use judgment and apply the guidelines below. If unsure, check with the TA Director. Document the request reason and any change on the job requisition. | Ensures upgrades follow agreed guardrails and remain within appropriate scope.

SCREENSHOTS / NOTES TABLE (Table 2):
  Single cell, full width. Starts with bold "Screenshots / Notes" on first line, then notes/guidelines below.

CHANGE LOG TABLE (Table 3):
  Header (fill #8DB3E2, bold): Effective Date | Change | Who
  Rows: date | description of change | author name

WRITING CONVENTIONS:
  - "What" column: short imperative noun phrases, not full sentences
  - "Who" column: exact role names used at Denver Health TA (TA Partner, TA Coordinator, Hiring Manager, HRIS, Finance, Recruiter, etc.)
  - "How" column: specific, actionable. Reference Workday task names, system names, forms. Use → for navigation paths.
  - "Why" column: one sentence explaining the business/compliance reason
  - Process Name: descriptive, starts with a verb or "How to..." (e.g. "Evaluate and process requests to upgrade an approved job requisition")
  - Overview paragraph: one sentence starting with "This Standard Work Procedure outlines how..."
`;

// ── Generate Denver Health .docx ──────────────────────────────────────────
async function generateDenverHealthDocx(sopData) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, BorderStyle, WidthType, ShadingType, AlignmentType } = await loadDocx();
  // Confirmed border color from template: 8DB3E2
  const B = { style: BorderStyle.SINGLE, size: 4, color: "8DB3E2" };
  const BORDERS = { top: B, bottom: B, left: B, right: B };
  const M = { top: 60, bottom: 60, left: 100, right: 100 };

  // Helper: cell with explicit width
  const wCell = (text, fill, bold, width) => new TableCell({
    borders: BORDERS,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: fill || "FFFFFF", type: ShadingType.CLEAR },
    margins: M,
    children: [new Paragraph({ children: [new TextRun({ text: String(text || ""), font: "Calibri", size: 20, bold: !!bold, color: "000000" })] })],
  });

  // Table 0 — Metadata (label col 1800, value col 7560)
  const metaTable = new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1800, 7560], rows:
    [["Department", sopData.department || ""], ["Job / Role", sopData.jobRole || ""], ["Process Name", sopData.processName || ""], ["Date Created", sopData.dateCreated || today()], ["Author", sopData.author || ""], ["Version", sopData.version || "1.0"]]
    .map(([k, v]) => new TableRow({ children: [wCell(k, "8DB3E2", true, 1800), wCell(v, "FFFFFF", false, 7560)] }))
  });

  // Table 1 — Process Steps
  // Widths scaled from source (434/1151/761/1679/975 → proportionally fill 9360 total)
  // Source sum = 5000, scale factor = 9360/5000 = 1.872
  const CW = [812, 2154, 1424, 3143, 1827]; // scaled, sum=9360
  const stepsTable = new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: CW, rows: [
    // Header row — #8DB3E2, bold, dark text (matches template exactly)
    new TableRow({ children: ["Step","What","Who","How","Why"].map((h, i) => new TableCell({
      borders: BORDERS, width: { size: CW[i], type: WidthType.DXA },
      shading: { fill: "8DB3E2", type: ShadingType.CLEAR }, margins: M,
      children: [new Paragraph({ alignment: i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT, children: [new TextRun({ text: h, font: "Calibri", size: 20, bold: true, color: "000000" })] })],
    })) }),
    // Data rows — white, no alternating shading (matches template)
    ...(sopData.steps || []).map((s, i) => new TableRow({ children: [
      new TableCell({ borders: BORDERS, width: { size: CW[0], type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: M, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(i+1), font: "Calibri", size: 20 })] })] }),
      ...[s.what, s.who, s.how, s.why].map((v, ci) => new TableCell({ borders: BORDERS, width: { size: CW[ci+1], type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: M, children: [new Paragraph({ children: [new TextRun({ text: String(v||""), font: "Calibri", size: 20 })] })] }))
    ] }))
  ]});

  // Table 2 — Screenshots / Notes (single cell, bold label inline, then content)
  const screenshotsTable = new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
    new TableRow({ children: [new TableCell({ borders: BORDERS, width: { size: 9360, type: WidthType.DXA }, margins: { top: 100, bottom: 300, left: 120, right: 120 }, children: [
      new Paragraph({ children: [new TextRun({ text: "Screenshots / Notes", font: "Calibri", size: 20, bold: true })] }),
      new Paragraph({ children: [new TextRun({ text: "", font: "Calibri", size: 20 })], spacing: { after: 600 } }),
    ] })] }),
  ]});

  // Table 3 — Change Log (header #8DB3E2 dark text, rows white)
  const clWidths = [1440, 6480, 1440];
  const changeLogTable = new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: clWidths, rows: [
    new TableRow({ children: ["Effective Date","Change","Who"].map((h, i) => new TableCell({ borders: BORDERS, width: { size: clWidths[i], type: WidthType.DXA }, shading: { fill: "8DB3E2", type: ShadingType.CLEAR }, margins: M, children: [new Paragraph({ children: [new TextRun({ text: h, font: "Calibri", size: 20, bold: true, color: "000000" })] })] })) }),
    new TableRow({ children: [
      new TableCell({ borders: BORDERS, width: { size: clWidths[0], type: WidthType.DXA }, margins: M, children: [new Paragraph({ children: [new TextRun({ text: today(), font: "Calibri", size: 20 })] })] }),
      new TableCell({ borders: BORDERS, width: { size: clWidths[1], type: WidthType.DXA }, margins: M, children: [new Paragraph({ children: [new TextRun({ text: "Created using SOP Assistant", font: "Calibri", size: 20 })] })] }),
      new TableCell({ borders: BORDERS, width: { size: clWidths[2], type: WidthType.DXA }, margins: M, children: [new Paragraph({ children: [new TextRun({ text: sopData.author || "", font: "Calibri", size: 20 })] })] }),
    ] }),
    // Blank row for future entries
    new TableRow({ children: clWidths.map(w => new TableCell({ borders: BORDERS, width: { size: w, type: WidthType.DXA }, margins: M, children: [new Paragraph({ children: [new TextRun({ text: "", font: "Calibri", size: 20 })] })] })) }),
  ]});

  const p = (text, bold) => new Paragraph({ spacing: { before: 140, after: 60 }, children: [new TextRun({ text, font: "Calibri", size: bold ? 22 : 20, bold: !!bold })] });

  const doc = new Document({ sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } }, children: [
    metaTable,
    p(sopData.overview || ""),
    p("Process Steps", true),
    stepsTable,
    p(""),
    screenshotsTable,
    p("Change Log", true),
    changeLogTable,
  ] }] });
  return Packer.toBlob(doc);
}

// ── Icons ─────────────────────────────────────────────────────────────────
const Ic = ({ d, sz = 18 }) => <svg width={sz} height={sz} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d={d} /></svg>;
const Icons = {
  upload: <Ic d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12" />,
  chat: <Ic d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z" />,
  fill: <Ic d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z" />,
  map: <Ic d="M3 3h7v7H3zM14 3h7v7h-7zM14 14h7v7h-7zM3 14h7v7H3z" />,
  doc: <Ic d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8zM14 2v6h6M16 13H8M16 17H8M10 9H8" />,
  trash: <Ic d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2" sz={14} />,
  send: <Ic d="M22 2L11 13M22 2l-7 20-4-9-9-4 20-7z" sz={16} />,
  dl: <Ic d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3" sz={16} />,
  plus: <Ic d="M12 5v14M5 12h14" sz={16} />,
};
const Spin = () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" style={{ animation: "spin .8s linear infinite", display: "inline-block", verticalAlign: "middle" }}><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83" /></svg>;

// ── App ───────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("upload");
  const [docs, setDocs] = useState([]);
  const [uploading, setUploading] = useState(false);

  const [msgs, setMsgs] = useState([{ role: "assistant", content: `Hi! Upload your Denver Health SOPs in the SOP Library tab, then ask me anything about your processes, generate new SOPs, or build process maps.` }]);
  const [chatIn, setChatIn] = useState("");
  const [chatBusy, setChatBusy] = useState(false);
  const chatEnd = useRef(null);

  const [form, setForm] = useState({ department: "", jobRole: "", processName: "", author: "", notes: "" });
  const [sop, setSop] = useState(null);
  const [fillBusy, setFillBusy] = useState(false);

  const [mapTopic, setMapTopic] = useState("");
  const [mapCode, setMapCode] = useState("");
  const [mapSvg, setMapSvg] = useState("");
  const [mapBusy, setMapBusy] = useState(false);

  useEffect(() => { chatEnd.current?.scrollIntoView({ behavior: "smooth" }); }, [msgs]);

  // Upload
  const handleFiles = useCallback(async (files) => {
    setUploading(true);
    const newDocs = [];
    for (const f of files) {
      try {
        const ab = await f.arrayBuffer();
        const { value } = await mammoth.extractRawText({ arrayBuffer: ab });
        newDocs.push({ name: f.name, text: value, chunks: chunkText(value), size: f.size, added: true });
      } catch(e) { console.error(e); }
    }
    setDocs(p => [...p, ...newDocs]);
    setUploading(false);
  }, []);

  const onDrop = e => { e.preventDefault(); const fs = [...e.dataTransfer.files].filter(f => f.name.endsWith(".docx")); if (fs.length) handleFiles(fs); };

  // Chat
  const sendChat = async () => {
    if (!chatIn.trim() || chatBusy) return;
    const msg = chatIn.trim(); setChatIn(""); setChatBusy(true);
    setMsgs(p => [...p, { role: "user", content: msg }]);
    try {
      const ctx = buildContext(docs, msg);
      const reply = await callClaude(
        `You are a Denver Health SOP assistant. Answer ONLY based on the provided SOP documents. Cite which SOP your answer comes from. Be concise and precise.\n\nSOPs:\n${ctx}`,
        msg
      );
      setMsgs(p => [...p, { role: "assistant", content: reply }]);
    } catch(e) { setMsgs(p => [...p, { role: "assistant", content: `⚠️ ${e.message}` }]); }
    setChatBusy(false);
  };

  // Generate SOP
  const runGenerate = async () => {
    if (!form.processName.trim() || fillBusy) return;
    setFillBusy(true); setSop(null);
    try {
      const ctx = docs.map(d => `[${d.name}]\n${d.text.slice(0, 2000)}`).join("\n\n---\n\n").slice(0, 12000);
      const raw = await callClaude(
        `You are an expert SOP writer for Denver Health Talent Acquisition. You write Standard Work Procedures in the exact Denver Health SWP format.\n\nCANONICAL TEMPLATE SPECS:\n${SOP_TEMPLATE_REFERENCE}\n\nEXISTING SOPs FOR TERMINOLOGY REFERENCE:\n${ctx}`,
        `Generate a complete Denver Health SWP for: "${form.processName}"
Department: ${form.department || "Talent Acquisition"}
Job/Role: ${form.jobRole || "TA Coordinator"}
Author: ${form.author || ""}
Notes: ${form.notes || "none"}

Return ONLY a JSON object (no markdown, no explanation) with:
{
  "department": "...",
  "jobRole": "...", 
  "processName": "...",
  "author": "...",
  "version": "1.0",
  "overview": "one sentence overview",
  "steps": [{"what":"...","who":"...","how":"...","why":"..."}]
}
Generate 6-14 realistic, detailed steps.`
      );
      const parsed = JSON.parse(raw.replace(/```json|```/g, "").trim());
      parsed.dateCreated = today();
      setSop(parsed);
    } catch(e) { alert("Error: " + e.message); }
    setFillBusy(false);
  };

  const downloadDocx = async () => {
    if (!sop) return;
    const blob = await generateDenverHealthDocx(sop);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${sop.processName.replace(/[^a-z0-9]/gi,"_")}_SWP.docx`;
    a.click();
  };

  // ── Canva process map reference (pulled 3/6/2026) ────────────────────────
  const CANVA_PROCESS_MAP = `
Denver Health Talent Acquisition - Candidate Lifecycle Map (Canva reference)

MAIN STAGES (left to right): Recruit → Screen → Interview → Select → Offer → Onboard → Career Advancement

RECRUIT STAGE:
- Requisition Intake
- Requirements Validation
- Tasks pertaining to building talent pipeline
- Resume Review (ATS)
- Resume Review (Manual)
- Credential Pre-Check
- Create Short List
- ATS Database Management
- ATS Backflow (screen recurring) for non-selected candidates
- (Maintain) Review Candidate Sourcing Strategy

SCREEN STAGE:
- Phone Screen
- Pre-Screen Decision (Y/N)
- Secondary Screening Window
- Skill Survey (Reference Check / HireRight)

INTERVIEW STAGE:
- Schedule Interview
- Send Candidate Interview Instructions
- Prep Interview Guide and Scorecard
- Conduct Interview
- Collect Feedback
- Interview Debrief Decision (Y/N)

SELECT STAGE:
- Final Candidate Selection Confirmation
- Candidate Selected
- ATS Workflow

OFFER STAGE:
- Compensation Alignment / Confirm Negotiation Parameters
- Complete Internal Approvals (HR/Finance/etc.)
- Offer Review (licensure/finance)
- Start date alignment (pay period)
- Extend Verbal Offer
- Generate Written Offer
- Track accept/decline and backup candidate plan
- Candidate Accepts → proceed
- Candidate Declines → Rescind Offer or loop back

ONBOARD STAGE (split into HR Onboard and Manager Onboard):
HR Onboard:
  - Start pre-employment checklist
  - Update ATS status
  - Verify licenses
  - Drug Screen
  - Background Check / HireRight
  - New hire paperwork (tax / direct deposit)
  - Schedule orientation
  - i-9 / E-verify
  - Badging (Badge Process / UserDash)
  - COSH Scheduling / COSH Clearance
  - Foreign Education Verification / Education Verification
  - Primary Source Verification
  - Onboarding Document Uploads
  Sub-types: Clinical vs Non-Clinical paths
  HR Onboarding handoff to Manager Onboarding

Manager Onboard:
  - Day 1 handoff to manager (orientation)
  - Employee complete with full onboard

CAREER ADVANCEMENT STAGE:
- Career development planning
- Internal mobility
- Performance management
- Promotions and career ladders
- Learning and training programs
- Compensation support
- Employee experience and retention

ROLES INVOLVED: Recruiter, TA Coordinator, HR, Manager, Candidate/Employee
`;

  // Process map
  const genMap = async () => {
    if (!mapTopic.trim() || mapBusy) return;
    setMapBusy(true); setMapCode(""); setMapSvg("");
    try {
      const ctx = buildContext(docs, mapTopic);
      const code = await callClaude(
        `You are a process documentation expert for Denver Health Talent Acquisition. Generate a Mermaid flowchart. Return ONLY valid Mermaid starting with "flowchart TD". No fences, no explanation, no extra text.

Use the Denver Health Candidate Lifecycle Map below as your primary reference for stage names, terminology, and process flow. Match the language and structure used by the TA team exactly.

DENVER HEALTH CANDIDATE LIFECYCLE REFERENCE:\n${CANVA_PROCESS_MAP}\n\nSOP DOCUMENTS:\n${ctx}`,
        `Create a detailed process flowchart for: ${mapTopic}`
      );
      const clean = code.replace(/```mermaid|```/g,"").trim();
      setMapCode(clean);
      try {
        const m = await loadMermaid();
        const { svg } = await m.render("map-"+Date.now(), clean);
        setMapSvg(svg);
      } catch(_) {}
    } catch(e) { setMapCode("Error: " + e.message); }
    setMapBusy(false);
  };

  const sf = k => e => setForm(p => ({ ...p, [k]: e.target.value }));

  const preloaded = docs.filter(d => !d.added);
  const uploaded = docs.filter(d => d.added);

  const S = {
    input: { width:"100%", padding:"9px 12px", border:"1px solid #dde6f0", borderRadius:7, fontFamily:"'Source Sans 3',sans-serif", fontSize:13, outline:"none", background:"white" },
    label: { fontFamily:"'Source Sans 3',sans-serif", fontWeight:600, fontSize:11, color:"#005587", display:"block", marginBottom:5, textTransform:"uppercase", letterSpacing:"0.07em" },
    card: { background:"white", border:"1px solid #dde6f0", borderRadius:10, padding:20 },
  };

  const TABS = [
    { id:"upload", label:"SOP Library", icon: Icons.upload },
    { id:"chat",   label:"Ask Assistant", icon: Icons.chat },
    { id:"fill",   label:"Generate SOP", icon: Icons.fill },
    { id:"map",    label:"Process Maps", icon: Icons.map },
  ];

  return (
    <div style={{ fontFamily:"'Source Sans 3',Georgia,sans-serif", minHeight:"100vh", background:"#f0f4f8" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=Source+Sans+3:wght@300;400;500;600&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes fu{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
        .fu{animation:fu .25s ease forwards}
        ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-thumb{background:#b0c4d8;border-radius:4px}
        input:focus,textarea:focus{border-color:#005587!important;box-shadow:0 0 0 3px rgba(0,85,135,.1)}
        .tb:hover{background:#f0f7fb!important;color:#005587!important}
        .dr:hover{background:#f5f9fb!important}
        .msg{white-space:pre-wrap;line-height:1.65}
        .st td,.st th{padding:5px 9px;font-family:Calibri,sans-serif;font-size:12px;border:1px solid #4F81BD}
        .st th{background:#4F81BD;color:white;font-weight:600}
        .st tr:nth-child(even) td{background:#EBF1F8}
      `}</style>

      {/* Header */}
      <header style={{ background:"white", borderBottom:"3px solid #E8521A", padding:"0 28px", display:"flex", alignItems:"center", gap:16, height:68, boxShadow:"0 2px 12px rgba(0,0,0,.1)" }}>
        {/* Denver Health logo mark — SVG recreation of the heart/mountain icon */}
        <svg width="42" height="42" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M50 85 C50 85 10 55 10 30 C10 18 19 8 30 8 C38 8 45 13 50 20 C55 13 62 8 70 8 C81 8 90 18 90 30 C90 55 50 85 50 85Z" fill="url(#hgrad)"/>
          <defs>
            <linearGradient id="hgrad" x1="10" y1="8" x2="90" y2="85" gradientUnits="userSpaceOnUse">
              <stop offset="0%" stopColor="#1AAFCE"/>
              <stop offset="100%" stopColor="#005587"/>
            </linearGradient>
          </defs>
          {/* Mountain peak */}
          <path d="M35 52 L50 32 L65 52Z" fill="white" opacity="0.25"/>
          <path d="M44 40 L50 32 L56 40 L53 37 L50 35 L47 37Z" fill="#E8521A"/>
          {/* Star/asterisk on peak */}
          <circle cx="50" cy="31" r="4" fill="#E8521A"/>
        </svg>

        {/* Wordmark */}
        <div style={{ borderRight:"1px solid #dde6f0", paddingRight:16, marginRight:4 }}>
          <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:900, fontSize:18, color:"#005587", letterSpacing:".04em", lineHeight:1.1, textTransform:"uppercase" }}>DENVER</div>
          <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:900, fontSize:18, color:"#005587", letterSpacing:".04em", lineHeight:1.1, textTransform:"uppercase" }}>HEALTH</div>
        </div>

        {/* Tool identity */}
        <div>
          <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:700, fontSize:15, color:"#1a1a2e", letterSpacing:".01em" }}>SOP Assistant</div>
          <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:400, fontSize:11, color:"#E8521A", letterSpacing:".1em", textTransform:"uppercase" }}>Talent Acquisition · Internal Tool</div>
        </div>

        <div style={{ marginLeft:"auto", display:"flex", gap:8, alignItems:"center" }}>
          <div style={{ background:"#EBF6FB", border:"1px solid #b0d8ea", borderRadius:16, padding:"4px 14px", fontFamily:"'Source Sans 3',sans-serif", fontSize:12, color:"#005587", fontWeight:600 }}>
            {docs.length} SOP{docs.length !== 1 ? 's' : ''} loaded
          </div>
        </div>
      </header>

      {/* Tabs */}
      <div style={{ background:"white", borderBottom:"1px solid #dde6f0", padding:"0 28px", display:"flex", boxShadow:"0 1px 4px rgba(0,0,0,.06)" }}>
        {TABS.map(t => (
          <button key={t.id} className="tb" onClick={() => setTab(t.id)} style={{ display:"flex", alignItems:"center", gap:7, padding:"13px 18px", border:"none", background:"none", cursor:"pointer", fontFamily:"'Source Sans 3',sans-serif", fontSize:13, fontWeight: tab===t.id?600:400, color: tab===t.id?"#005587":"#666", borderBottom: tab===t.id?"2.5px solid #E8521A":"2.5px solid transparent", transition:"all .15s" }}>
            {t.icon} {t.label}
          </button>
        ))}
      </div>

      <main style={{ maxWidth:980, margin:"0 auto", padding:"28px 20px" }}>

        {/* ══ LIBRARY ══ */}
        {tab === "upload" && (
          <div className="fu">
            <h2 style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:800, fontSize:22, color:"#005587", marginBottom:4, textTransform:"uppercase", letterSpacing:".02em" }}>SOP Library</h2>
            <p style={{ fontFamily:"'Source Sans 3',sans-serif", color:"#777", fontSize:13, marginBottom:20 }}>Upload your .docx SOP files. The assistant will use them to answer questions, generate new SOPs, and build process maps.</p>

            <div onDrop={onDrop} onDragOver={e=>e.preventDefault()} onClick={() => document.getElementById("fi").click()}
              style={{ border:"2px dashed #1AAFCE", borderRadius:10, padding:"28px 24px", textAlign:"center", background:"white", marginBottom:20, cursor:"pointer" }}>
              <input id="fi" type="file" accept=".docx" multiple style={{ display:"none" }} onChange={e => handleFiles([...e.target.files])} />
              <div style={{ fontSize:24, marginBottom:6 }}>📄</div>
              <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:600, color:"#005587", marginBottom:2 }}>{uploading ? "Processing…" : "Drop .docx files here or click to upload more SOPs"}</div>
              <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontSize:11, color:"#aaa" }}>Supports multiple files at once</div>
            </div>

            {docs.length > 0 ? (
              <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
                {docs.map((d,i) => (
                  <div key={i} className="dr fu" style={{ background:"white", border:"1px solid #e0eaf4", borderRadius:7, padding:"9px 14px", display:"flex", alignItems:"center", gap:10 }}>
                    <div style={{ color:"#005587" }}>{Icons.doc}</div>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:600, fontSize:13, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{d.name}</div>
                      <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontSize:11, color:"#aaa" }}>{d.chunks.length} chunks · {(d.text.length/1000).toFixed(1)}k chars</div>
                    </div>
                    <button onClick={() => setDocs(p => p.filter((_,j) => j !== i))} style={{ background:"none", border:"none", cursor:"pointer", color:"#cc4444", opacity:.7, padding:4 }}>{Icons.trash}</button>
                  </div>
                ))}
              </div>
            ) : (
              <div style={{ textAlign:"center", color:"#bbb", fontFamily:"'Source Sans 3',sans-serif", fontSize:13, padding:"16px 0" }}>No SOPs uploaded yet.</div>
            )}
          </div>
        )}

        {/* ══ CHAT ══ */}
        {tab === "chat" && (
          <div className="fu" style={{ display:"flex", flexDirection:"column", height:"calc(100vh - 200px)", minHeight:500 }}>
            <h2 style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:800, fontSize:22, color:"#005587", marginBottom:4, textTransform:"uppercase", letterSpacing:".02em" }}>Ask the SOP Assistant</h2>
            <p style={{ fontFamily:"'Source Sans 3',sans-serif", color:"#777", fontSize:13, marginBottom:14 }}>Ask questions about any process. Answers are sourced from your loaded SOPs.</p>
            <div style={{ flex:1, overflowY:"auto", ...S.card, display:"flex", flexDirection:"column", gap:14, marginBottom:14 }}>
              {msgs.map((m,i) => (
                <div key={i} className="fu" style={{ display:"flex", gap:8, flexDirection: m.role==="user"?"row-reverse":"row" }}>
                  <div style={{ width:28, height:28, borderRadius:"50%", background: m.role==="user"?"#005587":"#EBF1F8", flexShrink:0, display:"flex", alignItems:"center", justifyContent:"center", fontSize:13 }}>{m.role==="user"?"👤":"⚕"}</div>
                  <div className="msg" style={{ maxWidth:"76%", background: m.role==="user"?"#005587":"#f4f7fb", color: m.role==="user"?"white":"#1a1a2e", borderRadius: m.role==="user"?"14px 3px 14px 14px":"3px 14px 14px 14px", padding:"9px 13px", fontFamily:"'Source Sans 3',sans-serif", fontSize:13 }}>{m.content}</div>
                </div>
              ))}
              {chatBusy && <div style={{ display:"flex", gap:8 }}><div style={{ width:28, height:28, borderRadius:"50%", background:"#EBF1F8", display:"flex", alignItems:"center", justifyContent:"center" }}>⚕</div><div style={{ background:"#f4f7fb", borderRadius:"3px 14px 14px 14px", padding:"9px 13px", color:"#999", fontFamily:"'Source Sans 3',sans-serif", fontSize:13 }}><Spin /> Searching SOPs…</div></div>}
              <div ref={chatEnd} />
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <input value={chatIn} onChange={e=>setChatIn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&sendChat()} placeholder="Ask about a process, policy, or step…" style={{ ...S.input, flex:1 }} />
              <button onClick={sendChat} disabled={chatBusy||!chatIn.trim()} style={{ background:"#005587", color:"white", border:"none", borderRadius:8, padding:"9px 18px", cursor:"pointer", display:"flex", alignItems:"center", gap:6, fontFamily:"'Source Sans 3',sans-serif", fontWeight:600, fontSize:13, opacity: chatBusy||!chatIn.trim()?0.5:1 }}>
                {Icons.send} Send
              </button>
            </div>
          </div>
        )}

        {/* ══ GENERATE SOP ══ */}
        {tab === "fill" && (
          <div className="fu">
            <h2 style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:800, fontSize:22, color:"#005587", marginBottom:4, textTransform:"uppercase", letterSpacing:".02em" }}>Generate SOP</h2>
            <p style={{ fontFamily:"'Source Sans 3',sans-serif", color:"#777", fontSize:13, marginBottom:20 }}>Describe the process and the assistant will generate a complete SOP, then export it as a formatted Denver Health .docx file.</p>
            <div style={{ display:"grid", gridTemplateColumns:"320px 1fr", gap:20 }}>
              <div style={S.card}>
                <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:700, fontSize:11, color:"#005587", marginBottom:16, textTransform:"uppercase", letterSpacing:".08em" }}>SOP Details</div>
                {[["processName","Process Name *","e.g. How to close a job requisition"],["department","Department","Talent Acquisition-Recruiting"],["jobRole","Job / Role","TA Coordinator"],["author","Author","Your name"]].map(([k,l,ph]) => (
                  <div key={k} style={{ marginBottom:13 }}>
                    <label style={S.label}>{l}</label>
                    <input value={form[k]} onChange={sf(k)} placeholder={ph} style={S.input} />
                  </div>
                ))}
                <div style={{ marginBottom:13 }}>
                  <label style={S.label}>Additional Notes</label>
                  <textarea value={form.notes} onChange={sf("notes")} placeholder="Specific steps, edge cases, or systems to mention…" style={{ ...S.input, height:80, resize:"vertical" }} />
                </div>
                <button onClick={runGenerate} disabled={fillBusy||!form.processName.trim()} style={{ width:"100%", background:"#005587", color:"white", border:"none", borderRadius:8, padding:11, cursor:"pointer", fontFamily:"'Source Sans 3',sans-serif", fontWeight:700, fontSize:13, display:"flex", alignItems:"center", justifyContent:"center", gap:7, opacity: fillBusy||!form.processName.trim()?0.5:1 }}>
                  {fillBusy ? <><Spin /> Generating…</> : <>{Icons.fill} Generate SOP</>}
                </button>
              </div>

              <div style={{ ...S.card, overflowY:"auto", maxHeight:"72vh" }}>
                {!sop && !fillBusy && <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", height:"100%", minHeight:300, color:"#bbb", fontFamily:"'Source Sans 3',sans-serif", fontSize:14 }}><div style={{ fontSize:36, marginBottom:10 }}>📋</div>Preview appears here</div>}
                {fillBusy && <div style={{ display:"flex", alignItems:"center", justifyContent:"center", height:300, color:"#005587", gap:10, fontFamily:"'Source Sans 3',sans-serif", fontSize:14 }}><Spin /> Building your SOP…</div>}
                {sop && !fillBusy && (
                  <div className="fu">
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
                      <div style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:700, fontSize:11, color:"#005587", textTransform:"uppercase", letterSpacing:".08em" }}>Preview</div>
                      <button onClick={downloadDocx} style={{ background:"#E8521A", color:"white", border:"none", borderRadius:7, padding:"7px 14px", cursor:"pointer", fontFamily:"'Source Sans 3',sans-serif", fontWeight:600, fontSize:12, display:"flex", alignItems:"center", gap:5 }}>{Icons.dl} Download .docx</button>
                    </div>
                    <table style={{ width:"100%", borderCollapse:"collapse", marginBottom:12, fontSize:12 }}>
                      {[["Department",sop.department],["Job / Role",sop.jobRole],["Process Name",sop.processName],["Date Created",sop.dateCreated],["Author",sop.author],["Version",sop.version]].map(([k,v])=>(
                        <tr key={k}><td style={{ background:"#8DB3E2", color:"white", padding:"4px 9px", fontWeight:600, border:"1px solid #4F81BD", width:115, fontFamily:"Calibri,sans-serif" }}>{k}</td><td style={{ padding:"4px 9px", border:"1px solid #4F81BD", fontFamily:"Calibri,sans-serif" }}>{v}</td></tr>
                      ))}
                    </table>
                    <div style={{ fontFamily:"Calibri,sans-serif", fontSize:12, color:"#555", marginBottom:10, fontStyle:"italic" }}>{sop.overview}</div>
                    <div style={{ overflowX:"auto", marginBottom:12 }}>
                      <table className="st" style={{ width:"100%", borderCollapse:"collapse" }}>
                        <thead><tr>{["#","What","Who","How","Why"].map(h=><th key={h}>{h}</th>)}</tr></thead>
                        <tbody>{sop.steps.map((s,i)=><tr key={i}><td style={{ textAlign:"center", width:28 }}>{i+1}</td><td>{s.what}</td><td>{s.who}</td><td>{s.how}</td><td>{s.why}</td></tr>)}</tbody>
                      </table>
                    </div>
                    <div style={{ border:"1px solid #4F81BD", borderRadius:4, overflow:"hidden" }}>
                      <div style={{ background:"#005587", color:"white", padding:"5px 9px", fontFamily:"Calibri,sans-serif", fontSize:12, fontWeight:600 }}>Screenshots / Notes</div>
                      <div style={{ padding:"10px 9px", fontFamily:"Calibri,sans-serif", fontSize:12, color:"#aaa", minHeight:50 }}>[ Insert screenshots here ]</div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* ══ PROCESS MAP ══ */}
        {tab === "map" && (
          <div className="fu">
            <h2 style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:800, fontSize:22, color:"#005587", marginBottom:4, textTransform:"uppercase", letterSpacing:".02em" }}>Process Map Generator</h2>
            <p style={{ fontFamily:"'Source Sans 3',sans-serif", color:"#777", fontSize:13, marginBottom:20 }}>Enter a process topic to generate a visual flowchart based on your SOPs.</p>
            <div style={{ display:"flex", gap:8, marginBottom:20 }}>
              <input value={mapTopic} onChange={e=>setMapTopic(e.target.value)} onKeyDown={e=>e.key==="Enter"&&genMap()} placeholder='e.g. "Start a job requisition" or "No-start process"' style={{ ...S.input, flex:1 }} />
              <button onClick={genMap} disabled={mapBusy||!mapTopic.trim()} style={{ background:"#005587", color:"white", border:"none", borderRadius:8, padding:"9px 20px", cursor:"pointer", fontFamily:"'Source Sans 3',sans-serif", fontWeight:600, fontSize:13, display:"flex", alignItems:"center", gap:6, whiteSpace:"nowrap", opacity: mapBusy||!mapTopic.trim()?0.5:1 }}>
                {mapBusy ? <><Spin /> Generating…</> : <>{Icons.map} Generate Map</>}
              </button>
            </div>
            {mapSvg && (
              <div className="fu" style={{ ...S.card, overflow:"auto" }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
                  <span style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:700, fontSize:11, color:"#005587", textTransform:"uppercase", letterSpacing:".08em" }}>Process Flowchart</span>
                  <button onClick={() => navigator.clipboard.writeText(mapCode)} style={{ background:"none", border:"1px solid #4F81BD", color:"#005587", borderRadius:6, padding:"4px 10px", cursor:"pointer", fontFamily:"'Source Sans 3',sans-serif", fontSize:11, fontWeight:600 }}>Copy Mermaid Code</button>
                </div>
                <div dangerouslySetInnerHTML={{ __html: mapSvg }} style={{ display:"flex", justifyContent:"center" }} />
              </div>
            )}
            {mapCode && !mapSvg && (
              <div className="fu" style={S.card}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
                  <span style={{ fontFamily:"'Source Sans 3',sans-serif", fontWeight:700, fontSize:11, color:"#005587", textTransform:"uppercase" }}>Mermaid Code</span>
                  <button onClick={() => navigator.clipboard.writeText(mapCode)} style={{ background:"none", border:"1px solid #4F81BD", color:"#005587", borderRadius:6, padding:"4px 10px", cursor:"pointer", fontFamily:"'Source Sans 3',sans-serif", fontSize:11, fontWeight:600 }}>Copy</button>
                </div>
                <pre style={{ fontFamily:"monospace", fontSize:12, background:"#f4f7fb", padding:14, borderRadius:7, overflowX:"auto", whiteSpace:"pre-wrap" }}>{mapCode}</pre>
                <p style={{ fontFamily:"'Source Sans 3',sans-serif", fontSize:12, color:"#999", marginTop:8 }}>Paste into <a href="https://mermaid.live" target="_blank" rel="noreferrer" style={{ color:"#005587" }}>mermaid.live</a> to view.</p>
              </div>
            )}
            {!mapCode && !mapBusy && <div style={{ textAlign:"center", padding:"60px 0", color:"#bbb", fontFamily:"'Source Sans 3',sans-serif", fontSize:14 }}><div style={{ fontSize:36, marginBottom:10 }}>🗺️</div>Enter a topic to generate a flowchart</div>}
          </div>
        )}
      </main>
    </div>
  );
}

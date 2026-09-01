const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { execSync } = require('child_process');

const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

function renderHtmlToPdf(htmlContent, outputPath) {
  const tempHtmlPath = path.resolve(`temp_render_${path.basename(outputPath, '.pdf')}.html`);
  fs.writeFileSync(tempHtmlPath, htmlContent, 'utf8');
  const absOutPath = path.resolve(outputPath);
  execSync(`"${edgePath}" --headless --disable-gpu --no-pdf-header-footer --print-to-pdf="${absOutPath}" "file:///${tempHtmlPath.replace(/\\/g, '/')}"`);
  if (fs.existsSync(tempHtmlPath)) fs.unlinkSync(tempHtmlPath);
  console.log(`PDF successfully generated: ${outputPath}`);
}

const commonStyles = `
  @page {
    size: A4;
    margin: 8mm 10mm 8mm 10mm;
  }
  body {
    font-family: 'Segoe UI', Arial, sans-serif;
    color: #0f172a;
    background: #ffffff;
    line-height: 1.3;
    font-size: 8pt;
  }
  .header {
    border-bottom: 3px solid #1e40af;
    padding-bottom: 5px;
    margin-bottom: 10px;
  }
  .header h1 {
    color: #1e3a8a;
    margin: 0 0 2px 0;
    font-size: 14pt;
    letter-spacing: 0.3px;
  }
  .header .badge {
    display: inline-block;
    background: #dbeafe;
    color: #1e40af;
    padding: 2px 6px;
    border-radius: 4px;
    font-size: 7.5pt;
    font-weight: bold;
    text-transform: uppercase;
  }
  .header .meta {
    float: right;
    font-size: 7.5pt;
    color: #64748b;
  }
  .section-title {
    background: #f1f5f9;
    padding: 4px 7px;
    border-left: 4px solid #2563eb;
    margin-top: 10px;
    margin-bottom: 6px;
    font-size: 9pt;
    font-weight: bold;
    color: #0f172a;
    page-break-after: avoid;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 8px;
    font-size: 7pt;
  }
  th, td {
    border: 1px solid #cbd5e1;
    padding: 3px 5px;
    text-align: left;
    vertical-align: top;
  }
  th {
    background: #f8fafc;
    color: #334155;
    font-weight: 600;
  }
  .code-cell {
    font-family: 'Consolas', 'Courier New', monospace;
    font-weight: bold;
    color: #b91c1c;
    background: #fef2f2;
  }
  .ans-cell {
    font-family: 'Consolas', 'Courier New', monospace;
    font-weight: bold;
    color: #047857;
    background: #ecfdf5;
  }
  .note-box {
    background: #fffbeb;
    border: 1px solid #fef3c7;
    border-left: 4px solid #f59e0b;
    padding: 5px 7px;
    border-radius: 4px;
    margin: 6px 0;
    font-size: 7.5pt;
  }
  .page-break {
    page-break-before: always;
  }
  .station-card {
    border: 1px solid #cbd5e1;
    border-radius: 5px;
    padding: 6px 8px;
    margin-bottom: 6px;
    background: #ffffff;
    page-break-inside: avoid;
  }
  .station-card h3 {
    margin: 0 0 3px 0;
    color: #1e40af;
    font-size: 8.2pt;
    display: flex;
    justify-content: space-between;
  }
  .station-badge {
    background: #eff6ff;
    color: #2563eb;
    border: 1px solid #bfdbfe;
    padding: 1px 5px;
    border-radius: 3px;
    font-size: 6.8pt;
    font-weight: bold;
  }
  .station-desc {
    color: #334155;
    font-size: 7.2pt;
    margin-bottom: 3px;
  }
  .station-meta-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 4px;
    background: #f8fafc;
    padding: 4px 6px;
    border-radius: 4px;
    font-size: 6.8pt;
  }
  .meta-item strong {
    color: #0f172a;
  }
  .table-title {
    font-weight: bold;
    color: #1e40af;
    background: #eff6ff;
    padding: 2px 5px;
    border: 1px solid #bfdbfe;
    margin-top: 5px;
    margin-bottom: 2px;
    font-size: 7.5pt;
  }
`;

function renderGenericTable(rows) {
  if (!rows || rows.length === 0) return '<p>No data</p>';
  const keys = Object.keys(rows[0]);
  return `
    <table>
      <thead>
        <tr>
          ${keys.map(k => `<th>${k}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
        ${rows.map(r => `
          <tr>
            ${keys.map(k => {
              const val = String(r[k] !== undefined ? r[k] : '');
              const isCode = val.includes('-') && (val.includes('PASS') || val.includes('START') || val.includes('FINAL') || val.includes('MECH'));
              const isAns = k.toLowerCase().includes('answer') || k.toLowerCase().includes('decoded') || k.toLowerCase().includes('solution') || k.toLowerCase().includes('target');
              let cls = '';
              if (isCode) cls = 'code-cell';
              else if (isAns) cls = 'ans-cell';
              return `<td class="${cls}">${val}</td>`;
            }).join('')}
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}

// -------------------------------------------------------------
// ROUTE 1 GENERATOR
// -------------------------------------------------------------
function getRoute1Html() {
  const xlsxPath = path.resolve(__dirname, '..', 'ROUTE_1_ADMIN_MECH_AUDI', 'ROUTE1_ANSWER_KEY.xlsx');
  const wb = XLSX.readFile(xlsxPath);
  const masterRows = XLSX.utils.sheet_to_json(wb.Sheets['Master Answer Key']);
  const nsRows = XLSX.utils.sheet_to_json(wb.Sheets['Number Shuffle Sets']);
  const caesarRows = XLSX.utils.sheet_to_json(wb.Sheets['Caesar Cipher Key']);

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Route 1 Organizer Master Answer Key</title>
<style>${commonStyles}</style>
</head>
<body>

<div class="header">
  <span class="meta">CONFIDENTIAL &#8226; OFFICIAL ORGANIZER MASTER MANUAL</span>
  <h1>ROUTE 1 — COMPLETE ORGANIZER ANSWER KEY &amp; OPERATIONS GUIDE</h1>
  <span class="badge">ROUTE: ADMIN &rarr; MECH &rarr; LIBRARY / FOOD COURT &rarr; CYBER SECURITY &rarr; AUDI</span>
</div>

<div class="note-box">
  <strong>⚠️ ORGANIZER INSTRUCTION:</strong> This official manual corresponds to <code>ROUTE1_FINAL_CLUE_COMPLETE_ODK.xlsx</code>. All typed answers strictly enforce <strong>MANDATORY UPPERCASE</strong>.
</div>

<div class="section-title">1. MASTER FLOW &amp; VOLUNTEER PASS CODES</div>
${renderGenericTable(masterRows)}

<div class="section-title">2. MECH MINI-CHALLENGES &amp; TESTING MODES</div>
<table>
  <thead><tr><th>Mode / Variant</th><th>Challenge Details &amp; Task Instructions</th><th>Pass Code</th></tr></thead>
  <tbody>
    <tr><td><strong>Variant 1 (HG)</strong></td><td><strong>Human Shape:</strong> Recreate 7 poses using all 4 team members (Quad Arch, Diamond Pillar, Synchronized X, Lever Link, Rotary Wheel, Bridge Support, Starburst)</td><td class="code-cell">MECH-HG-7</td></tr>
    <tr><td><strong>Variant 2 (GAR)</strong></td><td><strong>Hidden Garland:</strong> Find 5 hidden objects (M12 Nut, Brass Gear Ring, Spacer, Copper Washer, Cotter Pin) &amp; tie with 1m rope into garland</td><td class="code-cell">MECH-GAR-5</td></tr>
    <tr><td><strong>Variant 3 (BOMB)</strong></td><td><strong>Bomb Defusal:</strong> Untangle knot puzzle box safely (Square Knot, Bowline Loop, Clove Hitch)</td><td class="code-cell">MECH-BOMB-DEFUSED</td></tr>
    <tr><td><strong>Multi-Phase 1</strong></td><td>Phase 1: Human Shape (7 poses with 4 team members)</td><td class="code-cell">MECH-P1-HG</td></tr>
    <tr><td><strong>Multi-Phase 2</strong></td><td>Phase 2: Hidden Garland (5 objects tied into rope garland)</td><td class="code-cell">MECH-P2-GAR</td></tr>
    <tr><td><strong>Multi-Phase 3</strong></td><td>Phase 3: Bomb Defusal (Untangle knot box)</td><td class="code-cell">MECH-P3-BOMB</td></tr>
    <tr><td><strong>Final MECH Gate</strong></td><td>Verified by MECH station head volunteer &rarr; Unlocks <strong>NEXT BLOCK: LIBRARY</strong></td><td class="code-cell">MECH-FINAL</td></tr>
  </tbody>
</table>

<div class="page-break"></div>

<div class="section-title">3. NUMBER SHUFFLE ANSWER BANK (ALL 56 QUESTIONS ACROSS 7 SETS)</div>
<p style="font-size:7pt; color:#475569; margin-bottom:4px;">Rule: 4=A, 5=B, 6=C... 29=Z. All 8 questions per set must be solved sequentially in UPPERCASE.</p>
${renderGenericTable(nsRows)}

<div class="section-title">4. CAESAR CIPHER DECRYPTION KEY (5 MANDATORY PHASES)</div>
<p style="font-size:7pt; color:#475569; margin-bottom:4px;">Shift Rule: Move each letter 5 positions BACKWARD (-5).</p>
${renderGenericTable(caesarRows)}

<div class="page-break"></div>

<div class="section-title">5. DETAILED CHALLENGE DESCRIPTIONS, MATERIALS &amp; PARALLEL STATION CAPACITIES</div>

<div class="station-card">
  <h3><span>Stage 1: Admin Hidden Object Search</span><span class="station-badge">PARALLEL STATIONS: 3 DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Teams search Admin/Gym outdoor perimeter to locate the physical clue token attached near the campus vending machine. Team presents token to volunteer desk for verification.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Admin Block Courtyard</div>
    <div class="meta-item"><strong>Materials:</strong> 3 Tagged Physical Object Cards</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Volunteers (1 per desk)</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Starting Teams (3-5 min each)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">R1-R1PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> MECH (Sound Clue)</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 2: Engine Sound Discovery Clue</span><span class="station-badge">PARALLEL STATIONS: DIGITAL / ALL PARALLEL</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Teams listen to <code>engine_sound.wav</code> in-app, identify sound as <code>ENGINE</code>, and deduce destination as <code>MECH</code>.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Mobile Device in Transit</div>
    <div class="meta-item"><strong>Materials:</strong> <code>engine_sound.wav</code> (Media)</div>
    <div class="meta-item"><strong>Staffing:</strong> Self-administered in ODK app</div>
    <div class="meta-item"><strong>Answers:</strong> <code>ENGINE</code> &rarr; <code>MECH</code></div>
    <div class="meta-item"><strong>Arrival Code:</strong> <code style="color:#b91c1c;">SOUND-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> MECH Workshop</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 3: MECH Mini-Challenges (Workshop)</span><span class="station-badge">PARALLEL STATIONS: 6 PARALLEL BAYS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 25 teams arrive (Top 15 qualify). Recreate 7 poses (Human Shape), collect &amp; tie 5 hardware parts (Hidden Garland), and untangle knot box (Bomb Defusal).</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> MECH Workshop Ground Floor</div>
    <div class="meta-item"><strong>Materials:</strong> 2 Pose Mats, 2 Rope &amp; Part Sets, 2 Knot Boxes</div>
    <div class="meta-item"><strong>Staffing:</strong> 6 Bay Judges + 1 Head Judge</div>
    <div class="meta-item"><strong>Capacity:</strong> 6 Teams simultaneously (8-10 min/team)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">MECH-FINAL</code> (Top 15 qualify)</div>
    <div class="meta-item"><strong>Next Block:</strong> Library / Food Court</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 4: Library / Food Court Number Shuffle</span><span class="station-badge">PARALLEL STATIONS: 4 VERIFICATION DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 15 qualified teams arrive (Top 7 qualify). Decode 8 cryptographic strings in assigned Set A–G using 4=A...29=Z rule.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Food Court / Library Foyer</div>
    <div class="meta-item"><strong>Materials:</strong> Number conversion reference sheets</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 15 Teams (5-7 min each)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">R3-FOOD-PASS</code> (Top 7 qualify)</div>
    <div class="meta-item"><strong>Next Block:</strong> Cyber Security</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 5: Cyber Security Caesar Cipher &amp; Seminar Agility</span><span class="station-badge">PARALLEL STATIONS: 3 DESKS + 2 AGILITY LANES</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 7 teams solve 5 sequential ciphers (-5 backward shift) and navigate physical obstacle course in Seminar Hall.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Cyber Security Seminar Hall</div>
    <div class="meta-item"><strong>Materials:</strong> Cipher wheels, cones, blindfolds</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Cipher Judges + 2 Lane Marshalls</div>
    <div class="meta-item"><strong>Capacity:</strong> 7 Teams (Top 5 qualify)</div>
    <div class="meta-item"><strong>Pass Codes:</strong> <code style="color:#b91c1c;">CAESAR-PASS</code> &bull; <code style="color:#b91c1c;">CY-PHYSICAL-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Main Auditorium</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 6: Main Auditorium Grand Finale</span><span class="station-badge">PARALLEL STATIONS: 2 STAGE ARENAS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Finalists solve Auditorium riddle (<code>AUDITORIUM</code>), Stage riddle (<code>STAGE</code>), and complete live physical coordination trial.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Main Auditorium Stage</div>
    <div class="meta-item"><strong>Materials:</strong> Coordination challenge props, scoreclocks</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Chief Judges + 2 Emcees</div>
    <div class="meta-item"><strong>Final Codes:</strong> Ans 1: <code>STAGE</code> &bull; Ans 2: <code>FINAL-ROUTE1</code></div>
    <div class="meta-item"><strong>Conclusion:</strong> 🏆 ROUTE 1 COMPLETE</div>
    <div class="meta-item"><strong>Podium:</strong> 1st, 2nd, 3rd Place Award Ceremony</div>
  </div>
</div>

</body>
</html>`;
}

// -------------------------------------------------------------
// ROUTE 2 GENERATOR
// -------------------------------------------------------------
function getRoute2Html() {
  const xlsxPath = path.resolve(__dirname, '..', 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'PATH2_ANSWER_KEY.xlsx');
  const wb = XLSX.readFile(xlsxPath);
  const gates = XLSX.utils.sheet_to_json(wb.Sheets['Master Gates & Codes'] || wb.Sheets[wb.SheetNames[0]]);
  const quickfire = XLSX.utils.sheet_to_json(wb.Sheets['MBA Quickfire Bank (70 Qs)'] || wb.Sheets[wb.SheetNames[1]]);
  const libData = XLSX.utils.sheet_to_json(wb.Sheets['Library Unscramble & Rangoli'] || wb.Sheets[wb.SheetNames[2]]);
  const aimlData = XLSX.utils.sheet_to_json(wb.Sheets['AIML 3-Phase Key'] || wb.Sheets[wb.SheetNames[3]]);

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Route 2 Organizer Master Answer Key</title>
<style>${commonStyles}</style>
</head>
<body>

<div class="header">
  <span class="meta">CONFIDENTIAL &#8226; OFFICIAL ORGANIZER MASTER MANUAL</span>
  <h1>ROUTE 2 (PATH 2) — COMPLETE ORGANIZER ANSWER KEY &amp; OPERATIONS GUIDE</h1>
  <span class="badge">ROUTE: ADMIN &rarr; MBA &rarr; LIBRARY &rarr; AIML &rarr; AUDI</span>
</div>

<div class="note-box">
  <strong>⚠️ ORGANIZER INSTRUCTION:</strong> This official manual corresponds to <code>PATH2_FINAL_ODK.xlsx</code>. All typed answers strictly enforce <strong>MANDATORY UPPERCASE</strong>.
</div>

<div class="section-title">1. MASTER FLOW &amp; VOLUNTEER PASS CODES</div>
${renderGenericTable(gates)}

<div class="section-title">2. AIML TRI-PHASE MASTER ANSWER BANK</div>
${renderGenericTable(aimlData)}

<div class="page-break"></div>

<div class="section-title">3. MBA QUICKFIRE COMPLETE QUESTION BANK (ALL 70 QUESTIONS ACROSS 7 SETS)</div>
${renderGenericTable(quickfire)}

<div class="page-break"></div>

<div class="section-title">4. LIBRARY MEMORY &amp; MISSING BOOK MYSTERY</div>
${renderGenericTable(libData)}

<div class="section-title">5. DETAILED CHALLENGE DESCRIPTIONS, MATERIALS &amp; PARALLEL STATION CAPACITIES</div>

<div class="station-card">
  <h3><span>Stage 1: Admin Object Search (Gym Area)</span><span class="station-badge">PARALLEL STATIONS: 3 DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Teams search the gym perimeter to find the hidden physical object token and present it to the volunteer desk.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Gym / Sports Courtyard</div>
    <div class="meta-item"><strong>Materials:</strong> Tagged physical object cards</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Starting Teams (3-5 min each)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P2-ADMIN-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> MBA (Clue: BETWEEN EXPERIENCE &amp; BEGINNINGS)</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 2: MBA Quickfire &amp; 15-Minute Sales Challenge</span><span class="station-badge">PARALLEL STATIONS: 4 VERIFICATION DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 20 qualified teams solve in-app business quickfire, then receive selling inventory pouches (ducks, pens, lollipops) to sell across campus for 15 minutes.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> MBA Foyer &amp; Campus Perimeter</div>
    <div class="meta-item"><strong>Materials:</strong> 20 Inventory Pouches, Cash Collection Boxes</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 20 Teams (Top 15 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P2-MBA-PASS</code> (Top 15 qualify)</div>
    <div class="meta-item"><strong>Next Block:</strong> Central Library</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 3: Central Library Visual Memory &amp; Catalog Mystery</span><span class="station-badge">PARALLEL STATIONS: 4 LIBRARY BAYS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 15 teams observe 16-book visual memory matrix for 30 seconds, answer recall questions, and solve missing book catalog riddle.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Central Library 1st Floor</div>
    <div class="meta-item"><strong>Materials:</strong> Memory matrix displays, stopwatches</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 15 Teams (Top 7 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P2-LIB-PASS</code> (Top 7 qualify)</div>
    <div class="meta-item"><strong>Next Block:</strong> AIML Computer Lab</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 4: AIML Tri-Phase Computer Lab Challenge</span><span class="station-badge">PARALLEL STATIONS: 7 WORKSTATIONS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 7 teams complete Optical Illusion Animal Count, Dataset Concept Table, and 10s AI Image Generation Prompting on dedicated lab PCs.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> AIML Main Computer Lab</div>
    <div class="meta-item"><strong>Materials:</strong> 7 Logged-in AI prompt terminals</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Lab Coordinators</div>
    <div class="meta-item"><strong>Capacity:</strong> 7 Workstations (Top 5 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P2-AIML-PASS</code> (Top 5 qualify)</div>
    <div class="meta-item"><strong>Next Block:</strong> Main Auditorium</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 5: Main Auditorium Grand Finale</span><span class="station-badge">PARALLEL STATIONS: 2 STAGE ARENAS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Finalists solve Auditorium riddle (<code>AUDITORIUM</code>), Stage riddle (<code>STAGE</code>), and complete live physical coordination trial.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Main Auditorium Stage</div>
    <div class="meta-item"><strong>Materials:</strong> Physical coordination trial props</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Chief Judges</div>
    <div class="meta-item"><strong>Final Codes:</strong> Ans 1: <code>STAGE</code> &bull; Ans 2: <code>FINAL-PATH2</code></div>
    <div class="meta-item"><strong>Conclusion:</strong> 🏆 PATH 2 COMPLETE</div>
    <div class="meta-item"><strong>Podium:</strong> 1st, 2nd, 3rd Place Award Ceremony</div>
  </div>
</div>

</body>
</html>`;
}

// -------------------------------------------------------------
// ROUTE 3 GENERATOR
// -------------------------------------------------------------
function getRoute3Html() {
  const xlsxPath = path.resolve(__dirname, '..', 'ROUTE_3_PATH3_MBA_ADMIN_ECE_LIB_FC_AUDI', 'PATH3_ANSWER_KEY.xlsx');
  const wb = XLSX.readFile(xlsxPath);
  const flow = XLSX.utils.sheet_to_json(wb.Sheets['Master Flow & Codes']);
  const adminData = XLSX.utils.sheet_to_json(wb.Sheets['Admin (Elemental, Count, Piece)']);
  const eceQuiz = XLSX.utils.sheet_to_json(wb.Sheets['ECE Science Quiz Bank']);
  const eceQr = XLSX.utils.sheet_to_json(wb.Sheets['ECE QR & Physical Hunt']);
  const libCaesar = XLSX.utils.sheet_to_json(wb.Sheets['Library Caesar Decryption']);

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Route 3 Organizer Master Answer Key</title>
<style>${commonStyles}</style>
</head>
<body>

<div class="header">
  <span class="meta">CONFIDENTIAL &#8226; OFFICIAL ORGANIZER MASTER MANUAL</span>
  <h1>ROUTE 3 (PATH 3) — COMPLETE ORGANIZER ANSWER KEY &amp; OPERATIONS GUIDE</h1>
  <span class="badge">ROUTE: MBA &rarr; ADMIN &rarr; ECE &rarr; LIBRARY &rarr; FOOD COURT &rarr; AUDI</span>
</div>

<div class="note-box">
  <strong>⚠️ ORGANIZER INSTRUCTION:</strong> This official manual corresponds to <code>PATH3_FINAL_ODK.xlsx</code>. All typed answers strictly enforce <strong>MANDATORY UPPERCASE</strong>.
</div>

<div class="section-title">1. MASTER FLOW &amp; VOLUNTEER PASS CODES</div>
${renderGenericTable(flow)}

<div class="section-title">2. ADMIN STAGE (ELEMENTAL, OBJECT COUNT &amp; PUZZLE PIECE)</div>
${renderGenericTable(adminData)}

<div class="page-break"></div>

<div class="section-title">3. ECE LOGIC &amp; CIRCUIT BREAKER QUESTION BANK</div>
${renderGenericTable(eceQuiz)}

<div class="section-title">4. ECE QR &amp; PHYSICAL HUNT DETAILS</div>
${renderGenericTable(eceQr)}

<div class="section-title">5. LIBRARY CAESAR DECRYPTION &amp; FOOD COURT KEYS</div>
${renderGenericTable(libCaesar)}

<div class="page-break"></div>

<div class="section-title">6. DETAILED CHALLENGE DESCRIPTIONS, MATERIALS &amp; PARALLEL STATION CAPACITIES</div>

<div class="station-card">
  <h3><span>Stage 1: MBA Sales Challenge &amp; Quickfire</span><span class="station-badge">PARALLEL STATIONS: 4 VERIFICATION DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 25 starting teams receive selling inventory pouches (15 mins campus selling) + solve MBA business quickfire.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> MBA Courtyard</div>
    <div class="meta-item"><strong>Materials:</strong> 25 Inventory pouches, cash logs</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Teams (Top 20 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P3-MBA-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Admin Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 2: Admin Physical Search (Gym Area)</span><span class="station-badge">PARALLEL STATIONS: 3 DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 20 teams search Admin outdoor perimeter for physical clue token. Present token to volunteer for verification.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Admin Block Courtyard</div>
    <div class="meta-item"><strong>Materials:</strong> Physical token cards</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 20 Teams (3-5 min each)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P3-ADMIN-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> ECE Hardware Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 3: ECE Logic Gate Circuit Breaker</span><span class="station-badge">PARALLEL STATIONS: 4 LAB TEST BENCHES</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Teams analyze Boolean logic gate matrices (AND/OR/XOR/NAND) across 7 variant sets on laboratory benches.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> ECE Hardware Lab</div>
    <div class="meta-item"><strong>Materials:</strong> 4 Circuit test benches, logic sheets</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Lab Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 20 Teams (Top 15 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P3-ECE-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Central Library</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 4: Central Library Book Memory &amp; Catalog Mystery</span><span class="station-badge">PARALLEL STATIONS: 4 LIBRARY BAYS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 15 teams observe 16-book visual matrix for 30s and solve missing book mystery.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Library 1st Floor</div>
    <div class="meta-item"><strong>Materials:</strong> Memory matrix boards, stopwatches</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 15 Teams (Top 7 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P3-LIB-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Food Court</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 5: Food Court Cryptogram Challenge</span><span class="station-badge">PARALLEL STATIONS: 3 VERIFICATION DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 7 teams decode 8 encrypted food &amp; dining words using 4=A...29=Z rule.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Food Court Foyer</div>
    <div class="meta-item"><strong>Materials:</strong> Cryptogram decoder sheets</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 7 Teams (Top 5 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P3-FC-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Main Auditorium</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 6: Main Auditorium Grand Finale</span><span class="station-badge">PARALLEL STATIONS: 2 STAGE ARENAS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Finalists solve Auditorium riddle (<code>AUDITORIUM</code>), Stage riddle (<code>STAGE</code>), and complete physical coordination trial.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Main Auditorium Stage</div>
    <div class="meta-item"><strong>Materials:</strong> Coordination challenge props</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Chief Judges</div>
    <div class="meta-item"><strong>Final Codes:</strong> Ans 1: <code>STAGE</code> &bull; Ans 2: <code>FINAL-PATH3</code></div>
    <div class="meta-item"><strong>Conclusion:</strong> 🏆 PATH 3 COMPLETE</div>
    <div class="meta-item"><strong>Podium:</strong> 1st, 2nd, 3rd Place Award Ceremony</div>
  </div>
</div>

</body>
</html>`;
}

// -------------------------------------------------------------
// ROUTE 4 GENERATOR
// -------------------------------------------------------------
function getRoute4Html() {
  const xlsxPath = path.resolve(__dirname, '..', 'ROUTE_4_PATH4_CANTEEN_CSE_CY_MBA_AUDI', 'PATH4_ANSWER_KEY.xlsx');
  const wb = XLSX.readFile(xlsxPath);
  const flow = XLSX.utils.sheet_to_json(wb.Sheets['Master Flow & Codes']);
  const cseBank = XLSX.utils.sheet_to_json(wb.Sheets['CSE Challenge Bank']);
  const cseLocks = XLSX.utils.sheet_to_json(wb.Sheets['CSE Lock Grid (42 Qs)']);
  const cseQr = XLSX.utils.sheet_to_json(wb.Sheets['CSE QR & Physical Tasks']);
  const cyCaesar = XLSX.utils.sheet_to_json(wb.Sheets['CY Caesar Decryption']);
  const mbaBank = XLSX.utils.sheet_to_json(wb.Sheets['MBA Quickfire Bank']);

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Route 4 Organizer Master Answer Key</title>
<style>${commonStyles}</style>
</head>
<body>

<div class="header">
  <span class="meta">CONFIDENTIAL &#8226; OFFICIAL ORGANIZER MASTER MANUAL</span>
  <h1>ROUTE 4 (PATH 4) — COMPLETE ORGANIZER ANSWER KEY &amp; OPERATIONS GUIDE</h1>
  <span class="badge">ROUTE: OLD CANTEEN &rarr; CSE &rarr; CY &rarr; MBA &rarr; AUDI</span>
</div>

<div class="note-box">
  <strong>⚠️ ORGANIZER INSTRUCTION:</strong> This official manual corresponds to <code>PATH4_FINAL_ODK.xlsx</code>. All typed answers strictly enforce <strong>MANDATORY UPPERCASE</strong>.
</div>

<div class="section-title">1. MASTER FLOW &amp; VOLUNTEER PASS CODES</div>
${renderGenericTable(flow)}

<div class="section-title">2. CSE GAUNTLET STAGES (EMOJI, AI SORTING, BINARY, MORSE)</div>
${renderGenericTable(cseBank)}

<div class="page-break"></div>

<div class="section-title">3. CSE LOCK GRID GAUNTLET (ALL 42 QUESTIONS ACROSS 7 SETS)</div>
${renderGenericTable(cseLocks)}

<div class="page-break"></div>

<div class="section-title">4. CSE BUILDING-WIDE QR CHAINS, DECOYS &amp; PHYSICAL TASKS</div>
${renderGenericTable(cseQr)}

<div class="section-title">5. CYBER CAESAR CIPHER &amp; MBA QUICKFIRE ANSWER BANKS</div>
${renderGenericTable(cyCaesar)}
<div class="table-title">MBA QUICKFIRE SAMPLE BANK</div>
${renderGenericTable(mbaBank.slice(0, 15))}

<div class="page-break"></div>

<div class="section-title">6. DETAILED CHALLENGE DESCRIPTIONS, MATERIALS &amp; PARALLEL STATION CAPACITIES</div>

<div class="station-card">
  <h3><span>Stage 1: Old Canteen Landmark Search</span><span class="station-badge">PARALLEL STATIONS: 3 DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 25 teams solve Old Canteen history riddle and retrieve physical clue token near the dining hall.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Old Canteen Courtyard</div>
    <div class="meta-item"><strong>Materials:</strong> Clue tokens</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Teams (3-5 min each)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P4-CANTEEN-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> CSE Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 2: CSE 6-Stage Building Gauntlet</span><span class="station-badge">PARALLEL STATIONS: 6 FLOOR CHECKPOINTS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Teams navigate Floors 1, 2, and 3: Emoji Math (Lab 101), AI Sorting (Room 204), Binary (Room 305 - 3 Tracks), Morse (Room 112), Lock Grid (Room 218), QR Hunt.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> CSE Building (Floors 1-3)</div>
    <div class="meta-item"><strong>Materials:</strong> QR code cards, decoy QRs, lock grids</div>
    <div class="meta-item"><strong>Staffing:</strong> 6 Floor Marshalls + 1 Head Judge</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Teams (Top 15 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P4-CSE-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Cyber Security Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 3: Cyber Security Caesar Cipher &amp; Seminar Agility</span><span class="station-badge">PARALLEL STATIONS: 3 DESKS + 2 AGILITY LANES</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 15 teams decrypt 5-phase cipher (-5 backward) and complete blindfolded physical agility trial in Seminar Hall.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Cyber Security &amp; Seminar Hall</div>
    <div class="meta-item"><strong>Materials:</strong> Cones, blindfolds, cipher wheels</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 15 Teams (Top 7 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P4-CY-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> MBA Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 4: MBA Rapid Business Pitch Challenge</span><span class="station-badge">PARALLEL STATIONS: 3 PITCH EVALUATION DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 7 teams formulate a 2-minute rapid business pitch and pricing strategy before volunteer judges.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> MBA Executive Classroom</div>
    <div class="meta-item"><strong>Materials:</strong> Pitch scoring rubrics</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 MBA Judges</div>
    <div class="meta-item"><strong>Capacity:</strong> 7 Teams (Top 5 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P4-MBA-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Main Auditorium</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 5: Main Auditorium Grand Finale</span><span class="station-badge">PARALLEL STATIONS: 2 STAGE ARENAS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Finalists solve Auditorium riddle (<code>AUDITORIUM</code>), Stage riddle (<code>STAGE</code>), and complete physical coordination trial.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Main Auditorium Stage</div>
    <div class="meta-item"><strong>Materials:</strong> Coordination challenge props</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Chief Judges</div>
    <div class="meta-item"><strong>Final Codes:</strong> Ans 1: <code>STAGE</code> &bull; Ans 2: <code>FINAL-PATH4</code></div>
    <div class="meta-item"><strong>Conclusion:</strong> 🏆 PATH 4 COMPLETE</div>
    <div class="meta-item"><strong>Podium:</strong> 1st, 2nd, 3rd Place Award Ceremony</div>
  </div>
</div>

</body>
</html>`;
}

// -------------------------------------------------------------
// ROUTE 5 GENERATOR
// -------------------------------------------------------------
function getRoute5Html() {
  const xlsxPath = path.resolve(__dirname, '..', 'ROUTE_5_PATH5_COE_AIML_CSE_MECH_AUDI', 'PATH5_ANSWER_KEY.xlsx');
  const wb = XLSX.readFile(xlsxPath);
  const flow = XLSX.utils.sheet_to_json(wb.Sheets['Master Flow & Codes']);
  const aimlBank = XLSX.utils.sheet_to_json(wb.Sheets['AIML Challenges Bank']);
  const cseBank = XLSX.utils.sheet_to_json(wb.Sheets['CSE Challenges Bank']);
  const cseLocks = XLSX.utils.sheet_to_json(wb.Sheets['CSE Lock Grid (42 Qs)']);
  const mechBank = XLSX.utils.sheet_to_json(wb.Sheets['MECH Challenges Bank']);

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Route 5 Organizer Master Answer Key</title>
<style>${commonStyles}</style>
</head>
<body>

<div class="header">
  <span class="meta">CONFIDENTIAL &#8226; OFFICIAL ORGANIZER MASTER MANUAL</span>
  <h1>ROUTE 5 (PATH 5) — COMPLETE ORGANIZER ANSWER KEY &amp; OPERATIONS GUIDE</h1>
  <span class="badge">ROUTE: COE (LIB) &rarr; AIML &rarr; CSE &rarr; MECH &rarr; AUDI</span>
</div>

<div class="note-box">
  <strong>⚠️ ORGANIZER INSTRUCTION:</strong> This official manual corresponds to <code>PATH5_FINAL_ODK.xlsx</code>. All typed answers strictly enforce <strong>MANDATORY UPPERCASE</strong>.
</div>

<div class="section-title">1. MASTER FLOW &amp; VOLUNTEER PASS CODES</div>
${renderGenericTable(flow)}

<div class="section-title">2. AIML TRI-PHASE MASTER ANSWER BANK</div>
${renderGenericTable(aimlBank)}

<div class="page-break"></div>

<div class="section-title">3. CSE 6-STAGE BUILDING GAUNTLET ANSWER BANK</div>
${renderGenericTable(cseBank)}

<div class="section-title">4. CSE LOCK GRID GAUNTLET (ALL 42 QUESTIONS ACROSS 7 SETS)</div>
${renderGenericTable(cseLocks)}

<div class="page-break"></div>

<div class="section-title">5. MECH SOUND &amp; MINI-CHALLENGES MASTER KEY</div>
${renderGenericTable(mechBank)}

<div class="section-title">6. DETAILED CHALLENGE DESCRIPTIONS, MATERIALS &amp; PARALLEL STATION CAPACITIES</div>

<div class="station-card">
  <h3><span>Stage 1: COE (Library Annex) Object Search</span><span class="station-badge">PARALLEL STATIONS: 3 DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 25 teams locate physical COE Board in Library Annex and present to volunteer desk.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Library Annex / COE Room</div>
    <div class="meta-item"><strong>Materials:</strong> COE Board object, tokens</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Teams (3-5 min each)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P5-COE-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> AIML Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 2: AIML Tri-Phase Computer Lab Challenge</span><span class="station-badge">PARALLEL STATIONS: 7 WORKSTATIONS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Optical illusion count (12 to 6), Dataset Concept Table, and 10s AI Image Generation Prompting on lab terminals.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> AIML Computer Lab</div>
    <div class="meta-item"><strong>Materials:</strong> 7 Logged-in AI prompt PCs</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Lab Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Teams (Top 15 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P5-AIML-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> CSE Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 3: CSE 6-Stage Building Gauntlet</span><span class="station-badge">PARALLEL STATIONS: 6 FLOOR STATIONS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 15 teams complete Emoji math, AI sorting, Binary identity (3 tracks A, B, C), Morse transmission, 42 lock grid, and QR hunt.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> CSE Building (Floors 1-3)</div>
    <div class="meta-item"><strong>Materials:</strong> QR code cards, lock stations</div>
    <div class="meta-item"><strong>Staffing:</strong> 6 CSE Marshalls</div>
    <div class="meta-item"><strong>Capacity:</strong> 15 Teams (Top 7 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P5-CSE-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> MECH Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 4: MECH Sound Clue &amp; Mini-Challenges</span><span class="station-badge">PARALLEL STATIONS: 6 WORKSHOP BAYS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 7 teams identify audio clue (<code>ENGINE</code> &rarr; <code>MECH</code>) and complete Human Shape, Hidden Garland, and Bomb Defusal knots.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> MECH Workshop Bay</div>
    <div class="meta-item"><strong>Materials:</strong> Pose mats, garland parts, knot boxes</div>
    <div class="meta-item"><strong>Staffing:</strong> 6 Bay Judges + 1 Head Judge</div>
    <div class="meta-item"><strong>Capacity:</strong> 7 Teams (Top 5 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P5-MECH-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Main Auditorium</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 5: Main Auditorium Grand Finale</span><span class="station-badge">PARALLEL STATIONS: 2 STAGE ARENAS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Finalists solve Auditorium riddle (<code>AUDITORIUM</code>), Stage riddle (<code>STAGE</code>), and complete live physical coordination trial.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Main Auditorium Stage</div>
    <div class="meta-item"><strong>Materials:</strong> Coordination challenge props</div>
    <div class="meta-item"><strong>Staffing:</strong> 3 Chief Judges</div>
    <div class="meta-item"><strong>Final Codes:</strong> Ans 1: <code>STAGE</code> &bull; Ans 2: <code>FINAL-PATH5</code></div>
    <div class="meta-item"><strong>Conclusion:</strong> 🏆 PATH 5 COMPLETE</div>
    <div class="meta-item"><strong>Podium:</strong> 1st, 2nd, 3rd Place Award Ceremony</div>
  </div>
</div>

</body>
</html>`;
}

function main() {
  const baseDir = path.resolve(__dirname, '..');

  console.log('Generating exhaustive organizer PDFs for ALL routes (1 to 5)...');

  // Route 1
  const r1Pdf = path.join(baseDir, 'ROUTE_1_ADMIN_MECH_AUDI', 'ROUTE1_ORGANIZER_ANSWER_KEY.pdf');
  renderHtmlToPdf(getRoute1Html(), r1Pdf);
  fs.copyFileSync(r1Pdf, path.join(baseDir, 'ROUTE1_ORGANIZER_ANSWER_KEY.pdf'));

  // Route 2
  const r2Pdf = path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'ROUTE2_PATH2_ORGANIZER_ANSWER_KEY.pdf');
  renderHtmlToPdf(getRoute2Html(), r2Pdf);
  fs.copyFileSync(r2Pdf, path.join(baseDir, 'ROUTE2_PATH2_ORGANIZER_ANSWER_KEY.pdf'));

  // Route 3
  const r3Pdf = path.join(baseDir, 'ROUTE_3_PATH3_MBA_ADMIN_ECE_LIB_FC_AUDI', 'ROUTE3_PATH3_ORGANIZER_ANSWER_KEY.pdf');
  renderHtmlToPdf(getRoute3Html(), r3Pdf);
  fs.copyFileSync(r3Pdf, path.join(baseDir, 'ROUTE3_PATH3_ORGANIZER_ANSWER_KEY.pdf'));

  // Route 4
  const r4Pdf = path.join(baseDir, 'ROUTE_4_PATH4_CANTEEN_CSE_CY_MBA_AUDI', 'ROUTE4_PATH4_ORGANIZER_ANSWER_KEY.pdf');
  renderHtmlToPdf(getRoute4Html(), r4Pdf);
  fs.copyFileSync(r4Pdf, path.join(baseDir, 'ROUTE4_PATH4_ORGANIZER_ANSWER_KEY.pdf'));

  // Route 5
  const r5Pdf = path.join(baseDir, 'ROUTE_5_PATH5_COE_AIML_CSE_MECH_AUDI', 'ROUTE5_PATH5_ORGANIZER_ANSWER_KEY.pdf');
  renderHtmlToPdf(getRoute5Html(), r5Pdf);
  fs.copyFileSync(r5Pdf, path.join(baseDir, 'ROUTE5_PATH5_ORGANIZER_ANSWER_KEY.pdf'));

  console.log('All 5 Route PDFs generated successfully!');
}

main();

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

function renderHtmlToPdf(htmlContent, outputPath) {
  const tempHtmlPath = path.resolve(`temp_render_${path.basename(outputPath, '.pdf')}.html`);
  fs.writeFileSync(tempHtmlPath, htmlContent, 'utf8');
  const absOutPath = path.resolve(outputPath);
  try {
    execSync(`"${edgePath}" --headless --disable-gpu --no-pdf-header-footer --print-to-pdf="${absOutPath}" "file:///${tempHtmlPath.replace(/\\/g, '/')}"`);
    console.log(`PDF successfully generated: ${outputPath}`);
  } catch (err) {
    console.error('PDF generation error:', err.message);
  } finally {
    if (fs.existsSync(tempHtmlPath)) fs.unlinkSync(tempHtmlPath);
  }
}

function getBase64Image(filePath) {
  if (fs.existsSync(filePath)) {
    const ext = path.extname(filePath).toLowerCase();
    const mime = ext === '.png' ? 'image/png' : 'image/jpeg';
    return `data:${mime};base64,${fs.readFileSync(filePath).toString('base64')}`;
  }
  return '';
}

function generateCombinedMasterGuideHtml(baseDir) {
  // Load Images as Base64
  const qrPath1 = getBase64Image(path.join(baseDir, 'Path 1.jpeg'));
  const qrPath2 = getBase64Image(path.join(baseDir, 'path 2.jpeg'));
  const qrPath3 = getBase64Image(path.join(baseDir, 'path 3.jpeg'));
  const qrPath4 = getBase64Image(path.join(baseDir, 'path 4.jpeg'));
  const qrPath5 = getBase64Image(path.join(baseDir, 'path 5.jpeg'));

  // Path 1 Media
  const p1Vending = getBase64Image(path.join(baseDir, 'ROUTE_1_PATH1_ECE_MECH_FC_CY_AUDI', 'media', 'vending machine hidden object.jpeg'));

  // Path 2 Media
  const p2Aiml = getBase64Image(path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'media', 'aiml.jpeg'));
  const p2Bambaloni = getBase64Image(path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'media', 'bambaloni.jpeg'));
  const p2Horrid = getBase64Image(path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'media', 'horrid henry.jpeg'));
  const p25Star = getBase64Image(path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'media', '5 star.jpeg'));
  const p2Nokia = getBase64Image(path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'media', 'nokia.jpeg'));
  const p2Durian = getBase64Image(path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'media', 'durian.jpeg'));
  const p2Michelin = getBase64Image(path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'media', 'michelin.jpeg'));
  const p2Palmonas = getBase64Image(path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI', 'media', 'palmonas.jpeg'));

  // Path 3 Media
  const p3Challenge = getBase64Image(path.join(baseDir, 'ROUTE_3_PATH3_MBA_ADMIN_ECE_LIB_AUDI', 'media', 'challenge.jpeg'));

  // Path 4 Media
  const p4Morse = getBase64Image(path.join(baseDir, 'ROUTE_4_PATH4_OLD_CANTEEN_CSE_CY_MBA_AUDI', 'media', 'morse.jpeg'));

  // Path 5 Media
  const p5Badminton = getBase64Image(path.join(baseDir, 'ROUTE_5_PATH5_COE_AIML_CSE_MECH_AUDI', 'media', 'badminton court.jpeg'));

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>FINAL CLUE 2026 — COMBINED MASTER ORGANIZER GUIDE</title>
<style>
  @page {
    size: A4;
    margin: 10mm 12mm 10mm 12mm;
  }
  body {
    font-family: 'Segoe UI', Arial, sans-serif;
    color: #0f172a;
    background: #ffffff;
    line-height: 1.35;
    font-size: 8.5pt;
  }
  .page-break {
    page-break-before: always;
  }
  .header-main {
    border-bottom: 3px solid #1e3a8a;
    padding-bottom: 8px;
    margin-bottom: 12px;
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
  }
  .header-main h1 {
    color: #1e3a8a;
    margin: 0;
    font-size: 16pt;
    letter-spacing: 0.5px;
  }
  .header-path {
    border-bottom: 3px solid;
    padding-bottom: 6px;
    margin-bottom: 10px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }
  .badge {
    display: inline-block;
    padding: 3px 8px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 8pt;
    color: #ffffff;
  }
  .code-badge {
    display: inline-block;
    background: #fee2e2;
    color: #991b1b;
    border: 1px solid #f87171;
    padding: 1px 6px;
    border-radius: 3px;
    font-family: 'Consolas', monospace;
    font-weight: bold;
    font-size: 8pt;
  }
  .start-code-badge {
    display: inline-block;
    background: #e0f2fe;
    color: #0369a1;
    border: 1px solid #7dd3fc;
    padding: 1px 6px;
    border-radius: 3px;
    font-family: 'Consolas', monospace;
    font-weight: bold;
    font-size: 8pt;
  }
  .section-title {
    color: #ffffff;
    padding: 4px 8px;
    font-size: 9.5pt;
    font-weight: bold;
    border-radius: 3px;
    margin: 10px 0 6px 0;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 10px;
    font-size: 8pt;
  }
  th, td {
    border: 1px solid #cbd5e1;
    padding: 4px 6px;
    text-align: left;
    vertical-align: top;
  }
  th {
    background: #f8fafc;
    color: #1e293b;
    font-weight: bold;
  }
  tr:nth-child(even) {
    background: #f8fafc;
  }
  .card-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
    margin-bottom: 8px;
  }
  .station-card {
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    padding: 6px 8px;
    background: #ffffff;
    page-break-inside: avoid;
    margin-bottom: 6px;
  }
  .station-card h3 {
    margin: 0 0 3px 0;
    font-size: 8.5pt;
    display: flex;
    justify-content: space-between;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 2px;
  }
  .rules-box {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    padding: 6px 10px;
    border-radius: 4px;
    margin-bottom: 10px;
    font-size: 8pt;
  }
  .qr-box {
    display: flex;
    align-items: center;
    gap: 12px;
    background: #f1f5f9;
    padding: 8px 12px;
    border-radius: 6px;
    border: 1px solid #cbd5e1;
    margin-bottom: 10px;
  }
  .qr-box img {
    width: 90px;
    height: 90px;
    object-fit: contain;
    background: #ffffff;
    padding: 4px;
    border-radius: 4px;
    border: 1px solid #94a3b8;
  }
  .image-gallery {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 6px;
    margin-top: 6px;
  }
  .image-card {
    border: 1px solid #cbd5e1;
    padding: 4px;
    border-radius: 4px;
    text-align: center;
    background: #ffffff;
  }
  .image-card img {
    max-width: 100%;
    height: 70px;
    object-fit: contain;
  }
  .image-card div {
    font-size: 7pt;
    font-weight: bold;
    margin-top: 2px;
  }
</style>
</head>
<body>

<!-- ============================================================= -->
<!-- COVER & EXECUTIVE SUMMARY -->
<!-- ============================================================= -->
<div class="header-main">
  <div>
    <h1>🏆 LUMINUS PRESENTS: FINAL CLUE 2026</h1>
    <div style="font-size: 10pt; color: #475569; font-weight: 600; margin-top: 2px;">
      CONFIDENTIAL MASTER ORGANIZER & VOLUNTEER OPERATIONAL GUIDE
    </div>
  </div>
  <div style="text-align: right;">
    <span class="badge" style="background: #1e3a8a;">ALL 5 ROUTES MASTER KEY</span>
  </div>
</div>

<div class="rules-box">
  <strong>📋 MASTER EVENT RULES & QUALIFICATION METRICS:</strong><br>
  • <strong>Strict Team Gating (5 Rounds per path):</strong>
    Round 1 ➔ Round 2: <strong>Top 25 Teams</strong> &bull;
    Round 2 ➔ Round 3: <strong>Top 15 Teams</strong> &bull;
    Round 3 ➔ Round 4: <strong>Top 7 Teams</strong> &bull;
    Round 4 ➔ Round 5: <strong>Top 2 Teams</strong> to Grand Finale.<br>
  • <strong>ODK Collect / KoboToolbox Constraints:</strong> 100% Mandatory on all questions, photo uploads, barcode scans, and codes. All text entries require <strong>UPPERCASE (CAPS ONLY)</strong>.
</div>

<div class="section-title" style="background: #1e3a8a;">Master Route Overview & Progression Summary</div>

<table>
  <thead>
    <tr>
      <th style="width: 10%;">Route</th>
      <th style="width: 25%;">Starting Point & Object</th>
      <th style="width: 45%;">Full Station Progression</th>
      <th style="width: 20%;">Key Media Attachments</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Path 1</strong></td>
      <td>ECE (Vending Machine) &rarr; <code>CYBER-112</code></td>
      <td><code>ECE → MECH → FOOD COURT → CYBER → AUDITORIUM</code></td>
      <td><code>engine_sound.mpeg</code></td>
    </tr>
    <tr>
      <td><strong>Path 2</strong></td>
      <td>ADMIN (Gym / Object) &rarr; <code>JOHN-CENA</code></td>
      <td><code>ADMIN → MBA → LIBRARY → AIML → AUDITORIUM</code></td>
      <td><code>aiml.jpeg</code>, Brand Quiz Images (8)</td>
    </tr>
    <tr>
      <td><strong>Path 3</strong></td>
      <td>MBA (Visvesvaraya Object) &rarr; <code>ENG-KID-119</code></td>
      <td><code>MBA → ADMIN → ECE → LIBRARY → AUDITORIUM</code></td>
      <td><code>challenge.jpeg</code></td>
    </tr>
    <tr>
      <td><strong>Path 4</strong></td>
      <td>OLD CANTEEN (Cat Board) &rarr; <code>MEOW-BOW-4</code></td>
      <td><code>OLD CANTEEN → CSE → CYBER → MBA → AUDITORIUM</code></td>
      <td><code>morse.jpeg</code></td>
    </tr>
    <tr>
      <td><strong>Path 5</strong></td>
      <td>Station 1 (Library / Object) &rarr; <code>YAAKE-GURU-5</code></td>
      <td><code>STATION 1 → AIML → CSE → MECH → AUDITORIUM</code></td>
      <td><code>badminton court.jpeg</code></td>
    </tr>
  </tbody>
</table>

<!-- ============================================================= -->
<!-- PATH 1 -->
<!-- ============================================================= -->
<div class="page-break"></div>

<div class="header-path" style="border-color: #1d4ed8;">
  <div>
    <h2 style="margin: 0; color: #1d4ed8; font-size: 14pt;">PATH 1 MASTER GUIDE &bull; ECE &rarr; MECH &rarr; FOOD COURT &rarr; CYBER &rarr; AUDI</h2>
    <small style="color: #64748b;">ROUTE 1 OFFICIAL FORM, MEDIA & CLEARANCE CODES</small>
  </div>
  <span class="badge" style="background: #1d4ed8;">PATH 1</span>
</div>

<div class="qr-box">
  <img src="${qrPath1}" alt="Path 1 Form QR">
  <div>
    <strong style="font-size: 9.5pt; color: #1d4ed8;">Path 1 KoboToolbox / ODK Collect Form QR</strong><br>
    <span style="font-size: 8pt; color: #475569;">Participants scan this QR code at registration desk to open Path 1 on their ODK Collect app.</span><br>
    <span style="font-size: 8pt; font-weight: bold; color: #1d4ed8;">Form ID: PATH1_TREASURE_HUNT</span>
  </div>
</div>

<div class="section-title" style="background: #1d4ed8;">Path 1 Station-by-Station Directory</div>

<table>
  <thead>
    <tr>
      <th style="width: 12%;">Stage</th>
      <th style="width: 14%;">Location</th>
      <th style="width: 32%;">Clue / Question Prompt</th>
      <th style="width: 27%;">Expected Answer &amp; Volunteer Codes</th>
      <th style="width: 15%;">Uploads</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>R1 Object</strong></td>
      <td>ECE</td>
      <td>Find Vending Machine physical object</td>
      <td>Code: <span class="code-badge">CYBER-112</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Clue</strong></td>
      <td>In-App (<code>engine_sound.mpeg</code>)</td>
      <td>Audio Clue: Engine Sound</td>
      <td>Destination: <code>MECH</code><br>Start: <span class="start-code-badge">SOUND-PASS</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R2 Challenge</strong></td>
      <td>MECH</td>
      <td>
        &bull; <strong>Variant A:</strong> Human Poses (4 photos) &rarr; <span class="code-badge">MECH-PS-1</span><br>
        &bull; <strong>Variant B:</strong> Paper Ball challenge &rarr; <span class="code-badge">MECH-PB-2</span>
      </td>
      <td><span class="code-badge">MECH-PS-1</span> / <span class="code-badge">MECH-PB-2</span></td>
      <td>Var A: 4 Photos (Mandatory)</td>
    </tr>
    <tr>
      <td><strong>R3 Clue</strong></td>
      <td>In-App</td>
      <td>Crossword Riddle (4 definitions) &rarr; <code>FOODCOURT</code></td>
      <td>Destination: <code>FOODCOURT</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>FOOD COURT</td>
      <td>Find hidden QR code in Food Court</td>
      <td>Barcode: <span class="code-badge">CRICKET_CHAMP</span><br>Code: <span class="code-badge">FC-GOOD-LUCK</span></td>
      <td>Scan + Code (Mandatory)</td>
    </tr>
    <tr>
      <td><strong>R4 Clue</strong></td>
      <td>In-App</td>
      <td>Caesar Cipher (-3): <code>FBEHUBF</code></td>
      <td>Destination: <code>CYBER</code><br>Start: <span class="start-code-badge">START-PHY</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R4 Challenge</strong></td>
      <td>CYBER</td>
      <td>Physical Relay Challenge with volunteer</td>
      <td>Status: <code>PASSED</code><br>Finish: <span class="code-badge">FINISH-PHY</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Final Stage</strong></td>
      <td>Main AUDI</td>
      <td>
        Riddle 1 (Dark hall, red seats): <code>AUDITORIUM</code><br>
        Riddle 2 (Elevated platform): <code>STAGE</code>
      </td>
      <td>Clearance: <span class="code-badge">FINAL-PATH1</span><br>&rarr; 🔔 <strong>RUN &amp; RING THE BELL!</strong></td>
      <td>📸 Solved Puzzle Photo (Mandatory)</td>
    </tr>
  </tbody>
</table>

${p1Vending ? `
<div style="margin-top: 4px;">
  <strong style="font-size: 8pt; color: #1d4ed8;">Path 1 Media Assets:</strong>
  <div class="image-gallery" style="grid-template-columns: repeat(2, 1fr);">
    <div class="image-card">
      <img src="${p1Vending}" alt="Vending Machine">
      <div>R1: vending machine hidden object.jpeg</div>
    </div>
  </div>
</div>
` : ''}

<!-- ============================================================= -->
<!-- PATH 2 -->
<!-- ============================================================= -->
<div class="page-break"></div>

<div class="header-path" style="border-color: #047857;">
  <div>
    <h2 style="margin: 0; color: #047857; font-size: 14pt;">PATH 2 MASTER GUIDE &bull; ADMIN &rarr; MBA &rarr; LIBRARY &rarr; AIML &rarr; AUDI</h2>
    <small style="color: #64748b;">ROUTE 2 OFFICIAL FORM, MEDIA & CLEARANCE CODES</small>
  </div>
  <span class="badge" style="background: #047857;">PATH 2</span>
</div>

<div class="qr-box">
  <img src="${qrPath2}" alt="Path 2 Form QR">
  <div>
    <strong style="font-size: 9.5pt; color: #047857;">Path 2 KoboToolbox / ODK Collect Form QR</strong><br>
    <span style="font-size: 8pt; color: #475569;">Participants scan this QR code at registration desk to open Path 2 on their ODK Collect app.</span><br>
    <span style="font-size: 8pt; font-weight: bold; color: #047857;">Form ID: PATH2_TREASURE_HUNT</span>
  </div>
</div>

<div class="section-title" style="background: #047857;">Path 2 Station-by-Station Directory</div>

<table>
  <thead>
    <tr>
      <th style="width: 12%;">Stage</th>
      <th style="width: 14%;">Location</th>
      <th style="width: 32%;">Clue / Question Prompt</th>
      <th style="width: 27%;">Expected Answer &amp; Volunteer Codes</th>
      <th style="width: 15%;">Uploads</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>R1 Object</strong></td>
      <td>ADMIN</td>
      <td>Find Gym / Weight Board physical object</td>
      <td>Code: <span class="code-badge">JOHN-CENA</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Clue</strong></td>
      <td>In-App</td>
      <td>Math Equations (12x4, 30/2, 2x7-13, 20-19) &rarr; <code>MBA</code></td>
      <td>Destination: <code>MBA</code><br>Start: <span class="start-code-badge">MBA-START</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R2 Challenge</strong></td>
      <td>MBA</td>
      <td>
        &bull; <strong>Variant A:</strong> MBA Quickfire (10 Questions) &rarr; <span class="code-badge">MBA-QF-1</span><br>
        &bull; <strong>Variant B:</strong> Brand Quiz (10 Questions) &rarr; <span class="code-badge">MBA-BQ-2</span>
      </td>
      <td><span class="code-badge">MBA-QF-1</span> / <span class="code-badge">MBA-BQ-2</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R3 Clue</strong></td>
      <td>In-App (<code>aiml.jpeg</code>)</td>
      <td>Spot the Differences &rarr; <code>LIBRARY</code></td>
      <td>Destination: <code>LIBRARY</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>LIBRARY</td>
      <td>Find hidden QR code in Library area</td>
      <td>Barcode: <span class="code-badge">DUMB_FAKE</span></td>
      <td>Scan Barcode</td>
    </tr>
    <tr>
      <td><strong>R4 Clue</strong></td>
      <td>In-App</td>
      <td>Memory Recall: First letters of 5 specific answers &rarr; <code>AIML</code></td>
      <td>Destination: <code>AIML</code><br>Start: <span class="start-code-badge">START-AI-PHY</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R4 Challenge</strong></td>
      <td>AIML</td>
      <td>Physical Fitness/Coordination Challenge with volunteer</td>
      <td>Status: <code>PASSED</code><br>Finish: <span class="code-badge">AI-FIN-123</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Final Stage</strong></td>
      <td>Main AUDI</td>
      <td>
        Riddle 1 (Dark hall, red seats): <code>AUDITORIUM</code><br>
        Riddle 2 (Elevated platform): <code>STAGE</code>
      </td>
      <td>Clearance: <span class="code-badge">FINAL-PATH2</span><br>&rarr; 🔔 <strong>RUN &amp; RING THE BELL!</strong></td>
      <td>📸 Solved Puzzle Photo (Mandatory)</td>
    </tr>
  </tbody>
</table>

<div style="margin-top: 4px;">
  <strong style="font-size: 8pt; color: #047857;">Path 2 Media Assets:</strong>
  <div class="image-gallery">
    ${p2Aiml ? `<div class="image-card"><img src="${p2Aiml}" alt="AIML Spot Difference"><div>R3: aiml.jpeg</div></div>` : ''}
    ${p2Bambaloni ? `<div class="image-card"><img src="${p2Bambaloni}" alt="Bambaloni"><div>R2: bambaloni.jpeg</div></div>` : ''}
    ${p2Horrid ? `<div class="image-card"><img src="${p2Horrid}" alt="Horrid Henry"><div>R2: horrid henry.jpeg</div></div>` : ''}
    ${p25Star ? `<div class="image-card"><img src="${p25Star}" alt="5 Star"><div>R2: 5 star.jpeg</div></div>` : ''}
    ${p2Nokia ? `<div class="image-card"><img src="${p2Nokia}" alt="Nokia"><div>R2: nokia.jpeg</div></div>` : ''}
    ${p2Durian ? `<div class="image-card"><img src="${p2Durian}" alt="Durian"><div>R2: durian.jpeg</div></div>` : ''}
    ${p2Michelin ? `<div class="image-card"><img src="${p2Michelin}" alt="Michelin"><div>R2: michelin.jpeg</div></div>` : ''}
    ${p2Palmonas ? `<div class="image-card"><img src="${p2Palmonas}" alt="Palmonas"><div>R2: palmonas.jpeg</div></div>` : ''}
  </div>
</div>

<!-- ============================================================= -->
<!-- PATH 3 -->
<!-- ============================================================= -->
<div class="page-break"></div>

<div class="header-path" style="border-color: #7c3aed;">
  <div>
    <h2 style="margin: 0; color: #6d28d9; font-size: 14pt;">PATH 3 MASTER GUIDE &bull; MBA &rarr; ADMIN &rarr; ECE &rarr; LIBRARY &rarr; AUDI</h2>
    <small style="color: #64748b;">ROUTE 3 OFFICIAL FORM, MEDIA & CLEARANCE CODES</small>
  </div>
  <span class="badge" style="background: #7c3aed;">PATH 3</span>
</div>

<div class="qr-box">
  <img src="${qrPath3}" alt="Path 3 Form QR">
  <div>
    <strong style="font-size: 9.5pt; color: #6d28d9;">Path 3 KoboToolbox / ODK Collect Form QR</strong><br>
    <span style="font-size: 8pt; color: #475569;">Participants scan this QR code at registration desk to open Path 3 on their ODK Collect app.</span><br>
    <span style="font-size: 8pt; font-weight: bold; color: #7c3aed;">Form ID: PATH3_TREASURE_HUNT</span>
  </div>
</div>

<div class="section-title" style="background: #7c3aed;">Path 3 Station-by-Station Directory</div>

<table>
  <thead>
    <tr>
      <th style="width: 12%;">Stage</th>
      <th style="width: 14%;">Location</th>
      <th style="width: 32%;">Clue / Question Prompt</th>
      <th style="width: 27%;">Expected Answer &amp; Volunteer Codes</th>
      <th style="width: 15%;">Uploads</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>R1 Object</strong></td>
      <td>MBA</td>
      <td>Find Sir M. Visvesvaraya Board physical object</td>
      <td>Code: <span class="code-badge">ENG-KID-119</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Clue</strong></td>
      <td>In-App (<code>challenge.jpeg</code>)</td>
      <td>Periodic Table: Ag(47), Dy(66), In(49), N(7) &rarr; <code>ADMIN</code></td>
      <td>Destination: <code>ADMIN</code><br>Start: <span class="start-code-badge">ADMIN-START</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R2 Challenge</strong></td>
      <td>ADMIN</td>
      <td>
        &bull; <strong>Var A (Logical 1):</strong> Tech Jargon (10) + Logic Puzzle (10) &rarr; <span class="code-badge">ADMIN-NEWTON-3</span><br>
        &bull; <strong>Var B (Logical 2):</strong> Brain Teasers (10) + Aptitude (10) &rarr; <span class="code-badge">ADMIN-PASCAL-3</span>
      </td>
      <td><span class="code-badge">ADMIN-NEWTON-3</span> / <span class="code-badge">ADMIN-PASCAL-3</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R3 Clue</strong></td>
      <td>In-App</td>
      <td>Morse Code: <code>. -.-. .</code> &rarr; <code>ECE</code></td>
      <td>Destination: <code>ECE</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>ECE</td>
      <td>Find hidden QR code in ECE</td>
      <td>Barcode: <span class="code-badge">HARRY_POTTER</span></td>
      <td>📸 Photo + Scan</td>
    </tr>
    <tr>
      <td><strong>R4 Clue</strong></td>
      <td>In-App</td>
      <td>Caesar Cipher (-3): <code>OLEUDUB</code></td>
      <td>Destination: <code>LIBRARY</code><br>Arrival: <span class="start-code-badge">TALL-NERD</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R4 Challenge</strong></td>
      <td>LIBRARY</td>
      <td>Physical Agility Challenge with volunteer</td>
      <td>Status: <code>PASSED</code><br>Finish: <span class="code-badge">LAST-LEG</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Final Stage</strong></td>
      <td>Main AUDI</td>
      <td>
        Riddle 1 (Dark hall, red seats): <code>AUDITORIUM</code><br>
        Riddle 2 (Elevated platform): <code>STAGE</code>
      </td>
      <td>Clearance: <span class="code-badge">FINAL-PATH3</span><br>&rarr; 🔔 <strong>RUN &amp; RING THE BELL!</strong></td>
      <td>📸 Solved Puzzle Photo (Mandatory)</td>
    </tr>
  </tbody>
</table>

${p3Challenge ? `
<div style="margin-top: 4px;">
  <strong style="font-size: 8pt; color: #5b21b6;">Path 3 Media Assets:</strong>
  <div class="image-gallery" style="grid-template-columns: repeat(2, 1fr);">
    <div class="image-card">
      <img src="${p3Challenge}" alt="Elemental Encryption">
      <div>R2: challenge.jpeg (Elemental Encryption)</div>
    </div>
  </div>
</div>
` : ''}

<!-- ============================================================= -->
<!-- PATH 4 -->
<!-- ============================================================= -->
<div class="page-break"></div>

<div class="header-path" style="border-color: #d97706;">
  <div>
    <h2 style="margin: 0; color: #b45309; font-size: 14pt;">PATH 4 MASTER GUIDE &bull; OLD CANTEEN &rarr; CSE &rarr; CYBER &rarr; MBA &rarr; AUDI</h2>
    <small style="color: #64748b;">ROUTE 4 OFFICIAL FORM, MEDIA & CLEARANCE CODES</small>
  </div>
  <span class="badge" style="background: #d97706;">PATH 4</span>
</div>

<div class="qr-box">
  <img src="${qrPath4}" alt="Path 4 Form QR">
  <div>
    <strong style="font-size: 9.5pt; color: #b45309;">Path 4 KoboToolbox / ODK Collect Form QR</strong><br>
    <span style="font-size: 8pt; color: #475569;">Participants scan this QR code at registration desk to open Path 4 on their ODK Collect app.</span><br>
    <span style="font-size: 8pt; font-weight: bold; color: #d97706;">Form ID: PATH4_TREASURE_HUNT</span>
  </div>
</div>

<div class="section-title" style="background: #d97706;">Path 4 Station-by-Station Directory</div>

<table>
  <thead>
    <tr>
      <th style="width: 12%;">Stage</th>
      <th style="width: 14%;">Location</th>
      <th style="width: 32%;">Clue / Question Prompt</th>
      <th style="width: 27%;">Expected Answer &amp; Volunteer Codes</th>
      <th style="width: 15%;">Uploads</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>R1 Object</strong></td>
      <td>OLD CANTEEN</td>
      <td>Find Cat Board physical object</td>
      <td>Code: <span class="code-badge">MEOW-BOW-4</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Clue</strong></td>
      <td>In-App</td>
      <td>Emoji Math: 🍎+🍎=10, 🍎+🍌=7, 🍌+🍇=6 -> 🍇+🍎=?</td>
      <td>Result: <code>9</code> &rarr; Destination: <code>CSE</code><br>Start: <span class="start-code-badge">CSE-START</span></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R2 Challenge</strong></td>
      <td>CSE</td>
      <td>
        &bull; <strong>Var A (Logical 1):</strong> Algo Relay (10) + Cyber Detective (10) &rarr; <span class="code-badge">CSE-EINSTEIN-4</span><br>
        &bull; <strong>Var B (Logical 2):</strong> Bug Hunter (10) + Tech Detective (10) &rarr; <span class="code-badge">CSE-ZUCKER-4</span>
      </td>
      <td><span class="code-badge">CSE-EINSTEIN-4</span> / <span class="code-badge">CSE-ZUCKER-4</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R3 Clue</strong></td>
      <td>In-App (<code>morse.jpeg</code>)</td>
      <td>Morse Code: <code>--. --- / - --- / -.-. -.-- -... . .-.</code></td>
      <td>Destination: <code>GO TO CYBER</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>CYBER</td>
      <td>Find hidden QR code in Cyber block</td>
      <td>Barcode: <span class="code-badge">JAMES_BOND</span></td>
      <td>📸 Photo + Scan</td>
    </tr>
    <tr>
      <td><strong>R4 Clue</strong></td>
      <td>In-App</td>
      <td>Riddle: <em>"Numbers become stories... Stories become strategies... Where managers learn..."</em></td>
      <td>Destination: <code>MBA</code><br>Start: <span class="start-code-badge">MONEY_BROTHA</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R4 Challenge</strong></td>
      <td>MBA</td>
      <td>Physical Coordination Challenge</td>
      <td>Status: <code>PASSED</code><br>Finish: <span class="code-badge">RAVI-KISHEN</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Final Stage</strong></td>
      <td>Main AUDI</td>
      <td>
        Riddle 1 (Dark hall, red seats): <code>AUDITORIUM</code><br>
        Riddle 2 (Elevated platform): <code>STAGE</code>
      </td>
      <td>Clearance: <span class="code-badge">FINAL-PATH4</span><br>&rarr; 🔔 <strong>RUN &amp; RING THE BELL!</strong></td>
      <td>📸 Solved Puzzle Photo (Mandatory)</td>
    </tr>
  </tbody>
</table>

${p4Morse ? `
<div style="margin-top: 4px;">
  <strong style="font-size: 8pt; color: #b45309;">Path 4 Media Assets:</strong>
  <div class="image-gallery" style="grid-template-columns: repeat(2, 1fr);">
    <div class="image-card">
      <img src="${p4Morse}" alt="Morse Code Key">
      <div>R3: morse.jpeg (Morse Code Key)</div>
    </div>
  </div>
</div>
` : ''}

<!-- ============================================================= -->
<!-- PATH 5 -->
<!-- ============================================================= -->
<div class="page-break"></div>

<div class="header-path" style="border-color: #0284c7;">
  <div>
    <h2 style="margin: 0; color: #0369a1; font-size: 14pt;">PATH 5 MASTER GUIDE &bull; STATION 1 &rarr; AIML &rarr; CSE &rarr; MECH &rarr; AUDI</h2>
    <small style="color: #64748b;">ROUTE 5 OFFICIAL FORM, MEDIA & CLEARANCE CODES</small>
  </div>
  <span class="badge" style="background: #0284c7;">PATH 5</span>
</div>

<div class="qr-box">
  <img src="${qrPath5}" alt="Path 5 Form QR">
  <div>
    <strong style="font-size: 9.5pt; color: #0369a1;">Path 5 KoboToolbox / ODK Collect Form QR</strong><br>
    <span style="font-size: 8pt; color: #475569;">Participants scan this QR code at registration desk to open Path 5 on their ODK Collect app.</span><br>
    <span style="font-size: 8pt; font-weight: bold; color: #0284c7;">Form ID: PATH5_TREASURE_HUNT</span>
  </div>
</div>

<div class="section-title" style="background: #0284c7;">Path 5 Station-by-Station Directory</div>

<table>
  <thead>
    <tr>
      <th style="width: 12%;">Stage</th>
      <th style="width: 14%;">Location</th>
      <th style="width: 32%;">Clue / Question Prompt</th>
      <th style="width: 27%;">Expected Answer &amp; Volunteer Codes</th>
      <th style="width: 15%;">Uploads</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>R1 Object</strong></td>
      <td>Station 1 (Library)</td>
      <td>Find assigned physical object</td>
      <td>Code: <span class="code-badge">YAAKE-GURU-5</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Clues</strong></td>
      <td>In-App</td>
      <td>
        Missing Concept (Slide-by-Slide):<br>
        (A)lgorithm + (I)nternet + (M)odel + (L)anguage &rarr; <code>AIML</code>
      </td>
      <td>Destination: <code>AIML</code></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R2 Challenge</strong></td>
      <td>AIML Lab</td>
      <td>Reverse Image Prompting (Recreate 3 images in 3 AI chats)</td>
      <td>Start: <span class="start-code-badge">GEMMA-V05</span><br>Final: <span class="code-badge">OLLAMA-V05</span></td>
      <td>📸 Photo 1 (Mandatory) + Photo 2 (Opt.)</td>
    </tr>
    <tr>
      <td><strong>R3 Clue</strong></td>
      <td>In-App</td>
      <td>Mathematical Code: (1+2)=3(C), (20-1)=19(S), (10/2)=5(E)</td>
      <td>Destination: <code>CSE</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>CSE</td>
      <td>Find hidden QR code in CSE area</td>
      <td>Barcode: <span class="code-badge">PETER_PARKER</span></td>
      <td>Scan Barcode</td>
    </tr>
    <tr>
      <td><strong>R4 Clue</strong></td>
      <td>In-App (<code>badminton court.jpeg</code>)</td>
      <td>MEC-02 Piece-by-Piece photograph puzzle</td>
      <td>Destination: <code>MECH</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R4 Challenge</strong></td>
      <td>MECH</td>
      <td>Physical Coordination Challenge with volunteer</td>
      <td>Start: <span class="start-code-badge">PIECE-START</span><br>Finish: <span class="code-badge">PIECE-BYE</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Final Stage</strong></td>
      <td>Main AUDI</td>
      <td>
        Riddle 1 (Dark hall, red seats): <code>AUDITORIUM</code><br>
        Riddle 2 (Elevated platform): <code>STAGE</code>
      </td>
      <td>Clearance: <span class="code-badge">FINAL-PATH5</span><br>&rarr; 🔔 <strong>RUN &amp; RING THE BELL!</strong></td>
      <td>📸 Solved Puzzle Photo (Mandatory)</td>
    </tr>
  </tbody>
</table>

${p5Badminton ? `
<div style="margin-top: 4px;">
  <strong style="font-size: 8pt; color: #0369a1;">Path 5 Media Assets:</strong>
  <div class="image-gallery" style="grid-template-columns: repeat(2, 1fr);">
    <div class="image-card">
      <img src="${p5Badminton}" alt="Badminton Court">
      <div>R4: badminton court.jpeg (Piece-by-Piece)</div>
    </div>
  </div>
</div>
` : ''}

</body>
</html>`;
}

function main() {
  const baseDir = path.resolve(__dirname, '..');
  console.log('Generating Combined Master Organizer Guide PDF with all 5 Form QRs, questions, codes, and embedded media...');
  const htmlContent = generateCombinedMasterGuideHtml(baseDir);
  const pdfPath = path.join(baseDir, 'FINAL_CLUE_MASTER_ORGANIZER_GUIDE.pdf');
  renderHtmlToPdf(htmlContent, pdfPath);
  console.log(`🎉 COMBINED MASTER ORGANIZER GUIDE GENERATED: ${pdfPath}`);
}

main();

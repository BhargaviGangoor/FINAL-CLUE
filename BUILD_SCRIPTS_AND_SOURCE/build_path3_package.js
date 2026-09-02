const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
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

function buildSurvey() {
  const survey = [];

  // =============================================================
  // START — TEAM INFORMATION & RULES
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'start_group',
    label: 'START — TEAM INFORMATION'
  });

  survey.push({
    type: 'note',
    name: 'start_rules_note',
    label: '🏆 FINAL CLUE — PATH 3\n\n### 🏴 TREASURE HUNT — RULES\n\n1. *No phones allowed* during the hunt. Keep them switched off and safely stored.\n2. *Stay with your team* and follow the designated route/instructions.\n3. *No cheating or interference* with other teams, clues, props, or college property.\n4. Complete all challenges *as instructed by the coordinators*. Safety comes first.\n5. Any violation of the rules may result in a *time penalty or disqualification*. The coordinator’s decision will be final.\n\n⚠️ IMPORTANT:\n• All questions, code entries, and photo uploads are MANDATORY.\n• All text answers and codes must be entered in UPPERCASE ONLY.',
    hint: 'Read all rules carefully before proceeding.'
  });

  survey.push({
    type: 'text',
    name: 'team_name',
    label: 'Enter Team Name\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "regex(., '^[A-Z0-9\\-_ ]+$')",
    constraint_message: '❌ Please enter Team Name in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'team_id',
    label: 'Enter Team ID\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "regex(., '^[A-Z0-9\\-_ ]+$')",
    constraint_message: '❌ Please enter Team ID in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'image',
    name: 'team_start_photo',
    label: '📸 Upload Team Verification Photo\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the start desk.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  const startPassed = "${team_name} != '' and ${team_id} != '' and ${team_start_photo} != ''";

  // =============================================================
  // R1 — MBA: Object (Visvesvaraya)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r1_mba_group',
    label: 'ROUND 1 — OBJECT FINDING',
    relevant: startPassed
  });

  survey.push({
    type: 'note',
    name: 'r1_mba_intro',
    label: '📍 ROUND 1: 🧩 OBJECT FINDING\n\nChallenge: Locate the assigned hidden object/code at this station.\n\nInstructions:\n1. Search the location to find the assigned hidden object.\n2. Take a clear photo of the object.\n3. Show the photo to the station volunteer to receive your verification code.\n\nEnter code in caps',
    hint: 'Find the assigned hidden object.'
  });

  survey.push({
    type: 'image',
    name: 'r1_mba_photo',
    label: '📸 Upload Photo of Discovered Hidden Object\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered object.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'r1_mba_code',
    label: 'Enter Volunteer Verification Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ENG-KID-119'",
    constraint_message: '❌ Incorrect code. Enter ENG-KID-119 in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r1Passed = `${startPassed} and translate(normalize-space(\${r1_mba_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ENG-KID-119' and \${r1_mba_photo} != ''`;

  // =============================================================
  // R2 — BLOCK CLUE (ADM-01 — Elemental Encryption with challenge.jpeg)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r2_admin_block_group',
    label: 'ROUND 2 LOCATION CLUE',
    relevant: r1Passed
  });

  survey.push({
    type: 'note',
    name: 'r2_admin_block_note',
    label: '📍 ROUND 2 LOCATION CLUE: ELEMENTAL ENCRYPTION\n\nExamine the periodic table elements in the puzzle image below.\nDecode the chemical symbols by taking the first letter of each element to reveal your next block destination!\n\nEnter code in caps',
    hint: 'Decode the chemical symbols in the image to find the destination.',
    'media::image': 'challenge.jpeg'
  });

  survey.push({
    type: 'text',
    name: 'r2_admin_block_answer',
    label: 'Which campus block does this elemental encryption direct your team to?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    'media::image': 'challenge.jpeg',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADMIN' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADMIN BLOCK'",
    constraint_message: '❌ Incorrect destination. Decode the chemical symbols to identify the block in CAPS.'
  });

  survey.push({
    type: 'end_group'
  });

  const r2BlockPassed = `${r1Passed} and (translate(normalize-space(\${r2_admin_block_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADMIN' or translate(normalize-space(\${r2_admin_block_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADMIN BLOCK')`;

  // =============================================================
  // ADMIN START CODE (ADMIN-START)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'admin_start_group',
    label: 'ADMIN ARRIVAL',
    relevant: r2BlockPassed
  });

  survey.push({
    type: 'note',
    name: 'admin_arrival_note',
    label: '🏃 PROCEED TO ADMIN BLOCK\n\nReport immediately to the ADMIN station and meet the volunteer to receive your ADMIN start code.\n\nEnter code in caps',
    hint: 'Meet volunteer at ADMIN station.'
  });

  survey.push({
    type: 'text',
    name: 'r2_admin_start_code',
    label: 'Enter ADMIN Start Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADMIN-START'",
    constraint_message: '❌ Incorrect start code. Enter ADMIN-START in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const adminStartPassed = `${r2BlockPassed} and translate(normalize-space(\${r2_admin_start_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADMIN-START'`;

  // =============================================================
  // ADMIN MINI CHALLENGE (Variants 1 & 2)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'admin_mini_group',
    label: 'ROUND 2 — MINI CHALLENGE',
    relevant: adminStartPassed
  });

  survey.push({
    type: 'select_one admin_variant_list',
    name: 'r2_admin_variant',
    label: 'Select Assigned Challenge Variant\nMandatory selection',
    hint: 'Choose the variant assigned by the station volunteer.',
    required: 'yes'
  });

  // Variant 1: ADM-02 — Count to Unlock
  survey.push({
    type: 'begin_group',
    name: 'admin_v1_group',
    label: 'VARIANT 1 — COUNT TO UNLOCK',
    relevant: "${r2_admin_variant} = 'var1'"
  });

  survey.push({
    type: 'note',
    name: 'admin_v1_intro',
    label: '🔢 MINI-CHALLENGE: COUNT TO UNLOCK\n\nTeams have 3–4 minutes to find and count the hidden objects in the decorated seminar hall without touching anything.\n\nObjects to find:\n• Caps\n• Chess pieces\n• UNO cards\n• Sunglasses\n• Keys/Keychains\n• Water bottles\n\nRules: Do not touch or move any objects. Submit counts to volunteer to clear at least 4 out of 6.\n\nEnter code in caps',
    hint: 'Count objects and submit to volunteer.'
  });

  survey.push({
    type: 'image',
    name: 'r2_v1_admin_photo',
    label: '📸 Upload Photo of Count to Unlock Station\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the station.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'r2_v1_code',
    label: 'Enter Volunteer Clearance Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADM-DUCK-3'",
    constraint_message: '❌ Incorrect code. Enter ADM-DUCK-3 in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  // Variant 2: ADM-02 — Memory Room
  survey.push({
    type: 'begin_group',
    name: 'admin_v2_group',
    label: 'VARIANT 2 — MEMORY ROOM',
    relevant: "${r2_admin_variant} = 'var2'"
  });

  survey.push({
    type: 'note',
    name: 'admin_v2_intro',
    label: '🧠 MINI-CHALLENGE: MEMORY ROOM\n\nObserve the decorated classroom and its details for 20 seconds. Then leave the room and answer the volunteer questions based on memory.\n\nEnter code in caps',
    hint: 'Complete memory questions with volunteer.'
  });

  survey.push({
    type: 'image',
    name: 'r2_v2_admin_photo',
    label: '📸 Upload Photo of Memory Room Challenge\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the station.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_code',
    label: 'Enter Volunteer Clearance Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADM-DISNEY-3'",
    constraint_message: '❌ Incorrect code. Enter ADM-DISNEY-3 in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  survey.push({
    type: 'end_group'
  });

  const adminMiniPassed = `${adminStartPassed} and ((${'\${r2_admin_variant}'} = 'var1' and translate(normalize-space(\${r2_v1_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADM-DUCK-3' and \${r2_v1_admin_photo} != '') or (${'\${r2_admin_variant}'} = 'var2' and translate(normalize-space(\${r2_v2_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ADM-DISNEY-3' and \${r2_v2_admin_photo} != ''))`;

  // =============================================================
  // R3 — BLOCK CLUE (ECE-01 — Electronics Riddle)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_ece_block_group',
    label: 'ROUND 3 LOCATION CLUE',
    relevant: adminMiniPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_ece_riddle_note',
    label: '📍 ROUND 3 LOCATION CLUE: ELECTRONICS RIDDLE\n\n"I deal in waves, both low and high,\nWhere analog signals never lie.\nWith resistors, chips, and wires in place,\nFind your next clue in this circuit space."\n\nEnter code in caps',
    hint: 'Solve the riddle to deduce the destination block.'
  });

  survey.push({
    type: 'text',
    name: 'r3_ece_block_answer',
    label: 'Which campus block does this circuit riddle direct your team to?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ECE' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ECE BLOCK'",
    constraint_message: '❌ Incorrect destination. Read the riddle and enter the block name in CAPS.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3BlockPassed = `${adminMiniPassed} and (translate(normalize-space(\${r3_ece_block_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ECE' or translate(normalize-space(\${r3_ece_block_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ECE BLOCK')`;

  // =============================================================
  // ECE ARRIVAL & START CODE (ECE-LETSGO)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'ece_arrival_group',
    label: 'ECE ARRIVAL',
    relevant: r3BlockPassed
  });

  survey.push({
    type: 'note',
    name: 'ece_arrival_note',
    label: '🏃 PROCEED TO ECE BLOCK\n\nReport immediately to the ECE station and meet the volunteer to receive your ECE start code.\n\nEnter code in caps',
    hint: 'Meet volunteer at ECE station.'
  });

  survey.push({
    type: 'text',
    name: 'r3_ece_start_code',
    label: 'Enter ECE Volunteer Start Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ECE-LETSGO'",
    constraint_message: '❌ Incorrect code. Enter ECE-LETSGO in CAPS provided by the ECE volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const eceStartPassed = `${r3BlockPassed} and translate(normalize-space(\${r3_ece_start_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'ECE-LETSGO'`;

  // =============================================================
  // R3 — ECE QR HUNT (HARRY_POTTER)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_ece_qr_group',
    label: 'ROUND 3 — QR HUNT',
    relevant: eceStartPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_ece_qr_note',
    label: '📍 ROUND 3: QR HUNT\n\nSearch the location physically to locate the hidden QR code.\nScan the QR code using the scanner below.\n\nEnter code in caps',
    hint: 'Locate and scan the hidden QR code.'
  });

  survey.push({
    type: 'barcode',
    name: 'r3_ece_qr_scan',
    label: 'Scan Discovered QR Code',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = 'HARRY_POTTER' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'HARRY_POTTER'",
    constraint_message: '❌ Incorrect QR code scanned. Search for the correct QR code at this station.'
  });

  survey.push({
    type: 'image',
    name: 'r3_ece_qr_photo',
    label: '📸 Upload Photo of Discovered QR Code / Station\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered QR code location.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  const r3QrPassed = `${eceStartPassed} and (normalize-space(\${r3_ece_qr_scan}) = 'HARRY_POTTER' or translate(normalize-space(\${r3_ece_qr_scan}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'HARRY_POTTER') and \${r3_ece_qr_photo} != ''`;

  // =============================================================
  // R4 — LOCATION CLUE (LIB-01 — Caesar Cipher)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_lib_cipher_group',
    label: 'ROUND 4 LOCATION CLUE',
    relevant: r3QrPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_lib_cipher_note',
    label: '🔐 ROUND 4 LOCATION CLUE: CAESAR CIPHER\n\nDecode the encrypted message by moving each letter 3 steps backward in the alphabet to reveal your next destination!\n\n📜 CIPHER INSTRUCTION:\n"OLEUDUB — Julius Caesar would have understood this. You probably won\'t — until you move every letter three steps backward. Decode the message. Your answer is where the next clue waits."\n\nExample Shift Rule (-3):\n• D → A\n• G → D\n• M → J\n• Z → W\n\nEnter code in caps',
    hint: 'Move each letter 3 steps backward.'
  });

  survey.push({
    type: 'text',
    name: 'r4_lib_cipher_answer',
    label: 'Decode Encrypted Message [ OLEUDUB ]:\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'LIBRARY' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CENTRAL LIBRARY'",
    constraint_message: '❌ Incorrect destination. Decode OLEUDUB by shifting each letter 3 steps backward.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4CipherPassed = `${r3QrPassed} and (translate(normalize-space(\${r4_lib_cipher_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'LIBRARY' or translate(normalize-space(\${r4_lib_cipher_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CENTRAL LIBRARY')`;

  // =============================================================
  // LIBRARY ARRIVAL CODE (TALL-NERD)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'lib_arrival_group',
    label: 'LIBRARY ARRIVAL',
    relevant: r4CipherPassed
  });

  survey.push({
    type: 'note',
    name: 'lib_arrival_note',
    label: '🏃 PROCEED TO LIBRARY\n\nReport immediately to the Library station and meet the volunteer to receive your arrival verification code.\n\nEnter code in caps',
    hint: 'Meet volunteer at Library station.'
  });

  survey.push({
    type: 'text',
    name: 'r4_lib_start_code',
    label: 'Enter Library Volunteer Arrival Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'TALL-NERD'",
    constraint_message: '❌ Incorrect code. Enter TALL-NERD in CAPS provided by the Library volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const libStartPassed = `${r4CipherPassed} and translate(normalize-space(\${r4_lib_start_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'TALL-NERD'`;

  // =============================================================
  // R4 — LIBRARY PHYSICAL CHALLENGE (LAST-LEG)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_lib_phy_group',
    label: 'ROUND 4 — PHYSICAL CHALLENGE',
    relevant: libStartPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_lib_phy_note',
    label: '📍 ROUND 4: PHYSICAL CHALLENGE\n\nReport to the station and complete the physical challenge under volunteer supervision.\n\nEnter code in caps',
    hint: 'Complete physical challenge with volunteer.'
  });

  survey.push({
    type: 'image',
    name: 'r4_lib_phy_photo',
    label: '📸 Upload Photo of Physical Challenge Completion\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team completing the physical challenge.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'r4_lib_phy_code',
    label: 'Enter Volunteer Clearance Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'LAST-LEG'",
    constraint_message: '❌ Incorrect code. Enter LAST-LEG in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4PhyPassed = `${libStartPassed} and translate(normalize-space(\${r4_lib_phy_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'LAST-LEG' and \${r4_lib_phy_photo} != ''`;

  // =============================================================
  // FINAL PUZZLES: AUDITORIUM & STAGE RIDDLES
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'final_puzzles_group',
    label: 'FINAL PUZZLES',
    relevant: r4PhyPassed
  });

  survey.push({
    type: 'note',
    name: 'final_audi_stage_intro',
    label: '🏛️ FINAL CHALLENGES\n\nProceed to the location and solve the two final riddles!\n\nEnter code in caps',
    hint: 'Solve the final riddles.'
  });

  survey.push({
    type: 'text',
    name: 'final_riddle1_answer',
    label: '🧩 RIDDLE 1:\n\"I am a place of darkness until the lights ignite. I hold hundreds of red seats, host grand orientations, and echo with voices through microphones. Where are you standing?\"\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AUDITORIUM' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AUDI'",
    constraint_message: '❌ Incorrect answer. Solve the riddle and enter in CAPS.'
  });

  const finalRiddle1Passed = `${r4PhyPassed} and (translate(normalize-space(\${final_riddle1_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AUDITORIUM' or translate(normalize-space(\${final_riddle1_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AUDI')`;

  survey.push({
    type: 'text',
    name: 'final_riddle2_stage_answer',
    label: '🎭 RIDDLE 2:\n\"I am elevated above the crowd, where performers stand and spotlights shine. Underneath my wooden floor or behind the curtains, the ultimate secret waits. What am I?\"\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    relevant: finalRiddle1Passed,
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'STAGE'",
    constraint_message: '❌ Incorrect answer. Solve the stage riddle and enter in CAPS.'
  });

  const finalRiddle2Passed = `${finalRiddle1Passed} and translate(normalize-space(\${final_riddle2_stage_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'STAGE'`;

  survey.push({
    type: 'image',
    name: 'final_solved_puzzle_photo',
    label: '📸 Upload Photo of Your Solved Puzzle\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team\'s completed/solved puzzle.',
    required: 'yes',
    relevant: finalRiddle2Passed
  });

  survey.push({
    type: 'text',
    name: 'final_stage_volunteer_code',
    label: 'Enter Final Volunteer Clearance Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    relevant: finalRiddle2Passed,
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH3' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH1' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH2'",
    constraint_message: '❌ Incorrect code. Enter the code in CAPS provided by the Chief Judge.'
  });

  survey.push({
    type: 'note',
    name: 'final_congratulations_screen',
    label: '🎉 CONGRATULATIONS! PATH 3 COMPLETED.\n\n🏆 You have successfully conquered every challenge, puzzle, and cipher on PATH 3!\n\nShow this completion screen immediately to the Chief Judge to lock in your finishing timestamp and rank!',
    hint: 'Report to Chief Judge to finalize completion.',
    relevant: `${finalRiddle2Passed} and (translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH3' or translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH1' or translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH2') and \${final_solved_puzzle_photo} != ''`
  });

  survey.push({
    type: 'end_group'
  });

  return survey;
}

function buildChoices() {
  return [
    {
      list_name: 'admin_variant_list',
      name: 'var1',
      label: 'Variant 1: ADM-02 — Count to Unlock'
    },
    {
      list_name: 'admin_variant_list',
      name: 'var2',
      label: 'Variant 2: ADM-02 — Memory Room'
    }
  ];
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — PATH 3',
      form_id: 'final_clue_path3',
      version: '20260902',
      default_language: 'default'
    }
  ];
}

function buildAnswerKeyPdfHtml() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>PATH 3 — ORGANIZER MASTER ANSWER KEY</title>
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
  .header {
    border-bottom: 3px solid #7c3aed;
    padding-bottom: 6px;
    margin-bottom: 12px;
  }
  .header h1 {
    color: #5b21b6;
    margin: 0 0 3px 0;
    font-size: 15pt;
    letter-spacing: 0.4px;
  }
  .header .badge {
    display: inline-block;
    background: #7c3aed;
    color: #ffffff;
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 8pt;
  }
  .meta-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
    margin-bottom: 12px;
  }
  .meta-card {
    background: #f5f3ff;
    padding: 6px 8px;
    border-radius: 4px;
    border-left: 3px solid #8b5cf6;
  }
  .meta-card strong {
    color: #5b21b6;
    display: block;
    font-size: 7.5pt;
    text-transform: uppercase;
  }
  .meta-card span {
    font-size: 8.5pt;
    font-weight: 600;
  }
  .section-title {
    background: #5b21b6;
    color: #ffffff;
    padding: 4px 8px;
    font-size: 9.5pt;
    font-weight: bold;
    border-radius: 3px;
    margin: 12px 0 6px 0;
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
    padding: 5px 6px;
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
  .code-badge {
    display: inline-block;
    background: #fee2e2;
    color: #991b1b;
    border: 1px solid #f87171;
    padding: 1px 5px;
    border-radius: 3px;
    font-family: 'Consolas', monospace;
    font-weight: bold;
    font-size: 8pt;
  }
  .placeholder-badge {
    display: inline-block;
    background: #fef3c7;
    color: #92400e;
    border: 1px solid #fcd34d;
    padding: 1px 6px;
    border-radius: 3px;
    font-weight: bold;
    font-size: 7.5pt;
  }
  .station-card {
    border: 1px solid #94a3b8;
    border-radius: 4px;
    margin-bottom: 8px;
    padding: 6px 8px;
    background: #ffffff;
    page-break-inside: avoid;
  }
  .station-card h3 {
    margin: 0 0 4px 0;
    color: #5b21b6;
    font-size: 9pt;
    display: flex;
    justify-content: space-between;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 3px;
  }
</style>
</head>
<body>

<div class="header">
  <div style="float: right; text-align: right;">
    <span class="badge">ORGANIZER MASTER ANSWER KEY</span><br>
    <small style="color: #64748b;">PATH 3 &bull; FOR ORGANIZER USE ONLY</small>
  </div>
  <h1>PATH 3 &bull; TREASURE HUNT MASTER ANSWER KEY</h1>
  <div style="font-size: 9pt; color: #475569; font-weight: 600;">
    CONFIDENTIAL ORGANIZER KEY &bull; PROGRESSION: MBA &rarr; ADMIN &rarr; ECE &rarr; LIBRARY &rarr; AUDI
  </div>
</div>

<div class="meta-grid">
  <div class="meta-card">
    <strong>Starting Station</strong>
    <span>MBA Block</span>
  </div>
  <div class="meta-card">
    <strong>Final Station</strong>
    <span>Main Auditorium Stage</span>
  </div>
  <div class="meta-card">
    <strong>Input Rule</strong>
    <span>CAPS ONLY (Case-Insensitive Validated)</span>
  </div>
  <div class="meta-card">
    <strong>Media Assets</strong>
    <span>challenge.jpeg (Elemental Encryption)</span>
  </div>
</div>

<div class="section-title">1. Master Station-by-Station Directory</div>

<table>
  <thead>
    <tr>
      <th style="width: 12%;">Round</th>
      <th style="width: 16%;">Location</th>
      <th style="width: 26%;">Challenge Description</th>
      <th style="width: 26%;">Question / Decoded Clue</th>
      <th style="width: 20%;">Correct Answer / Code</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Registration</strong></td>
      <td>Start Desk</td>
      <td>START</td>
      <td>Team Name, Team ID &amp; Team Photo</td>
      <td><em>Team Info + 📸 Photo</em></td>
    </tr>
    <tr>
      <td><strong>R1</strong></td>
      <td>MBA</td>
      <td>🧩 Object (Visvesvaraya)</td>
      <td>Initial object/code finding</td>
      <td><span class="code-badge">ENG-KID-119</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Clue</strong></td>
      <td>In-App (<code>challenge.jpeg</code>)</td>
      <td>ADM-01 Elemental Encryption</td>
      <td>Periodic table symbols &rarr; first letters</td>
      <td><strong>ADMIN</strong></td>
    </tr>
    <tr>
      <td><strong>ADMIN Start</strong></td>
      <td>ADMIN</td>
      <td>Start Code</td>
      <td>Arrival verification at ADMIN block</td>
      <td><span class="code-badge">ADMIN-START</span></td>
    </tr>
    <tr>
      <td><strong>R2 V1</strong></td>
      <td>ADMIN</td>
      <td>ADM-02 Count to Unlock</td>
      <td>Count hidden objects (Caps, Chess, UNO, etc.)</td>
      <td><span class="code-badge">ADM-DUCK-3</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>R2 V2</strong></td>
      <td>ADMIN</td>
      <td>ADM-02 Memory Room</td>
      <td>Observe room 20s, answer memory questions</td>
      <td><span class="code-badge">ADM-DISNEY-3</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>R3 Clue</strong></td>
      <td>Clue &rarr; ECE</td>
      <td>ECE-01 Electronics Riddle</td>
      <td><em>"I deal in waves, both low and high..."</em></td>
      <td><strong>ECE</strong></td>
    </tr>
    <tr>
      <td><strong>ECE Start</strong></td>
      <td>ECE</td>
      <td>Start Code</td>
      <td>Arrival verification at ECE block</td>
      <td><span class="code-badge">ECE-LETSGO</span></td>
    </tr>
    <tr>
      <td><strong>R3</strong></td>
      <td>ECE</td>
      <td>QR Hunt</td>
      <td>Physical QR code scan at ECE station</td>
      <td><span class="code-badge">HARRY_POTTER</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>R4 Clue</strong></td>
      <td>Clue &rarr; LIBRARY</td>
      <td>LIB-01 Caesar Cipher</td>
      <td>Decode <code>OLEUDUB</code> (-3 shift backward)</td>
      <td><strong>LIBRARY</strong></td>
    </tr>
    <tr>
      <td><strong>Library Start</strong></td>
      <td>LIBRARY</td>
      <td>Arrival Code</td>
      <td>Arrival verification at Library</td>
      <td><span class="code-badge">TALL-NERD</span></td>
    </tr>
    <tr>
      <td><strong>R4 Physical</strong></td>
      <td>LIBRARY</td>
      <td>Physical Challenge</td>
      <td>Physical coordination task clearance</td>
      <td><span class="code-badge">LAST-LEG</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>Final Stage 1</strong></td>
      <td>AUDI</td>
      <td>AUDI-01</td>
      <td>Auditorium Riddle: Dark hall, red seats, mic, orientations</td>
      <td><strong>AUDITORIUM</strong> (or <strong>AUDI</strong>)</td>
    </tr>
    <tr>
      <td><strong>Final Stage 2</strong></td>
      <td>AUDI Stage</td>
      <td>AUDI-02</td>
      <td>Stage Riddle: Elevated platform, spotlights, curtains</td>
      <td><strong>STAGE</strong> + 📸 Solved Puzzle Photo<br>Code: <span class="code-badge">FINAL-PATH3</span></td>
    </tr>
    <tr>
      <td><strong>Completion</strong></td>
      <td>AUDI Stage</td>
      <td>FINISH</td>
      <td>Confirmation Screen</td>
      <td><strong>CONGRATULATIONS! PATH 3 COMPLETED.</strong></td>
    </tr>
  </tbody>
</table>

<div class="section-title">2. Detailed Clues &amp; Riddle Texts for Volunteers</div>

<div class="station-card">
  <h3><span>R2 Location Clue: Elemental Encryption (ADM-01)</span><span class="code-badge">ANSWER: ADMIN &bull; MEDIA: challenge.jpeg</span></h3>
  <div>
    Participants inspect the periodic table elemental symbols in the image, take the first letter of each element, and decode the target word &rarr; <strong>ADMIN</strong>.
  </div>
</div>

<div class="station-card">
  <h3><span>R3 Location Clue: Electronics Riddle (ECE-01)</span><span class="code-badge">ANSWER: ECE &bull; ARRIVAL: ECE-LETSGO</span></h3>
  <div>
    <em>"I deal in waves, both low and high,<br>
    Where analog signals never lie.<br>
    With resistors, chips, and wires in place,<br>
    Find your next clue in this circuit space."</em>
  </div>
</div>

<div class="station-card">
  <h3><span>R4 Location Clue: Caesar Cipher (LIB-01)</span><span class="code-badge">ANSWER: LIBRARY &bull; ARRIVAL: TALL-NERD &bull; CLEARANCE: LAST-LEG</span></h3>
  <div>
    <strong>Ciphertext:</strong> <code>OLEUDUB</code><br>
    <strong>Instruction:</strong> <em>"OLEUDUB — Julius Caesar would have understood this. You probably won't — until you move every letter three steps backward. Decode the message. Your answer is where the next clue waits."</em><br>
    <strong>Shift Mapping (-3):</strong> O&rarr;L, L&rarr;I, E&rarr;B, U&rarr;R, D&rarr;A, U&rarr;R, B&rarr;Y &rarr; <strong>LIBRARY</strong>
  </div>
</div>

<div class="section-title">3. Volunteer Station Instructions &amp; Codes</div>

<div class="station-card">
  <h3><span>ADMIN Station: Count to Unlock (V1) &amp; Memory Room (V2)</span><span class="code-badge">START: ADMIN-START</span></h3>
  <div>
    &bull; <strong>Arrival:</strong> Verify team arrival and provide start code <code>ADMIN-START</code>.<br>
    &bull; <strong>Variant 1 (Count to Unlock — ANU &amp; SHREEYA D):</strong> Teams have 3–4 mins to count 6 object types without touching. Must get &ge;4 correct. Provide code <code>ADM-DUCK-3</code>.<br>
    &bull; <strong>Variant 2 (Memory Room — ANU &amp; SHREEYA D):</strong> Teams observe classroom for 20s and answer questions outside. Provide code <code>ADM-DISNEY-3</code>.
  </div>
</div>

<div class="station-card">
  <h3><span>ECE Station: QR Hunt (Anushree &amp; Koushik)</span><span class="code-badge">START: ECE-LETSGO &bull; SCANNED QR: HARRY_POTTER</span></h3>
  <div>Verify arrival with <code>ECE-LETSGO</code>. Ensure the QR code with payload <code>HARRY_POTTER</code> is hidden in ECE area. Teams scan with in-app scanner.</div>
</div>

<div class="station-card">
  <h3><span>LIBRARY Station: Physical Challenge (SANCHITH &amp; PRUTHVI)</span><span class="code-badge">ARRIVAL: TALL-NERD &bull; CLEARANCE: LAST-LEG</span></h3>
  <div>Verify arrival with <code>TALL-NERD</code>. Supervise team completing physical agility task and provide clearance code <code>LAST-LEG</code> after photo upload.</div>
</div>

<div class="station-card">
  <h3><span>Main Auditorium Stage: Grand Finale</span><span class="code-badge">CODE: FINAL-PATH3</span></h3>
  <div>Teams solve Riddle 1 (<code>AUDITORIUM</code>) and Riddle 2 (<code>STAGE</code>), upload a clear photo of their solved puzzle sheet, and present to Chief Judges for code <code>FINAL-PATH3</code>!</div>
</div>

</body>
</html>`;
}

function main() {
  const baseDir = path.resolve(__dirname, '..');
  const route3Dir = path.join(baseDir, 'ROUTE_3_PATH3_MBA_ADMIN_ECE_LIB_AUDI');
  if (!fs.existsSync(route3Dir)) fs.mkdirSync(route3Dir, { recursive: true });

  const mediaDir = path.join(route3Dir, 'media');
  if (!fs.existsSync(mediaDir)) fs.mkdirSync(mediaDir, { recursive: true });

  // If challenge.jpeg exists in route3Dir, copy to mediaDir
  const sourceImg = path.join(route3Dir, 'challenge.jpeg');
  const destImg = path.join(mediaDir, 'challenge.jpeg');
  if (fs.existsSync(sourceImg)) {
    fs.copyFileSync(sourceImg, destImg);
  }

  console.log('Generating PATH3_FINAL_ODK.xlsx (Updated with challenge.jpeg for Elemental Encryption)...');
  const survey = buildSurvey();
  const choices = buildChoices();
  const settings = buildSettings();

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(survey), 'survey');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(choices), 'choices');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settings), 'settings');

  const xlsxPath = path.join(route3Dir, 'PATH3_FINAL_ODK.xlsx');
  XLSX.writeFile(wb, xlsxPath);
  console.log(`Successfully created: ${xlsxPath}`);

  console.log('Generating PATH3_ANSWER_KEY.pdf...');
  const pdfHtml = buildAnswerKeyPdfHtml();
  const pdfPath = path.join(route3Dir, 'PATH3_ANSWER_KEY.pdf');
  renderHtmlToPdf(pdfHtml, pdfPath);
  console.log(`Successfully created: ${pdfPath}`);

  console.log('Packaging PATH3_COMPLETE_PACKAGE.zip and PATH3_MEDIA.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route3Dir}\\PATH3_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route3Dir}\\PATH3_FINAL_ODK.xlsx', '${route3Dir}\\PATH3_ANSWER_KEY.pdf', '${mediaDir}' -DestinationPath '${route3Dir}\\PATH3_COMPLETE_PACKAGE.zip' -Force"`);

  console.log('🎉 PATH 3 PACKAGE REBUILT & PACKAGED SUCCESSFULLY!');
}

main();

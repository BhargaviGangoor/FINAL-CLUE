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

function getWordSearchCondition(varName) {
  const words = [
    'MACHINE LEARNING',
    'NEURAL NETWORK',
    'PYTHON',
    'DATASET',
    'ALGORITHM',
    'TRAINING DATA',
    'DEEP LEARNING',
    'DATA MINING',
    'MODELING',
    'REGRESSION',
    'CLASSIFY'
  ];
  return words
    .map(w => `translate(normalize-space(${varName}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '${w}'`)
    .join(' or ');
}

function buildSurvey() {
  const survey = [];

  // =============================================================
  // START — REGISTRATION & RULES
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'start_group',
    label: 'START — REGISTRATION'
  });

  survey.push({
    type: 'note',
    name: 'start_rules_note',
    label: '🏆 FINAL CLUE — PATH 2\n\n### 🏴 TREASURE HUNT — RULES\n\n1. *No phones allowed* during the hunt. Keep them switched off and safely stored.\n2. *Stay with your team* and follow the designated route/instructions.\n3. *No cheating or interference* with other teams, clues, props, or college property.\n4. Complete all challenges *as instructed by the coordinators*. Safety comes first.\n5. Any violation of the rules may result in a *time penalty or disqualification*. The coordinator’s decision will be final.\n\n⚠️ IMPORTANT:\n• All questions, code entries, and photo uploads are MANDATORY.\n• All text answers and codes must be entered in UPPERCASE ONLY.',
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
    name: 'player_id',
    label: 'Enter Team / Player Identification\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "regex(., '^[A-Z0-9\\-_ ]+$')",
    constraint_message: '❌ Please enter Player/Team Identification in UPPERCASE (CAPS ONLY).'
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

  const startPassed = "${team_name} != '' and ${player_id} != '' and ${team_start_photo} != ''";

  // =============================================================
  // R1 — ADMIN (Object / Gym)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r1_admin_group',
    label: 'ROUND 1 — OBJECT FINDING',
    relevant: startPassed
  });

  survey.push({
    type: 'note',
    name: 'r1_admin_note',
    label: '📍 ROUND 1: 🧩 OBJECT FINDING\n\nChallenge: Search the location to find the assigned hidden object/code.\n\nInstructions:\n1. Search the location to find the assigned hidden object.\n2. Take a clear photo of the object.\n3. Show the photo to the station volunteer to receive your verification code.\n\nEnter code in caps',
    hint: 'Find the assigned hidden object.'
  });

  survey.push({
    type: 'image',
    name: 'r1_admin_photo',
    label: '📸 Upload Photo of the Discovered Hidden Object\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered object.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'r1_admin_code',
    label: 'Enter Volunteer Verification Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'JOHN-CENA'",
    constraint_message: '❌ Incorrect code. Enter the code in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r1Passed = `${startPassed} and translate(normalize-space(\${r1_admin_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'JOHN-CENA' and \${r1_admin_photo} != ''`;

  // =============================================================
  // R2 — MBA LOCATION CLUE (MBA-03)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r2_mba_loc_group',
    label: 'ROUND 2 LOCATION CLUE',
    relevant: r1Passed
  });

  survey.push({
    type: 'note',
    name: 'r2_mba_loc_note',
    label: '📍 ROUND 2 LOCATION CLUE\n\nOne neighbour has already chosen its path.\nThe other is still preparing for the journey ahead.\nBetween the experienced and the yet-to-begin,\nyour next destination quietly stands.\n\nEnter code in caps',
    hint: 'Read the riddle to deduce the destination block.'
  });

  survey.push({
    type: 'text',
    name: 'r2_mba_loc_answer',
    label: 'Enter destination location\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA BLOCK'",
    constraint_message: '❌ Incorrect destination. Read the clue carefully and enter in CAPS.'
  });

  survey.push({
    type: 'end_group'
  });

  const r2LocPassed = `${r1Passed} and (translate(normalize-space(\${r2_mba_loc_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA' or translate(normalize-space(\${r2_mba_loc_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA BLOCK')`;

  // =============================================================
  // MBA CHALLENGE (Variants 1 & 2 -> Common Clearance Code MBA-QF-1)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'mba_challenge_group',
    label: 'ROUND 2 — MINI CHALLENGE',
    relevant: r2LocPassed
  });

  survey.push({
    type: 'select_one mba_variants',
    name: 'mba_variant_select',
    label: 'Select Assigned Challenge Variant\nMandatory selection',
    hint: 'Select the variant set assigned by the station volunteer.',
    required: 'yes'
  });

  // Variant 1: MBA Quickfire (Set A)
  survey.push({
    type: 'begin_group',
    name: 'mba_v1_group',
    label: 'VARIANT 1 — MBA QUICKFIRE (SET A)',
    relevant: "${mba_variant_select} = 'mba_01'"
  });

  survey.push({
    type: 'note',
    name: 'mba_v1_intro',
    label: '📊 MINI-CHALLENGE: MBA QUICKFIRE (SET A)\n\nAnswer all 10 quickfire business & management questions below.\n\nEnter code in caps',
    hint: 'Answer all 10 questions in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q1',
    label: '1. A market with only one seller is called?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MONOPOLY'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q2',
    label: '2. What is the currency of China?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'YUAN'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q3',
    label: '3. What does CEO stand for?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CHIEF EXECUTIVE OFFICER'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q4',
    label: '4. What does ROI stand for in business finance?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'RETURN ON INVESTMENT'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q5',
    label: '5. What term describes business transactions conducted between two companies (abbreviation)?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'B2B'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q6',
    label: '6. In accounting: Assets minus Liabilities equals what?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'EQUITY'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q7',
    label: '7. Which animal represents a rising, optimistic financial market?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'BULL'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q8',
    label: '8. What does IPO stand for when a company goes public?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'INITIAL PUBLIC OFFERING'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q9',
    label: '9. What is the standard 3-letter abbreviation for Gross Domestic Product?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'GDP'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q10',
    label: '10. In the 4 Ps of Marketing (Product, Price, Place), what is the 4th P?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'PROMOTION'",
    constraint_message: '❌ Incorrect answer. Please enter in CAPS.'
  });

  survey.push({
    type: 'image',
    name: 'r2_var1_quickfire_photo',
    label: '📸 Upload Photo of Team Completing Quickfire Challenge\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the station.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  // Variant 2: MBA Quickfire (Set B / Brand Quiz)
  survey.push({
    type: 'begin_group',
    name: 'mba_v2_group',
    label: 'VARIANT 2 — MBA QUICKFIRE (SET B)',
    relevant: "${mba_variant_select} = 'mba_02'"
  });

  survey.push({
    type: 'note',
    name: 'mba_v2_placeholder',
    label: 'BRAND QUIZ / QUICKFIRE SET B CONTENT TO BE ADDED\n\nParticipants solve the assigned Variant 2 questions.\n\n[CONTENT TO BE ADDED BY ORGANIZER BEFORE DEPLOYMENT]',
    hint: 'Organizer placeholder — to be updated before deployment.'
  });

  survey.push({
    type: 'image',
    name: 'r2_var2_quickfire_photo',
    label: '📸 Upload Photo of Team Completing Variant 2 Challenge\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the station.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  const mba1AllPassed = "(translate(normalize-space(${mba_q1}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MONOPOLY' and translate(normalize-space(${mba_q2}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'YUAN' and translate(normalize-space(${mba_q3}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CHIEF EXECUTIVE OFFICER' and translate(normalize-space(${mba_q4}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'RETURN ON INVESTMENT' and translate(normalize-space(${mba_q5}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'B2B' and translate(normalize-space(${mba_q6}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'EQUITY' and translate(normalize-space(${mba_q7}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'BULL' and translate(normalize-space(${mba_q8}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'INITIAL PUBLIC OFFERING' and translate(normalize-space(${mba_q9}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'GDP' and translate(normalize-space(${mba_q10}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'PROMOTION' and ${r2_var1_quickfire_photo} != '')";

  const variantAnswered = `((${mba1AllPassed} and \${mba_variant_select} = 'mba_01') or (\${mba_variant_select} = 'mba_02' and \${r2_var2_quickfire_photo} != ''))`;

  survey.push({
    type: 'text',
    name: 'r2_mba_clearance_code',
    label: 'Enter Volunteer Clearance Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    relevant: variantAnswered,
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA-QF-1'",
    constraint_message: '❌ Incorrect clearance code. Obtain clearance code in CAPS from the MBA volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3Unlocked = `${r2LocPassed} and ${variantAnswered} and translate(normalize-space(\${r2_mba_clearance_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA-QF-1'`;

  // =============================================================
  // R3 — LIBRARY LOCATION CLUE (LIB-04)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_lib_loc_group',
    label: 'ROUND 3 LOCATION CLUE',
    relevant: r3Unlocked
  });

  survey.push({
    type: 'note',
    name: 'r3_lib_loc_note',
    label: '📍 ROUND 3 LOCATION CLUE\n\nVEHSELS\nGIDRAEN\nIFIW\nSISCUDISNO GGUONLE\nIGIDLAT BILRYRA\n\nEnter code in caps',
    hint: 'Decode the anagrams to identify the next destination.'
  });

  survey.push({
    type: 'text',
    name: 'r3_lib_loc_answer',
    label: 'Enter confirmed destination location\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'LIBRARY' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CENTRAL LIBRARY'",
    constraint_message: '❌ Incorrect destination. Decode the anagrams to identify the location in CAPS.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3LocPassed = `${r3Unlocked} and (translate(normalize-space(\${r3_lib_loc_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'LIBRARY' or translate(normalize-space(\${r3_lib_loc_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CENTRAL LIBRARY')`;

  // =============================================================
  // R3 — LIBRARY QR HUNT
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_lib_qr_group',
    label: 'ROUND 3 — QR HUNT',
    relevant: r3LocPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_lib_qr_note',
    label: '📍 ROUND 3: QR HUNT\n\nSearch the location physically to locate the hidden QR code.\nScan the QR code using the scanner below.\n\nEnter code in caps',
    hint: 'Locate and scan the hidden QR code.'
  });

  survey.push({
    type: 'barcode',
    name: 'r3_lib_qr_scan',
    label: 'Scan Discovered QR Code',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = 'DUMB_FAKE' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'DUMB_FAKE'",
    constraint_message: '❌ Incorrect QR code scanned. Search for the correct QR code at this station.'
  });

  survey.push({
    type: 'image',
    name: 'r3_lib_qr_photo',
    label: '📸 Upload Photo of Discovered QR Code / Station\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered QR code location.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  const r3QrPassed = `${r3LocPassed} and (normalize-space(\${r3_lib_qr_scan}) = 'DUMB_FAKE' or translate(normalize-space(\${r3_lib_qr_scan}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'DUMB_FAKE') and \${r3_lib_qr_photo} != ''`;

  // =============================================================
  // R4 LOCATION CLUE: WORD SEARCH (aiml.jpeg in EVERY question)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_wordsearch_group',
    label: 'ROUND 4 LOCATION CLUE: WORD SEARCH',
    relevant: r3QrPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_wordsearch_intro',
    label: '🧩 ROUND 4 LOCATION CLUE: WORD SEARCH PUZZLE\n\nExamine the word search puzzle image below.\nFind and enter ANY 5 hidden technology concepts from the grid!\n\nInstructions:\n1. Spot at least 5 hidden words in the puzzle image.\n2. Enter each word in the fields below.\n3. After finding 5 valid words, deduce the destination campus block!\n\nEnter code in caps',
    hint: 'Examine puzzle image and enter 5 discovered words in CAPS.',
    'media::image': 'aiml.jpeg'
  });

  const wCond = getWordSearchCondition('.');

  survey.push({
    type: 'text',
    name: 'r4_ws_word1',
    label: 'Enter Discovered Word 1 (from puzzle image):\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    'media::image': 'aiml.jpeg',
    constraint: wCond,
    constraint_message: '❌ Invalid word. Enter a valid word found in the puzzle in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'r4_ws_word2',
    label: 'Enter Discovered Word 2 (from puzzle image):\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    'media::image': 'aiml.jpeg',
    constraint: `${wCond} and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word1}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')`,
    constraint_message: '❌ Invalid or duplicate word. Enter a different valid word from the puzzle in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'r4_ws_word3',
    label: 'Enter Discovered Word 3 (from puzzle image):\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    'media::image': 'aiml.jpeg',
    constraint: `${wCond} and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word1}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word2}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')`,
    constraint_message: '❌ Invalid or duplicate word. Enter a different valid word from the puzzle in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'r4_ws_word4',
    label: 'Enter Discovered Word 4 (from puzzle image):\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    'media::image': 'aiml.jpeg',
    constraint: `${wCond} and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word1}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word2}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word3}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')`,
    constraint_message: '❌ Invalid or duplicate word. Enter a different valid word from the puzzle in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'r4_ws_word5',
    label: 'Enter Discovered Word 5 (from puzzle image):\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    'media::image': 'aiml.jpeg',
    constraint: `${wCond} and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word1}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word2}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word3}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') and translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') != translate(normalize-space(\${r4_ws_word4}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')`,
    constraint_message: '❌ Invalid or duplicate word. Enter a different valid word from the puzzle in CAPS.'
  });

  const w1Val = getWordSearchCondition('${r4_ws_word1}');
  const w2Val = getWordSearchCondition('${r4_ws_word2}');
  const w3Val = getWordSearchCondition('${r4_ws_word3}');
  const w4Val = getWordSearchCondition('${r4_ws_word4}');
  const w5Val = getWordSearchCondition('${r4_ws_word5}');

  const all5WordsValid = `(${w1Val}) and (${w2Val}) and (${w3Val}) and (${w4Val}) and (${w5Val})`;

  survey.push({
    type: 'text',
    name: 'r4_aiml_destination',
    label: 'Which campus block do these 5 discovered concepts direct your team to?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    'media::image': 'aiml.jpeg',
    relevant: all5WordsValid,
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AIML' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AI ML' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AIML BLOCK' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AI/ML' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AI & ML'",
    constraint_message: '❌ Incorrect destination. Deduce the campus block in CAPS from the 5 discovered concepts.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4DestPassed = `${r3QrPassed} and ${all5WordsValid} and (translate(normalize-space(\${r4_aiml_destination}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AIML' or translate(normalize-space(\${r4_aiml_destination}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AI ML' or translate(normalize-space(\${r4_aiml_destination}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AIML BLOCK' or translate(normalize-space(\${r4_aiml_destination}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AI/ML' or translate(normalize-space(\${r4_aiml_destination}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AI & ML')`;

  // =============================================================
  // R4 — PHYSICAL CHALLENGE (AIML & Cyber Pathway)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_phy_group',
    label: 'ROUND 4 — PHYSICAL CHALLENGE',
    relevant: r4DestPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_phy_note',
    label: '📍 ROUND 4: PHYSICAL CHALLENGE\n\nReport immediately to the checkpoint along the pathway!\n\nInstructions:\n1. Meet the station volunteers.\n2. Complete the physical coordination challenge.\n3. Take a verification photo.\n4. Enter clearance code from the volunteer.\n\nEnter code in caps',
    hint: 'Complete physical challenge with volunteer.'
  });

  survey.push({
    type: 'image',
    name: 'r4_phy_photo',
    label: '📸 Upload Photo of Physical Challenge Completion\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team completing the physical coordination challenge.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'r4_phy_code',
    label: 'Enter Volunteer Verification Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'PHY-CY'",
    constraint_message: '❌ Incorrect verification code. Obtain the code from the station volunteer and enter in CAPS.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4Passed = `${r4DestPassed} and translate(normalize-space(\${r4_phy_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'PHY-CY' and \${r4_phy_photo} != ''`;

  // =============================================================
  // FINAL PUZZLES: AUDITORIUM & STAGE RIDDLES
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'final_puzzles_group',
    label: 'FINAL PUZZLES',
    relevant: r4Passed
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

  const finalRiddle1Passed = `${r4Passed} and (translate(normalize-space(\${final_riddle1_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AUDITORIUM' or translate(normalize-space(\${final_riddle1_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'AUDI')`;

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
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH2' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH1'",
    constraint_message: '❌ Incorrect code. Enter the code in CAPS provided by the Chief Judge.'
  });

  survey.push({
    type: 'note',
    name: 'final_congratulations_screen',
    label: '🎉 CONGRATULATIONS! PATH 2 COMPLETED.\n\n🏆 You have successfully conquered every challenge, puzzle, and cipher on PATH 2!\n\nShow this completion screen immediately to the Chief Judge to lock in your finishing timestamp and rank!',
    hint: 'Report to Chief Judge to finalize completion.',
    relevant: `${finalRiddle2Passed} and (translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH2' or translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH1') and \${final_solved_puzzle_photo} != ''`
  });

  survey.push({
    type: 'end_group'
  });

  return survey;
}

function buildChoices() {
  return [
    {
      list_name: 'mba_variants',
      name: 'mba_01',
      label: 'Variant 1: MBA Quickfire (Set A)'
    },
    {
      list_name: 'mba_variants',
      name: 'mba_02',
      label: 'Variant 2: MBA Quickfire (Set B)'
    }
  ];
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — PATH 2',
      form_id: 'final_clue_path2',
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
<title>PATH 2 — ORGANIZER MASTER ANSWER KEY</title>
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
    border-bottom: 3px solid #047857;
    padding-bottom: 6px;
    margin-bottom: 12px;
  }
  .header h1 {
    color: #065f46;
    margin: 0 0 3px 0;
    font-size: 15pt;
    letter-spacing: 0.4px;
  }
  .header .badge {
    display: inline-block;
    background: #047857;
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
    background: #f0fdf4;
    padding: 6px 8px;
    border-radius: 4px;
    border-left: 3px solid #10b981;
  }
  .meta-card strong {
    color: #065f46;
    display: block;
    font-size: 7.5pt;
    text-transform: uppercase;
  }
  .meta-card span {
    font-size: 8.5pt;
    font-weight: 600;
  }
  .section-title {
    background: #065f46;
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
    color: #065f46;
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
    <small style="color: #64748b;">PATH 2 &bull; FOR ORGANIZER USE ONLY</small>
  </div>
  <h1>PATH 2 &bull; TREASURE HUNT MASTER ANSWER KEY</h1>
  <div style="font-size: 9pt; color: #475569; font-weight: 600;">
    CONFIDENTIAL ORGANIZER KEY &bull; PROGRESSION: ADMIN &rarr; MBA &rarr; LIBRARY &rarr; AIML &rarr; AUDI
  </div>
</div>

<div class="meta-grid">
  <div class="meta-card">
    <strong>Starting Station</strong>
    <span>ADMIN Block</span>
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
    <span>aiml.jpeg (Word Search Grid)</span>
  </div>
</div>

<div class="section-title">1. Master Station-by-Station Directory</div>

<table>
  <thead>
    <tr>
      <th style="width: 14%;">Round / Stage</th>
      <th style="width: 16%;">Location</th>
      <th style="width: 15%;">Challenge ID</th>
      <th style="width: 30%;">Clue / Question Description</th>
      <th style="width: 25%;">Correct Answer / Code</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Registration</strong></td>
      <td>Start Desk</td>
      <td>START</td>
      <td>Team Name, Player ID &amp; Mandatory Team Photo</td>
      <td><em>Participant Info + 📸 Photo</em></td>
    </tr>
    <tr>
      <td><strong>Round 1 (R1)</strong></td>
      <td>ADMIN</td>
      <td>Object / Gym</td>
      <td>Find the required object/code in Admin/Gym area</td>
      <td><span class="code-badge">JOHN-CENA</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Location Clue</strong></td>
      <td>Clue &rarr; MBA</td>
      <td>MBA-03</td>
      <td>
        <em>"One neighbour has already chosen its path.<br>
        The other is still preparing for the journey ahead.<br>
        Between the experienced and the yet-to-begin,<br>
        your next destination quietly stands."</em>
      </td>
      <td><strong>MBA</strong></td>
    </tr>
    <tr>
      <td><strong>MBA Mini Challenge</strong></td>
      <td>MBA Block</td>
      <td>MBA Quickfire<br>(Variants 1 &amp; 2)</td>
      <td>
        &bull; <strong>Variant 1 (Set A):</strong> 10 Business &amp; Finance Questions + 📸 Photo<br>
        &bull; <strong>Variant 2 (Set B):</strong> Variant 2 Challenge + 📸 Photo<br>
        <strong>Common Clearance Code:</strong>
      </td>
      <td><span class="code-badge">MBA-QF-1</span></td>
    </tr>
    <tr>
      <td><strong>R3 Location Clue</strong></td>
      <td>Clue &rarr; LIBRARY</td>
      <td>LIB-04</td>
      <td>
        Anagrams: VEHSELS, GIDRAEN, IFIW, SISCUDISNO GGUONLE, IGIDLAT BILRYRA<br>
        <em>(Decoded: SHELVES, READING, WIFI, DISCUSSION LOUNGE, DIGITAL LIBRARY)</em>
      </td>
      <td><strong>LIBRARY</strong></td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>LIBRARY</td>
      <td>QR Hunt</td>
      <td>Physically search library for hidden QR code and scan barcode</td>
      <td><span class="code-badge">DUMB_FAKE</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>R4 Location Clue</strong></td>
      <td>In-App Image (<code>aiml.jpeg</code>)</td>
      <td>Word Search</td>
      <td>
        Find ANY 5 valid words from puzzle grid:<br>
        • MACHINE LEARNING<br>
        • NEURAL NETWORK<br>
        • PYTHON<br>
        • DATASET<br>
        • ALGORITHM<br>
        • TRAINING DATA<br>
        • DEEP LEARNING<br>
        • DATA MINING<br>
        • MODELING<br>
        • REGRESSION<br>
        • CLASSIFY
      </td>
      <td><strong>5 Valid Words</strong> &rarr; Destination: <strong>AIML</strong></td>
    </tr>
    <tr>
      <td><strong>Round 4 (R4)</strong></td>
      <td>AIML &amp; Cyber Pathway</td>
      <td>Physical Challenge</td>
      <td>Complete physical agility challenge along pathway with volunteer</td>
      <td><span class="code-badge">PHY-CY</span> + 📸 Photo</td>
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
      <td><strong>STAGE</strong> + 📸 Solved Puzzle Photo<br>Code: <span class="code-badge">FINAL-PATH2</span></td>
    </tr>
    <tr>
      <td><strong>Completion</strong></td>
      <td>AUDI Stage</td>
      <td>FINISH</td>
      <td>Confirmation Screen</td>
      <td><strong>CONGRATULATIONS! PATH 2 COMPLETED.</strong></td>
    </tr>
  </tbody>
</table>

<div class="section-title">2. MBA-01 Quickfire Set A Question &amp; Answer Breakdown</div>

<table>
  <thead>
    <tr>
      <th style="width: 8%;">#</th>
      <th style="width: 62%;">Question Text</th>
      <th style="width: 30%;">Exact Expected Answer</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>1</td><td>A market with only one seller is called?</td><td><strong>MONOPOLY</strong></td></tr>
    <tr><td>2</td><td>What is the currency of China?</td><td><strong>YUAN</strong></td></tr>
    <tr><td>3</td><td>What does CEO stand for?</td><td><strong>CHIEF EXECUTIVE OFFICER</strong></td></tr>
    <tr><td>4</td><td>What does ROI stand for in business finance?</td><td><strong>RETURN ON INVESTMENT</strong></td></tr>
    <tr><td>5</td><td>What term describes business transactions conducted between two companies (abbreviation)?</td><td><strong>B2B</strong></td></tr>
    <tr><td>6</td><td>In accounting: Assets minus Liabilities equals what?</td><td><strong>EQUITY</strong></td></tr>
    <tr><td>7</td><td>Which animal represents a rising, optimistic financial market?</td><td><strong>BULL</strong></td></tr>
    <tr><td>8</td><td>What does IPO stand for when a company goes public?</td><td><strong>INITIAL PUBLIC OFFERING</strong></td></tr>
    <tr><td>9</td><td>What is the standard 3-letter abbreviation for Gross Domestic Product?</td><td><strong>GDP</strong></td></tr>
    <tr><td>10</td><td>In the 4 Ps of Marketing (Product, Price, Place), what is the 4th P?</td><td><strong>PROMOTION</strong></td></tr>
  </tbody>
</table>

<div class="section-title">3. Volunteer Station Instructions &amp; Verification Procedures</div>

<div class="station-card">
  <h3><span>ADMIN Station: Object Finding</span><span class="code-badge">CODE: JOHN-CENA</span></h3>
  <div><strong>Volunteer Instructions:</strong> Participants search the area for the assigned hidden item. Once found, they upload a photo and enter code <code>JOHN-CENA</code> into their KoboCollect form in capital letters.</div>
</div>

<div class="station-card">
  <h3><span>MBA Station: Quickfire Challenges</span><span class="code-badge">CLEARANCE CODE: MBA-QF-1</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams choose their assigned variant (Set A or Set B), answer their questions, and upload a station photo. Once verified, provide the final clearance code <code>MBA-QF-1</code> regardless of the variant chosen.</div>
</div>

<div class="station-card">
  <h3><span>LIBRARY Station: QR Hunt</span><span class="code-badge">SCANNED QR: DUMB_FAKE</span></h3>
  <div><strong>Volunteer Instructions:</strong> Ensure the physical QR code with payload <code>DUMB_FAKE</code> is placed in the reading/shelves section. Participants scan it using the in-app barcode question and upload a location photo.</div>
</div>

<div class="station-card">
  <h3><span>AIML &amp; Cyber Pathway Station: Physical Challenge</span><span class="code-badge">VOLUNTEER CODE: PHY-CY</span></h3>
  <div><strong>Volunteer Instructions:</strong> Supervise participants completing the physical trial along the pathway. Upon successful completion and photo upload, provide the verification code <code>PHY-CY</code>.</div>
</div>

<div class="station-card">
  <h3><span>Main Auditorium Stage: Grand Finale</span><span class="code-badge">CODE: FINAL-PATH2</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams solve Riddle 1 (<code>AUDITORIUM</code>) and Riddle 2 (<code>STAGE</code>), upload a clear photo of their solved puzzle sheet, and present to Chief Judges for code <code>FINAL-PATH2</code>!</div>
</div>

</body>
</html>`;
}

function main() {
  const baseDir = path.resolve(__dirname, '..');
  const route2Dir = path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI');
  if (!fs.existsSync(route2Dir)) fs.mkdirSync(route2Dir, { recursive: true });

  const mediaDir = path.join(route2Dir, 'media');
  if (!fs.existsSync(mediaDir)) fs.mkdirSync(mediaDir, { recursive: true });

  // If aiml.jpeg exists in route2Dir, copy to mediaDir
  const sourceImg = path.join(route2Dir, 'aiml.jpeg');
  const destImg = path.join(mediaDir, 'aiml.jpeg');
  if (fs.existsSync(sourceImg)) {
    fs.copyFileSync(sourceImg, destImg);
  }

  console.log('Generating PATH2_FINAL_ODK.xlsx (Updated with MBA-QF-1 clearance & image on every word search question)...');
  const survey = buildSurvey();
  const choices = buildChoices();
  const settings = buildSettings();

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(survey), 'survey');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(choices), 'choices');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settings), 'settings');

  const xlsxPath = path.join(route2Dir, 'PATH2_FINAL_ODK.xlsx');
  XLSX.writeFile(wb, xlsxPath);
  console.log(`Successfully created: ${xlsxPath}`);

  console.log('Generating PATH2_ANSWER_KEY.pdf...');
  const pdfHtml = buildAnswerKeyPdfHtml();
  const pdfPath = path.join(route2Dir, 'PATH2_ANSWER_KEY.pdf');
  renderHtmlToPdf(pdfHtml, pdfPath);
  console.log(`Successfully created: ${pdfPath}`);

  console.log('Packaging PATH2_COMPLETE_PACKAGE.zip and PATH2_MEDIA.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route2Dir}\\PATH2_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route2Dir}\\PATH2_FINAL_ODK.xlsx', '${route2Dir}\\PATH2_ANSWER_KEY.pdf', '${mediaDir}' -DestinationPath '${route2Dir}\\PATH2_COMPLETE_PACKAGE.zip' -Force"`);

  console.log('🎉 PATH 2 PACKAGE GENERATED SUCCESSFULLY!');
}

main();

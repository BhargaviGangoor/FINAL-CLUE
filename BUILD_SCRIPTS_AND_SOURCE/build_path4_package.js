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
    label: '🏆 FINAL CLUE — PATH 4\n\n### 🏴 TREASURE HUNT — RULES\n\n1. *No phones allowed* during the hunt. Keep them switched off and safely stored.\n2. *Stay with your team* and follow the designated route/instructions.\n3. *No cheating or interference* with other teams, clues, props, or college property.\n4. Complete all challenges *as instructed by the coordinators*. Safety comes first.\n5. Any violation of the rules may result in a *time penalty or disqualification*. The coordinator’s decision will be final.\n\n⚠️ IMPORTANT:\n• All questions, code entries, and photo uploads are MANDATORY.\n• All text answers and codes must be entered in UPPERCASE ONLY.',
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
  // R1 — OLD CANTEEN (Object: cat board)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r1_canteen_group',
    label: 'ROUND 1 — OBJECT FINDING',
    relevant: startPassed
  });

  survey.push({
    type: 'note',
    name: 'r1_canteen_intro',
    label: '📍 ROUND 1: 🧩 OBJECT FINDING\n\nChallenge: Locate the assigned hidden object/code at this station.\n\nInstructions:\n1. Search the location to find the assigned hidden object.\n2. Take a clear photo of the object.\n3. Show the photo to the station volunteer to receive your verification code.\n\nEnter code in caps',
    hint: 'Find the assigned hidden object.'
  });

  survey.push({
    type: 'image',
    name: 'r1_canteen_photo',
    label: '📸 Upload Photo of Discovered Hidden Object\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered object.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'r1_canteen_code',
    label: 'Enter Volunteer Verification Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MEOW-BOW-4'",
    constraint_message: '❌ Incorrect code. Enter MEOW-BOW-4 in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r1Passed = `${startPassed} and translate(normalize-space(\${r1_canteen_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MEOW-BOW-4' and \${r1_canteen_photo} != ''`;

  // =============================================================
  // R2 — LOCATION CLUE (CSE-02 — Emoji Math Riddle)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r2_math_group',
    label: 'ROUND 2 LOCATION CLUE',
    relevant: r1Passed
  });

  survey.push({
    type: 'note',
    name: 'r2_math_note',
    label: '📍 ROUND 2 LOCATION CLUE: EMOJI MATH RIDDLE\n\nSolve the emoji equations to calculate the destination number pointing to your next block on the station campus map!\n\n🍎 + 🍎 = 10\n🍎 + 🍌 = 7\n🍌 + 🍇 = 6\n\nFinal Equation:\n🍇 + 🍎 = ?\n\nEnter code in caps',
    hint: 'Solve the equations for the numeric value.'
  });

  survey.push({
    type: 'text',
    name: 'r2_math_answer',
    label: 'Enter the calculated value for [ 🍇 + 🍎 ]:\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '9' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '9' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'NINE'",
    constraint_message: '❌ Incorrect answer. Solve the equations to find the numeric answer.'
  });

  const r2MathPassed = `${r1Passed} and (normalize-space(\${r2_math_answer}) = '9' or translate(normalize-space(\${r2_math_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '9' or translate(normalize-space(\${r2_math_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'NINE')`;

  survey.push({
    type: 'text',
    name: 'r2_cse_loc_answer',
    label: 'Which campus block does this number direct your team to on the map?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    relevant: r2MathPassed,
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE BLOCK'",
    constraint_message: '❌ Incorrect destination. Enter the campus block name in CAPS.'
  });

  survey.push({
    type: 'end_group'
  });

  const r2CseLocPassed = `${r2MathPassed} and (translate(normalize-space(\${r2_cse_loc_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE' or translate(normalize-space(\${r2_cse_loc_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE BLOCK')`;

  // =============================================================
  // CSE START CODE (CSE-START)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'cse_start_group',
    label: 'CSE ARRIVAL',
    relevant: r2CseLocPassed
  });

  survey.push({
    type: 'note',
    name: 'cse_arrival_note',
    label: '🏃 PROCEED TO CSE BLOCK\n\nReport immediately to the CSE station and meet the volunteer to receive your CSE start code.\n\nEnter code in caps',
    hint: 'Meet volunteer at CSE station.'
  });

  survey.push({
    type: 'text',
    name: 'r2_cse_start_code',
    label: 'Enter CSE Volunteer Start Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE-START'",
    constraint_message: '❌ Incorrect start code. Enter CSE-START in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const cseStartPassed = `${r2CseLocPassed} and translate(normalize-space(\${r2_cse_start_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE-START'`;

  // =============================================================
  // CSE CHALLENGES (Variant 1 vs Variant 2)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'cse_challenge_select_group',
    label: 'ROUND 2 — CSE MINI CHALLENGES',
    relevant: cseStartPassed
  });

  survey.push({
    type: 'select_one cse_variant_list',
    name: 'r2_cse_variant_select',
    label: 'Select Assigned Challenge Variant\nMandatory selection',
    hint: 'Choose the variant assigned by the station volunteer.',
    required: 'yes'
  });

  // -------------------------------------------------------------
  // VARIANT 1: ALGO RELAY + BINARY BREAKER
  // -------------------------------------------------------------
  survey.push({
    type: 'begin_group',
    name: 'cse_v1_container',
    label: 'VARIANT 1 — ALGO RELAY & BINARY BREAKER',
    relevant: "${r2_cse_variant_select} = 'var1'"
  });

  // V1 Part 1: Algo Relay (10 Tasks)
  survey.push({
    type: 'begin_group',
    name: 'cse_v1_algo_group',
    label: 'CHALLENGE 1: 🏃 ALGO RELAY (10 TASKS)'
  });

  survey.push({
    type: 'note',
    name: 'cse_algo_intro',
    label: '🏃 ALGO RELAY — 10 TASKS\n\nArrange the algorithmic steps correctly for each scenario below. Watch out for traps and fake steps!\n\nEnter code in caps',
    hint: 'Complete each algorithmic arrangement in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'algo_q1',
    label: 'Q1. Make Tea\nArrange steps: Add tea powder, Boil water, Add milk, Add sugar, Strain, Serve\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q2',
    label: 'Q2. ATM Withdrawal\nArrange: Enter PIN, Insert card, Select withdrawal, Enter amount, Collect cash, Remove card (Trap: Collect cash before amount)\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q3',
    label: 'Q3. Find the Largest Number\nGiven 5 numbers (17, 42, 9, 63, 31), write the logic to find the largest:\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q4',
    label: 'Q4. Login System\nArrange logic: Enter username → Enter password → Check credentials → Correct? → Allow/Deny access\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q5',
    label: 'Q5. Find a Name\nDesign the steps to find whether "RAHUL" exists in a list of 100 student names:\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q6',
    label: 'Q6. Odd or Even\nCreate an algorithm that takes an input number and determines EVEN / ODD:\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q7',
    label: 'Q7. Sort Books\nArrange the algorithm to sort 6 books of different heights from shortest → tallest:\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q8',
    label: 'Q8. Traffic Light Error Detection\nIdentify the fake condition among: RED→STOP, YELLOW→WAIT, GREEN→GO, BLUE→GO\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q9',
    label: 'Q9. Password Attempts\nWrite the algorithm for 3 attempts (Correct→Login, Wrong→Retry, 3 Wrong→Locked):\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'algo_q10',
    label: 'Q10. Final Boss: Second-Largest Number\nWrite an algorithm to find the second-largest number among 10 numbers without sorting:\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'image',
    name: 'r2_v1_algo_photo',
    label: '📸 Upload Photo of Solved Algo Relay Cards\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your arranged algorithm cards.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  // V1 Part 2: Binary Breaker (10 Questions)
  survey.push({
    type: 'begin_group',
    name: 'cse_v1_binary_group',
    label: 'CHALLENGE 2: 🔢 BINARY BREAKER (10 QUESTIONS)'
  });

  survey.push({
    type: 'note',
    name: 'cse_binary_intro',
    label: '🔢 BINARY BREAKER — 10 QUESTIONS\n\nKey: A=1, B=2, C=3 ... Z=26 (5-bit binary representations).\nSolve all 10 binary challenges below!\n\nEnter code in caps',
    hint: 'Answer all 10 questions in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q1',
    label: 'Q1. Decode: 00010 00101 00111\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'BEG'",
    constraint_message: '❌ Incorrect. Decode using A=1 to Z=26.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q2',
    label: 'Q2. Decode: 01000 00001 00110 00110\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'HALL'",
    constraint_message: '❌ Incorrect. Decode using A=1 to Z=26.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q3',
    label: 'Q3. Encode "CODE" into 5-bit binary (A=1 to Z=26)\nEnter binary string (e.g. 00011 01111 00100 00101)',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), ' ', '') = '00011011110010000101'",
    constraint_message: '❌ Incorrect. Encode C=3, O=15, D=4, E=5 in 5-bit binary.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q4',
    label: 'Q4. Decimal to Binary: Convert 13 to binary\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '1101'",
    constraint_message: '❌ Incorrect. Convert 13 to binary.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q5',
    label: 'Q5. Binary to Decimal: Convert 10110 to decimal\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '22'",
    constraint_message: '❌ Incorrect. Convert 10110 to decimal.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q6',
    label: 'Q6. Binary Addition: 1011 + 0110 = ?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '10001'",
    constraint_message: '❌ Incorrect. Perform binary addition.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q7',
    label: 'Q7. Find the Odd One Out among: 00001, 00010, 00011, 00101, 00110\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '00101'",
    constraint_message: '❌ Incorrect. Look at decimal sequence 1, 2, 3, 4, 5.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q8',
    label: 'Q8. ASCII Challenge: Decode 01001000 01001001 (ASCII table)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'HI'",
    constraint_message: '❌ Incorrect. Decode ASCII bytes 72 and 73.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q9',
    label: 'Q9. Binary Location: Decode 01000 00101 00001 01100 01100\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'HEALL' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'HELL'",
    constraint_message: '❌ Incorrect. Decode using A=1 to Z=26.'
  });

  survey.push({
    type: 'text',
    name: 'binary_q10',
    label: 'Q10. Final Binary Message: Decode 01000 00001 00110 00100 01001 01010 (A=1..26)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'HAFDIJ' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'IBGEJK'",
    constraint_message: '❌ Incorrect. Decode the 6 letters using A=1 to Z=26.'
  });

  survey.push({
    type: 'image',
    name: 'r2_v1_binary_photo',
    label: '📸 Upload Photo of Team Solving Binary Breaker\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team solving Binary Breaker.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  survey.push({
    type: 'text',
    name: 'r2_v1_code',
    label: 'Enter Volunteer Clearance Code (Variant 1)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE-EINSTEIN-4'",
    constraint_message: '❌ Incorrect code. Enter CSE-EINSTEIN-4 provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  // -------------------------------------------------------------
  // VARIANT 2: PATTERN HACK + BUG HUNTER
  // -------------------------------------------------------------
  survey.push({
    type: 'begin_group',
    name: 'cse_v2_container',
    label: 'VARIANT 2 — PATTERN HACK & BUG HUNTER',
    relevant: "${r2_cse_variant_select} = 'var2'"
  });

  // V2 Part 1: Pattern Hack (10 Questions)
  survey.push({
    type: 'begin_group',
    name: 'cse_v2_pattern_group',
    label: 'CHALLENGE 1: 🧠 PATTERN HACK (10 QUESTIONS)'
  });

  survey.push({
    type: 'note',
    name: 'cse_pattern_intro',
    label: '🧠 PATTERN HACK — 10 QUESTIONS\n\nFind the missing number, letter, or rule for each sequence below!\n\nEnter code in caps',
    hint: 'Answer all 10 questions in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q1',
    label: 'Q1. Find next: 2, 4, 6, 8, ?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '10'",
    constraint_message: '❌ Incorrect answer.'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q2',
    label: 'Q2. Find next: 3, 6, 12, 24, ?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '48'",
    constraint_message: '❌ Incorrect answer.'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q3',
    label: 'Q3. Find next: 1, 4, 9, 16, 25, ?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '36'",
    constraint_message: '❌ Incorrect answer.'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q4',
    label: 'Q4. Find next: 1, 1, 2, 3, 5, 8, ?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '13'",
    constraint_message: '❌ Incorrect answer.'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q5',
    label: 'Q5. Find next: A, C, F, J, O, ?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'U'",
    constraint_message: '❌ Incorrect answer. Follow letter positions +2, +3, +4, +5, +6.'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q6',
    label: 'Q6. Rule: 2→6, 3→12, 4→20, 5→30, 6→?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '42'",
    constraint_message: '❌ Incorrect answer. Follow n * (n+1).'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q7',
    label: 'Q7. What comes next: 1A, 2B, 4D, 8H, 16P, ?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '32F'",
    constraint_message: '❌ Incorrect answer. Numbers double, letters double modulo 26.'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q8',
    label: 'Q8. Find missing: 2, 3, 5, 9, 17, ?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '33'",
    constraint_message: '❌ Incorrect answer. Rule: *2 - 1.'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q9',
    label: 'Q9. Visual Logic: If ★=10, ▲=3, ■=4. Calculate [ ★ + ▲ × ■ ]:\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = '22'",
    constraint_message: '❌ Incorrect answer. Remember operator precedence (* before +).'
  });

  survey.push({
    type: 'text',
    name: 'pattern_q10',
    label: 'Q10. What does the series 1, 2, 4, 8, 16 represent?\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'POWERS OF 2' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'POWER OF 2' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'POWERS OF TWO' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'BINARY'",
    constraint_message: '❌ Incorrect answer. Think about powers in computing.'
  });

  survey.push({
    type: 'image',
    name: 'r2_v2_pattern_photo',
    label: '📸 Upload Photo of Team Solving Pattern Hack\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team solving Pattern Hack.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  // V2 Part 2: Bug Hunter (9/10 Tasks)
  survey.push({
    type: 'begin_group',
    name: 'cse_v2_bug_group',
    label: 'CHALLENGE 2: 🐛 BUG HUNTER (DEBUG THE SYSTEM)'
  });

  survey.push({
    type: 'note',
    name: 'cse_bug_intro',
    label: '🐛 BUG HUNTER — DEBUG THE SYSTEM\n\nExamine the programs, pseudocode, and flowcharts below. Find the logic and syntax bugs!\n\nEnter code in caps',
    hint: 'Identify and describe the bugs in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'bug_q1',
    label: 'Q1. Output Bug: x=10, y=20. IF x>y PRINT "10 is greater" ELSE PRINT "20 is greater". What is wrong?\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'bug_q2',
    label: 'Q2. Infinite Loop: x=1. WHILE x<=5 PRINT x. What is the bug?\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'bug_q3',
    label: 'Q3. Swapped Variables: a=10, b=20. a=b; b=a. After execution, what are a and b?\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'bug_q4',
    label: 'Q4. Wrong Output: FOR i=1 TO 5 PRINT i. System displays: 1 2 3 4 6. Where is the bug?\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'bug_q5',
    label: 'Q5. Password Bug: IF password="CSE123" PRINT "Access Granted" ELSE PRINT "Access Granted". Find the bug:\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'bug_q6',
    label: 'Q6. Array Bug: [10, 20, 30, 40, 50]. Program says PRINT array[5]. What is wrong?\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'bug_q7',
    label: 'Q7. Flowchart Bug: Start → Input number → number % 2 == 0? YES→ODD, NO→EVEN. Find the bug:\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'bug_q8',
    label: 'Q8. Detective Debugging: Input: 5. x=1; FOR i=1 TO x { x=x+1 }; PRINT x. Why doesn\'t this behave as expected?\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'bug_q9',
    label: 'Q9. Find TWO Bugs: a=5, b=10. IF a>b PRINT "b is larger" ELSE PRINT "a is larger". What are the 2 bugs?\nEnter in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes'
  });

  survey.push({
    type: 'image',
    name: 'r2_v2_bug_photo',
    label: '📸 Upload Photo of Team Solving Bug Hunter\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team solving Bug Hunter.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_code',
    label: 'Enter Volunteer Clearance Code (Variant 2)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE-ZUCKERBERG-4' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE-ZUCKER-4'",
    constraint_message: '❌ Incorrect code. Enter CSE-ZUCKERBERG-4 provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  survey.push({
    type: 'end_group'
  });

  const cseChallengesPassed = `${cseStartPassed} and ((${'\${r2_cse_variant_select}'} = 'var1' and translate(normalize-space(\${r2_v1_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE-EINSTEIN-4' and \${r2_v1_algo_photo} != '' and \${r2_v1_binary_photo} != '') or (${'\${r2_cse_variant_select}'} = 'var2' and (translate(normalize-space(\${r2_v2_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE-ZUCKERBERG-4' or translate(normalize-space(\${r2_v2_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CSE-ZUCKER-4') and \${r2_v2_pattern_photo} != '' and \${r2_v2_bug_photo} != ''))`;

  // =============================================================
  // R3 — CYBER LOCATION CLUE (MORSE CODE with morse.jpeg)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_cyber_loc_group',
    label: 'ROUND 3 LOCATION CLUE',
    relevant: cseChallengesPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_morse_note',
    label: '📍 ROUND 3 LOCATION CLUE: MORSE CODE\n\nExamine the Morse code chart in the image below and decode the encrypted message to reveal your next destination!\n\n📜 CIPHER TEXT:\n--.  ---\n-  ---\n-.-.  -.--  -...  .  .-.\n\nEnter code in caps',
    hint: 'Decode the Morse code message using the attached image chart.',
    'media::image': 'morse.jpeg'
  });

  survey.push({
    type: 'text',
    name: 'r3_morse_answer',
    label: 'Enter decoded message:\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    'media::image': 'morse.jpeg',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'GO TO CYBER' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CYBER'",
    constraint_message: '❌ Incorrect decoded message. Decode the Morse code and enter in CAPS.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3MorsePassed = `${cseChallengesPassed} and (translate(normalize-space(\${r3_morse_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'GO TO CYBER' or translate(normalize-space(\${r3_morse_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CYBER')`;

  // =============================================================
  // CYBER ARRIVAL CODE (CYBER-THIEF)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'cyber_arrival_group',
    label: 'CYBER ARRIVAL',
    relevant: r3MorsePassed
  });

  survey.push({
    type: 'note',
    name: 'cyber_arrival_note',
    label: '🏃 PROCEED TO CYBER BLOCK\n\nReport immediately to the Cyber station and meet the volunteer to receive your arrival verification code.\n\nEnter code in caps',
    hint: 'Meet volunteer at Cyber station.'
  });

  survey.push({
    type: 'text',
    name: 'r3_cyber_start_code',
    label: 'Enter Cyber Volunteer Start Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CYBER-THIEF'",
    constraint_message: '❌ Incorrect code. Enter CYBER-THIEF in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const cyberArrivalPassed = `${r3MorsePassed} and translate(normalize-space(\${r3_cyber_start_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'CYBER-THIEF'`;

  // =============================================================
  // R3 — CYBER QR HUNT (JAMES_BOND)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_cy_qr_group',
    label: 'ROUND 3 — QR HUNT',
    relevant: cyberArrivalPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_cy_qr_note',
    label: '📍 ROUND 3: QR HUNT\n\nSearch the location physically to locate the hidden QR code.\nScan the QR code using the scanner below.\n\nEnter code in caps',
    hint: 'Locate and scan the hidden QR code.'
  });

  survey.push({
    type: 'barcode',
    name: 'r3_cy_qr_scan',
    label: 'Scan Discovered QR Code',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "normalize-space(.) = 'JAMES_BOND' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'JAMES_BOND'",
    constraint_message: '❌ Incorrect QR code scanned. Search for the correct QR code at this station.'
  });

  survey.push({
    type: 'image',
    name: 'r3_cy_qr_photo',
    label: '📸 Upload Photo of Discovered QR Code / Station\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered QR code location.',
    required: 'yes'
  });

  survey.push({
    type: 'end_group'
  });

  const r3QrPassed = `${cyberArrivalPassed} and (normalize-space(\${r3_cy_qr_scan}) = 'JAMES_BOND' or translate(normalize-space(\${r3_cy_qr_scan}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'JAMES_BOND') and \${r3_cy_qr_photo} != ''`;

  // =============================================================
  // R4 — LOCATION CLUE (MBA-05 — Wisdom, Below & Beyond)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_mba_loc_group',
    label: 'ROUND 4 LOCATION CLUE',
    relevant: r3QrPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_mba_loc_note',
    label: '📍 ROUND 4 LOCATION CLUE: THREE STEPS\n\nFollow the three clues to identify the next campus destination:\n1. The one who gives wisdom\n2. Look below\n3. Where students prepare for their next career step\n\nEnter code in caps',
    hint: 'Follow the 3 clues to deduce the destination block.'
  });

  survey.push({
    type: 'text',
    name: 'r4_mba_loc_answer',
    label: 'Enter destination location\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA BLOCK'",
    constraint_message: '❌ Incorrect destination. Follow the clues and enter the block in CAPS.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4MbaLocPassed = `${r3QrPassed} and (translate(normalize-space(\${r4_mba_loc_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA' or translate(normalize-space(\${r4_mba_loc_answer}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MBA BLOCK')`;

  // =============================================================
  // MBA ARRIVAL CODE (MONEY-BROTHA)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'mba_arrival_group',
    label: 'MBA ARRIVAL',
    relevant: r4MbaLocPassed
  });

  survey.push({
    type: 'note',
    name: 'mba_arrival_note',
    label: '🏃 PROCEED TO MBA BLOCK\n\nReport immediately to the MBA station and meet the volunteer to receive your arrival verification code.\n\nEnter code in caps',
    hint: 'Meet volunteer at MBA station.'
  });

  survey.push({
    type: 'text',
    name: 'r4_mba_start_code',
    label: 'Enter MBA Volunteer Arrival Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MONEY-BROTHA'",
    constraint_message: '❌ Incorrect code. Enter MONEY-BROTHA in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const mbaArrivalPassed = `${r4MbaLocPassed} and translate(normalize-space(\${r4_mba_start_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'MONEY-BROTHA'`;

  // =============================================================
  // R4 — MBA PHYSICAL CHALLENGE (RAVI-KISHEN)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_mba_phy_group',
    label: 'ROUND 4 — PHYSICAL CHALLENGE',
    relevant: mbaArrivalPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_mba_phy_note',
    label: '📍 ROUND 4: PHYSICAL CHALLENGE (COLOURED CUPS)\n\nReport to the MBA station and complete the 7 coloured paper cups physical coordination challenge under volunteer supervision.\n\nEnter code in caps',
    hint: 'Complete physical challenge with volunteer.'
  });

  survey.push({
    type: 'image',
    name: 'r4_mba_phy_photo',
    label: '📸 Upload Photo of Physical Challenge Completion\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team completing the physical challenge.',
    required: 'yes'
  });

  survey.push({
    type: 'text',
    name: 'r4_mba_phy_code',
    label: 'Enter Volunteer Clearance Code\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'RAVI-KISHEN'",
    constraint_message: '❌ Incorrect code. Enter RAVI-KISHEN in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4PhyPassed = `${mbaArrivalPassed} and translate(normalize-space(\${r4_mba_phy_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'RAVI-KISHEN' and \${r4_mba_phy_photo} != ''`;

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
    constraint: "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH4' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH1' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH2' or translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH3'",
    constraint_message: '❌ Incorrect code. Enter the code in CAPS provided by the Chief Judge.'
  });

  survey.push({
    type: 'note',
    name: 'final_congratulations_screen',
    label: '🎉 CONGRATULATIONS! PATH 4 COMPLETED.\n\n🏆 You have successfully conquered every challenge, puzzle, and cipher on PATH 4!\n\nShow this completion screen immediately to the Chief Judge to lock in your finishing timestamp and rank!',
    hint: 'Report to Chief Judge to finalize completion.',
    relevant: `${finalRiddle2Passed} and (translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH4' or translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH1' or translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH2' or translate(normalize-space(\${final_stage_volunteer_code}), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = 'FINAL-PATH3') and \${final_solved_puzzle_photo} != ''`
  });

  survey.push({
    type: 'end_group'
  });

  return survey;
}

function buildChoices() {
  return [
    {
      list_name: 'cse_variant_list',
      name: 'var1',
      label: 'Variant 1: Algo Relay + Binary Breaker'
    },
    {
      list_name: 'cse_variant_list',
      name: 'var2',
      label: 'Variant 2: Pattern Hack + Bug Hunter'
    }
  ];
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — PATH 4',
      form_id: 'final_clue_path4',
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
<title>PATH 4 — ORGANIZER MASTER ANSWER KEY</title>
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
    border-bottom: 3px solid #b91c1c;
    padding-bottom: 6px;
    margin-bottom: 12px;
  }
  .header h1 {
    color: #991b1b;
    margin: 0 0 3px 0;
    font-size: 15pt;
    letter-spacing: 0.4px;
  }
  .header .badge {
    display: inline-block;
    background: #b91c1c;
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
    background: #fef2f2;
    padding: 6px 8px;
    border-radius: 4px;
    border-left: 3px solid #ef4444;
  }
  .meta-card strong {
    color: #991b1b;
    display: block;
    font-size: 7.5pt;
    text-transform: uppercase;
  }
  .meta-card span {
    font-size: 8.5pt;
    font-weight: 600;
  }
  .section-title {
    background: #991b1b;
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
    color: #991b1b;
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
    <small style="color: #64748b;">PATH 4 &bull; FOR ORGANIZER USE ONLY</small>
  </div>
  <h1>PATH 4 &bull; TREASURE HUNT MASTER ANSWER KEY</h1>
  <div style="font-size: 9pt; color: #475569; font-weight: 600;">
    CONFIDENTIAL ORGANIZER KEY &bull; PROGRESSION: OLD CANTEEN &rarr; CSE &rarr; CYBER &rarr; MBA &rarr; AUDI
  </div>
</div>

<div class="meta-grid">
  <div class="meta-card">
    <strong>Starting Station</strong>
    <span>Old Canteen</span>
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
    <span>morse.jpeg (Morse Code Key)</span>
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
      <td>OLD CANTEEN</td>
      <td>🧩 Object (cat board)</td>
      <td>Initial object finding at Old Canteen</td>
      <td><span class="code-badge">MEOW-BOW-4</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Clue</strong></td>
      <td>Clue &rarr; CSE</td>
      <td>CSE-02 Emoji Math Riddle</td>
      <td>🍎+🍎=10, 🍎+🍌=7, 🍌+🍇=6 &rarr; 🍇+🍎=?</td>
      <td><strong>9</strong> &rarr; <strong>CSE</strong></td>
    </tr>
    <tr>
      <td><strong>CSE Start</strong></td>
      <td>CSE</td>
      <td>Start Code</td>
      <td>Arrival verification at CSE block</td>
      <td><span class="code-badge">CSE-START</span></td>
    </tr>
    <tr>
      <td><strong>R2 Variant 1</strong></td>
      <td>CSE</td>
      <td>Algo Relay + Binary Breaker</td>
      <td>10 Algo Tasks + 10 Binary Questions + 📸 Photos</td>
      <td><span class="code-badge">CSE-EINSTEIN-4</span></td>
    </tr>
    <tr>
      <td><strong>R2 Variant 2</strong></td>
      <td>CSE</td>
      <td>Pattern Hack + Bug Hunter</td>
      <td>10 Pattern Questions + 9 Debug Tasks + 📸 Photos</td>
      <td><span class="code-badge">CSE-ZUCKERBERG-4</span></td>
    </tr>
    <tr>
      <td><strong>R3 Clue</strong></td>
      <td>In-App (<code>morse.jpeg</code>)</td>
      <td>CYB-01 Morse Code</td>
      <td><code>--. ---</code> / <code>- ---</code> / <code>-.-. -.-- -... . .-.</code></td>
      <td><strong>GO TO CYBER</strong></td>
    </tr>
    <tr>
      <td><strong>Cyber Start</strong></td>
      <td>CYBER</td>
      <td>Start Code</td>
      <td>Arrival verification at Cyber block</td>
      <td><span class="code-badge">CYBER-THIEF</span></td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>CYBER</td>
      <td>QR Hunt</td>
      <td>Physical QR code scan at Cyber station</td>
      <td><span class="code-badge">JAMES_BOND</span> + 📸 Photo</td>
    </tr>
    <tr>
      <td><strong>R4 Clue</strong></td>
      <td>Clue &rarr; MBA</td>
      <td>MBA-05 Wisdom, Below &amp; Beyond</td>
      <td>Wisdom giver + Look below + Career step</td>
      <td><strong>MBA</strong></td>
    </tr>
    <tr>
      <td><strong>MBA Start</strong></td>
      <td>MBA</td>
      <td>Start Code</td>
      <td>Arrival verification at MBA block</td>
      <td><span class="code-badge">MONEY-BROTHA</span></td>
    </tr>
    <tr>
      <td><strong>R4 Physical</strong></td>
      <td>MBA</td>
      <td>Physical (7 Coloured Cups)</td>
      <td>Coloured cups physical coordination trial</td>
      <td><span class="code-badge">RAVI-KISHEN</span> + 📸 Photo</td>
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
      <td><strong>STAGE</strong> + 📸 Solved Puzzle Photo<br>Code: <span class="code-badge">FINAL-PATH4</span></td>
    </tr>
    <tr>
      <td><strong>Completion</strong></td>
      <td>AUDI Stage</td>
      <td>FINISH</td>
      <td>Confirmation Screen</td>
      <td><strong>CONGRATULATIONS! PATH 4 COMPLETED.</strong></td>
    </tr>
  </tbody>
</table>

<div class="section-title">2. Binary Breaker &amp; Pattern Hack Master Keys</div>

<table>
  <thead>
    <tr>
      <th style="width: 50%;">Binary Breaker (Variant 1)</th>
      <th style="width: 50%;">Pattern Hack (Variant 2)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>
        <strong>Q1:</strong> <code>00010 00101 00111</code> &rarr; <strong>BEG</strong><br>
        <strong>Q2:</strong> <code>01000 00001 00110 00110</code> &rarr; <strong>HALL</strong><br>
        <strong>Q3:</strong> Encode "CODE" &rarr; <strong>00011 01111 00100 00101</strong><br>
        <strong>Q4:</strong> Decimal 13 to Binary &rarr; <strong>1101</strong><br>
        <strong>Q5:</strong> Binary 10110 to Decimal &rarr; <strong>22</strong><br>
        <strong>Q6:</strong> 1011 + 0110 &rarr; <strong>10001</strong><br>
        <strong>Q7:</strong> Odd one out &rarr; <strong>00101</strong><br>
        <strong>Q8:</strong> ASCII bytes 72 73 &rarr; <strong>HI</strong><br>
        <strong>Q9:</strong> <code>01000 00101 00001 01100 01100</code> &rarr; <strong>HEALL</strong><br>
        <strong>Q10:</strong> <code>01000 00001 00110 00100 01001 01010</code> &rarr; <strong>HAFDIJ</strong>
      </td>
      <td>
        <strong>Q1:</strong> 2, 4, 6, 8, ? &rarr; <strong>10</strong><br>
        <strong>Q2:</strong> 3, 6, 12, 24, ? &rarr; <strong>48</strong><br>
        <strong>Q3:</strong> 1, 4, 9, 16, 25, ? &rarr; <strong>36</strong><br>
        <strong>Q4:</strong> 1, 1, 2, 3, 5, 8, ? &rarr; <strong>13</strong><br>
        <strong>Q5:</strong> A, C, F, J, O, ? &rarr; <strong>U</strong> (+2, +3, +4, +5, +6)<br>
        <strong>Q6:</strong> 6 &rarr; <strong>42</strong> (n × (n+1))<br>
        <strong>Q7:</strong> 1A, 2B, 4D, 8H, 16P, ? &rarr; <strong>32F</strong><br>
        <strong>Q8:</strong> 2, 3, 5, 9, 17, ? &rarr; <strong>33</strong> (×2 − 1)<br>
        <strong>Q9:</strong> ★=10, ▲=3, ■=4 &rarr; 10 + 3×4 = <strong>22</strong><br>
        <strong>Q10:</strong> 1, 2, 4, 8, 16 &rarr; <strong>POWERS OF 2</strong>
      </td>
    </tr>
  </tbody>
</table>

<div class="section-title">3. Volunteer Station Instructions &amp; Codes</div>

<div class="station-card">
  <h3><span>CSE Station: Mini Challenges</span><span class="code-badge">START: CSE-START &bull; V1: CSE-EINSTEIN-4 &bull; V2: CSE-ZUCKERBERG-4</span></h3>
  <div>Verify arrival with <code>CSE-START</code>. Administer chosen variant (Variant 1: Algo Relay + Binary Breaker OR Variant 2: Pattern Hack + Bug Hunter). Provide respective clearance code upon verified completion.</div>
</div>

<div class="station-card">
  <h3><span>CYBER Station: QR Hunt (Anushree &amp; Koushik)</span><span class="code-badge">START: CYBER-THIEF &bull; SCANNED QR: JAMES_BOND</span></h3>
  <div>Verify arrival with <code>CYBER-THIEF</code>. Ensure QR code with payload <code>JAMES_BOND</code> is scanned and location photo uploaded.</div>
</div>

<div class="station-card">
  <h3><span>MBA Station: 7 Coloured Cups Challenge (Pruthvi &amp; Sanchith)</span><span class="code-badge">START: MONEY-BROTHA &bull; CLEARANCE: RAVI-KISHEN</span></h3>
  <div>Verify arrival with <code>MONEY-BROTHA</code>. Administer the 7 coloured paper cups coordination game. Provide clearance code <code>RAVI-KISHEN</code> after photo upload.</div>
</div>

<div class="station-card">
  <h3><span>Main Auditorium Stage: Grand Finale</span><span class="code-badge">CODE: FINAL-PATH4</span></h3>
  <div>Teams solve Riddle 1 (<code>AUDITORIUM</code>) and Riddle 2 (<code>STAGE</code>), upload a clear photo of their solved puzzle sheet, and present to Chief Judges for code <code>FINAL-PATH4</code>!</div>
</div>

</body>
</html>`;
}

function main() {
  const baseDir = path.resolve(__dirname, '..');
  const route4Dir = path.join(baseDir, 'ROUTE_4_PATH4_OLD_CANTEEN_CSE_CY_MBA_AUDI');
  if (!fs.existsSync(route4Dir)) fs.mkdirSync(route4Dir, { recursive: true });

  const mediaDir = path.join(route4Dir, 'media');
  if (!fs.existsSync(mediaDir)) fs.mkdirSync(mediaDir, { recursive: true });

  // If morse.jpeg exists in route4Dir, copy to mediaDir
  const sourceImg = path.join(route4Dir, 'morse.jpeg');
  const destImg = path.join(mediaDir, 'morse.jpeg');
  if (fs.existsSync(sourceImg)) {
    fs.copyFileSync(sourceImg, destImg);
  }

  console.log('Generating PATH4_FINAL_ODK.xlsx (Updated with morse.jpeg and clean puzzle prompts)...');
  const survey = buildSurvey();
  const choices = buildChoices();
  const settings = buildSettings();

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(survey), 'survey');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(choices), 'choices');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settings), 'settings');

  const xlsxPath = path.join(route4Dir, 'PATH4_FINAL_ODK.xlsx');
  XLSX.writeFile(wb, xlsxPath);
  console.log(`Successfully created: ${xlsxPath}`);

  console.log('Generating PATH4_ANSWER_KEY.pdf...');
  const pdfHtml = buildAnswerKeyPdfHtml();
  const pdfPath = path.join(route4Dir, 'PATH4_ANSWER_KEY.pdf');
  renderHtmlToPdf(pdfHtml, pdfPath);
  console.log(`Successfully created: ${pdfPath}`);

  console.log('Packaging PATH4_COMPLETE_PACKAGE.zip and PATH4_MEDIA.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route4Dir}\\PATH4_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route4Dir}\\PATH4_FINAL_ODK.xlsx', '${route4Dir}\\PATH4_ANSWER_KEY.pdf', '${mediaDir}' -DestinationPath '${route4Dir}\\PATH4_COMPLETE_PACKAGE.zip' -Force"`);

  console.log('🎉 PATH 4 PACKAGE REBUILT & PACKAGED SUCCESSFULLY!');
}

main();

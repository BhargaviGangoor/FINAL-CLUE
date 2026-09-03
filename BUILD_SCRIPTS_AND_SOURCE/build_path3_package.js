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
  // START — INSTRUCTIONS, RULES & REGISTRATION
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'start_group',
    label: 'START — REGISTRATION'
  });

  survey.push({
    type: 'note',
    name: 'instructions_rules_note',
    label: `🏆 FINAL CLUE — PATH 3
TREASURE HUNT FOR FRESHERS 2026

Welcome to PATH 3 of the Final Clue Treasure Hunt!

Instructions:

Rules and Regulations:

* Participants must report to the designated starting point 5–10 minutes before the event begins.
* Each team must consist of 3–4 members.
* At least one member of the team should have an Android phone.
* Each team will receive the first clue at the beginning of the event.
* Teams must solve each clue to find the location of the next clue.
* Clues must be solved in the given sequence.
* Teams are not allowed to take, hide, damage, or tamper with clues belonging to other teams.
* Teams must remain within the designated event area.
* Running in unsafe areas or restricted zones is prohibited.
* Participants must not enter restricted areas or disturb ongoing events/classes.
* Physical force, pushing, blocking, or interfering with other teams is strictly prohibited.
* Only one device containing ODK Collect is allowed per team.
* No use of Wi-Fi unless specifically specified. Otherwise, the team may be disqualified.
* Participants must not damage or move any property while searching for clues.
* Teams must follow instructions given by volunteers and organizers at all times.
* Asking people outside the team for answers or assistance is not allowed.
* Teams must not follow, copy, or deliberately interfere with another team's progress.
* Tampering with clues, cheating, entering restricted areas, or intentionally misleading other teams may result in immediate disqualification.
* The Organizing Committee will not be responsible for the loss or damage of any personal belongings of participants.

🏆 TEAM QUALIFICATION RULES (PER PATH):
There are 5 rounds in each path.

From each path:
* First 25 teams proceed to Round 2.
* Next 15 teams proceed to Round 3.
* Next 7 teams proceed to Round 4.
* Next 2 teams proceed to Round 5.

⚠️ MANDATORY RESPONSE & CAPS ONLY RULES:
• Every single question, photo upload, and code entry is strictly MANDATORY.
• All text answers and volunteer codes must be entered in UPPERCASE (CAPS ONLY).
• Lowercase letters will be rejected by validation.`,
    hint: 'Read all rules and instructions carefully.'
  });

  survey.push({
    type: 'text',
    name: 'team_name',
    label: 'Enter Team Name (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Team Name is mandatory. Please enter in UPPERCASE.',
    constraint: "regex(., '^[A-Z0-9\\-_ ]+$')",
    constraint_message: '❌ Please enter Team Name in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'player_id',
    label: 'Enter Team / Player Identification (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Player / Team ID is mandatory. Please enter in UPPERCASE.',
    constraint: "regex(., '^[A-Z0-9\\-_ ]+$')",
    constraint_message: '❌ Please enter Player/Team Identification in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'image',
    name: 'team_start_photo',
    label: '📸 Upload Team Verification Photo (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the start desk.',
    required: 'yes',
    required_message: '❌ Team verification photo is strictly mandatory.'
  });

  survey.push({
    type: 'end_group'
  });

  const startPassed = "${team_name} != '' and ${player_id} != '' and ${team_start_photo} != ''";

  // =============================================================
  // ROUND 1 — MBA: OBJECT (Visvesvaraya)
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
    label: '📍 ROUND 1: 🧩 OBJECT FINDING\n\nInstructions:\n1. Search the location to find the designated Visvesvaraya-related physical object.\n2. Take a mandatory photo of the object.\n3. Show the object/photo to the nearby Luminus volunteer.\n4. Enter the verification code given by the volunteer.\n\nEnter code in caps',
    hint: 'Find the Visvesvaraya object, take photo, and ask volunteer for code.'
  });

  survey.push({
    type: 'image',
    name: 'r1_mba_photo',
    label: '📸 Upload Photo of Discovered Hidden Object (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered Visvesvaraya object.',
    required: 'yes',
    required_message: '❌ Photo upload of the object is strictly mandatory.'
  });

  survey.push({
    type: 'text',
    name: 'r1_mba_code',
    label: 'Enter Volunteer Verification Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer verification code is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'ENG-KID-119'",
    constraint_message: '❌ Incorrect code. Check the object again and enter the code exactly as displayed.'
  });

  survey.push({
    type: 'end_group'
  });

  const r1Passed = `${startPassed} and normalize-space(\${r1_mba_code}) = 'ENG-KID-119' and \${r1_mba_photo} != ''`;

  // =============================================================
  // ROUND 2 — LOCATION CLUE: ELEMENTAL ENCRYPTION (ADMIN)
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
    label: `📍 ROUND 2 LOCATION CLUE: ELEMENTAL ENCRYPTION

Decode the chemical symbols using the periodic table. Take the first letter of each identified element and atomic numbers to reveal your next location.

Examine the image attached below carefully.

Enter code in caps`,
    hint: 'Decode the chemical symbols in the image to discover your next location.',
    'media::image': 'challenge.jpeg'
  });

  survey.push({
    type: 'text',
    name: 'r2_admin_block_answer',
    label: 'Enter your decoded destination location: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Destination location is mandatory.',
    'media::image': 'challenge.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'ADMIN' or normalize-space(.) = 'ADMIN BLOCK')",
    constraint_message: '❌ That answer is not correct. Recheck the chemical symbols and try again.'
  });

  const r2LocGuessed = "normalize-space(${r2_admin_block_answer}) = 'ADMIN' or normalize-space(${r2_admin_block_answer}) = 'ADMIN BLOCK'";

  survey.push({
    type: 'note',
    name: 'r2_admin_proceed_note',
    label: '🏃 Correct! Proceed to your next location and enter the START CODE provided there.\n\nEnter code in caps',
    hint: 'Proceed to the location and ask volunteer for start code.',
    relevant: r2LocGuessed
  });

  survey.push({
    type: 'text',
    name: 'r2_admin_start_code',
    label: 'Enter START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    relevant: r2LocGuessed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'ADMIN-START'",
    constraint_message: '❌ Incorrect START CODE. Check with the location setup and enter the code exactly as provided.'
  });

  survey.push({
    type: 'end_group'
  });

  const adminStartPassed = `${r1Passed} and (${r2LocGuessed}) and normalize-space(\${r2_admin_start_code}) = 'ADMIN-START'`;

  // =============================================================
  // ROUND 2 — MINI CHALLENGE (VARIANTS 1 & 2)
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
    label: 'Select your assigned challenge variant: (MANDATORY)\nMandatory selection',
    hint: 'Choose the variant assigned by the station volunteer.',
    required: 'yes',
    required_message: '❌ Selecting your assigned variant is mandatory.'
  });

  // Variant 1: Count to Unlock
  survey.push({
    type: 'begin_group',
    name: 'admin_v1_group',
    label: 'VARIANT A — COUNT TO UNLOCK',
    relevant: "${r2_admin_variant} = 'var1'"
  });

  survey.push({
    type: 'note',
    name: 'admin_v1_intro',
    label: `🔢 MINI-CHALLENGE: COUNT TO UNLOCK

Look carefully around the decorated hall and count the specified objects. You may divide the search among your team members, but do not touch, move, or pick up anything.

Enter code in caps`,
    hint: 'Count all hidden objects accurately.'
  });

  survey.push({
    type: 'integer',
    name: 'r2_v1_q1_chess',
    label: '1. ♟️ How many chess pieces were hidden? (MANDATORY)',
    hint: 'Enter the exact integer count.',
    required: 'yes',
    required_message: '❌ Question 1 count is mandatory.',
    constraint: '. = 5',
    constraint_message: '❌ Incorrect count for chess pieces.'
  });

  survey.push({
    type: 'integer',
    name: 'r2_v1_q2_uno',
    label: '2. 🎴 How many UNO cards were hidden? (MANDATORY)',
    hint: 'Enter the exact integer count.',
    required: 'yes',
    required_message: '❌ Question 2 count is mandatory.',
    constraint: '. = 7',
    constraint_message: '❌ Incorrect count for UNO cards.'
  });

  survey.push({
    type: 'integer',
    name: 'r2_v1_q3_appy',
    label: '3. 🧃 How many Appy juice bottles/packs were hidden? (MANDATORY)',
    hint: 'Enter the exact integer count.',
    required: 'yes',
    required_message: '❌ Question 3 count is mandatory.',
    constraint: '. = 3',
    constraint_message: '❌ Incorrect count for Appy juice.'
  });

  survey.push({
    type: 'integer',
    name: 'r2_v1_q4_smoodh',
    label: '4. 🥛 How many Smoodh bottles were hidden? (MANDATORY)',
    hint: 'Enter the exact integer count.',
    required: 'yes',
    required_message: '❌ Question 4 count is mandatory.',
    constraint: '. = 4',
    constraint_message: '❌ Incorrect count for Smoodh bottles.'
  });

  survey.push({
    type: 'integer',
    name: 'r2_v1_q5_files',
    label: '5. 📁 How many files were hidden? (MANDATORY)',
    hint: 'Enter the exact integer count.',
    required: 'yes',
    required_message: '❌ Question 5 count is mandatory.',
    constraint: '. = 2',
    constraint_message: '❌ Incorrect count for files.'
  });

  const v1AllCorrect = '${r2_v1_q1_chess} = 5 and ${r2_v1_q2_uno} = 7 and ${r2_v1_q3_appy} = 3 and ${r2_v1_q4_smoodh} = 4 and ${r2_v1_q5_files} = 2';

  survey.push({
    type: 'image',
    name: 'r2_v1_admin_photo',
    label: '📸 Upload Photo of Count to Unlock Station (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the station.',
    required: 'yes',
    required_message: '❌ Station photo upload is mandatory.',
    relevant: v1AllCorrect
  });

  survey.push({
    type: 'text',
    name: 'r2_v1_code',
    label: 'Enter Volunteer Clearance Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer clearance code is mandatory.',
    relevant: v1AllCorrect,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'ADM-DUCK-3'",
    constraint_message: '❌ Incorrect clearance code. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  // Variant 2: Memory Room (ADM-03)
  survey.push({
    type: 'begin_group',
    name: 'admin_v2_group',
    label: 'VARIANT B — MEMORY ROOM',
    relevant: "${r2_admin_variant} = 'var2'"
  });

  survey.push({
    type: 'note',
    name: 'admin_v2_intro',
    label: `🧠 MINI-CHALLENGE: MEMORY ROOM

You will enter a decorated classroom and observe the room for approximately 20 seconds. You may not record or photograph the room. After leaving, answer the questions from memory.

Enter code in caps`,
    hint: 'Answer all 7 memory questions in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_q1_id',
    label: '1. What colour was the ID card hanging in the classroom? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 1 is mandatory.',
    constraint: "regex(., '^[A-Z]+$') and normalize-space(.) = 'RED'",
    constraint_message: '❌ Incorrect colour. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'select_one clock_time_choices',
    name: 'r2_v2_q2_clock',
    label: '2. What time was shown on the classroom clock? (MANDATORY)',
    hint: 'Select the correct time observed on the clock.',
    required: 'yes',
    required_message: '❌ Question 2 is mandatory.',
    constraint: ". = 't_1015' or . = 't_1200' or . = 't_0230' or . = 't_0445'",
    constraint_message: '❌ Please select an option.'
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_q3_toy',
    label: '3. What stuffed toy was present in the classroom? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 3 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'MANGO'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_q4_poster',
    label: '4. What movie poster was pasted on the wall? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 4 is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\- ]+$') and (normalize-space(.) = 'SPIDER MAN BRAND NEW DAY' or normalize-space(.) = 'SPIDER-MAN BRAND NEW DAY' or normalize-space(.) = 'SPIDERMAN BRAND NEW DAY' or normalize-space(.) = 'SPIDER MAN')",
    constraint_message: '❌ Incorrect movie poster name. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_q5_novel',
    label: '5. Name of the novel present on the ground? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 5 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'ATOMIC HABITS'",
    constraint_message: '❌ Incorrect novel name. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_q6_board',
    label: '6. What was written on the classroom board? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 6 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'PYTHAGORAS THEOREM' or normalize-space(.) = 'PYTHAGOREAN THEOREM')",
    constraint_message: '❌ Incorrect board text. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_q7_benches',
    label: '7. What object was placed on the benches? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 7 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'TABLE TENNIS' or normalize-space(.) = 'TABLE TENNIS RACKET' or normalize-space(.) = 'TABLE TENNIS BALL' or normalize-space(.) = 'TT RACKET')",
    constraint_message: '❌ Incorrect object name. Please enter in UPPERCASE (CAPS ONLY).'
  });

  const v2AllCorrect = "normalize-space(${r2_v2_q1_id}) = 'RED' and ${r2_v2_q2_clock} != '' and normalize-space(${r2_v2_q3_toy}) = 'MANGO' and (normalize-space(${r2_v2_q4_poster}) = 'SPIDER MAN BRAND NEW DAY' or normalize-space(${r2_v2_q4_poster}) = 'SPIDER-MAN BRAND NEW DAY' or normalize-space(${r2_v2_q4_poster}) = 'SPIDERMAN BRAND NEW DAY' or normalize-space(${r2_v2_q4_poster}) = 'SPIDER MAN') and normalize-space(${r2_v2_q5_novel}) = 'ATOMIC HABITS' and (normalize-space(${r2_v2_q6_board}) = 'PYTHAGORAS THEOREM' or normalize-space(${r2_v2_q6_board}) = 'PYTHAGOREAN THEOREM') and (normalize-space(${r2_v2_q7_benches}) = 'TABLE TENNIS' or normalize-space(${r2_v2_q7_benches}) = 'TABLE TENNIS RACKET' or normalize-space(${r2_v2_q7_benches}) = 'TABLE TENNIS BALL' or normalize-space(${r2_v2_q7_benches}) = 'TT RACKET')";

  survey.push({
    type: 'note',
    name: 'r2_v2_success_note',
    label: '🎉 Memory challenge cleared. Enter the challenge code from volunteer to continue.\n\nEnter code in caps',
    hint: 'Ask volunteer for challenge code.',
    relevant: v2AllCorrect
  });

  survey.push({
    type: 'image',
    name: 'r2_v2_admin_photo',
    label: '📸 Upload Photo of Memory Room Challenge (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the station.',
    required: 'yes',
    required_message: '❌ Station photo upload is mandatory.',
    relevant: v2AllCorrect
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_code',
    label: 'Enter Volunteer Clearance Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer clearance code is mandatory.',
    relevant: v2AllCorrect,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'ADM-DISNEY-3'",
    constraint_message: '❌ Incorrect code. Enter ADM-DISNEY-3 in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  survey.push({
    type: 'end_group'
  });

  const adminMiniPassed = `${adminStartPassed} and ((${'\${r2_admin_variant}'} = 'var1' and normalize-space(\${r2_v1_code}) = 'ADM-DUCK-3' and \${r2_v1_admin_photo} != '') or (${'\${r2_admin_variant}'} = 'var2' and normalize-space(\${r2_v2_code}) = 'ADM-DISNEY-3' and \${r2_v2_admin_photo} != ''))`;

  // =============================================================
  // ROUND 3 — LOCATION CLUE (Destination: ECE)
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
    label: `📍 ROUND 3 LOCATION CLUE

"I deal in waves, both low and high,
Where analog signals never lie.
With resistors, chips, and wires in place,
Find your next clue in this circuit space."

Enter code in caps`,
    hint: 'Solve the riddle to discover your next location.'
  });

  survey.push({
    type: 'text',
    name: 'r3_ece_block_answer',
    label: 'Enter your deduced destination location: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Destination location is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'ECE' or normalize-space(.) = 'ECE BLOCK')",
    constraint_message: '❌ Not quite. Think about the department described by the electronics-related clues.'
  });

  const r3LocGuessed = "normalize-space(${r3_ece_block_answer}) = 'ECE' or normalize-space(${r3_ece_block_answer}) = 'ECE BLOCK'";

  survey.push({
    type: 'note',
    name: 'r3_ece_proceed_note',
    label: '🏃 Correct! Proceed to your next location and look for the next challenge.\n\nEnter code in caps',
    hint: 'Proceed to the location for the QR hunt.',
    relevant: r3LocGuessed
  });

  survey.push({
    type: 'end_group'
  });

  const r3BlockPassed = `${adminMiniPassed} and (${r3LocGuessed})`;

  // =============================================================
  // ROUND 3 — QR HUNT (ECE)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_ece_qr_group',
    label: 'ROUND 3 — QR HUNT',
    relevant: r3BlockPassed
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
    label: 'Scan Discovered QR Code (MANDATORY)',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ QR code scanning is mandatory.',
    constraint: "normalize-space(.) = 'HARRY_POTTER'",
    constraint_message: '❌ Incorrect code. Continue the QR hunt and check the code carefully.'
  });

  survey.push({
    type: 'image',
    name: 'r3_ece_qr_photo',
    label: '📸 Upload Photo of Discovered QR Code / Station (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered QR code location.',
    required: 'yes',
    required_message: '❌ Station photo upload is mandatory.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3QrPassed = `${r3BlockPassed} and normalize-space(\${r3_ece_qr_scan}) = 'HARRY_POTTER' and \${r3_ece_qr_photo} != ''`;

  // =============================================================
  // ROUND 4 — LOCATION CLUE (Destination: LIBRARY)
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
    label: `🔐 ROUND 4 LOCATION CLUE: CAESAR CIPHER

"Julius Caesar would have understood this.

Every letter has been shifted three places.

Move every letter three steps backward and decode the message.

Your answer is where the next clue waits."

Cipher:
OLEUDUB

Enter code in caps`,
    hint: 'Shift every letter 3 places backward.'
  });

  survey.push({
    type: 'text',
    name: 'r4_lib_cipher_answer',
    label: 'Enter your decoded destination location: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Destination location is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'LIBRARY' or normalize-space(.) = 'CENTRAL LIBRARY')",
    constraint_message: '❌ Incorrect. Remember that every letter must be shifted three places backward.'
  });

  const r4LocGuessed = "normalize-space(${r4_lib_cipher_answer}) = 'LIBRARY' or normalize-space(${r4_lib_cipher_answer}) = 'CENTRAL LIBRARY'";

  survey.push({
    type: 'note',
    name: 'r4_lib_proceed_note',
    label: '🏃 Correct! Proceed to your next location and look for the physical challenge.\n\nEnter code in caps',
    hint: 'Proceed to the location and ask volunteer for arrival code.',
    relevant: r4LocGuessed
  });

  survey.push({
    type: 'text',
    name: 'r4_lib_start_code',
    label: 'Enter Arrival Code from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Arrival code is mandatory.',
    relevant: r4LocGuessed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'TALL-NERD'",
    constraint_message: '❌ Incorrect code. Enter TALL-NERD in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const libStartPassed = `${r3QrPassed} and (${r4LocGuessed}) and normalize-space(\${r4_lib_start_code}) = 'TALL-NERD'`;

  // =============================================================
  // ROUND 4 — PHYSICAL CHALLENGE (LIBRARY)
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
    type: 'select_one challenge_status_list',
    name: 'r4_challenge_status',
    label: 'Challenge Status (MANDATORY)',
    hint: 'Select PASSED once completed with volunteer.',
    required: 'yes',
    required_message: '❌ Challenge status selection is mandatory.'
  });

  survey.push({
    type: 'note',
    name: 'r4_not_passed_note',
    label: '⚠️ Please follow the organizer\'s instructions before continuing.',
    hint: 'Complete the physical challenge as instructed.',
    relevant: "${r4_challenge_status} = 'not_passed'"
  });

  survey.push({
    type: 'image',
    name: 'r4_lib_phy_photo',
    label: '📸 Upload Photo of Physical Challenge Completion (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team completing the physical challenge.',
    required: 'yes',
    required_message: '❌ Challenge photo upload is mandatory.',
    relevant: "${r4_challenge_status} = 'passed'"
  });

  survey.push({
    type: 'text',
    name: 'r4_lib_phy_code',
    label: 'Enter Volunteer Completion Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer completion code is mandatory.',
    relevant: "${r4_challenge_status} = 'passed'",
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'LAST-LEG'",
    constraint_message: '❌ Incorrect code. Enter LAST-LEG in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4PhyPassed = `${libStartPassed} and \${r4_challenge_status} = 'passed' and normalize-space(\${r4_lib_phy_code}) = 'LAST-LEG' and \${r4_lib_phy_photo} != ''`;

  // =============================================================
  // FINAL ROUND: AUDITORIUM & STAGE RIDDLES & GRAND FINALE
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'final_puzzles_group',
    label: 'FINAL ROUND',
    relevant: r4PhyPassed
  });

  survey.push({
    type: 'note',
    name: 'final_audi_stage_intro',
    label: '🏛️ FINAL ROUND — RIDDLES\n\nSolve the two final riddles to reveal the final destination!\n\nEnter code in caps',
    hint: 'Solve the final riddles in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'final_riddle1_answer',
    label: `🧩 AUDITORIUM RIDDLE: (MANDATORY)
"I am empty, yet I am built for crowds.
I have a stage, but no actors of my own.
I have countless seats, but none are meant to sleep.
When a voice rises before me, silence falls behind me.
When the lights awaken, all eyes face one direction.
What am I?"
Enter code in caps`,
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Solving Riddle 1 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'AUDITORIUM' or normalize-space(.) = 'AUDI')",
    constraint_message: '❌ Incorrect answer. Solve Riddle 1 and enter in UPPERCASE (CAPS ONLY).'
  });

  const finalRiddle1Passed = `${r4PhyPassed} and (normalize-space(\${final_riddle1_answer}) = 'AUDITORIUM' or normalize-space(\${final_riddle1_answer}) = 'AUDI')`;

  survey.push({
    type: 'text',
    name: 'final_riddle2_stage_answer',
    label: `🎭 STAGE RIDDLE: (MANDATORY)
"I am elevated above the crowd, where performers stand and spotlights shine.
Underneath my wooden floor or behind the curtains, the ultimate secret waits.
What am I?"
Enter code in caps`,
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Solving Riddle 2 is mandatory.',
    relevant: finalRiddle1Passed,
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'STAGE'",
    constraint_message: '❌ Incorrect answer. Solve Riddle 2 and enter in UPPERCASE (CAPS ONLY).'
  });

  const finalRiddle2Passed = `${finalRiddle1Passed} and normalize-space(\${final_riddle2_stage_answer}) = 'STAGE'`;

  survey.push({
    type: 'note',
    name: 'final_proceed_stage_note',
    label: '🏃 Proceed to the final location identified!\n\nSolve the final puzzle at the stage, upload a photo of the completed puzzle, and get your clearance code from the Chief Judge!\n\nEnter code in caps',
    hint: 'Go to the stage to solve the final puzzle.',
    relevant: finalRiddle2Passed
  });

  survey.push({
    type: 'image',
    name: 'final_solved_puzzle_photo',
    label: '📸 Upload Photo of Your Solved Puzzle (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team\'s completed/solved puzzle.',
    required: 'yes',
    required_message: '❌ Photo of the solved puzzle is mandatory.',
    relevant: finalRiddle2Passed
  });

  survey.push({
    type: 'text',
    name: 'final_stage_volunteer_code',
    label: 'Enter Final Volunteer Clearance Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Final clearance code is mandatory.',
    relevant: finalRiddle2Passed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'FINAL-PATH3' or normalize-space(.) = 'FINAL-PATH1' or normalize-space(.) = 'FINAL-PATH2')",
    constraint_message: '❌ Incorrect code. Enter the code in UPPERCASE (CAPS ONLY) provided by the Chief Judge.'
  });

  survey.push({
    type: 'note',
    name: 'final_congratulations_screen',
    label: '🎉 CONGRATULATIONS! YOU HAVE COMPLETED PATH 3!\n\n🏆 You have successfully entered the final clearance code!\n\n🔔 NOW RUN TO GO RING THE BELL TO WIN THE GAME! 🔔🏃💨',
    hint: 'Run to ring the bell to claim victory!',
    relevant: `${finalRiddle2Passed} and (normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH3' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH1' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH2') and \${final_solved_puzzle_photo} != ''`
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
      label: 'Variant A: Count to Unlock'
    },
    {
      list_name: 'admin_variant_list',
      name: 'var2',
      label: 'Variant B: Memory Room'
    },
    {
      list_name: 'clock_time_choices',
      name: 't_1015',
      label: '10:15'
    },
    {
      list_name: 'clock_time_choices',
      name: 't_1200',
      label: '12:00'
    },
    {
      list_name: 'clock_time_choices',
      name: 't_0230',
      label: '02:30'
    },
    {
      list_name: 'clock_time_choices',
      name: 't_0445',
      label: '04:45'
    },
    {
      list_name: 'challenge_status_list',
      name: 'passed',
      label: 'PASSED'
    },
    {
      list_name: 'challenge_status_list',
      name: 'not_passed',
      label: 'NOT PASSED'
    }
  ];
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — PATH 3',
      form_id: 'PATH3_TREASURE_HUNT',
      version: '20260904',
      default_language: 'default'
    }
  ];
}

function buildAnswerKeyWorkbook() {
  const data = [
    {
      'Stage / Round': 'Registration',
      'Location': 'Start Desk',
      'Challenge / Item': 'Team Setup & Instructions',
      'Question / Prompt': 'Rules, Qualification Rules, Team Name, Player ID & Photo',
      'Media Attached': 'None',
      'Expected Answer / Code': 'TEAM-XX (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex(., \'^[A-Z0-9\\-_ ]+$\')',
      'Mandatory Upload': 'Yes (Team Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 1 (R1)',
      'Location': 'MBA',
      'Challenge / Item': 'Object Finding (Visvesvaraya)',
      'Question / Prompt': 'Find assigned Visvesvaraya object, upload photo, enter volunteer code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'ENG-KID-119 (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'ENG-KID-119\'',
      'Mandatory Upload': 'Yes (Discovered Object Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 2 Location Clue',
      'Location': 'In-App Elemental Clue',
      'Challenge / Item': 'Elemental Encryption (challenge.jpeg)',
      'Question / Prompt': 'Decode chemical symbols using periodic table -> ADMIN -> Start Code: ADMIN-START',
      'Media Attached': 'challenge.jpeg',
      'Expected Answer / Code': 'ADMIN & Start Code: ADMIN-START (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'ADMIN\' & \'ADMIN-START\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 2 (Var A)',
      'Location': 'ADMIN',
      'Challenge / Item': 'Variant A: Count to Unlock',
      'Question / Prompt': 'Count hidden items: 5 chess, 7 uno, 3 appy, 4 smoodh, 2 files + Photo + Code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Chess: 5, UNO: 7, Appy: 3, Smoodh: 4, Files: 2 -> Code: ADM-DUCK-3',
      'Verification / Constraint Rule': 'Integer equality & Code regex & normalize-space(.) = \'ADM-DUCK-3\'',
      'Mandatory Upload': 'Yes (Station Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 2 (Var B)',
      'Location': 'ADMIN',
      'Challenge / Item': 'Variant B: Memory Room (ADM-03)',
      'Question / Prompt': 'Memory questions: ID: RED, Clock: 10:15, Toy: MANGO, Movie: SPIDER MAN BRAND NEW DAY, Novel: ATOMIC HABITS, Board: PYTHAGORAS THEOREM, Bench: TABLE TENNIS + Photo + Code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'RED, 10:15, MANGO, SPIDER MAN BRAND NEW DAY, ATOMIC HABITS, PYTHAGORAS THEOREM, TABLE TENNIS -> Code: ADM-DISNEY-3',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'ADM-DISNEY-3\'',
      'Mandatory Upload': 'Yes (Station Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 3 Location Clue',
      'Location': 'In-App Electronics Riddle',
      'Challenge / Item': 'Circuit & Waves Riddle',
      'Question / Prompt': 'I deal in waves, both low and high... resistors, chips and wires -> Destination: ECE',
      'Media Attached': 'None',
      'Expected Answer / Code': 'ECE (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'ECE\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 3 Checkpoint (R3)',
      'Location': 'ECE',
      'Challenge / Item': 'QR Hunt',
      'Question / Prompt': 'Scan hidden QR code in ECE, upload photo',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Barcode: HARRY_POTTER (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Scanned barcode = \'HARRY_POTTER\'',
      'Mandatory Upload': 'Yes (Barcode Scan & Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 4 Location Clue',
      'Location': 'In-App Caesar Cipher',
      'Challenge / Item': 'LIB-01 Caesar Cipher (-3)',
      'Question / Prompt': 'Decode OLEUDUB by shifting letters 3 steps backward -> Destination: LIBRARY -> Arrival Code: TALL-NERD',
      'Media Attached': 'None',
      'Expected Answer / Code': 'LIBRARY & Arrival Code: TALL-NERD (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'LIBRARY\' & \'TALL-NERD\'',
      'Mandatory Upload': 'No (Inputs Mandatory)'
    },
    {
      'Stage / Round': 'Round 4 Checkpoint (R4)',
      'Location': 'LIBRARY',
      'Challenge / Item': 'Physical Challenge',
      'Question / Prompt': 'Complete physical challenge with volunteer, select PASSED, upload photo, enter finish code LAST-LEG',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Status: PASSED, Finish Code: LAST-LEG (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Status = PASSED & normalize-space(.) = \'LAST-LEG\'',
      'Mandatory Upload': 'Yes (Challenge Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Final Round (Riddle 1)',
      'Location': 'Main Auditorium',
      'Challenge / Item': 'Auditorium Riddle',
      'Question / Prompt': 'Solve riddle: Empty yet built for crowds, seats not meant to sleep...',
      'Media Attached': 'None',
      'Expected Answer / Code': 'AUDITORIUM (or AUDI) (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'AUDITORIUM\' or \'AUDI\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Final Round (Riddle 2)',
      'Location': 'Main Auditorium Stage',
      'Challenge / Item': 'Stage Riddle',
      'Question / Prompt': 'Solve stage riddle: Elevated above crowd, wooden floor...',
      'Media Attached': 'None',
      'Expected Answer / Code': 'STAGE (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'STAGE\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Grand Finale',
      'Location': 'Main Auditorium Stage',
      'Challenge / Item': 'Solved Puzzle & Bell Ring',
      'Question / Prompt': 'Solve puzzle, upload photo of solved puzzle, enter clearance code FINAL-PATH3, run to ring the bell to win',
      'Media Attached': 'None',
      'Expected Answer / Code': 'FINAL-PATH3 (UPPERCASE ONLY) -> Ring Bell',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'FINAL-PATH3\'',
      'Mandatory Upload': 'Yes (Solved Puzzle Photo MANDATORY)'
    }
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'PATH3_MASTER_KEY');
  return wb;
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
  .rules-box {
    background: #f5f3ff;
    border: 1px solid #c4b5fd;
    padding: 6px 10px;
    border-radius: 4px;
    margin-bottom: 10px;
    font-size: 8pt;
  }
</style>
</head>
<body>

<div class="header">
  <div style="float: right; text-align: right;">
    <span class="badge">PATH 3 OFFICIAL MASTER KEY</span><br>
    <small style="color: #64748b;">FINAL CLUE 2026</small>
  </div>
  <h1>FINAL CLUE &bull; ROUTE 3 / PATH 3</h1>
  <div style="font-size: 9pt; color: #475569; font-weight: 600;">
    CONFIDENTIAL ORGANIZER MASTER KEY &bull; PROGRESSION: MBA &rarr; ADMIN &rarr; ECE &rarr; LIBRARY &rarr; AUDITORIUM
  </div>
</div>

<div class="meta-grid">
  <div class="meta-card">
    <strong>Starting Station</strong>
    <span>MBA</span>
  </div>
  <div class="meta-card">
    <strong>Final Station</strong>
    <span>Main Auditorium Stage</span>
  </div>
  <div class="meta-card">
    <strong>Input Constraints</strong>
    <span>STRICT UPPERCASE (CAPS ONLY) &bull; 100% MANDATORY</span>
  </div>
  <div class="meta-card">
    <strong>Media Assets</strong>
    <span>challenge.jpeg (Elemental Encryption)</span>
  </div>
</div>

<div class="rules-box">
  <strong>🏆 Path 3 Team Qualification Rules:</strong><br>
  • Round 1 ➔ Round 2: First 25 teams proceed &bull; Round 2 ➔ Round 3: Next 15 teams proceed &bull; Round 3 ➔ Round 4: Next 7 teams proceed &bull; Round 4 ➔ Round 5: Next 2 teams proceed to the Grand Finale!<br>
  <strong>🔒 Strict Policy:</strong> Every question, photo upload, barcode scan, and code entry is <strong>100% MANDATORY</strong> and requires <strong>STRICT UPPERCASE ONLY</strong>.
</div>

<div class="section-title">1. Master Station-by-Station Directory</div>

<table>
  <thead>
    <tr>
      <th style="width: 14%;">Round / Stage</th>
      <th style="width: 16%;">Location</th>
      <th style="width: 35%;">Clue / Question Description</th>
      <th style="width: 20%;">Correct Answer / Code</th>
      <th style="width: 15%;">Mandatory Upload</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Registration</strong></td>
      <td>Start Desk</td>
      <td>Team Name, Player ID &amp; Team Verification Photo</td>
      <td><em>Participant Info</em></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Round 1 (R1)</strong></td>
      <td>MBA</td>
      <td>Find assigned Visvesvaraya object &amp; verify with volunteer</td>
      <td><span class="code-badge">ENG-KID-119</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Location Clue</strong></td>
      <td>In-App (<code>challenge.jpeg</code>)</td>
      <td>Elemental Encryption using periodic table</td>
      <td>Destination: <code>ADMIN</code><br>Start Code: <span class="code-badge">ADMIN-START</span></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R2 Mini Challenge</strong></td>
      <td>ADMIN</td>
      <td>
        &bull; <strong>Variant A:</strong> Count to Unlock (5 chess, 7 uno, 3 appy, 4 smoodh, 2 files) &rarr; Code: <span class="code-badge">ADM-DUCK-3</span><br>
        &bull; <strong>Variant B:</strong> Memory Room (7 Questions) &rarr; Code: <span class="code-badge">ADM-DISNEY-3</span>
      </td>
      <td><span class="code-badge">ADM-DUCK-3</span> / <span class="code-badge">ADM-DISNEY-3</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R3 Location Clue</strong></td>
      <td>In-App</td>
      <td>Circuit &amp; Waves Riddle: Analog signals, resistors, chips</td>
      <td>Destination: <code>ECE</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>ECE</td>
      <td>Physically search ECE for hidden QR code and scan</td>
      <td>Barcode: <span class="code-badge">HARRY_POTTER</span></td>
      <td>📸 Photo + Scan</td>
    </tr>
    <tr>
      <td><strong>R4 Location Clue</strong></td>
      <td>In-App</td>
      <td>Caesar Cipher (-3): <code>OLEUDUB</code></td>
      <td>Destination: <code>LIBRARY</code><br>Arrival Code: <span class="code-badge">TALL-NERD</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>Round 4 (R4)</strong></td>
      <td>LIBRARY</td>
      <td>Complete physical challenge with volunteer</td>
      <td>Status: <code>PASSED</code><br>Finish Code: <span class="code-badge">LAST-LEG</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Final Stage</strong></td>
      <td>AUDI Stage</td>
      <td>
        Riddle 1 (Dark hall, red seats): <code>AUDITORIUM</code><br>
        Riddle 2 (Elevated platform): <code>STAGE</code><br>
        Solve puzzle &rarr; Enter Code &rarr; Ring the Bell! 🔔
      </td>
      <td><span class="code-badge">FINAL-PATH3</span><br>&rarr; Ring Bell! 🔔</td>
      <td>📸 Mandatory Photo of Solved Puzzle</td>
    </tr>
  </tbody>
</table>

<div class="section-title">2. Volunteer Station Instructions &amp; Verification Procedures</div>

<div class="station-card">
  <h3><span>MBA Station: Object Finding</span><span class="code-badge">CODE: ENG-KID-119</span></h3>
  <div><strong>Volunteer Instructions:</strong> Participants locate the assigned Visvesvaraya object. Once verified with photo, provide code <code>ENG-KID-119</code>.</div>
</div>

<div class="station-card">
  <h3><span>ADMIN Station: Arrival &amp; Mini-Challenges</span><span class="code-badge">START: ADMIN-START &bull; VAR A: ADM-DUCK-3 &bull; VAR B: ADM-DISNEY-3</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams arrive at ADMIN. Provide start code <code>ADMIN-START</code>.
  <br>&bull; <strong>Variant A (Count to Unlock):</strong> Team counts items (5 chess, 7 uno, 3 appy, 4 smoodh, 2 files) and enters <code>ADM-DUCK-3</code>.
  <br>&bull; <strong>Variant B (Memory Room):</strong> Team observes room for 20s, answers 7 questions, and enters <code>ADM-DISNEY-3</code>.</div>
</div>

<div class="station-card">
  <h3><span>ECE Station: QR Hunt</span><span class="code-badge">SCANNED QR: HARRY_POTTER</span></h3>
  <div><strong>Volunteer Instructions:</strong> Ensure the physical QR code with payload <code>HARRY_POTTER</code> is hidden in the ECE area. Participants scan it using the in-app barcode scanner.</div>
</div>

<div class="station-card">
  <h3><span>LIBRARY Station: Physical Challenge</span><span class="code-badge">ARRIVAL: TALL-NERD &bull; FINISH: LAST-LEG</span></h3>
  <div><strong>Volunteer Instructions:</strong> Provide arrival code <code>TALL-NERD</code>. Supervise the physical challenge. Upon successful completion and photo upload, provide finish code <code>LAST-LEG</code>.</div>
</div>

<div class="station-card">
  <h3><span>Main Auditorium Stage: Grand Finale</span><span class="code-badge">CODE: FINAL-PATH3 &bull; 🔔 RING THE BELL</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams solve Riddle 1 (<code>AUDITORIUM</code>) and Riddle 2 (<code>STAGE</code>), upload a clear photo of their solved puzzle sheet, and receive clearance code <code>FINAL-PATH3</code> from Chief Judges before running to ring the victory bell!</div>
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
  if (fs.existsSync(sourceImg) && !fs.existsSync(destImg)) {
    fs.copyFileSync(sourceImg, destImg);
  }

  console.log('Generating PATH3_FINAL_ODK.xlsx with STRICT UPPERCASE ONLY & 100% MANDATORY enforcement...');
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

  console.log('Generating PATH3_ANSWER_KEY.xlsx...');
  const keyWb = buildAnswerKeyWorkbook();
  const keyPath = path.join(route3Dir, 'PATH3_ANSWER_KEY.xlsx');
  XLSX.writeFile(keyWb, keyPath);
  console.log(`Successfully created: ${keyPath}`);

  console.log('Generating PATH3_ANSWER_KEY.pdf...');
  const pdfHtml = buildAnswerKeyPdfHtml();
  const pdfPath = path.join(route3Dir, 'PATH3_ANSWER_KEY.pdf');
  renderHtmlToPdf(pdfHtml, pdfPath);
  console.log(`Successfully created: ${pdfPath}`);

  console.log('Packaging PATH3_COMPLETE_PACKAGE.zip and PATH3_MEDIA.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route3Dir}\\PATH3_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route3Dir}\\PATH3_FINAL_ODK.xlsx', '${route3Dir}\\PATH3_ANSWER_KEY.xlsx', '${route3Dir}\\PATH3_ANSWER_KEY.pdf', '${mediaDir}' -DestinationPath '${route3Dir}\\PATH3_COMPLETE_PACKAGE.zip' -Force"`);

  console.log('🎉 PATH 3 PACKAGE GENERATED SUCCESSFULLY!');
}

main();

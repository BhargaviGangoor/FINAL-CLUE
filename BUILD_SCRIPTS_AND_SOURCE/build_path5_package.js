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
    label: `🏆 FINAL CLUE — PATH 5
TREASURE HUNT FOR FRESHERS 2026

Welcome to PATH 5 of the Final Clue Treasure Hunt!

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
  // ROUND 1 — COE (LIB): OBJECT (COE Board)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r1_coe_group',
    label: 'ROUND 1 — OBJECT FINDING',
    relevant: startPassed
  });

  survey.push({
    type: 'note',
    name: 'r1_coe_intro',
    label: '📍 ROUND 1: 🧩 OBJECT FINDING\n\nInstructions:\n1. Search the location to find the designated COE board/object.\n2. Take a mandatory photo of the object.\n3. Show the object/photo to the nearby Luminus volunteer.\n4. Enter the verification code given by the volunteer.\n\nEnter code in caps',
    hint: 'Find the COE Board object, take photo, and ask volunteer for code.'
  });

  survey.push({
    type: 'image',
    name: 'r1_coe_photo',
    label: '📸 Upload Photo of Discovered COE Board Object (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered COE Board object.',
    required: 'yes',
    required_message: '❌ Photo upload of the object is strictly mandatory.'
  });

  survey.push({
    type: 'text',
    name: 'r1_coe_code',
    label: 'Enter Volunteer Verification Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer verification code is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'YAAKE-GURU-5'",
    constraint_message: '❌ Incorrect code. Check the object again and enter the code exactly as displayed.'
  });

  survey.push({
    type: 'end_group'
  });

  const r1Passed = `${startPassed} and normalize-space(\${r1_coe_code}) = 'YAAKE-GURU-5' and \${r1_coe_photo} != ''`;

  // =============================================================
  // ROUND 2 — LOCATION CLUE (Destination: AIML)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r2_aiml_loc_group',
    label: 'ROUND 2 LOCATION CLUE',
    relevant: r1Passed
  });

  survey.push({
    type: 'note',
    name: 'r2_aiml_loc_note',
    label: `📍 ROUND 2 LOCATION CLUE: RIDDLE OF THE MISSING CONCEPT

Four clues describe four concepts.
Solve each clue and take the first letter of each answer.
Combine the four letters to discover your next location.

Clue 1:
"Step by step, I show the way,
solving problems every day."

Clue 2:
"I connect the world without a sound,
through wires and waves, I'm all around."

Clue 3:
"I learn from data, again and again,
getting smarter with every gain."

Clue 4:
"I'm the code you write, line by line,
Python or Java, the choice is fine."

Enter code in caps`,
    hint: 'Take the first letter of each answer and combine them.'
  });

  survey.push({
    type: 'text',
    name: 'r2_aiml_loc_answer',
    label: 'Enter your 4-letter combined destination location: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Destination location is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\/& ]+$') and (normalize-space(.) = 'AIML' or normalize-space(.) = 'AI ML' or normalize-space(.) = 'AIML BLOCK' or normalize-space(.) = 'AI/ML' or normalize-space(.) = 'AI & ML')",
    constraint_message: '❌ Incorrect. Solve all four clues and take the first letter of each answer.'
  });

  const r2LocGuessed = "normalize-space(${r2_aiml_loc_answer}) = 'AIML' or normalize-space(${r2_aiml_loc_answer}) = 'AI ML' or normalize-space(${r2_aiml_loc_answer}) = 'AIML BLOCK' or normalize-space(${r2_aiml_loc_answer}) = 'AI/ML' or normalize-space(${r2_aiml_loc_answer}) = 'AI & ML'";

  survey.push({
    type: 'note',
    name: 'r2_aiml_proceed_note',
    label: '🏃 Correct! Proceed to the location identified by your answer and enter the START CODE provided there.\n\nEnter code in caps',
    hint: 'Proceed to the location and ask volunteer for start code.',
    relevant: r2LocGuessed
  });

  survey.push({
    type: 'text',
    name: 'r2_aiml_start_code',
    label: 'Enter START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    relevant: r2LocGuessed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'AIML-START'",
    constraint_message: '❌ Incorrect START CODE. Check the code at your current location and enter it exactly as provided.'
  });

  survey.push({
    type: 'end_group'
  });

  const aimlStartPassed = `${r1Passed} and (${r2LocGuessed}) and normalize-space(\${r2_aiml_start_code}) = 'AIML-START'`;

  // =============================================================
  // ROUND 2 — CHALLENGES (AIML: Reverse Image Prompting)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r2_aiml_challenge_group',
    label: 'ROUND 2 — AIML CHALLENGES',
    relevant: aimlStartPassed
  });

  survey.push({
    type: 'note',
    name: 'r2_aiml_prompt_rules_note',
    label: `🎨 AIML CHALLENGE: REVERSE IMAGE PROMPTING

Event Format:
• One lab system will be provided per team.
• Images will be displayed for 10 seconds each in a slideshow.

Rules & Regulations:
• Each team must recreate at least 3 images from the slideshow.
• The 3 images must be created on 3 different AI chats / tabs - not in the same chat.
• Teams can use any AI image generation tool (ChatGPT, Midjourney, Gemini, Leonardo, etc.).
• Google Image Search, reverse image search, or uploading / scanning the displayed image is strictly prohibited. Direct disqualification if found.
• Only original prompting is allowed - img-to-img is not allowed.
• Volunteers will come to your system and verify your generated images directly.

Enter code in caps`,
    hint: 'Follow lab instructions and volunteer guidance.'
  });

  survey.push({
    type: 'select_one challenge_status_list',
    name: 'r2_aiml_challenge_status',
    label: 'AIML Image Prompting Challenge Status (MANDATORY)',
    hint: 'Select PASSED once verified by station volunteers.',
    required: 'yes',
    required_message: '❌ Challenge status selection is mandatory.'
  });

  survey.push({
    type: 'note',
    name: 'r2_aiml_not_passed_note',
    label: '⚠️ Challenge not cleared. Follow the organizer\'s instructions before continuing.',
    hint: 'Complete the image generation prompt challenge.',
    relevant: "${r2_aiml_challenge_status} = 'not_passed'"
  });

  survey.push({
    type: 'note',
    name: 'r2_aiml_passed_note',
    label: '🎉 Challenge cleared. Enter the code provided by the organizer.\n\nEnter code in caps',
    hint: 'Ask volunteer for clearance code.',
    relevant: "${r2_aiml_challenge_status} = 'passed'"
  });

  survey.push({
    type: 'image',
    name: 'r2_aiml_photo',
    label: '📸 Upload Photo of Generated Images / Station (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team\'s generated images on the lab screen.',
    required: 'yes',
    required_message: '❌ Station verification photo is mandatory.',
    relevant: "${r2_aiml_challenge_status} = 'passed'"
  });

  survey.push({
    type: 'text',
    name: 'r2_aiml_code1',
    label: 'Enter Volunteer Clearance Code 1 (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer clearance code is mandatory.',
    relevant: "${r2_aiml_challenge_status} = 'passed'",
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'GEMMA-V05'",
    constraint_message: '❌ Incorrect code. Check the code obtained after completing the challenge.'
  });

  const aimlCode1Passed = `${aimlStartPassed} and \${r2_aiml_challenge_status} = 'passed' and normalize-space(\${r2_aiml_code1}) = 'GEMMA-V05'`;

  survey.push({
    type: 'text',
    name: 'r2_aiml_end_code',
    label: 'Enter Station Finish Code 2 from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Station finish code is mandatory.',
    relevant: aimlCode1Passed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'OLLAMA-V05' or normalize-space(.) = 'OLLAMA-v05')",
    constraint_message: '❌ Incorrect finish code. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r2AllPassed = `${aimlCode1Passed} and (normalize-space(\${r2_aiml_end_code}) = 'OLLAMA-V05' or normalize-space(\${r2_aiml_end_code}) = 'OLLAMA-v05') and \${r2_aiml_photo} != ''`;

  // =============================================================
  // ROUND 3 — LOCATION CLUE (Destination: CSE)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_cse_loc_group',
    label: 'ROUND 3 LOCATION CLUE',
    relevant: r2AllPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_cse_math_note',
    label: `📍 ROUND 3 LOCATION CLUE: MATHEMATICAL CODE

Three numbers are hidden inside the equations.
Solve each equation.

Then use the alphabet as your key:
A = 1, B = 2, C = 3 ... Z = 26

What three-letter code do you uncover?

Code:
(1 + 2) — (20 − 1) — (10 ÷ 2)

Enter code in caps`,
    hint: 'Solve each equation and map the 3 numbers to letters A=1 to Z=26.'
  });

  survey.push({
    type: 'text',
    name: 'r3_cse_loc_answer',
    label: 'Enter your uncovered 3-letter destination location: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Destination location is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'CSE' or normalize-space(.) = 'CSE BLOCK')",
    constraint_message: '❌ Incorrect. Solve all three equations first, then convert the resulting numbers using A=1, B=2, C=3...Z=26.'
  });

  const r3LocGuessed = "normalize-space(${r3_cse_loc_answer}) = 'CSE' or normalize-space(${r3_cse_loc_answer}) = 'CSE BLOCK'";

  survey.push({
    type: 'note',
    name: 'r3_cse_proceed_note',
    label: '🏃 Correct! You have uncovered the three-letter code. Proceed to the location and look for the QR challenge.\n\nEnter code in caps',
    hint: 'Proceed to the location for the QR hunt.',
    relevant: r3LocGuessed
  });

  survey.push({
    type: 'end_group'
  });

  const r3BlockPassed = `${r2AllPassed} and (${r3LocGuessed})`;

  // =============================================================
  // ROUND 3 — QR HUNT (CSE)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_cse_qr_group',
    label: 'ROUND 3 — QR HUNT',
    relevant: r3BlockPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_cse_qr_note',
    label: '📍 ROUND 3: QR HUNT\n\nFind the designated QR code at this location and scan it.\nFollow the instructions provided by the QR challenge.\n\nEnter code in caps',
    hint: 'Locate and scan the hidden QR code.'
  });

  survey.push({
    type: 'barcode',
    name: 'r3_cse_qr_scan',
    label: 'Scan Discovered QR Code (MANDATORY)',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ QR code scanning is mandatory.',
    constraint: "normalize-space(.) = 'PETER_PARKER'",
    constraint_message: '❌ Incorrect code. Continue the QR hunt and check the code carefully.'
  });

  survey.push({
    type: 'image',
    name: 'r3_cse_qr_photo',
    label: '📸 Upload Photo of Discovered QR Code / Station (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered QR code location.',
    required: 'yes',
    required_message: '❌ Station photo upload is mandatory.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3QrPassed = `${r3BlockPassed} and normalize-space(\${r3_cse_qr_scan}) = 'PETER_PARKER' and \${r3_cse_qr_photo} != ''`;

  // =============================================================
  // ROUND 4 — LOCATION CLUE (MEC-02 — Piece-by-Piece with badminton court.jpeg)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_mech_loc_group',
    label: 'ROUND 4 LOCATION CLUE',
    relevant: r3QrPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_mech_piece_note',
    label: `📍 ROUND 4 LOCATION CLUE: PIECE-BY-PIECE

Piece by piece, an image has been hidden from you.
Study the available pieces in the image attached below and determine what place or object they form.
Your answer will help identify your next location.

Enter code in caps`,
    hint: 'Examine the image pieces to identify the destination.',
    'media::image': 'badminton court.jpeg'
  });

  survey.push({
    type: 'text',
    name: 'r4_mech_loc_answer',
    label: 'Enter the location/object identified from the photograph pieces: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Location answer is mandatory.',
    'media::image': 'badminton court.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'MECH' or normalize-space(.) = 'MECH BLOCK' or normalize-space(.) = 'MECHANICAL' or normalize-space(.) = 'BADMINTON COURT')",
    constraint_message: '❌ Incorrect. Examine the image pieces carefully to identify the campus location.'
  });

  const r4LocGuessed = "normalize-space(${r4_mech_loc_answer}) = 'MECH' or normalize-space(${r4_mech_loc_answer}) = 'MECH BLOCK' or normalize-space(${r4_mech_loc_answer}) = 'MECHANICAL' or normalize-space(${r4_mech_loc_answer}) = 'BADMINTON COURT'";

  survey.push({
    type: 'note',
    name: 'r4_mech_proceed_note',
    label: '🏃 Correct! Proceed to the location identified by your answer and look for the physical challenge.\n\nEnter code in caps',
    hint: 'Proceed to the location and ask volunteer for start code.',
    relevant: r4LocGuessed
  });

  survey.push({
    type: 'text',
    name: 'r4_mech_start_code',
    label: 'Enter START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    relevant: r4LocGuessed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'PIECE-START' or normalize-space(.) = 'PIECE_START')",
    constraint_message: '❌ Incorrect START CODE. Enter PIECE-START in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const mechStartPassed = `${r3QrPassed} and (${r4LocGuessed}) and (normalize-space(\${r4_mech_start_code}) = 'PIECE-START' or normalize-space(\${r4_mech_start_code}) = 'PIECE_START')`;

  // =============================================================
  // ROUND 4 — PHYSICAL CHALLENGE (MECH)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_mech_phy_group',
    label: 'ROUND 4 — PHYSICAL CHALLENGE',
    relevant: mechStartPassed
  });

  survey.push({
    type: 'select_one yes_no_list',
    name: 'r4_enable_optional_piece',
    label: 'Enable Piece-by-Piece Station Challenge? (Organizer Option)',
    hint: 'Select YES to record station challenge or NO to proceed directly to physical challenge.',
    required: 'yes',
    required_message: '❌ Selection is mandatory.'
  });

  survey.push({
    type: 'note',
    name: 'r4_mech_phy_note',
    label: '📍 ROUND 4: PHYSICAL CHALLENGE\n\nReport to the station and complete the physical challenge under volunteer supervision.\n\nEnter code in caps',
    hint: 'Complete physical challenge with volunteer.'
  });

  survey.push({
    type: 'select_one challenge_status_list',
    name: 'r4_mech_status',
    label: 'Physical Challenge Status (MANDATORY)',
    hint: 'Select PASSED once completed with volunteer.',
    required: 'yes',
    required_message: '❌ Challenge status selection is mandatory.'
  });

  survey.push({
    type: 'note',
    name: 'r4_mech_not_passed_note',
    label: '⚠️ Please follow the organizer\'s instructions before continuing.',
    hint: 'Complete the physical challenge as instructed.',
    relevant: "${r4_mech_status} = 'not_passed'"
  });

  survey.push({
    type: 'image',
    name: 'r4_mech_phy_photo',
    label: '📸 Upload Photo of Physical Challenge Completion (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team completing the physical challenge.',
    required: 'yes',
    required_message: '❌ Challenge photo upload is mandatory.',
    relevant: "${r4_mech_status} = 'passed'"
  });

  survey.push({
    type: 'text',
    name: 'r4_mech_phy_code',
    label: 'Enter Volunteer Completion Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer completion code is mandatory.',
    relevant: "${r4_mech_status} = 'passed'",
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'PIECE-BYE' or normalize-space(.) = 'PIECE-BYE---' or normalize-space(.) = 'PIECE_BYE')",
    constraint_message: '❌ Incorrect code. Enter PIECE-BYE in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4PhyPassed = `${mechStartPassed} and \${r4_mech_status} = 'passed' and (normalize-space(\${r4_mech_phy_code}) = 'PIECE-BYE' or normalize-space(\${r4_mech_phy_code}) = 'PIECE-BYE---' or normalize-space(\${r4_mech_phy_code}) = 'PIECE_BYE') and \${r4_mech_phy_photo} != ''`;

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
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'FINAL-PATH5' or normalize-space(.) = 'FINAL-PATH1' or normalize-space(.) = 'FINAL-PATH2' or normalize-space(.) = 'FINAL-PATH3' or normalize-space(.) = 'FINAL-PATH4')",
    constraint_message: '❌ Incorrect code. Enter the code in UPPERCASE (CAPS ONLY) provided by the Chief Judge.'
  });

  survey.push({
    type: 'note',
    name: 'final_congratulations_screen',
    label: '🎉 CONGRATULATIONS! YOU HAVE COMPLETED PATH 5!\n\n🏆 You have successfully entered the final clearance code!\n\n🔔 NOW RUN TO GO RING THE BELL TO WIN THE GAME! 🔔🏃💨',
    hint: 'Run to ring the bell to claim victory!',
    relevant: `${finalRiddle2Passed} and (normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH5' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH1' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH2' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH3' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH4') and \${final_solved_puzzle_photo} != ''`
  });

  survey.push({
    type: 'end_group'
  });

  return survey;
}

function buildChoices() {
  return [
    {
      list_name: 'challenge_status_list',
      name: 'passed',
      label: 'PASSED'
    },
    {
      list_name: 'challenge_status_list',
      name: 'not_passed',
      label: 'NOT PASSED'
    },
    {
      list_name: 'yes_no_list',
      name: 'yes',
      label: 'YES'
    },
    {
      list_name: 'yes_no_list',
      name: 'no',
      label: 'NO'
    }
  ];
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — PATH 5',
      form_id: 'PATH5_TREASURE_HUNT',
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
      'Location': 'COE (LIB)',
      'Challenge / Item': 'Object Finding (COE Board)',
      'Question / Prompt': 'Find assigned COE Board object, upload photo, enter volunteer code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'YAAKE-GURU-5 (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'YAAKE-GURU-5\'',
      'Mandatory Upload': 'Yes (Discovered Object Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 2 Location Clue',
      'Location': 'In-App Riddle',
      'Challenge / Item': 'Riddle of Missing Concept (A+I+M+L)',
      'Question / Prompt': 'Algorithm (A) + Internet (I) + Model (M) + Language (L) -> Destination: AIML -> Start Code: AIML-START',
      'Media Attached': 'None',
      'Expected Answer / Code': 'AIML & Start Code: AIML-START (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'AIML\' & \'AIML-START\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 2 Challenge (R2)',
      'Location': 'AIML Lab',
      'Challenge / Item': 'Reverse Image Prompting Challenge',
      'Question / Prompt': 'Recreate 3 slideshow images in 3 different AI chats + Photo + Codes',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Status: PASSED -> Code 1: GEMMA-V05 -> Finish Code: OLLAMA-V05',
      'Verification / Constraint Rule': 'Status = PASSED & normalize-space(.) = \'GEMMA-V05\' & \'OLLAMA-V05\'',
      'Mandatory Upload': 'Yes (Station Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 3 Location Clue',
      'Location': 'In-App Math Code',
      'Challenge / Item': 'Mathematical Code: (1+2)-(20-1)-(10/2)',
      'Question / Prompt': '3-19-5 -> Using A=1..Z=26 -> C-S-E -> Destination: CSE',
      'Media Attached': 'None',
      'Expected Answer / Code': 'CSE (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'CSE\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 3 Checkpoint (R3)',
      'Location': 'CSE',
      'Challenge / Item': 'QR Hunt',
      'Question / Prompt': 'Scan hidden QR code in CSE, upload photo',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Barcode: PETER_PARKER (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Scanned barcode = \'PETER_PARKER\'',
      'Mandatory Upload': 'Yes (Barcode Scan & Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 4 Location Clue',
      'Location': 'In-App Image (badminton court.jpeg)',
      'Challenge / Item': 'MEC-02 Piece-by-Piece',
      'Question / Prompt': 'Identify location from photograph pieces -> Destination: MECH -> Start Code: PIECE-START',
      'Media Attached': 'badminton court.jpeg',
      'Expected Answer / Code': 'MECH & Start Code: PIECE-START (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'MECH\' & \'PIECE-START\'',
      'Mandatory Upload': 'No (Inputs Mandatory)'
    },
    {
      'Stage / Round': 'Round 4 Checkpoint (R4)',
      'Location': 'MECH',
      'Challenge / Item': 'Physical Challenge',
      'Question / Prompt': 'Complete physical challenge with volunteer, select PASSED, upload photo, enter finish code PIECE-BYE',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Status: PASSED, Finish Code: PIECE-BYE (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Status = PASSED & normalize-space(.) = \'PIECE-BYE\'',
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
      'Question / Prompt': 'Solve puzzle, upload photo of solved puzzle, enter clearance code FINAL-PATH5, run to ring the bell to win',
      'Media Attached': 'None',
      'Expected Answer / Code': 'FINAL-PATH5 (UPPERCASE ONLY) -> Ring Bell',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'FINAL-PATH5\'',
      'Mandatory Upload': 'Yes (Solved Puzzle Photo MANDATORY)'
    }
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'PATH5_MASTER_KEY');
  return wb;
}

function buildAnswerKeyPdfHtml() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>PATH 5 — ORGANIZER MASTER ANSWER KEY</title>
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
    border-bottom: 3px solid #0284c7;
    padding-bottom: 6px;
    margin-bottom: 12px;
  }
  .header h1 {
    color: #0369a1;
    margin: 0 0 3px 0;
    font-size: 15pt;
    letter-spacing: 0.4px;
  }
  .header .badge {
    display: inline-block;
    background: #0284c7;
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
    background: #f0f9ff;
    padding: 6px 8px;
    border-radius: 4px;
    border-left: 3px solid #38bdf8;
  }
  .meta-card strong {
    color: #0369a1;
    display: block;
    font-size: 7.5pt;
    text-transform: uppercase;
  }
  .meta-card span {
    font-size: 8.5pt;
    font-weight: 600;
  }
  .section-title {
    background: #0369a1;
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
    color: #0369a1;
    font-size: 9pt;
    display: flex;
    justify-content: space-between;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 3px;
  }
  .rules-box {
    background: #f0f9ff;
    border: 1px solid #bae6fd;
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
    <span class="badge">PATH 5 OFFICIAL MASTER KEY</span><br>
    <small style="color: #64748b;">FINAL CLUE 2026</small>
  </div>
  <h1>FINAL CLUE &bull; ROUTE 5 / PATH 5</h1>
  <div style="font-size: 9pt; color: #475569; font-weight: 600;">
    CONFIDENTIAL ORGANIZER MASTER KEY &bull; PROGRESSION: COE (LIB) &rarr; AIML &rarr; CSE &rarr; MECH &rarr; AUDITORIUM
  </div>
</div>

<div class="meta-grid">
  <div class="meta-card">
    <strong>Starting Station</strong>
    <span>COE (Library)</span>
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
    <span>badminton court.jpeg (Piece-by-Piece)</span>
  </div>
</div>

<div class="rules-box">
  <strong>🏆 Path 5 Team Qualification Rules:</strong><br>
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
      <td>COE (LIB)</td>
      <td>Find COE Board object &amp; verify with volunteer</td>
      <td><span class="code-badge">YAAKE-GURU-5</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Location Clue</strong></td>
      <td>In-App</td>
      <td>Missing Concept Riddle: Algorithm (A) + Internet (I) + Model (M) + Language (L)</td>
      <td>Destination: <code>AIML</code><br>Start Code: <span class="code-badge">AIML-START</span></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R2 Challenge</strong></td>
      <td>AIML Lab</td>
      <td>Reverse Image Prompting (3 images in 3 different AI chats)</td>
      <td>Code 1: <span class="code-badge">GEMMA-V05</span><br>Finish Code: <span class="code-badge">OLLAMA-V05</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R3 Location Clue</strong></td>
      <td>In-App</td>
      <td>Mathematical Code: (1+2)=3(C), (20-1)=19(S), (10/2)=5(E)</td>
      <td>Destination: <code>CSE</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>CSE</td>
      <td>Physically search CSE for hidden QR code and scan</td>
      <td>Barcode: <span class="code-badge">PETER_PARKER</span></td>
      <td>📸 Photo + Scan</td>
    </tr>
    <tr>
      <td><strong>R4 Location Clue</strong></td>
      <td>In-App (<code>badminton court.jpeg</code>)</td>
      <td>MEC-02 Piece-by-Piece Location Clue</td>
      <td>Destination: <code>MECH</code><br>Start Code: <span class="code-badge">PIECE-START</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>Round 4 (R4)</strong></td>
      <td>MECH</td>
      <td>Complete physical challenge with volunteer</td>
      <td>Status: <code>PASSED</code><br>Finish Code: <span class="code-badge">PIECE-BYE</span></td>
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
      <td><span class="code-badge">FINAL-PATH5</span><br>&rarr; Ring Bell! 🔔</td>
      <td>📸 Mandatory Photo of Solved Puzzle</td>
    </tr>
  </tbody>
</table>

<div class="section-title">2. Volunteer Station Instructions &amp; Verification Procedures</div>

<div class="station-card">
  <h3><span>COE (LIB) Station: Object Finding</span><span class="code-badge">CODE: YAAKE-GURU-5</span></h3>
  <div><strong>Volunteer Instructions:</strong> Participants locate the assigned COE Board object. Once verified with photo, provide code <code>YAAKE-GURU-5</code>.</div>
</div>

<div class="station-card">
  <h3><span>AIML Lab Station: Reverse Image Prompting</span><span class="code-badge">START: AIML-START &bull; CODE 1: GEMMA-V05 &bull; FINISH: OLLAMA-V05</span></h3>
  <div><strong>Volunteer Instructions:</strong> Provide start code <code>AIML-START</code>. Participants recreate at least 3 slideshow images in 3 different AI tabs. Once verified, provide clearance code <code>GEMMA-V05</code> and station finish code <code>OLLAMA-V05</code>.</div>
</div>

<div class="station-card">
  <h3><span>CSE Station: QR Hunt</span><span class="code-badge">SCANNED QR: PETER_PARKER</span></h3>
  <div><strong>Volunteer Instructions:</strong> Ensure the physical QR code with payload <code>PETER_PARKER</code> is hidden in the CSE area. Participants scan it using the in-app barcode scanner.</div>
</div>

<div class="station-card">
  <h3><span>MECH Station: Physical Challenge</span><span class="code-badge">START: PIECE-START &bull; FINISH: PIECE-BYE</span></h3>
  <div><strong>Volunteer Instructions:</strong> Provide start code <code>PIECE-START</code>. Supervise the physical challenge. Upon successful completion and photo upload, provide finish code <code>PIECE-BYE</code>.</div>
</div>

<div class="station-card">
  <h3><span>Main Auditorium Stage: Grand Finale</span><span class="code-badge">CODE: FINAL-PATH5 &bull; 🔔 RING THE BELL</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams solve Riddle 1 (<code>AUDITORIUM</code>) and Riddle 2 (<code>STAGE</code>), upload a clear photo of their solved puzzle sheet, and receive clearance code <code>FINAL-PATH5</code> from Chief Judges before running to ring the victory bell!</div>
</div>

</body>
</html>`;
}

function main() {
  const baseDir = path.resolve(__dirname, '..');
  const route5Dir = path.join(baseDir, 'ROUTE_5_PATH5_COE_AIML_CSE_MECH_AUDI');
  if (!fs.existsSync(route5Dir)) fs.mkdirSync(route5Dir, { recursive: true });

  const mediaDir = path.join(route5Dir, 'media');
  if (!fs.existsSync(mediaDir)) fs.mkdirSync(mediaDir, { recursive: true });

  // Copy badminton court.jpeg from baseDir to mediaDir
  const sourceImg = path.join(baseDir, 'badminton court.jpeg');
  const destImg = path.join(mediaDir, 'badminton court.jpeg');
  if (fs.existsSync(sourceImg)) {
    fs.copyFileSync(sourceImg, destImg);
  }

  console.log('Generating PATH5_FINAL_ODK.xlsx with STRICT UPPERCASE ONLY & 100% MANDATORY enforcement...');
  const survey = buildSurvey();
  const choices = buildChoices();
  const settings = buildSettings();

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(survey), 'survey');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(choices), 'choices');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settings), 'settings');

  const xlsxPath = path.join(route5Dir, 'PATH5_FINAL_ODK.xlsx');
  XLSX.writeFile(wb, xlsxPath);
  console.log(`Successfully created: ${xlsxPath}`);

  console.log('Generating PATH5_ANSWER_KEY.xlsx...');
  const keyWb = buildAnswerKeyWorkbook();
  const keyPath = path.join(route5Dir, 'PATH5_ANSWER_KEY.xlsx');
  XLSX.writeFile(keyWb, keyPath);
  console.log(`Successfully created: ${keyPath}`);

  console.log('Generating PATH5_ANSWER_KEY.pdf...');
  const pdfHtml = buildAnswerKeyPdfHtml();
  const pdfPath = path.join(route5Dir, 'PATH5_ANSWER_KEY.pdf');
  renderHtmlToPdf(pdfHtml, pdfPath);
  console.log(`Successfully created: ${pdfPath}`);

  console.log('Packaging PATH5_COMPLETE_PACKAGE.zip and PATH5_MEDIA.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route5Dir}\\PATH5_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route5Dir}\\PATH5_FINAL_ODK.xlsx', '${route5Dir}\\PATH5_ANSWER_KEY.xlsx', '${route5Dir}\\PATH5_ANSWER_KEY.pdf', '${mediaDir}' -DestinationPath '${route5Dir}\\PATH5_COMPLETE_PACKAGE.zip' -Force"`);

  console.log('🎉 PATH 5 PACKAGE GENERATED SUCCESSFULLY!');
}

main();

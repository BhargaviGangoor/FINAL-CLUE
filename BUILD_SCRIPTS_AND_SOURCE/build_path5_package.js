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

  const progressionWarning = 'Read the question properly as you cannot come back to the question once you go ahead.';

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

⚠️ STRICT PROGRESSION & MANDATORY RULES:
• Read each question properly as you cannot come back to the question once you go ahead!
• Every single question, photo upload, and code entry is strictly MANDATORY.
• All text answers and volunteer codes must be entered in UPPERCASE (CAPS ONLY).
• Lowercase letters will be rejected by validation.`,
    hint: `Read all rules and instructions carefully. ${progressionWarning}`
  });

  survey.push({
    type: 'text',
    name: 'team_name',
    label: 'Enter Team Name (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Team Name is mandatory. Please enter in UPPERCASE.',
    constraint: "regex(., '^[A-Z0-9\\-_ ]+$')",
    constraint_message: '❌ Please enter Team Name in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'player_id',
    label: 'Enter Team / Player Identification (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Player / Team ID is mandatory. Please enter in UPPERCASE.',
    constraint: "regex(., '^[A-Z0-9\\-_ ]+$')",
    constraint_message: '❌ Please enter Player/Team Identification in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'image',
    name: 'team_start_photo',
    label: '📸 Upload Team Verification Photo (MANDATORY)\nPhoto upload is mandatory',
    hint: `Take a clear photo of your team at the start desk. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Team verification photo is strictly mandatory.'
  });

  survey.push({
    type: 'end_group'
  });

  const startPassed = "${team_name} != '' and ${player_id} != '' and ${team_start_photo} != ''";

  // =============================================================
  // ROUND 1: OBJECT FINDING
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r1_group',
    label: 'ROUND 1',
    relevant: startPassed
  });

  survey.push({
    type: 'note',
    name: 'r1_intro',
    label: '📍 ROUND 1\n\nInstructions:\n1. Search the location to find the designated board/object.\n2. Take a mandatory photo of the object.\n3. Show the object/photo to the nearby Luminus volunteer.\n4. Enter the verification code given by the volunteer.\n\nEnter code in caps',
    hint: `Find the object, take photo, and ask volunteer for code. ${progressionWarning}`
  });

  survey.push({
    type: 'image',
    name: 'r1_photo',
    label: '📸 Upload Photo of Discovered Object (MANDATORY)\nPhoto upload is mandatory',
    hint: `Take a clear photo of the discovered object. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Photo upload of the object is strictly mandatory.'
  });

  survey.push({
    type: 'text',
    name: 'r1_code',
    label: 'Enter Volunteer Verification Code (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Volunteer verification code is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'YAAKE-GURU-5'",
    constraint_message: '❌ Incorrect code. Check the object again and enter the code exactly as displayed.'
  });

  survey.push({
    type: 'end_group'
  });

  const r1Passed = `${startPassed} and normalize-space(\${r1_code}) = 'YAAKE-GURU-5' and \${r1_photo} != ''`;

  // =============================================================
  // ROUND 2: LOCATION CLUES
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r2_aiml_loc_group',
    label: 'ROUND 2',
    relevant: r1Passed
  });

  // Clue 1: Algorithm (A)
  survey.push({
    type: 'note',
    name: 'r2_clue1_note',
    label: `📍 ROUND 2\n\nFour clues describe four concepts.\nSolve each clue slide-by-slide and enter the concept answer.\nAt the end, you will take the first letter of each answer to deduce your next destination block!\n\n🧩 Riddle 1 of 4:\n"Step by step, I show the way,\nsolving problems every day."\n\nEnter code in caps`,
    hint: `Solve Riddle 1 for the first concept. ${progressionWarning}`
  });

  survey.push({
    type: 'text',
    name: 'r2_clue1_answer',
    label: 'Enter Concept 1 Answer: (MANDATORY)\nEnter code in caps',
    hint: `ENTER ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Riddle 1 answer is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'ALGORITHM' or normalize-space(.) = 'ALGORITHMS')",
    constraint_message: '❌ Incorrect answer. Solve Riddle 1 and enter in UPPERCASE (CAPS ONLY).'
  });

  // Clue 2: Internet (I)
  survey.push({
    type: 'note',
    name: 'r2_clue2_note',
    label: `🧩 Riddle 2 of 4:\n"I connect the world without a sound,\nthrough wires and waves, I'm all around."\n\nEnter code in caps`,
    hint: `Solve Riddle 2 for the second concept. ${progressionWarning}`
  });

  survey.push({
    type: 'text',
    name: 'r2_clue2_answer',
    label: 'Enter Concept 2 Answer: (MANDATORY)\nEnter code in caps',
    hint: `ENTER ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Riddle 2 answer is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'INTERNET'",
    constraint_message: '❌ Incorrect answer. Solve Riddle 2 and enter in UPPERCASE (CAPS ONLY).'
  });

  // Clue 3: Model (M)
  survey.push({
    type: 'note',
    name: 'r2_clue3_note',
    label: `🧩 Riddle 3 of 4:\n"I learn from data, again and again,\ngetting smarter with every gain."\n\nEnter code in caps`,
    hint: `Solve Riddle 3 for the third concept. ${progressionWarning}`
  });

  survey.push({
    type: 'text',
    name: 'r2_clue3_answer',
    label: 'Enter Concept 3 Answer: (MANDATORY)\nEnter code in caps',
    hint: `ENTER ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Riddle 3 answer is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'MODEL' or normalize-space(.) = 'MODELS' or normalize-space(.) = 'MACHINE LEARNING')",
    constraint_message: '❌ Incorrect answer. Solve Riddle 3 and enter in UPPERCASE (CAPS ONLY).'
  });

  // Clue 4: Language (L)
  survey.push({
    type: 'note',
    name: 'r2_clue4_note',
    label: `🧩 Riddle 4 of 4:\n"I'm the code you write, line by line,\nPython or Java, the choice is fine."\n\nEnter code in caps`,
    hint: `Solve Riddle 4 for the fourth concept. ${progressionWarning}`
  });

  survey.push({
    type: 'text',
    name: 'r2_clue4_answer',
    label: 'Enter Concept 4 Answer: (MANDATORY)\nEnter code in caps',
    hint: `ENTER ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Riddle 4 answer is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'LANGUAGE' or normalize-space(.) = 'LANGUAGES' or normalize-space(.) = 'PROGRAMMING LANGUAGE')",
    constraint_message: '❌ Incorrect answer. Solve Riddle 4 and enter in UPPERCASE (CAPS ONLY).'
  });

  // Deduction Slide: First Letters -> AIML (No giveaway in question text)
  survey.push({
    type: 'note',
    name: 'r2_aiml_deduce_note',
    label: `🧠 DEDUCE THE NEXT DESTINATION BLOCK!\n\nNow, recall and take the FIRST LETTER of each of the 4 concept answers you just solved in sequence.\n\nCombine those four initial letters to deduce and reveal your next destination block!\n\nEnter code in caps`,
    hint: `Deduce the 4-letter block name from your answers. ${progressionWarning}`
  });

  survey.push({
    type: 'text',
    name: 'r2_aiml_loc_answer',
    label: 'Enter your 4-letter combined destination location: (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Destination location is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\/& ]+$') and (normalize-space(.) = 'AIML' or normalize-space(.) = 'AI ML' or normalize-space(.) = 'AIML BLOCK' or normalize-space(.) = 'AI/ML' or normalize-space(.) = 'AI & ML')",
    constraint_message: '❌ Incorrect. Take the first letter of each concept to find the block name.'
  });

  const r2LocGuessed = "normalize-space(${r2_aiml_loc_answer}) = 'AIML' or normalize-space(${r2_aiml_loc_answer}) = 'AI ML' or normalize-space(${r2_aiml_loc_answer}) = 'AIML BLOCK' or normalize-space(${r2_aiml_loc_answer}) = 'AI/ML' or normalize-space(${r2_aiml_loc_answer}) = 'AI & ML'";

  survey.push({
    type: 'note',
    name: 'r2_aiml_proceed_note',
    label: '🏃 Correct! Proceed to the destination block and report to the station volunteer to begin your challenge.\n\nEnter code in caps',
    hint: `Proceed to the station. ${progressionWarning}`,
    relevant: r2LocGuessed
  });

  survey.push({
    type: 'end_group'
  });

  // =============================================================
  // ROUND 2: AIML CHALLENGES (Reverse Image Prompting)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r2_aiml_challenge_group',
    label: 'ROUND 2',
    relevant: r2LocGuessed
  });

  survey.push({
    type: 'text',
    name: 'r2_aiml_start_code',
    label: 'Enter Challenge START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'GEMMA-V05'",
    constraint_message: '❌ Incorrect START CODE. Enter GEMMA-V05 in CAPS provided by the volunteer.'
  });

  const aimlStartPassed = `${r1Passed} and (${r2LocGuessed}) and normalize-space(\${r2_aiml_start_code}) = 'GEMMA-V05'`;

  survey.push({
    type: 'note',
    name: 'r2_aiml_prompt_rules_note',
    label: `🎨 ROUND 2\n\nEvent Format:\n• One lab system will be provided per team.\n• Images will be displayed for 10 seconds each in a slideshow.\n\nRules & Regulations:\n• Each team must recreate at least 3 images from the slideshow.\n• The 3 images must be created on 3 different AI chats / tabs - not in the same chat.\n• Teams can use any AI image generation tool (ChatGPT, Midjourney, Gemini, Leonardo, etc.).\n• Google Image Search, reverse image search, or uploading / scanning the displayed image is strictly prohibited. Direct disqualification if found.\n• Only original prompting is allowed - img-to-img is not allowed.\n• Volunteers will come to your system and verify your generated images directly.\n\nEnter code in caps`,
    hint: `Follow lab instructions and volunteer guidance. ${progressionWarning}`,
    relevant: aimlStartPassed
  });

  survey.push({
    type: 'image',
    name: 'r2_aiml_photo1',
    label: '📸 Upload Photo 1 of Generated Images / Station Screen (MANDATORY)\nPhoto upload is mandatory',
    hint: `Take a clear photo of your team's generated images on the lab screen. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Station verification photo 1 is mandatory.',
    relevant: aimlStartPassed
  });

  survey.push({
    type: 'image',
    name: 'r2_aiml_photo2',
    label: '📸 Upload Photo 2 of Additional Generated Images (OPTIONAL)',
    hint: `Optional: Upload a second photo of your team's generated images. ${progressionWarning}`,
    required: 'no',
    relevant: aimlStartPassed
  });

  survey.push({
    type: 'text',
    name: 'r2_aiml_end_code',
    label: 'Enter Volunteer FINAL Clearance Code (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Station final clearance code is mandatory.',
    relevant: aimlStartPassed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'OLLAMA-V05' or normalize-space(.) = 'OLLAMA-v05')",
    constraint_message: '❌ Incorrect final code. Enter OLLAMA-V05 in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r2AllPassed = `${aimlStartPassed} and (normalize-space(\${r2_aiml_end_code}) = 'OLLAMA-V05' or normalize-space(\${r2_aiml_end_code}) = 'OLLAMA-v05') and \${r2_aiml_photo1} != ''`;

  // =============================================================
  // ROUND 3: LOCATION CLUE (Destination: CSE)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_cse_loc_group',
    label: 'ROUND 3',
    relevant: r2AllPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_cse_math_note',
    label: `📍 ROUND 3\n\nThree numbers are hidden inside the equations.\nSolve each equation.\n\nThen use the alphabet as your key:\nA = 1, B = 2, C = 3 ... Z = 26\n\nWhat three-letter code do you uncover?\n\nCode:\n(1 + 2) — (20 − 1) — (10 ÷ 2)\n\nEnter code in caps`,
    hint: `Solve each equation and map the 3 numbers to letters A=1 to Z=26. ${progressionWarning}`
  });

  survey.push({
    type: 'text',
    name: 'r3_cse_loc_answer',
    label: 'Enter your uncovered 3-letter destination location: (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
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
    hint: `Proceed to the location for the QR hunt. ${progressionWarning}`,
    relevant: r3LocGuessed
  });

  survey.push({
    type: 'end_group'
  });

  const r3BlockPassed = `${r2AllPassed} and (${r3LocGuessed})`;

  // =============================================================
  // ROUND 3: QR HUNT (CSE) — Photo Removed, Link Removed
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_cse_qr_group',
    label: 'ROUND 3',
    relevant: r3BlockPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_cse_qr_note',
    label: '📍 ROUND 3\n\nFind the designated QR code at this location and scan it.\nFollow the instructions provided by the QR challenge.\n\nEnter code in caps',
    hint: `Locate and scan the hidden QR code. ${progressionWarning}`
  });

  survey.push({
    type: 'barcode',
    name: 'r3_cse_qr_scan',
    label: 'Scan Discovered QR Code (MANDATORY)',
    hint: `Scan the QR code found at the location. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ QR code scanning is mandatory.',
    constraint: "normalize-space(.) = 'PETER_PARKER' or contains(., 'PETER_PARKER')",
    constraint_message: '❌ Incorrect code. Continue the QR hunt and check the code carefully.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3QrPassed = `${r3BlockPassed} and (normalize-space(\${r3_cse_qr_scan}) = 'PETER_PARKER' or contains(\${r3_cse_qr_scan}, 'PETER_PARKER'))`;

  // =============================================================
  // ROUND 4: LOCATION CLUE (Piece-by-Piece)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_mech_loc_group',
    label: 'ROUND 4',
    relevant: r3QrPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_mech_piece_note',
    label: `📍 ROUND 4\n\nPiece by piece, an image has been hidden from you.\nStudy the available pieces in the image attached below and determine what place or object they form.\nYour answer will help identify your next location.\n\nEnter code in caps`,
    hint: `Examine the image pieces to identify the destination. ${progressionWarning}`,
    'media::image': 'badminton court.jpeg'
  });

  survey.push({
    type: 'text',
    name: 'r4_mech_loc_answer',
    label: 'Enter the location/object identified from the photograph pieces: (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
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
    label: '🏃 Correct! Proceed to the destination block and report to the station volunteer for your physical challenge.\n\nEnter code in caps',
    hint: `Proceed to the station. ${progressionWarning}`,
    relevant: r4LocGuessed
  });

  survey.push({
    type: 'end_group'
  });

  // =============================================================
  // ROUND 4: PHYSICAL CHALLENGE
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_mech_phy_group',
    label: 'ROUND 4',
    relevant: r4LocGuessed
  });

  survey.push({
    type: 'text',
    name: 'r4_mech_start_code',
    label: 'Enter Physical Challenge START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'PIECE-START' or normalize-space(.) = 'PIECE_START')",
    constraint_message: '❌ Incorrect START CODE. Enter PIECE-START in CAPS provided by the volunteer.'
  });

  const mechStartPassed = `${r3QrPassed} and (${r4LocGuessed}) and (normalize-space(\${r4_mech_start_code}) = 'PIECE-START' or normalize-space(\${r4_mech_start_code}) = 'PIECE_START')`;

  survey.push({
    type: 'note',
    name: 'r4_mech_phy_note',
    label: '📍 ROUND 4\n\nReport to the station and complete the physical challenge under volunteer supervision.\nOnce completed, take a team photo and ask the volunteer for the END CODE.\n\nEnter code in caps',
    hint: `Complete physical challenge with volunteer. ${progressionWarning}`,
    relevant: mechStartPassed
  });

  survey.push({
    type: 'image',
    name: 'r4_mech_phy_photo',
    label: '📸 Upload Photo of Physical Challenge Completion (MANDATORY)\nPhoto upload is mandatory',
    hint: `Take a clear photo of your team completing the physical challenge. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Challenge photo upload is mandatory.',
    relevant: mechStartPassed
  });

  survey.push({
    type: 'text',
    name: 'r4_mech_phy_code',
    label: 'Enter Volunteer END / Completion Code (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Volunteer completion code is mandatory.',
    relevant: mechStartPassed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'PIECE-BYE' or normalize-space(.) = 'PIECE-BYE---' or normalize-space(.) = 'PIECE_BYE')",
    constraint_message: '❌ Incorrect code. Enter PIECE-BYE in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4PhyPassed = `${mechStartPassed} and (normalize-space(\${r4_mech_phy_code}) = 'PIECE-BYE' or normalize-space(\${r4_mech_phy_code}) = 'PIECE-BYE---' or normalize-space(\${r4_mech_phy_code}) = 'PIECE_BYE') and \${r4_mech_phy_photo} != ''`;

  // =============================================================
  // FINAL ROUND (ROUND 5): RIDDLES & GRAND FINALE
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'final_puzzles_group',
    label: 'ROUND 5',
    relevant: r4PhyPassed
  });

  survey.push({
    type: 'note',
    name: 'final_audi_stage_intro',
    label: '🏛️ ROUND 5\n\nSolve the two final riddles to reveal the final destination!\n\nEnter code in caps',
    hint: `Solve the final riddles in CAPS. ${progressionWarning}`
  });

  survey.push({
    type: 'text',
    name: 'final_riddle1_answer',
    label: `"I am empty, yet I am built for crowds.
I have a stage, but no actors of my own.
I have countless seats, but none are meant to sleep.
When a voice rises before me, silence falls behind me.
When the lights awaken, all eyes face one direction.
What am I?"

Enter code in caps`,
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Solving Riddle 1 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'AUDITORIUM' or normalize-space(.) = 'AUDI')",
    constraint_message: '❌ Incorrect answer. Solve Riddle 1 and enter in UPPERCASE (CAPS ONLY).'
  });

  const finalRiddle1Passed = `${r4PhyPassed} and (normalize-space(\${final_riddle1_answer}) = 'AUDITORIUM' or normalize-space(\${final_riddle1_answer}) = 'AUDI')`;

  survey.push({
    type: 'text',
    name: 'final_riddle2_stage_answer',
    label: `"I am elevated above the crowd, where performers stand and spotlights shine.
Underneath my wooden floor or behind the curtains, the ultimate secret waits.
What am I?"

Enter code in caps`,
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
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
    hint: `Go to the stage to solve the final puzzle. ${progressionWarning}`,
    relevant: finalRiddle2Passed
  });

  survey.push({
    type: 'image',
    name: 'final_solved_puzzle_photo',
    label: '📸 Upload Photo of Your Solved Puzzle (MANDATORY)\nPhoto upload is mandatory',
    hint: `Take a clear photo of your team's completed/solved puzzle. ${progressionWarning}`,
    required: 'yes',
    required_message: '❌ Photo of the solved puzzle is mandatory.',
    relevant: finalRiddle2Passed
  });

  survey.push({
    type: 'text',
    name: 'final_stage_volunteer_code',
    label: 'Enter Final Volunteer Clearance Code (MANDATORY)\nEnter code in caps',
    hint: `ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY. ${progressionWarning}`,
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
  return [];
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
      'Location': 'Station 1 (Library)',
      'Challenge / Item': 'Object Finding',
      'Question / Prompt': 'Find assigned object, upload photo, enter volunteer code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'YAAKE-GURU-5 (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'YAAKE-GURU-5\'',
      'Mandatory Upload': 'Yes (Discovered Object Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 2 Location Clues',
      'Location': 'In-App Riddles',
      'Challenge / Item': 'Missing Concept Riddles (Slide-by-Slide)',
      'Question / Prompt': '1. ALGORITHM (A) | 2. INTERNET (I) | 3. MODEL (M) | 4. LANGUAGE (L) -> Deduced: AIML',
      'Media Attached': 'None',
      'Expected Answer / Code': 'ALGORITHM, INTERNET, MODEL, LANGUAGE -> Destination: AIML (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) for all 4 concepts and AIML',
      'Mandatory Upload': 'No (Inputs Mandatory)'
    },
    {
      'Stage / Round': 'Round 2 Challenge (R2)',
      'Location': 'AIML Lab',
      'Challenge / Item': 'Reverse Image Prompting Challenge',
      'Question / Prompt': 'Start Code: GEMMA-V05 -> Recreate 3 slideshow images in 3 AI chats + Photo (1 Mandatory, 1 Optional) -> Final Code: OLLAMA-V05',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Start: GEMMA-V05 -> Photo -> Final Code: OLLAMA-V05',
      'Verification / Constraint Rule': 'Start Code = \'GEMMA-V05\' & Final Code = \'OLLAMA-V05\'',
      'Mandatory Upload': 'Yes (Photo 1 Mandatory, Photo 2 Optional)'
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
      'Question / Prompt': 'Scan hidden QR code in CSE',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Barcode: PETER_PARKER (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Scanned barcode = \'PETER_PARKER\'',
      'Mandatory Upload': 'Yes (Barcode Scan MANDATORY)'
    },
    {
      'Stage / Round': 'Round 4 Location Clue',
      'Location': 'In-App Image (badminton court.jpeg)',
      'Challenge / Item': 'Piece-by-Piece Clue',
      'Question / Prompt': 'Identify location from photograph pieces -> Destination: MECH',
      'Media Attached': 'badminton court.jpeg',
      'Expected Answer / Code': 'MECH (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'MECH\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 4 Checkpoint (R4)',
      'Location': 'MECH',
      'Challenge / Item': 'Physical Challenge',
      'Question / Prompt': 'Enter Start Code: PIECE-START, complete physical challenge, upload photo, enter End Code: PIECE-BYE',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Start: PIECE-START -> Photo -> End Code: PIECE-BYE (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'normalize-space(.) = \'PIECE-START\' & \'PIECE-BYE\'',
      'Mandatory Upload': 'Yes (Challenge Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 5 (Riddle 1)',
      'Location': 'Main Auditorium',
      'Challenge / Item': 'Riddle 1',
      'Question / Prompt': 'Solve riddle: Empty yet built for crowds, seats not meant to sleep...',
      'Media Attached': 'None',
      'Expected Answer / Code': 'AUDITORIUM (or AUDI) (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'AUDITORIUM\' or \'AUDI\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 5 (Riddle 2)',
      'Location': 'Main Auditorium Stage',
      'Challenge / Item': 'Riddle 2',
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
  .start-code-badge {
    display: inline-block;
    background: #e0f2fe;
    color: #0369a1;
    border: 1px solid #7dd3fc;
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
    CONFIDENTIAL ORGANIZER MASTER KEY &bull; PROGRESSION: STATION 1 &rarr; AIML &rarr; CSE &rarr; MECH &rarr; AUDITORIUM
  </div>
</div>

<div class="meta-grid">
  <div class="meta-card">
    <strong>Starting Station</strong>
    <span>Station 1 (Library)</span>
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
  <strong>🔒 Strict Policy:</strong> Every question, photo upload, barcode scan, and code entry is <strong>100% MANDATORY</strong> and requires <strong>STRICT UPPERCASE ONLY</strong>. Progression is irreversible once submitted.
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
      <td>Station 1 (Library)</td>
      <td>Find assigned object &amp; verify with volunteer</td>
      <td><span class="code-badge">YAAKE-GURU-5</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Round 2 (R2)</strong></td>
      <td>In-App</td>
      <td>
        Missing Concept (Slide-by-Slide):<br>
        1. <code>ALGORITHM</code> (A)<br>
        2. <code>INTERNET</code> (I)<br>
        3. <code>MODEL</code> (M)<br>
        4. <code>LANGUAGE</code> (L)<br>
        &rarr; Deduced Destination: <code>AIML</code>
      </td>
      <td>Destination: <code>AIML</code></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>Round 2 Challenge</strong></td>
      <td>AIML Lab</td>
      <td>Reverse Image Prompting (Recreate 3 images in 3 AI chats)</td>
      <td>Start: <span class="start-code-badge">GEMMA-V05</span><br>Final Code: <span class="code-badge">OLLAMA-V05</span></td>
      <td>📸 Mandatory Photo 1 (+ Opt. Photo 2)</td>
    </tr>
    <tr>
      <td><strong>Round 3 (R3)</strong></td>
      <td>In-App</td>
      <td>Mathematical Code: (1+2)=3(C), (20-1)=19(S), (10/2)=5(E)</td>
      <td>Destination: <code>CSE</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>Round 3 QR</strong></td>
      <td>CSE</td>
      <td>Physically search CSE for hidden QR code and scan</td>
      <td>Barcode: <span class="code-badge">PETER_PARKER</span></td>
      <td>Scan Only</td>
    </tr>
    <tr>
      <td><strong>Round 4 (R4)</strong></td>
      <td>In-App (<code>badminton court.jpeg</code>)</td>
      <td>Piece-by-Piece Location Clue</td>
      <td>Destination: <code>MECH</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>Round 4 Challenge</strong></td>
      <td>MECH</td>
      <td>Complete physical challenge with volunteer</td>
      <td>Start: <span class="start-code-badge">PIECE-START</span><br>Finish Code: <span class="code-badge">PIECE-BYE</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>Round 5 (Final)</strong></td>
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
  <h3><span>Station 1 (Library): Object Finding</span><span class="code-badge">CODE: YAAKE-GURU-5</span></h3>
  <div><strong>Volunteer Instructions:</strong> Participants locate the assigned object. Once verified with photo, provide code <code>YAAKE-GURU-5</code>.</div>
</div>

<div class="station-card">
  <h3><span>AIML Lab Station: Reverse Image Prompting</span><span class="code-badge">START: GEMMA-V05 &bull; FINAL: OLLAMA-V05</span></h3>
  <div><strong>Volunteer Instructions:</strong> Provide start code <code>GEMMA-V05</code> to begin. Participants recreate at least 3 slideshow images in 3 different AI tabs. Once verified with screen photo upload, provide final clearance code <code>OLLAMA-V05</code>.</div>
</div>

<div class="station-card">
  <h3><span>CSE Station: QR Hunt</span><span class="code-badge">SCANNED QR: PETER_PARKER</span></h3>
  <div><strong>Volunteer Instructions:</strong> Ensure the physical QR code with payload <code>PETER_PARKER</code> is hidden in the CSE area. Participants scan it using the in-app scanner.</div>
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

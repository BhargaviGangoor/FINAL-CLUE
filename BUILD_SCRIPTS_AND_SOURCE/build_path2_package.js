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
    .map(w => `normalize-space(${varName}) = '${w}'`)
    .join(' or ');
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
    label: `🏆 FINAL CLUE — PATH 2
TREASURE HUNT FOR FRESHERS 2026

Welcome to PATH 2 of the Final Clue Treasure Hunt!

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
  // ROUND 1 — OBJECT FINDING (ADMIN)
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
    label: '📍 ROUND 1: OBJECT FINDING\n\nInstructions:\n1. Find the assigned object.\n2. Take a mandatory photo of the object.\n3. Show the object/photo to a nearby Luminus ID-card volunteer.\n4. Enter the Round 1 completion code given by the volunteer.\n\nEnter code in caps',
    hint: 'Find the assigned object, take photo, and ask volunteer for code.'
  });

  survey.push({
    type: 'image',
    name: 'r1_admin_photo',
    label: '📸 Upload Photo of the Discovered Hidden Object (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered object.',
    required: 'yes',
    required_message: '❌ Photo upload of the object is mandatory.'
  });

  survey.push({
    type: 'text',
    name: 'r1_admin_code',
    label: 'Enter Volunteer Verification Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer verification code is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'JOHN-CENA'",
    constraint_message: '❌ Incorrect code. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r1Passed = `${startPassed} and normalize-space(\${r1_admin_code}) = 'JOHN-CENA' and \${r1_admin_photo} != ''`;

  // =============================================================
  // ROUND 2 — LOCATION CLUE (NO CHALLENGE NAME)
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
    label: `At 14:30, the security team detected an unusual login attempt on the college network. The account was accessed from an unknown device shortly after the user received a suspicious email. The security team immediately changed the account credentmials and checked the systbem logs for unusual activity. No confidential files appeared to have been downloaded during the incident. The affected user was advised to enable multi-factor authentication and avoid opening links from unknown senders. The incident was then reported to the network administrataor for further investigation.The report has been tampered with. The intruder left three traces behind.

Find the words that don't belong and recover what was hidden: credentials → credentmials, system → systbem ,administrator → administrataor The inserted letters are:

Enter code in caps`,
    hint: 'Deduce the inserted letters and enter the destination.'
  });

  survey.push({
    type: 'text',
    name: 'r2_mba_loc_answer',
    label: 'Enter destination location (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Destination location is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'MBA' or normalize-space(.) = 'MBA BLOCK')",
    constraint_message: '❌ Incorrect destination. Read the clue carefully and enter in UPPERCASE (CAPS ONLY).'
  });

  const r2LocGuessed = "normalize-space(${r2_mba_loc_answer}) = 'MBA' or normalize-space(${r2_mba_loc_answer}) = 'MBA BLOCK'";

  survey.push({
    type: 'note',
    name: 'r2_proceed_note',
    label: '🏃 Go to the location you identified and ask the Luminus volunteer there for the START CODE.\n\nEnter code in caps',
    hint: 'Go to the location and ask volunteer for start code.',
    relevant: r2LocGuessed
  });

  survey.push({
    type: 'text',
    name: 'r2_start_code',
    label: 'Enter START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    relevant: r2LocGuessed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'MBA-START'",
    constraint_message: '❌ Incorrect START CODE. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r2StartPassed = `${r1Passed} and (${r2LocGuessed}) and normalize-space(\${r2_start_code}) = 'MBA-START'`;

  // =============================================================
  // ROUND 2 — MINI CHALLENGE (VARIANT A & VARIANT B)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'mba_challenge_group',
    label: 'ROUND 2 — MINI CHALLENGE',
    relevant: r2StartPassed
  });

  survey.push({
    type: 'select_one mba_variants',
    name: 'mba_variant_select',
    label: 'Select your assigned challenge variant: (MANDATORY)\nMandatory selection',
    hint: 'Choose Variant A or Variant B as assigned by the volunteer.',
    required: 'yes',
    required_message: '❌ Selecting your assigned variant is mandatory.'
  });

  // Variant A: MBA Quickfire
  survey.push({
    type: 'begin_group',
    name: 'mba_v1_group',
    label: 'VARIANT A',
    relevant: "${mba_variant_select} = 'var_a'"
  });

  survey.push({
    type: 'note',
    name: 'mba_v1_intro',
    label: '📊 VARIANT A: MBA QUICKFIRE\n\nAnswer all 10 quickfire questions below.\n\nEnter code in caps',
    hint: 'Answer all 10 questions in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'mba_q1',
    label: '1. A market with only one seller is called? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 1 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'MONOPOLY'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q2',
    label: '2. What is the currency of China? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 2 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'YUAN'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q3',
    label: '3. What does CEO stand for? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 3 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'CHIEF EXECUTIVE OFFICER'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q4',
    label: '4. What does ROI stand for in business finance? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 4 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'RETURN ON INVESTMENT'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q5',
    label: '5. What term describes business transactions conducted between two companies (abbreviation)? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 5 is mandatory.',
    constraint: "regex(., '^[A-Z0-9]+$') and normalize-space(.) = 'B2B'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q6',
    label: '6. In accounting: Assets minus Liabilities equals what? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 6 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'EQUITY'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q7',
    label: '7. Which animal represents a rising, optimistic financial market? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 7 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'BULL'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q8',
    label: '8. What does IPO stand for when a company goes public? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 8 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'INITIAL PUBLIC OFFERING'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q9',
    label: '9. What is the standard 3-letter abbreviation for Gross Domestic Product? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 9 is mandatory.',
    constraint: "regex(., '^[A-Z]+$') and normalize-space(.) = 'GDP'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'mba_q10',
    label: '10. In the 4 Ps of Marketing (Product, Price, Place), what is the 4th P? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 10 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'PROMOTION'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  const mba1AllPassed = "(normalize-space(${mba_q1}) = 'MONOPOLY' and normalize-space(${mba_q2}) = 'YUAN' and normalize-space(${mba_q3}) = 'CHIEF EXECUTIVE OFFICER' and normalize-space(${mba_q4}) = 'RETURN ON INVESTMENT' and normalize-space(${mba_q5}) = 'B2B' and normalize-space(${mba_q6}) = 'EQUITY' and normalize-space(${mba_q7}) = 'BULL' and normalize-space(${mba_q8}) = 'INITIAL PUBLIC OFFERING' and normalize-space(${mba_q9}) = 'GDP' and normalize-space(${mba_q10}) = 'PROMOTION')";

  survey.push({
    type: 'text',
    name: 'r2_var_a_code',
    label: 'Enter Completion Code from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Completion code is mandatory.',
    relevant: mba1AllPassed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'MBA-QF-1'",
    constraint_message: '❌ Incorrect completion code. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  // Variant B: Brand Quiz
  survey.push({
    type: 'begin_group',
    name: 'mba_v2_group',
    label: 'VARIANT B',
    relevant: "${mba_variant_select} = 'var_b'"
  });

  survey.push({
    type: 'note',
    name: 'mba_v2_intro',
    label: '🏷️ VARIANT B: BRAND QUIZ\n\nAnswer all 10 brand quiz questions below.\n\nEnter code in caps',
    hint: 'Answer all 10 brand questions in CAPS.'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q1',
    label: '1. Name of dessert? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 1 is mandatory.',
    'media::image': 'bambaloni.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'BAMBALONI'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q2',
    label: '2. Which brand is associated with this tagline? “THINK DIFFERENT” (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 2 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'APPLE'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q3',
    label: '3. What is the name of the cartoon this character represents? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 3 is mandatory.',
    'media::image': 'horrid henry.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'HORRID HENRY'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q4',
    label: '4. Identify the brand from this famous ad screenshot. (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 4 is mandatory.',
    'media::image': '5 star.jpeg',
    constraint: "regex(., '^[A-Z0-9 ]+$') and (normalize-space(.) = '5 STAR' or normalize-space(.) = '5STAR' or normalize-space(.) = 'FIVE STAR')",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q5',
    label: '5. Can you identify the brand from this zoomed-in logo? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 5 is mandatory.',
    'media::image': 'nokia.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'NOKIA'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q6',
    label: '6. What is the name of this fruit? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 6 is mandatory.',
    'media::image': 'durian.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'DURIAN'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q7',
    label: '7. Which brand is associated with this tagline? “WHEREVER YOU GO, OUR NETWORK FOLLOWS” (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 7 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'VODAFONE'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q8',
    label: '8. Which brand does this mascot represent? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 8 is mandatory.',
    'media::image': 'michelin.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'MICHELIN'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q9',
    label: '9. What is the full name of RN Shetty? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 9 is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\.\\- ]+$') and (normalize-space(.) = 'DR. RAMA NAGAPPA SHETTY' or normalize-space(.) = 'RAMA NAGAPPA SHETTY' or normalize-space(.) = 'DR RAMA NAGAPPA SHETTY')",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r2_b_q10',
    label: '10. This actress holds a brand of herself. What is the brand name? (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Question 10 is mandatory.',
    'media::image': 'palmonas.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'PALMONAS'",
    constraint_message: '❌ Incorrect answer. Please enter in UPPERCASE (CAPS ONLY).'
  });

  const mba2AllPassed = "(normalize-space(${r2_b_q1}) = 'BAMBALONI' and normalize-space(${r2_b_q2}) = 'APPLE' and normalize-space(${r2_b_q3}) = 'HORRID HENRY' and (normalize-space(${r2_b_q4}) = '5 STAR' or normalize-space(${r2_b_q4}) = '5STAR' or normalize-space(${r2_b_q4}) = 'FIVE STAR') and normalize-space(${r2_b_q5}) = 'NOKIA' and normalize-space(${r2_b_q6}) = 'DURIAN' and normalize-space(${r2_b_q7}) = 'VODAFONE' and normalize-space(${r2_b_q8}) = 'MICHELIN' and (normalize-space(${r2_b_q9}) = 'DR. RAMA NAGAPPA SHETTY' or normalize-space(${r2_b_q9}) = 'RAMA NAGAPPA SHETTY' or normalize-space(${r2_b_q9}) = 'DR RAMA NAGAPPA SHETTY') and normalize-space(${r2_b_q10}) = 'PALMONAS')";

  survey.push({
    type: 'text',
    name: 'r2_var_b_code',
    label: 'Enter Completion Code from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Completion code is mandatory.',
    relevant: mba2AllPassed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'MB-BQ-2'",
    constraint_message: '❌ Incorrect completion code. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  survey.push({
    type: 'end_group'
  });

  const r2ClearedRel = "((${mba_variant_select} = 'var_a' and normalize-space(${r2_var_a_code}) = 'MBA-QF-1') or (${mba_variant_select} = 'var_b' and normalize-space(${r2_var_b_code}) = 'MB-BQ-2'))";

  // =============================================================
  // ROUND 3 — LOCATION CLUE (NO CHALLENGE NAME)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_lib_loc_group',
    label: 'ROUND 3',
    relevant: `${r2StartPassed} and ${r2ClearedRel}`
  });

  survey.push({
    type: 'note',
    name: 'r3_shuffled_intro_note',
    label: '📍 ROUND 3\n\nDecode each of the shuffled words below.\n\nE.g., OOBK ➔ BOOK, APPRE ➔ PAPER\n\nEnter code in caps',
    hint: 'Decode each shuffled word.'
  });

  survey.push({
    type: 'text',
    name: 'r3_decode_w1',
    label: 'Decode Shuffled Word 1: [ VEHSELS ] (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Decoding Word 1 is mandatory.',
    constraint: "regex(., '^[A-Z]+$') and normalize-space(.) = 'SHELVES'",
    constraint_message: '❌ Incorrect decoded word. Enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r3_decode_w2',
    label: 'Decode Shuffled Word 2: [ GIDRAEN ] (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Decoding Word 2 is mandatory.',
    constraint: "regex(., '^[A-Z]+$') and normalize-space(.) = 'READING'",
    constraint_message: '❌ Incorrect decoded word. Enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r3_decode_w3',
    label: 'Decode Shuffled Word 3: [ IFIW ] (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Decoding Word 3 is mandatory.',
    constraint: "regex(., '^[A-Z]+$') and normalize-space(.) = 'WIFI'",
    constraint_message: '❌ Incorrect decoded word. Enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r3_decode_w4',
    label: 'Decode Shuffled Word 4: [ SISCUDISNO GGUONLE ] (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Decoding Word 4 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'DISCUSSION LOUNGE'",
    constraint_message: '❌ Incorrect decoded word. Enter in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r3_decode_w5',
    label: 'Decode Shuffled Word 5: [ IGIDLAT BILRYRA ] (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Decoding Word 5 is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and normalize-space(.) = 'DIGITAL LIBRARY'",
    constraint_message: '❌ Incorrect decoded word. Enter in UPPERCASE (CAPS ONLY).'
  });

  const r3AllDecoded = "(normalize-space(${r3_decode_w1}) = 'SHELVES' and normalize-space(${r3_decode_w2}) = 'READING' and normalize-space(${r3_decode_w3}) = 'WIFI' and normalize-space(${r3_decode_w4}) = 'DISCUSSION LOUNGE' and normalize-space(${r3_decode_w5}) = 'DIGITAL LIBRARY')";

  survey.push({
    type: 'text',
    name: 'r3_lib_loc_answer',
    label: 'Decoded Clues:\n1. ${r3_decode_w1}\n2. ${r3_decode_w2}\n3. ${r3_decode_w3}\n4. ${r3_decode_w4}\n5. ${r3_decode_w5}\n\nBased on your decoded words, identify the next location: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Location answer is mandatory.',
    relevant: r3AllDecoded,
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'LIBRARY' or normalize-space(.) = 'CENTRAL LIBRARY')",
    constraint_message: '❌ Incorrect location. Enter the location in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'end_group'
  });

  const r3LocPassed = `${r2StartPassed} and ${r2ClearedRel} and ${r3AllDecoded} and (normalize-space(\${r3_lib_loc_answer}) = 'LIBRARY' or normalize-space(\${r3_lib_loc_answer}) = 'CENTRAL LIBRARY')`;

  // =============================================================
  // ROUND 3 — QR HUNT (LIBRARY)
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
    label: '📍 ROUND 3: QR HUNT\n\nGo to the location you identified and ask the Luminus volunteer for the next instruction/code.\n\nScan the available QR codes and identify the correct one in a fun manner!\n\nEnter code in caps',
    hint: 'Locate and scan the correct QR code.'
  });

  survey.push({
    type: 'barcode',
    name: 'r3_lib_qr_scan',
    label: 'Scan Discovered QR Code (MANDATORY)',
    hint: 'Scan the correct discovered QR code.',
    required: 'yes',
    required_message: '❌ QR scan is mandatory.',
    constraint: "normalize-space(.) = 'DUMB_FAKE'",
    constraint_message: '❌ Incorrect QR code scanned. Search for the correct QR code at this station.'
  });

  const r3QrScanned = `${r3LocPassed} and normalize-space(\${r3_lib_qr_scan}) = 'DUMB_FAKE'`;

  survey.push({
    type: 'text',
    name: 'r3_lib_volunteer_code',
    label: 'Enter Volunteer Completion Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer completion code is mandatory.',
    relevant: r3QrScanned,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'DUMB_FAKE' or normalize-space(.) = 'LIB-HUNT-GOOD' or normalize-space(.) = 'LIB-PASS' or normalize-space(.) = 'QR-LIB-2')",
    constraint_message: '❌ Incorrect code. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3Completed = `${r3QrScanned} and (normalize-space(\${r3_lib_volunteer_code}) = 'DUMB_FAKE' or normalize-space(\${r3_lib_volunteer_code}) = 'LIB-HUNT-GOOD' or normalize-space(\${r3_lib_volunteer_code}) = 'LIB-PASS' or normalize-space(\${r3_lib_volunteer_code}) = 'QR-LIB-2')`;

  // =============================================================
  // ROUND 4 — LOCATION CLUE (AIML WORD SEARCH)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_wordsearch_group',
    label: 'ROUND 4 LOCATION CLUE',
    relevant: r3Completed
  });

  survey.push({
    type: 'note',
    name: 'r4_wordsearch_intro',
    label: `Examine the puzzle image below.
Find and enter ANY 5 words from the grid amongst:
• MACHINE LEARNING
• NEURAL NETWORK
• PYTHON
• DATASET
• ALGORITHM
• TRAINING DATA
• DEEP LEARNING
• DATA MINING
• MODELING
• REGRESSION
• CLASSIFY

Enter code in caps`,
    hint: 'Find 5 words in the image and enter below in CAPS.',
    'media::image': 'aiml.jpeg'
  });

  const wCond = getWordSearchCondition('.');

  survey.push({
    type: 'text',
    name: 'r4_ws_word1',
    label: 'Enter Discovered Word 1: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Word 1 is mandatory.',
    'media::image': 'aiml.jpeg',
    constraint: `regex(., '^[A-Z ]+$') and (${wCond})`,
    constraint_message: '❌ Invalid word. Enter a valid word found in the puzzle in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r4_ws_word2',
    label: 'Enter Discovered Word 2: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Word 2 is mandatory.',
    'media::image': 'aiml.jpeg',
    constraint: `regex(., '^[A-Z ]+$') and (${wCond}) and normalize-space(.) != normalize-space(\${r4_ws_word1})`,
    constraint_message: '❌ Invalid or duplicate word. Enter a different valid word from the puzzle in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r4_ws_word3',
    label: 'Enter Discovered Word 3: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Word 3 is mandatory.',
    'media::image': 'aiml.jpeg',
    constraint: `regex(., '^[A-Z ]+$') and (${wCond}) and normalize-space(.) != normalize-space(\${r4_ws_word1}) and normalize-space(.) != normalize-space(\${r4_ws_word2})`,
    constraint_message: '❌ Invalid or duplicate word. Enter a different valid word from the puzzle in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r4_ws_word4',
    label: 'Enter Discovered Word 4: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Word 4 is mandatory.',
    'media::image': 'aiml.jpeg',
    constraint: `regex(., '^[A-Z ]+$') and (${wCond}) and normalize-space(.) != normalize-space(\${r4_ws_word1}) and normalize-space(.) != normalize-space(\${r4_ws_word2}) and normalize-space(.) != normalize-space(\${r4_ws_word3})`,
    constraint_message: '❌ Invalid or duplicate word. Enter a different valid word from the puzzle in UPPERCASE (CAPS ONLY).'
  });

  survey.push({
    type: 'text',
    name: 'r4_ws_word5',
    label: 'Enter Discovered Word 5: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Word 5 is mandatory.',
    'media::image': 'aiml.jpeg',
    constraint: `regex(., '^[A-Z ]+$') and (${wCond}) and normalize-space(.) != normalize-space(\${r4_ws_word1}) and normalize-space(.) != normalize-space(\${r4_ws_word2}) and normalize-space(.) != normalize-space(\${r4_ws_word3}) and normalize-space(.) != normalize-space(\${r4_ws_word4})`,
    constraint_message: '❌ Invalid or duplicate word. Enter a different valid word from the puzzle in UPPERCASE (CAPS ONLY).'
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
    label: 'Deduce the location related to these words: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Location answer is mandatory.',
    'media::image': 'aiml.jpeg',
    relevant: all5WordsValid,
    constraint: "regex(., '^[A-Z0-9\\/& ]+$') and (normalize-space(.) = 'AIML' or normalize-space(.) = 'AI ML' or normalize-space(.) = 'AIML BLOCK' or normalize-space(.) = 'AI/ML' or normalize-space(.) = 'AI & ML')",
    constraint_message: '❌ Incorrect destination. Enter the location in UPPERCASE (CAPS ONLY).'
  });

  const r4DestIdentified = "normalize-space(${r4_aiml_destination}) = 'AIML' or normalize-space(${r4_aiml_destination}) = 'AI ML' or normalize-space(${r4_aiml_destination}) = 'AIML BLOCK' or normalize-space(${r4_aiml_destination}) = 'AI/ML' or normalize-space(${r4_aiml_destination}) = 'AI & ML'";

  survey.push({
    type: 'note',
    name: 'r4_proceed_note',
    label: '🏃 Go to the location you identified and ask the Luminus volunteer for the START CODE.\n\nEnter code in caps',
    hint: 'Go to the location and ask volunteer for start code.',
    relevant: `${all5WordsValid} and (${r4DestIdentified})`
  });

  survey.push({
    type: 'text',
    name: 'r4_start_code',
    label: 'Enter START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    relevant: `${all5WordsValid} and (${r4DestIdentified})`,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'START-AI-PHY' or normalize-space(.) = 'START-PHY')",
    constraint_message: '❌ Incorrect START CODE. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4StartPassed = `${r3Completed} and ${all5WordsValid} and (${r4DestIdentified}) and (normalize-space(\${r4_start_code}) = 'START-AI-PHY' or normalize-space(\${r4_start_code}) = 'START-PHY')`;

  // =============================================================
  // ROUND 4 — PHYSICAL CHALLENGE (AIML)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_phy_group',
    label: 'ROUND 4 — PHYSICAL CHALLENGE',
    relevant: r4StartPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_phy_note',
    label: '📍 ROUND 4: PHYSICAL CHALLENGE\n\nPerform the physical challenge as instructed by the station volunteer.\nOnce completed, collect the finish code from the volunteer.\n\nEnter code in caps',
    hint: 'Complete physical challenge with volunteer.'
  });

  survey.push({
    type: 'image',
    name: 'r4_phy_photo',
    label: '📸 Upload Photo of Physical Challenge Completion (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team completing the physical coordination challenge.',
    required: 'yes',
    required_message: '❌ Photo upload of the challenge is mandatory.'
  });

  survey.push({
    type: 'text',
    name: 'r4_phy_code',
    label: 'Enter END CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ END CODE is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'PHY-AIML' or normalize-space(.) = 'PHY-CY')",
    constraint_message: '❌ Incorrect END CODE. Enter the code in UPPERCASE (CAPS ONLY) provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4Passed = `${r4StartPassed} and (normalize-space(\${r4_phy_code}) = 'PHY-AIML' or normalize-space(\${r4_phy_code}) = 'PHY-CY') and \${r4_phy_photo} != ''`;

  // =============================================================
  // FINAL ROUND: AUDITORIUM & STAGE RIDDLES
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'final_puzzles_group',
    label: 'FINAL ROUND',
    relevant: r4Passed
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

  const finalRiddle1Passed = `${r4Passed} and (normalize-space(\${final_riddle1_answer}) = 'AUDITORIUM' or normalize-space(\${final_riddle1_answer}) = 'AUDI')`;

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
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'FINAL-PATH2' or normalize-space(.) = 'FINAL-PATH1')",
    constraint_message: '❌ Incorrect code. Enter the code in UPPERCASE (CAPS ONLY) provided by the Chief Judge.'
  });

  survey.push({
    type: 'note',
    name: 'final_congratulations_screen',
    label: '🎉 CONGRATULATIONS! YOU HAVE COMPLETED PATH 2!\n\n🏆 You have successfully entered the final clearance code!\n\n🔔 NOW RUN TO GO RING THE BELL TO WIN THE GAME! 🔔🏃💨',
    hint: 'Run to ring the bell to claim victory!',
    relevant: `${finalRiddle2Passed} and (normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH2' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH1') and \${final_solved_puzzle_photo} != ''`
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
      name: 'var_a',
      label: 'Variant A'
    },
    {
      list_name: 'mba_variants',
      name: 'var_b',
      label: 'Variant B'
    }
  ];
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — PATH 2',
      form_id: 'final_clue_path2',
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
      'Location': 'ADMIN',
      'Challenge / Item': 'Object Finding',
      'Question / Prompt': 'Find assigned hidden object, upload photo, enter volunteer code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'JOHN-CENA (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'JOHN-CENA\'',
      'Mandatory Upload': 'Yes (Discovered Object Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 2 Location Clue',
      'Location': 'In-App Security Log Clue',
      'Challenge / Item': 'Security Report Intruder Clue',
      'Question / Prompt': 'Read security tampering clue, deduce letters MBA & enter start code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'MBA & Start Code: MBA-START (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'MBA\' & \'MBA-START\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 2 (R2 Var A)',
      'Location': 'MBA',
      'Challenge / Item': 'Variant A: MBA Quickfire',
      'Question / Prompt': 'Answer 10 business/finance questions, enter volunteer clearance code',
      'Media Attached': 'None',
      'Expected Answer / Code': '10 Answers & Code: MBA-QF-1 (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'MBA-QF-1\'',
      'Mandatory Upload': 'No (All 10 Inputs MANDATORY)'
    },
    {
      'Stage / Round': 'Round 2 (R2 Var B)',
      'Location': 'MBA',
      'Challenge / Item': 'Variant B: Brand Quiz',
      'Question / Prompt': 'Answer 10 brand quiz questions, enter volunteer clearance code',
      'Media Attached': 'None',
      'Expected Answer / Code': '10 Answers & Code: MB-BQ-2 (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'MB-BQ-2\'',
      'Mandatory Upload': 'No (All 10 Inputs MANDATORY)'
    },
    {
      'Stage / Round': 'Round 3 Location Clue',
      'Location': 'In-App Shuffled Words',
      'Challenge / Item': 'Shuffled Words (Displayed as ROUND 3)',
      'Question / Prompt': 'Decode 5 shuffled words: VEHSELS, GIDRAEN, IFIW, SISCUDISNO GGUONLE, IGIDLAT BILRYRA -> Destination guess',
      'Media Attached': 'None',
      'Expected Answer / Code': 'SHELVES, READING, WIFI, DISCUSSION LOUNGE, DIGITAL LIBRARY -> Destination: LIBRARY (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) on all 5 decoded words + LIBRARY',
      'Mandatory Upload': 'No (All 6 Inputs MANDATORY)'
    },
    {
      'Stage / Round': 'Round 3 Checkpoint (R3)',
      'Location': 'LIBRARY',
      'Challenge / Item': 'QR Hunt & Volunteer Code',
      'Question / Prompt': 'Scan hidden QR code in library, enter volunteer completion code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Barcode: DUMB_FAKE & Code: DUMB_FAKE (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Scanned value = \'DUMB_FAKE\' & normalize-space(.) = \'DUMB_FAKE\'',
      'Mandatory Upload': 'Yes (Barcode Scan MANDATORY)'
    },
    {
      'Stage / Round': 'Round 4 Location Clue',
      'Location': 'In-App Word Search (aiml.jpeg)',
      'Challenge / Item': 'Word Search Grid',
      'Question / Prompt': 'Find 5 words from grid -> Deduce location AIML -> Enter start code START-AI-PHY',
      'Media Attached': 'aiml.jpeg',
      'Expected Answer / Code': '5 Words, Location: AIML, Start Code: START-AI-PHY (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'AIML\' & \'START-AI-PHY\'',
      'Mandatory Upload': 'No (Inputs Mandatory)'
    },
    {
      'Stage / Round': 'Round 4 Checkpoint (R4)',
      'Location': 'AIML Pathway',
      'Challenge / Item': 'Physical Challenge',
      'Question / Prompt': 'Complete physical challenge, upload photo, enter finish code PHY-AIML',
      'Media Attached': 'None',
      'Expected Answer / Code': 'PHY-AIML (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'PHY-AIML\'',
      'Mandatory Upload': 'Yes (Photo MANDATORY)'
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
      'Question / Prompt': 'Solve puzzle, upload photo of solved puzzle, enter clearance code FINAL-PATH2, run to ring the bell to win',
      'Media Attached': 'None',
      'Expected Answer / Code': 'FINAL-PATH2 (UPPERCASE ONLY) -> Ring Bell',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'FINAL-PATH2\'',
      'Mandatory Upload': 'Yes (Solved Puzzle Photo MANDATORY)'
    }
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'PATH2_MASTER_KEY');
  return wb;
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
  .rules-box {
    background: #f0fdf4;
    border: 1px solid #86efac;
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
    <span class="badge">PATH 2 OFFICIAL MASTER KEY</span><br>
    <small style="color: #64748b;">FINAL CLUE 2026</small>
  </div>
  <h1>FINAL CLUE &bull; ROUTE 2 / PATH 2</h1>
  <div style="font-size: 9pt; color: #475569; font-weight: 600;">
    CONFIDENTIAL ORGANIZER MASTER KEY &bull; PROGRESSION: ADMIN &rarr; MBA &rarr; LIBRARY &rarr; AIML &rarr; AUDITORIUM
  </div>
</div>

<div class="meta-grid">
  <div class="meta-card">
    <strong>Starting Station</strong>
    <span>ADMIN</span>
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
    <span>aiml.jpeg (Word Search Grid)</span>
  </div>
</div>

<div class="rules-box">
  <strong>🏆 Path 2 Team Qualification Rules:</strong><br>
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
      <td>ADMIN</td>
      <td>Find assigned hidden object/code &amp; verify with volunteer</td>
      <td><span class="code-badge">JOHN-CENA</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Location Clue</strong></td>
      <td>In-App</td>
      <td>Security Report Intruder Clue: credentmials, systbem, administrataor</td>
      <td>Destination: <code>MBA</code><br>Start Code: <span class="code-badge">MBA-START</span></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R2 Mini Challenge</strong></td>
      <td>MBA</td>
      <td>
        &bull; <strong>Variant A:</strong> 10 MBA Quickfire Questions &rarr; Code: <span class="code-badge">MBA-QF-1</span><br>
        &bull; <strong>Variant B:</strong> 10 Brand Quiz Questions &rarr; Code: <span class="code-badge">MB-BQ-2</span>
      </td>
      <td><span class="code-badge">MBA-QF-1</span> / <span class="code-badge">MB-BQ-2</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>R3 Location Clue</strong></td>
      <td>In-App</td>
      <td>Shuffled Words: VEHSELS, GIDRAEN, IFIW, SISCUDISNO GGUONLE, IGIDLAT BILRYRA</td>
      <td>Destination: <code>LIBRARY</code></td>
      <td>All 5 Words Mandatory</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>LIBRARY</td>
      <td>Physically search library for hidden QR code and scan</td>
      <td>Barcode: <span class="code-badge">DUMB_FAKE</span></td>
      <td>Native Barcode Scanner</td>
    </tr>
    <tr>
      <td><strong>R4 Location Clue</strong></td>
      <td>In-App (<code>aiml.jpeg</code>)</td>
      <td>Find ANY 5 words from word search grid</td>
      <td>Destination: <code>AIML</code><br>Start Code: <span class="code-badge">START-AI-PHY</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>Round 4 (R4)</strong></td>
      <td>AIML Pathway</td>
      <td>Complete physical challenge along pathway with volunteer</td>
      <td>Finish Code: <span class="code-badge">PHY-AIML</span></td>
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
      <td><span class="code-badge">FINAL-PATH2</span><br>&rarr; Ring Bell! 🔔</td>
      <td>📸 Mandatory Photo of Solved Puzzle</td>
    </tr>
  </tbody>
</table>

<div class="section-title">2. Volunteer Station Instructions &amp; Verification Procedures</div>

<div class="station-card">
  <h3><span>ADMIN Station: Object Finding</span><span class="code-badge">CODE: JOHN-CENA</span></h3>
  <div><strong>Volunteer Instructions:</strong> Participants locate the assigned hidden item. Once found, they upload a photo and enter code <code>JOHN-CENA</code> in capital letters.</div>
</div>

<div class="station-card">
  <h3><span>MBA Station: Arrival &amp; Challenges</span><span class="code-badge">START: MBA-START &bull; VAR A: MBA-QF-1 &bull; VAR B: MB-BQ-2</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams arrive at MBA. Volunteer gives start code <code>MBA-START</code>.
  <br>&bull; <strong>Variant A (Quickfire):</strong> Team answers 10 business questions and enters <code>MBA-QF-1</code>.
  <br>&bull; <strong>Variant B (Brand Quiz):</strong> Team answers 10 brand quiz questions and enters <code>MB-BQ-2</code>.</div>
</div>

<div class="station-card">
  <h3><span>LIBRARY Station: QR Hunt</span><span class="code-badge">SCANNED QR: DUMB_FAKE</span></h3>
  <div><strong>Volunteer Instructions:</strong> Participants scan the physical QR code with payload <code>DUMB_FAKE</code> using the in-app barcode scanner.</div>
</div>

<div class="station-card">
  <h3><span>AIML Station: Physical Challenge</span><span class="code-badge">START: START-AI-PHY &bull; END: PHY-AIML</span></h3>
  <div><strong>Volunteer Instructions:</strong> Volunteer gives start code <code>START-AI-PHY</code>. Team completes physical challenge, uploads photo, and receives finish code <code>PHY-AIML</code>.</div>
</div>

<div class="station-card">
  <h3><span>Main Auditorium Stage: Grand Finale</span><span class="code-badge">CODE: FINAL-PATH2 &bull; 🔔 RING THE BELL</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams solve Riddle 1 (<code>AUDITORIUM</code>) and Riddle 2 (<code>STAGE</code>), upload a clear photo of their solved puzzle sheet, and receive clearance code <code>FINAL-PATH2</code> from Chief Judges before running to ring the victory bell!</div>
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

  console.log('Generating PATH2_FINAL_ODK.xlsx with STRICT UPPERCASE ONLY & 100% MANDATORY enforcement...');
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

  console.log('Generating PATH2_ANSWER_KEY.xlsx...');
  const keyWb = buildAnswerKeyWorkbook();
  const keyPath = path.join(route2Dir, 'PATH2_ANSWER_KEY.xlsx');
  XLSX.writeFile(keyWb, keyPath);
  console.log(`Successfully created: ${keyPath}`);

  console.log('Generating PATH2_ANSWER_KEY.pdf...');
  const pdfHtml = buildAnswerKeyPdfHtml();
  const pdfPath = path.join(route2Dir, 'PATH2_ANSWER_KEY.pdf');
  renderHtmlToPdf(pdfHtml, pdfPath);
  console.log(`Successfully created: ${pdfPath}`);

  console.log('Packaging PATH2_COMPLETE_PACKAGE.zip and PATH2_MEDIA.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route2Dir}\\PATH2_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route2Dir}\\PATH2_FINAL_ODK.xlsx', '${route2Dir}\\PATH2_ANSWER_KEY.xlsx', '${route2Dir}\\PATH2_ANSWER_KEY.pdf', '${mediaDir}' -DestinationPath '${route2Dir}\\PATH2_COMPLETE_PACKAGE.zip' -Force"`);

  console.log('🎉 PATH 2 PACKAGE GENERATED SUCCESSFULLY!');
}

main();

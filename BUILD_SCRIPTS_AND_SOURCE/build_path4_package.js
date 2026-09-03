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
    label: `🏆 FINAL CLUE — PATH 4
TREASURE HUNT FOR FRESHERS 2026

Welcome to PATH 4 of the Final Clue Treasure Hunt!

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
  // ROUND 1 — OLD CANTEEN: OBJECT (Cat Board)
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
    label: '📍 ROUND 1: 🧩 OBJECT FINDING\n\nInstructions:\n1. Search the location to find the designated Cat Board physical object.\n2. Take a mandatory photo of the object.\n3. Show the object/photo to the nearby Luminus volunteer.\n4. Enter the verification code given by the volunteer.\n\nEnter code in caps',
    hint: 'Find the Cat Board object, take photo, and ask volunteer for code.'
  });

  survey.push({
    type: 'image',
    name: 'r1_canteen_photo',
    label: '📸 Upload Photo of Discovered Cat Board Object (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered Cat Board object.',
    required: 'yes',
    required_message: '❌ Photo upload of the object is strictly mandatory.'
  });

  survey.push({
    type: 'text',
    name: 'r1_canteen_code',
    label: 'Enter Volunteer Verification Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer verification code is mandatory.',
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'MEOW-BOW-4'",
    constraint_message: '❌ Incorrect code. Check the object again and enter the code exactly as displayed.'
  });

  survey.push({
    type: 'end_group'
  });

  const r1Passed = `${startPassed} and normalize-space(\${r1_canteen_code}) = 'MEOW-BOW-4' and \${r1_canteen_photo} != ''`;

  // =============================================================
  // ROUND 2 — LOCATION CLUE (CSE-02 — Emoji Math Riddle)
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
    label: `📍 ROUND 2 LOCATION CLUE: EMOJI MATH RIDDLE

Solve the emoji equations below. Your final answer corresponds to a number used at the station. Use that number to identify your next location.

🍎 + 🍎 = 10
🍎 + 🍌 = 7
🍌 + 🍇 = 6

Final equation:
🍇 + 🍎 = ?

Enter code in caps`,
    hint: 'Solve the emoji equations to find the numeric value.'
  });

  survey.push({
    type: 'integer',
    name: 'r2_math_answer',
    label: 'Enter the calculated value for [ 🍇 + 🍎 ]: (MANDATORY)',
    hint: 'Enter the exact integer solution.',
    required: 'yes',
    required_message: '❌ Solution is mandatory.',
    constraint: '. = 9',
    constraint_message: '❌ Incorrect. Recheck the three equations and calculate the final value.'
  });

  const r2MathPassed = `${r1Passed} and \${r2_math_answer} = 9`;

  survey.push({
    type: 'note',
    name: 'r2_cse_proceed_note',
    label: '🏃 Correct! Go to the computational block of the campus---the house of three computing blocks, and ask the Luminus volunteer there for the START CODE.\n\nEnter code in caps',
    hint: 'Proceed to the location and ask volunteer for start code.',
    relevant: r2MathPassed
  });

  survey.push({
    type: 'text',
    name: 'r2_cse_start_code',
    label: 'Enter START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    relevant: r2MathPassed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'CSE-START'",
    constraint_message: '❌ Incorrect START CODE. Check the code at your current location and enter it exactly as provided.'
  });

  survey.push({
    type: 'end_group'
  });

  const cseStartPassed = `${r2MathPassed} and normalize-space(\${r2_cse_start_code}) = 'CSE-START'`;

  // =============================================================
  // ROUND 2 — MINI CHALLENGES (VARIANTS 1 & 2)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'cse_challenge_select_group',
    label: 'ROUND 2 — MINI CHALLENGES',
    relevant: cseStartPassed
  });

  survey.push({
    type: 'select_one cse_variant_list',
    name: 'r2_cse_variant_select',
    label: 'Select your assigned challenge variant: (MANDATORY)\nMandatory selection',
    hint: 'Choose the variant assigned by the station volunteer.',
    required: 'yes',
    required_message: '❌ Selecting your assigned variant is mandatory.'
  });

  // -------------------------------------------------------------
  // VARIANT 1: CSE-06 LOGICAL CHALLENGE 1 (ALGO RELAY + CYBER DETECTIVE)
  // -------------------------------------------------------------
  survey.push({
    type: 'begin_group',
    name: 'cse_v1_container',
    label: 'VARIANT A — LOGICAL CHALLENGE 1',
    relevant: "${r2_cse_variant_select} = 'var1'"
  });

  // Challenge 1: Algo Relay (10 Questions)
  survey.push({
    type: 'begin_group',
    name: 'cse_v1_algo_group',
    label: 'CHALLENGE 1: 🏃 ALGO RELAY'
  });

  survey.push({
    type: 'note',
    name: 'cse_algo_intro',
    label: '🏃 CHALLENGE 1 — ALGO RELAY\n\nAnswer the 10 algorithmic reasoning questions below.\n\nEnter code in caps',
    hint: 'Select the correct option for each question.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q1',
    label: 'Q1. A program needs to calculate the average of 3 numbers. Which should happen first?\nA. Print the average\nB. Input the 3 numbers\nC. Divide by 3\nD. Add the numbers',
    required: 'yes',
    required_message: '❌ Question 1 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q1.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q2',
    label: 'Q2. Which is the correct order to find the largest of three numbers?\nA. Print → Input → Compare\nB. Input → Compare → Print\nC. Compare → Print → Input\nD. Print → Compare → Input',
    required: 'yes',
    required_message: '❌ Question 2 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q2.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q3',
    label: 'Q3. What is the output?\nx = 5\nx = x + 3\nx = x × 2\nx = x - 4\nPRINT x\n\nA. 10\nB. 12\nC. 16\nD. 20',
    required: 'yes',
    required_message: '❌ Question 3 is mandatory.',
    constraint: ". = 'c'",
    constraint_message: '❌ Incorrect answer for Q3.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q4',
    label: 'Q4. A robot starts at position 0.\n+4 → -2 → +7 → -3\nWhere does it finish?\n\nA. 4\nB. 5\nC. 6\nD. 7',
    required: 'yes',
    required_message: '❌ Question 4 is mandatory.',
    constraint: ". = 'c'",
    constraint_message: '❌ Incorrect answer for Q4.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q5',
    label: 'Q5. Which condition correctly checks whether n is divisible by both 3 and 5?\nA. n % 3 == 0 OR n % 5 == 0\nB. n % 3 == 0 AND n % 5 == 0\nC. n / 3 == 0 AND n / 5 == 0\nD. n % 15 == 1',
    required: 'yes',
    required_message: '❌ Question 5 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q5.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q6',
    label: 'Q6. What will this algorithm print?\nx = 10\nIF x > 5\n    PRINT "A"\nELSE\n    PRINT "B"\n\nA. A\nB. B\nC. A B\nD. Nothing',
    required: 'yes',
    required_message: '❌ Question 6 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q6.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q7',
    label: 'Q7. How many times will HELLO be printed?\nFOR i = 1 TO 5\n    PRINT "HELLO"\n\nA. 4\nB. 5\nC. 6\nD. 10',
    required: 'yes',
    required_message: '❌ Question 7 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q7.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q8',
    label: 'Q8. What is the output?\nx = 1\nREPEAT 4 TIMES\n    x = x × 2\nPRINT x\n\nA. 4\nB. 8\nC. 16\nD. 32',
    required: 'yes',
    required_message: '❌ Question 8 is mandatory.',
    constraint: ". = 'c'",
    constraint_message: '❌ Incorrect answer for Q8.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q9',
    label: 'Q9. A sorted list is: 5, 10, 15, 20, 25, 30, 35. You want to find 30 using binary search. Which number should you check first?\nA. 5\nB. 10\nC. 20\nD. 30',
    required: 'yes',
    required_message: '❌ Question 9 is mandatory.',
    constraint: ". = 'c'",
    constraint_message: '❌ Incorrect answer for Q9.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_algo_q10',
    label: 'Q10. You need to find a name in a sorted list of 1,000 names. Which method is generally more efficient?\nA. Check every name from the beginning\nB. Binary search\nC. Pick names randomly\nD. Check only the first name',
    required: 'yes',
    required_message: '❌ Question 10 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q10.'
  });

  survey.push({
    type: 'end_group'
  });

  // Challenge 2: Cyber Detective (10 Questions)
  survey.push({
    type: 'begin_group',
    name: 'cse_v1_cyber_group',
    label: 'CHALLENGE 2: 🕵️ CYBER DETECTIVE'
  });

  survey.push({
    type: 'note',
    name: 'cse_cyber_intro',
    label: '🕵️ CHALLENGE 2 — CYBER DETECTIVE\n\nAnswer the 10 cyber security and detection questions below.\n\nEnter code in caps',
    hint: 'Select the correct option for each question.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q1',
    label: 'Q1. You receive: "Your bank account will be closed today. Click this link and enter your password." What is the most likely threat?\nA. Phishing\nB. Bluetooth\nC. Backup\nD. Encryption',
    required: 'yes',
    required_message: '❌ Question 1 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q1.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q2',
    label: 'Q2. Which password is strongest?\nA. 12345678\nB. password123\nC. College2005\nD. R7#kP2!mQ9@x',
    required: 'yes',
    required_message: '❌ Question 2 is mandatory.',
    constraint: ". = 'd'",
    constraint_message: '❌ Incorrect answer for Q2.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q3',
    label: 'Q3. Someone calls pretending to be from your bank and asks for your OTP. What should you do?\nA. Give the OTP\nB. Give only half the OTP\nC. Do not share it\nD. Send your password instead',
    required: 'yes',
    required_message: '❌ Question 3 is mandatory.',
    constraint: ". = 'c'",
    constraint_message: '❌ Incorrect answer for Q3.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q4',
    label: 'Q4. Which URL is most suspicious if the real website is mybank.com?\nA. mybank.com/login\nB. secure.mybank.com\nC. mybank.com\nD. mybank-login-security.com',
    required: 'yes',
    required_message: '❌ Question 4 is mandatory.',
    constraint: ". = 'd'",
    constraint_message: '❌ Incorrect answer for Q4.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q5',
    label: 'Q5. A student uses the same password for email, Instagram and banking. What is the biggest risk?\nA. The password becomes longer\nB. One stolen password can compromise multiple accounts\nC. The internet becomes slower\nD. The phone battery drains faster',
    required: 'yes',
    required_message: '❌ Question 5 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q5.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q6',
    label: 'Q6. A message says: "Send this message to 20 people or your account will be deleted." What is the safest response?\nA. Forward it\nB. Send your OTP\nC. Ignore it\nD. Give the sender your password',
    required: 'yes',
    required_message: '❌ Question 6 is mandatory.',
    constraint: ". = 'c'",
    constraint_message: '❌ Incorrect answer for Q6.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q7',
    label: 'Q7. What is the main purpose of two-factor authentication?\nA. Make the screen brighter\nB. Add another layer of account security\nC. Increase internet speed\nD. Store more files',
    required: 'yes',
    required_message: '❌ Question 7 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q7.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q8',
    label: 'Q8. A student clicks a fake login link and enters their username and password. What has most likely happened?\nA. Their credentials may have been stolen\nB. Their phone automatically upgraded\nC. Their password became stronger\nD. Their Wi-Fi was repaired',
    required: 'yes',
    required_message: '❌ Question 8 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q8.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q9',
    label: 'Q9. Which is the safest action when using a public computer?\nA. Save your password in the browser\nB. Leave your account logged in\nC. Log out after use\nD. Disable the screen lock',
    required: 'yes',
    required_message: '❌ Question 9 is mandatory.',
    constraint: ". = 'c'",
    constraint_message: '❌ Incorrect answer for Q9.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v1_cyb_q10',
    label: 'Q10. A student\'s account was hacked. Evidence shows: Student received a fake password-reset email, clicked the link, entered credentials, and an unknown device logged in. What is the most likely cause?\nA. Weak Wi-Fi signal\nB. Phishing\nC. Low battery\nD. Bluetooth connection',
    required: 'yes',
    required_message: '❌ Question 10 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q10.'
  });

  survey.push({
    type: 'end_group'
  });

  const v1AllPassed = "${v1_algo_q1} = 'b' and ${v1_algo_q2} = 'b' and ${v1_algo_q3} = 'c' and ${v1_algo_q4} = 'c' and ${v1_algo_q5} = 'b' and ${v1_algo_q6} = 'a' and ${v1_algo_q7} = 'b' and ${v1_algo_q8} = 'c' and ${v1_algo_q9} = 'c' and ${v1_algo_q10} = 'b' and ${v1_cyb_q1} = 'a' and ${v1_cyb_q2} = 'd' and ${v1_cyb_q3} = 'c' and ${v1_cyb_q4} = 'd' and ${v1_cyb_q5} = 'b' and ${v1_cyb_q6} = 'c' and ${v1_cyb_q7} = 'b' and ${v1_cyb_q8} = 'a' and ${v1_cyb_q9} = 'c' and ${v1_cyb_q10} = 'b'";

  survey.push({
    type: 'image',
    name: 'r2_v1_cse_photo',
    label: '📸 Upload Photo of Team Solving Variant 1 Challenge (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the station.',
    required: 'yes',
    required_message: '❌ Station photo upload is mandatory.',
    relevant: v1AllPassed
  });

  survey.push({
    type: 'text',
    name: 'r2_v1_code',
    label: 'Enter Volunteer Clearance Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer clearance code is mandatory.',
    relevant: v1AllPassed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'CSE-EINSTEIN-4'",
    constraint_message: '❌ Incorrect code. Check the code obtained after completing the challenge.'
  });

  survey.push({
    type: 'end_group'
  });

  // -------------------------------------------------------------
  // VARIANT 2: CSE-07 LOGICAL CHALLENGE 2 (BUG HUNTER + TECH DETECTIVE)
  // -------------------------------------------------------------
  survey.push({
    type: 'begin_group',
    name: 'cse_v2_container',
    label: 'VARIANT B — LOGICAL CHALLENGE 2',
    relevant: "${r2_cse_variant_select} = 'var2'"
  });

  // Challenge 3: Bug Hunter (10 Questions)
  survey.push({
    type: 'begin_group',
    name: 'cse_v2_bug_group',
    label: 'CHALLENGE 3: 🐛 BUG HUNTER'
  });

  survey.push({
    type: 'note',
    name: 'cse_bug_intro',
    label: '🐛 CHALLENGE 3 — BUG HUNTER\n\nFind the bug in each of the 10 code snippets below.\n\nEnter code in caps',
    hint: 'Select the correct option for each question.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q1',
    label: 'Q1. What is wrong with this program?\nIF age > 18\n    PRINT "ELIGIBLE"\nELSE\n    PRINT "NOT ELIGIBLE"\nA person who is exactly 18 is being rejected. What should change?\nA. > to >=\nB. > to <\nC. 18 to 19\nD. Remove ELSE',
    required: 'yes',
    required_message: '❌ Question 1 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q1.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q2',
    label: 'Q2. What will this program print?\nx = 2\nWHILE x < 10\n    PRINT x\n    x = x + 2\n\nA. 2 4 6 8\nB. 2 4 6 8 10\nC. 1 2 3 4\nD. Infinite loop',
    required: 'yes',
    required_message: '❌ Question 2 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q2.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q3',
    label: 'Q3. What is the bug?\nIF marks >= 40\n    PRINT "FAIL"\nELSE\n    PRINT "PASS"\n\nA. >= should be <=\nB. PASS and FAIL are reversed\nC. Marks cannot be compared\nD. ELSE must be removed',
    required: 'yes',
    required_message: '❌ Question 3 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q3.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q4',
    label: 'Q4. What will be the output?\na = 5\nb = 3\na = a + b\nb = a - b\na = a - b\nPRINT a\nPRINT b\n\nA. 5 3\nB. 3 5\nC. 8 3\nD. 8 5',
    required: 'yes',
    required_message: '❌ Question 4 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q4.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q5',
    label: 'Q5. A program should print numbers from 1 to 5.\ni = 1\nWHILE i <= 5\n    PRINT i\nWhat is missing?\n\nA. i = i + 1\nB. i = i - 1\nC. i = 0\nD. PRINT i',
    required: 'yes',
    required_message: '❌ Question 5 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q5.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q6',
    label: 'Q6. A program checks whether a number is positive:\nIF number > 0\n    PRINT "POSITIVE"\nELSE\n    PRINT "NEGATIVE"\nWhat happens when the input is 0?\n\nA. POSITIVE\nB. NEGATIVE\nC. Neither is mathematically correct\nD. It always prints both',
    required: 'yes',
    required_message: '❌ Question 6 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q6.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q7',
    label: 'Q7. A calculator produces: 8 × 0 = 8. What is the bug?\nA. Multiplication is incorrect\nB. Addition is incorrect\nC. Division is incorrect\nD. Nothing is wrong',
    required: 'yes',
    required_message: '❌ Question 7 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q7.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q8',
    label: 'Q8. A program should calculate: 10 + 20 × 2. The programmer calculates: (10 + 20) × 2 = 60. What rule was ignored?\nA. Multiplication has priority over addition\nB. Addition has priority over multiplication\nC. Division must happen first\nD. Subtraction must happen first',
    required: 'yes',
    required_message: '❌ Question 8 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q8.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q9',
    label: 'Q9. A program is supposed to find the largest number.\nlargest = 0\nIF number < largest\n    largest = number\nWhat is the main problem?\n\nA. It updates when a smaller number is found\nB. It updates when a larger number is found\nC. largest should always be 100\nD. There is no comparison',
    required: 'yes',
    required_message: '❌ Question 9 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q9.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_bug_q10',
    label: 'Q10. A discount should be given only when: Amount ≥ ₹1,000 AND customer is a member.\nThe programmer writes:\nIF amount >= 1000 OR member == YES\n    PRINT "DISCOUNT"\nWhat is the bug?\n\nA. OR should be AND\nB. AND should be OR\nC. Amount should be ₹100\nD. Member status should be ignored',
    required: 'yes',
    required_message: '❌ Question 10 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q10.'
  });

  survey.push({
    type: 'end_group'
  });

  // Challenge 4: Tech Detective (10 Questions)
  survey.push({
    type: 'begin_group',
    name: 'cse_v2_tech_group',
    label: 'CHALLENGE 4: 💻 TECH DETECTIVE'
  });

  survey.push({
    type: 'note',
    name: 'cse_tech_intro',
    label: '💻 CHALLENGE 4 — TECH DETECTIVE\n\nAnswer the 10 technology and networking questions below.\n\nEnter code in caps',
    hint: 'Select the correct option for each question.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q1',
    label: 'Q1. Which device connects different networks and directs data between them?\nA. Monitor\nB. Router\nC. Keyboard\nD. Printer',
    required: 'yes',
    required_message: '❌ Question 1 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q1.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q2',
    label: 'Q2. You type: www.example.com. Which system finds the IP address associated with this domain?\nA. DNS\nB. USB\nC. HDMI\nD. RAM',
    required: 'yes',
    required_message: '❌ Question 2 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q2.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q3',
    label: 'Q3. Which of these is an operating system?\nA. Chrome\nB. Windows\nC. YouTube\nD. Google',
    required: 'yes',
    required_message: '❌ Question 3 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q3.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q4',
    label: 'Q4. Which statement about RAM is correct?\nA. It is mainly used as temporary working memory\nB. It permanently stores files even without power\nC. It is used only for printing\nD. It is the same as a keyboard',
    required: 'yes',
    required_message: '❌ Question 4 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q4.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q5',
    label: 'Q5. Which protocol is used for secure web browsing?\nA. HTTP\nB. HTTPS\nC. FTP\nD. SMTP',
    required: 'yes',
    required_message: '❌ Question 5 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q5.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q6',
    label: 'Q6. A device automatically receives an IP address when joining a network. Which service commonly provides it?\nA. DNS\nB. DHCP\nC. HTTP\nD. HTML',
    required: 'yes',
    required_message: '❌ Question 6 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q6.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q7',
    label: 'Q7. Which of these has the largest storage capacity?\nA. 500 KB\nB. 2 MB\nC. 1 GB\nD. 5 TB',
    required: 'yes',
    required_message: '❌ Question 7 is mandatory.',
    constraint: ". = 'd'",
    constraint_message: '❌ Incorrect answer for Q7.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q8',
    label: 'Q8. Which is the correct order from smallest to largest?\nA. KB → MB → GB → TB\nB. MB → KB → TB → GB\nC. GB → MB → KB → TB\nD. TB → GB → MB → KB',
    required: 'yes',
    required_message: '❌ Question 8 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q8.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q9',
    label: 'Q9. A website address begins with: https://. What does the S mainly indicate?\nA. The website is a search engine\nB. The connection uses security/encryption\nC. The website is stored on your computer\nD. The website has no password',
    required: 'yes',
    required_message: '❌ Question 9 is mandatory.',
    constraint: ". = 'b'",
    constraint_message: '❌ Incorrect answer for Q9.'
  });

  survey.push({
    type: 'select_one mcq_opts',
    name: 'v2_tech_q10',
    label: 'Q10. You type a website address into your browser. Which sequence is the most appropriate?\nA. DNS finds the IP → Browser connects to the server → Website data is requested\nB. Browser requests the webpage → DNS shuts down → Server connects\nC. RAM finds the website → Keyboard connects to DNS → Router prints it\nD. DHCP creates the webpage → Monitor finds the IP → Browser shuts down',
    required: 'yes',
    required_message: '❌ Question 10 is mandatory.',
    constraint: ". = 'a'",
    constraint_message: '❌ Incorrect answer for Q10.'
  });

  survey.push({
    type: 'end_group'
  });

  const v2AllPassed = "${v2_bug_q1} = 'a' and ${v2_bug_q2} = 'a' and ${v2_bug_q3} = 'b' and ${v2_bug_q4} = 'b' and ${v2_bug_q5} = 'a' and ${v2_bug_q6} = 'b' and ${v2_bug_q7} = 'a' and ${v2_bug_q8} = 'a' and ${v2_bug_q9} = 'a' and ${v2_bug_q10} = 'a' and ${v2_tech_q1} = 'b' and ${v2_tech_q2} = 'a' and ${v2_tech_q3} = 'b' and ${v2_tech_q4} = 'a' and ${v2_tech_q5} = 'b' and ${v2_tech_q6} = 'b' and ${v2_tech_q7} = 'd' and ${v2_tech_q8} = 'a' and ${v2_tech_q9} = 'b' and ${v2_tech_q10} = 'a'";

  survey.push({
    type: 'image',
    name: 'r2_v2_cse_photo',
    label: '📸 Upload Photo of Team Solving Variant 2 Challenge (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team at the station.',
    required: 'yes',
    required_message: '❌ Station photo upload is mandatory.',
    relevant: v2AllPassed
  });

  survey.push({
    type: 'text',
    name: 'r2_v2_code',
    label: 'Enter Volunteer Clearance Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer clearance code is mandatory.',
    relevant: v2AllPassed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and normalize-space(.) = 'CSE-ZUCKER-4'",
    constraint_message: '❌ Incorrect code. Check the code obtained after completing the challenge.'
  });

  survey.push({
    type: 'end_group'
  });

  survey.push({
    type: 'end_group'
  });

  const cseMiniPassed = `${cseStartPassed} and ((${'\${r2_cse_variant_select}'} = 'var1' and normalize-space(\${r2_v1_code}) = 'CSE-EINSTEIN-4' and \${r2_v1_cse_photo} != '') or (${'\${r2_cse_variant_select}'} = 'var2' and normalize-space(\${r2_v2_code}) = 'CSE-ZUCKER-4' and \${r2_v2_cse_photo} != ''))`;

  // =============================================================
  // ROUND 3 — LOCATION CLUE (CYB-01 — Morse Code with morse.jpeg)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_cyber_loc_group',
    label: 'ROUND 3 LOCATION CLUE',
    relevant: cseMiniPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_cyber_morse_note',
    label: `📍 ROUND 3 LOCATION CLUE: MORSE CODE MESSAGE

Decode the Morse-code message to reveal your next location.

Morse Message:
--. --- / - --- / -.-. -.-- -... . .-.

Refer to the Morse key reference image attached below.

Enter code in caps`,
    hint: 'Decode each Morse group to discover the next location.',
    'media::image': 'morse.jpeg'
  });

  survey.push({
    type: 'text',
    name: 'r3_cyber_loc_answer',
    label: 'Enter your decoded message: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Decoded message is mandatory.',
    'media::image': 'morse.jpeg',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'GO TO CYBER' or normalize-space(.) = 'CYBER' or normalize-space(.) = 'CYBER BLOCK')",
    constraint_message: '❌ Incorrect. Decode each Morse-code group carefully and try again.'
  });

  const r3LocGuessed = "normalize-space(${r3_cyber_loc_answer}) = 'GO TO CYBER' or normalize-space(${r3_cyber_loc_answer}) = 'CYBER' or normalize-space(${r3_cyber_loc_answer}) = 'CYBER BLOCK'";

  survey.push({
    type: 'note',
    name: 'r3_cyber_proceed_note',
    label: '🏃 Correct! Proceed to the location identified by the decoded message.\n\nEnter code in caps',
    hint: 'Proceed to the location for the QR hunt.',
    relevant: r3LocGuessed
  });

  survey.push({
    type: 'end_group'
  });

  const r3BlockPassed = `${cseMiniPassed} and (${r3LocGuessed})`;

  // =============================================================
  // ROUND 3 — QR HUNT (CYBER)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r3_cyber_qr_group',
    label: 'ROUND 3 — QR HUNT',
    relevant: r3BlockPassed
  });

  survey.push({
    type: 'note',
    name: 'r3_cyber_qr_note',
    label: '📍 ROUND 3: QR HUNT\n\nFind the designated QR code at this location and scan it. Follow the instructions provided by the QR challenge.\n\nEnter code in caps',
    hint: 'Locate and scan the hidden QR code.'
  });

  survey.push({
    type: 'barcode',
    name: 'r3_cyber_qr_scan',
    label: 'Scan Discovered QR Code (MANDATORY)',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ QR code scanning is mandatory.',
    constraint: "normalize-space(.) = 'JAMES_BOND'",
    constraint_message: '❌ Incorrect code. Continue the QR hunt and check the code carefully.'
  });

  survey.push({
    type: 'image',
    name: 'r3_cyber_qr_photo',
    label: '📸 Upload Photo of Discovered QR Code / Station (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of the discovered QR code location.',
    required: 'yes',
    required_message: '❌ Station photo upload is mandatory.'
  });

  survey.push({
    type: 'end_group'
  });

  const r3QrPassed = `${r3BlockPassed} and normalize-space(\${r3_cyber_qr_scan}) = 'JAMES_BOND' and \${r3_cyber_qr_photo} != ''`;

  // =============================================================
  // ROUND 4 — LOCATION CLUE (MBA-05 — Wisdom, Below & Beyond)
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
    label: `📍 ROUND 4 LOCATION CLUE

"Numbers become stories here.
Stories become strategies.
Strategies become presentations.
Where managers learn to turn information into decisions,
find me."

Enter code in caps`,
    hint: 'Solve the riddle to identify your next destination.'
  });

  survey.push({
    type: 'text',
    name: 'r4_mba_loc_answer',
    label: 'Enter your deduced destination location: (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Destination location is mandatory.',
    constraint: "regex(., '^[A-Z ]+$') and (normalize-space(.) = 'MBA' or normalize-space(.) = 'MBA BLOCK')",
    constraint_message: '❌ Not quite. Reconsider each of the three clues and how they connect.'
  });

  const r4LocGuessed = "normalize-space(${r4_mba_loc_answer}) = 'MBA' or normalize-space(${r4_mba_loc_answer}) = 'MBA BLOCK'";

  survey.push({
    type: 'note',
    name: 'r4_mba_proceed_note',
    label: '🏃 Correct! Proceed to the location identified by your answer and look for the physical challenge.\n\nEnter code in caps',
    hint: 'Proceed to the location and ask volunteer for start code.',
    relevant: r4LocGuessed
  });

  survey.push({
    type: 'text',
    name: 'r4_mba_start_code',
    label: 'Enter START CODE from Volunteer (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ START CODE is mandatory.',
    relevant: r4LocGuessed,
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'MONEY_BROTHA' or normalize-space(.) = 'MONEY-BROTHA')",
    constraint_message: '❌ Incorrect START CODE. Enter MONEY_BROTHA in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const mbaStartPassed = `${r3QrPassed} and (${r4LocGuessed}) and (normalize-space(\${r4_mba_start_code}) = 'MONEY_BROTHA' or normalize-space(\${r4_mba_start_code}) = 'MONEY-BROTHA')`;

  // =============================================================
  // ROUND 4 — PHYSICAL CHALLENGE (MBA)
  // =============================================================
  survey.push({
    type: 'begin_group',
    name: 'r4_mba_phy_group',
    label: 'ROUND 4 — PHYSICAL CHALLENGE',
    relevant: mbaStartPassed
  });

  survey.push({
    type: 'note',
    name: 'r4_mba_phy_note',
    label: '📍 ROUND 4: PHYSICAL CHALLENGE\n\nReport to the station and complete the physical challenge under volunteer supervision.\n\nEnter code in caps',
    hint: 'Complete physical challenge with volunteer.'
  });

  survey.push({
    type: 'select_one challenge_status_list',
    name: 'r4_mba_status',
    label: 'Challenge Status (MANDATORY)',
    hint: 'Select PASSED once completed with volunteer.',
    required: 'yes',
    required_message: '❌ Challenge status selection is mandatory.'
  });

  survey.push({
    type: 'note',
    name: 'r4_mba_not_passed_note',
    label: '⚠️ Please follow the organizer\'s instructions before continuing.',
    hint: 'Complete the physical challenge as instructed.',
    relevant: "${r4_mba_status} = 'not_passed'"
  });

  survey.push({
    type: 'image',
    name: 'r4_mba_phy_photo',
    label: '📸 Upload Photo of Physical Challenge Completion (MANDATORY)\nPhoto upload is mandatory',
    hint: 'Take a clear photo of your team completing the physical challenge.',
    required: 'yes',
    required_message: '❌ Challenge photo upload is mandatory.',
    relevant: "${r4_mba_status} = 'passed'"
  });

  survey.push({
    type: 'text',
    name: 'r4_mba_phy_code',
    label: 'Enter Volunteer Completion Code (MANDATORY)\nEnter code in caps',
    hint: 'ENTER CODE/ANSWER IN CAPITAL LETTERS ONLY.',
    required: 'yes',
    required_message: '❌ Volunteer completion code is mandatory.',
    relevant: "${r4_mba_status} = 'passed'",
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'RAVI-KISHEN' or normalize-space(.) = 'RAVI_KISHEN')",
    constraint_message: '❌ Incorrect code. Enter RAVI-KISHEN in CAPS provided by the volunteer.'
  });

  survey.push({
    type: 'end_group'
  });

  const r4PhyPassed = `${mbaStartPassed} and \${r4_mba_status} = 'passed' and (normalize-space(\${r4_mba_phy_code}) = 'RAVI-KISHEN' or normalize-space(\${r4_mba_phy_code}) = 'RAVI_KISHEN') and \${r4_mba_phy_photo} != ''`;

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
    constraint: "regex(., '^[A-Z0-9\\-_]+$') and (normalize-space(.) = 'FINAL-PATH4' or normalize-space(.) = 'FINAL-PATH1' or normalize-space(.) = 'FINAL-PATH2' or normalize-space(.) = 'FINAL-PATH3')",
    constraint_message: '❌ Incorrect code. Enter the code in UPPERCASE (CAPS ONLY) provided by the Chief Judge.'
  });

  survey.push({
    type: 'note',
    name: 'final_congratulations_screen',
    label: '🎉 CONGRATULATIONS! YOU HAVE COMPLETED PATH 4!\n\n🏆 You have successfully entered the final clearance code!\n\n🔔 NOW RUN TO GO RING THE BELL TO WIN THE GAME! 🔔🏃💨',
    hint: 'Run to ring the bell to claim victory!',
    relevant: `${finalRiddle2Passed} and (normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH4' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH1' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH2' or normalize-space(\${final_stage_volunteer_code}) = 'FINAL-PATH3') and \${final_solved_puzzle_photo} != ''`
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
      label: 'Variant A: Logical Challenge 1 (Algo Relay + Cyber Detective)'
    },
    {
      list_name: 'cse_variant_list',
      name: 'var2',
      label: 'Variant B: Logical Challenge 2 (Bug Hunter + Tech Detective)'
    },
    {
      list_name: 'mcq_opts',
      name: 'a',
      label: 'A'
    },
    {
      list_name: 'mcq_opts',
      name: 'b',
      label: 'B'
    },
    {
      list_name: 'mcq_opts',
      name: 'c',
      label: 'C'
    },
    {
      list_name: 'mcq_opts',
      name: 'd',
      label: 'D'
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
      form_title: 'FINAL CLUE — PATH 4',
      form_id: 'PATH4_TREASURE_HUNT',
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
      'Location': 'OLD CANTEEN',
      'Challenge / Item': 'Object Finding (Cat Board)',
      'Question / Prompt': 'Find assigned Cat Board object, upload photo, enter volunteer code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'MEOW-BOW-4 (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'MEOW-BOW-4\'',
      'Mandatory Upload': 'Yes (Discovered Object Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 2 Location Clue',
      'Location': 'In-App Emoji Math',
      'Challenge / Item': 'Emoji Math Riddle (CSE-02)',
      'Question / Prompt': '🍎+🍎=10, 🍎+🍌=7, 🍌+🍇=6 -> 🍇+🍎=? -> Result: 9 -> Start Code: CSE-START',
      'Media Attached': 'None',
      'Expected Answer / Code': '9 & Start Code: CSE-START (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Integer = 9 & normalize-space(.) = \'CSE-START\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 2 (Var A)',
      'Location': 'CSE',
      'Challenge / Item': 'Variant A: Logical Challenge 1',
      'Question / Prompt': 'Algo Relay (10 MCQs) + Cyber Detective (10 MCQs) + Photo + Code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Algo: B,B,C,C,B,A,B,C,C,B | Cyber: A,D,C,D,B,C,B,A,C,B -> Code: CSE-EINSTEIN-4',
      'Verification / Constraint Rule': 'MCQ selections & Code regex & normalize-space(.) = \'CSE-EINSTEIN-4\'',
      'Mandatory Upload': 'Yes (Station Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 2 (Var B)',
      'Location': 'CSE',
      'Challenge / Item': 'Variant B: Logical Challenge 2',
      'Question / Prompt': 'Bug Hunter (10 MCQs) + Tech Detective (10 MCQs) + Photo + Code',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Bug: A,A,B,B,A,B,A,A,A,A | Tech: B,A,B,A,B,B,D,A,B,A -> Code: CSE-ZUCKER-4',
      'Verification / Constraint Rule': 'MCQ selections & Code regex & normalize-space(.) = \'CSE-ZUCKER-4\'',
      'Mandatory Upload': 'Yes (Station Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 3 Location Clue',
      'Location': 'In-App Morse Code',
      'Challenge / Item': 'CYB-01 Morse Code Message (morse.jpeg)',
      'Question / Prompt': '--. --- / - --- / -.-. -.-- -... . .-. -> Destination: GO TO CYBER',
      'Media Attached': 'morse.jpeg',
      'Expected Answer / Code': 'GO TO CYBER (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'GO TO CYBER\'',
      'Mandatory Upload': 'No (Input Mandatory)'
    },
    {
      'Stage / Round': 'Round 3 Checkpoint (R3)',
      'Location': 'CYBER',
      'Challenge / Item': 'QR Hunt',
      'Question / Prompt': 'Scan hidden QR code in CYBER, upload photo',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Barcode: JAMES_BOND (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Scanned barcode = \'JAMES_BOND\'',
      'Mandatory Upload': 'Yes (Barcode Scan & Photo MANDATORY)'
    },
    {
      'Stage / Round': 'Round 4 Location Clue',
      'Location': 'In-App Riddle',
      'Challenge / Item': 'MBA Location Riddle',
      'Question / Prompt': 'Numbers become stories... Stories become strategies... Managers learn -> Destination: MBA -> Start Code: MONEY_BROTHA',
      'Media Attached': 'None',
      'Expected Answer / Code': 'MBA & Start Code: MONEY_BROTHA (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'MBA\' & \'MONEY_BROTHA\'',
      'Mandatory Upload': 'No (Inputs Mandatory)'
    },
    {
      'Stage / Round': 'Round 4 Checkpoint (R4)',
      'Location': 'MBA',
      'Challenge / Item': 'Physical Challenge',
      'Question / Prompt': 'Complete physical challenge with volunteer, select PASSED, upload photo, enter finish code RAVI-KISHEN',
      'Media Attached': 'None',
      'Expected Answer / Code': 'Status: PASSED, Finish Code: RAVI-KISHEN (UPPERCASE ONLY)',
      'Verification / Constraint Rule': 'Status = PASSED & normalize-space(.) = \'RAVI-KISHEN\'',
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
      'Question / Prompt': 'Solve puzzle, upload photo of solved puzzle, enter clearance code FINAL-PATH4, run to ring the bell to win',
      'Media Attached': 'None',
      'Expected Answer / Code': 'FINAL-PATH4 (UPPERCASE ONLY) -> Ring Bell',
      'Verification / Constraint Rule': 'Strict UPPERCASE regex & normalize-space(.) = \'FINAL-PATH4\'',
      'Mandatory Upload': 'Yes (Solved Puzzle Photo MANDATORY)'
    }
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'PATH4_MASTER_KEY');
  return wb;
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
    border-bottom: 3px solid #d97706;
    padding-bottom: 6px;
    margin-bottom: 12px;
  }
  .header h1 {
    color: #b45309;
    margin: 0 0 3px 0;
    font-size: 15pt;
    letter-spacing: 0.4px;
  }
  .header .badge {
    display: inline-block;
    background: #d97706;
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
    background: #fffbeb;
    padding: 6px 8px;
    border-radius: 4px;
    border-left: 3px solid #f59e0b;
  }
  .meta-card strong {
    color: #b45309;
    display: block;
    font-size: 7.5pt;
    text-transform: uppercase;
  }
  .meta-card span {
    font-size: 8.5pt;
    font-weight: 600;
  }
  .section-title {
    background: #b45309;
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
    color: #b45309;
    font-size: 9pt;
    display: flex;
    justify-content: space-between;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 3px;
  }
  .rules-box {
    background: #fffbeb;
    border: 1px solid #fde68a;
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
    <span class="badge">PATH 4 OFFICIAL MASTER KEY</span><br>
    <small style="color: #64748b;">FINAL CLUE 2026</small>
  </div>
  <h1>FINAL CLUE &bull; ROUTE 4 / PATH 4</h1>
  <div style="font-size: 9pt; color: #475569; font-weight: 600;">
    CONFIDENTIAL ORGANIZER MASTER KEY &bull; PROGRESSION: OLD CANTEEN &rarr; CSE &rarr; CYBER &rarr; MBA &rarr; AUDITORIUM
  </div>
</div>

<div class="meta-grid">
  <div class="meta-card">
    <strong>Starting Station</strong>
    <span>OLD CANTEEN</span>
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
    <span>morse.jpeg (Morse Code Key)</span>
  </div>
</div>

<div class="rules-box">
  <strong>🏆 Path 4 Team Qualification Rules:</strong><br>
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
      <td>OLD CANTEEN</td>
      <td>Find Cat Board object &amp; verify with volunteer</td>
      <td><span class="code-badge">MEOW-BOW-4</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R2 Location Clue</strong></td>
      <td>In-App</td>
      <td>Emoji Math: 🍎+🍎=10, 🍎+🍌=7, 🍌+🍇=6 -> 🍇+🍎=?</td>
      <td>Result: <code>9</code><br>Start Code: <span class="code-badge">CSE-START</span></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R2 Mini Challenge</strong></td>
      <td>CSE</td>
      <td>
        &bull; <strong>Variant A:</strong> Algo Relay (10 MCQs) + Cyber Detective (10 MCQs) &rarr; Code: <span class="code-badge">CSE-EINSTEIN-4</span><br>
        &bull; <strong>Variant B:</strong> Bug Hunter (10 MCQs) + Tech Detective (10 MCQs) &rarr; Code: <span class="code-badge">CSE-ZUCKER-4</span>
      </td>
      <td><span class="code-badge">CSE-EINSTEIN-4</span> / <span class="code-badge">CSE-ZUCKER-4</span></td>
      <td>📸 Mandatory Photo</td>
    </tr>
    <tr>
      <td><strong>R3 Location Clue</strong></td>
      <td>In-App (<code>morse.jpeg</code>)</td>
      <td>Morse Code: <code>--. --- / - --- / -.-. -.-- -... . .-.</code></td>
      <td>Destination: <code>GO TO CYBER</code></td>
      <td>Mandatory Input</td>
    </tr>
    <tr>
      <td><strong>R3 QR Hunt</strong></td>
      <td>CYBER</td>
      <td>Physically search CYBER for hidden QR code and scan</td>
      <td>Barcode: <span class="code-badge">JAMES_BOND</span></td>
      <td>📸 Photo + Scan</td>
    </tr>
    <tr>
      <td><strong>R4 Location Clue</strong></td>
      <td>In-App</td>
      <td>Wisdom + Look below + Next career step</td>
      <td>Destination: <code>MBA</code><br>Start Code: <span class="code-badge">MONEY_BROTHA</span></td>
      <td>Mandatory Inputs</td>
    </tr>
    <tr>
      <td><strong>Round 4 (R4)</strong></td>
      <td>MBA</td>
      <td>Complete physical challenge with volunteer</td>
      <td>Status: <code>PASSED</code><br>Finish Code: <span class="code-badge">RAVI-KISHEN</span></td>
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
      <td><span class="code-badge">FINAL-PATH4</span><br>&rarr; Ring Bell! 🔔</td>
      <td>📸 Mandatory Photo of Solved Puzzle</td>
    </tr>
  </tbody>
</table>

<div class="section-title">2. Volunteer Station Instructions &amp; Verification Procedures</div>

<div class="station-card">
  <h3><span>OLD CANTEEN Station: Object Finding</span><span class="code-badge">CODE: MEOW-BOW-4</span></h3>
  <div><strong>Volunteer Instructions:</strong> Participants locate the assigned Cat Board object. Once verified with photo, provide code <code>MEOW-BOW-4</code>.</div>
</div>

<div class="station-card">
  <h3><span>CSE Station: Arrival &amp; Mini-Challenges</span><span class="code-badge">START: CSE-START &bull; VAR A: CSE-EINSTEIN-4 &bull; VAR B: CSE-ZUCKER-4</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams arrive at CSE. Provide start code <code>CSE-START</code>.
  <br>&bull; <strong>Variant A (Logical Challenge 1):</strong> Team completes Algo Relay (10) + Cyber Detective (10) and enters <code>CSE-EINSTEIN-4</code>.
  <br>&bull; <strong>Variant B (Logical Challenge 2):</strong> Team completes Bug Hunter (10) + Tech Detective (10) and enters <code>CSE-ZUCKER-4</code>.</div>
</div>

<div class="station-card">
  <h3><span>CYBER Station: QR Hunt</span><span class="code-badge">SCANNED QR: JAMES_BOND</span></h3>
  <div><strong>Volunteer Instructions:</strong> Ensure the physical QR code with payload <code>JAMES_BOND</code> is hidden in the CYBER area. Participants scan it using the in-app barcode scanner.</div>
</div>

<div class="station-card">
  <h3><span>MBA Station: Physical Challenge</span><span class="code-badge">START: MONEY_BROTHA &bull; FINISH: RAVI-KISHEN</span></h3>
  <div><strong>Volunteer Instructions:</strong> Provide start code <code>MONEY_BROTHA</code>. Supervise the physical challenge. Upon successful completion and photo upload, provide finish code <code>RAVI-KISHEN</code>.</div>
</div>

<div class="station-card">
  <h3><span>Main Auditorium Stage: Grand Finale</span><span class="code-badge">CODE: FINAL-PATH4 &bull; 🔔 RING THE BELL</span></h3>
  <div><strong>Volunteer Instructions:</strong> Teams solve Riddle 1 (<code>AUDITORIUM</code>) and Riddle 2 (<code>STAGE</code>), upload a clear photo of their solved puzzle sheet, and receive clearance code <code>FINAL-PATH4</code> from Chief Judges before running to ring the victory bell!</div>
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
  if (fs.existsSync(sourceImg) && !fs.existsSync(destImg)) {
    fs.copyFileSync(sourceImg, destImg);
  }

  console.log('Generating PATH4_FINAL_ODK.xlsx with STRICT UPPERCASE ONLY & 100% MANDATORY enforcement...');
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

  console.log('Generating PATH4_ANSWER_KEY.xlsx...');
  const keyWb = buildAnswerKeyWorkbook();
  const keyPath = path.join(route4Dir, 'PATH4_ANSWER_KEY.xlsx');
  XLSX.writeFile(keyWb, keyPath);
  console.log(`Successfully created: ${keyPath}`);

  console.log('Generating PATH4_ANSWER_KEY.pdf...');
  const pdfHtml = buildAnswerKeyPdfHtml();
  const pdfPath = path.join(route4Dir, 'PATH4_ANSWER_KEY.pdf');
  renderHtmlToPdf(pdfHtml, pdfPath);
  console.log(`Successfully created: ${pdfPath}`);

  console.log('Packaging PATH4_COMPLETE_PACKAGE.zip and PATH4_MEDIA.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route4Dir}\\PATH4_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route4Dir}\\PATH4_FINAL_ODK.xlsx', '${route4Dir}\\PATH4_ANSWER_KEY.xlsx', '${route4Dir}\\PATH4_ANSWER_KEY.pdf', '${mediaDir}' -DestinationPath '${route4Dir}\\PATH4_COMPLETE_PACKAGE.zip' -Force"`);

  console.log('🎉 PATH 4 PACKAGE GENERATED SUCCESSFULLY!');
}

main();

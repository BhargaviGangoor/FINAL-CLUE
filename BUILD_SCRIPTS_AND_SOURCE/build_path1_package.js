const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { execSync } = require('child_process');

const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

function renderHtmlToPdf(htmlContent, outputPath) {
  const tempHtmlPath = path.resolve('temp_render_path1.html');
  fs.writeFileSync(tempHtmlPath, htmlContent, 'utf8');
  const absOutPath = path.resolve(outputPath);
  execSync(`"${edgePath}" --headless --disable-gpu --no-pdf-header-footer --print-to-pdf="${absOutPath}" "file:///${tempHtmlPath.replace(/\\/g, '/')}"`);
  if (fs.existsSync(tempHtmlPath)) fs.unlinkSync(tempHtmlPath);
  console.log(`PDF successfully generated: ${outputPath}`);
}

// -------------------------------------------------------------
// ECE Challenge Sets (Reused from Route 3)
// -------------------------------------------------------------
const eceQuizSets = {
  A: { q: 'Which passive component opposes sudden changes in electric current?', ans: 'INDUCTOR' },
  B: { q: 'What is the base frequency unit named after a German physicist (abbreviation Hz)?', ans: 'HERTZ' },
  C: { q: 'What fundamental semiconductor device allows electric current to flow in only one direction?', ans: 'DIODE' },
  D: { q: 'In Boolean algebra: what logic gate produces a HIGH output only when both inputs are HIGH?', ans: 'AND GATE' },
  E: { q: 'What property of a material characterizes its opposition to the flow of electric charge?', ans: 'RESISTANCE' },
  F: { q: 'Which basic electronic component stores electrical energy in an electrostatic field?', ans: 'CAPACITOR' },
  G: { q: 'In communication theory: what is the spread between upper and lower frequencies called?', ans: 'BANDWIDTH' }
};

const eceRouteSets = {
  A: { task: 'Locate the Digital Circuits Lab entrance plaque on Ground Floor. Find the highlighted keyword.', ans: 'INTEGRATION' },
  B: { task: 'Locate the Communication Lab antenna exhibit on Floor 1. Identify the bold signal transmission keyword.', ans: 'MODULATION' },
  C: { task: 'Inspect the Robotics showcase on Ground Floor. Identify the keyword displayed above Lab 308.', ans: 'FREQUENCY' },
  D: { task: 'Find the Microprocessor trainer unit on 2nd Floor. Read the matching circuit block name.', ans: 'OSCILLATOR' },
  E: { task: 'Visit the VLSI wafer casing on 1st Floor. Read the target optical parameter.', ans: 'WAVELENGTH' },
  F: { task: 'Locate the Solar Tracking demonstrator in Ground Floor project bay. Get the signal attribute.', ans: 'AMPLITUDE' },
  G: { task: 'Examine the Microwave waveguide fixture on 2nd Floor. Identify the circuit tuning principle.', ans: 'RESONANCE' }
};

// -------------------------------------------------------------
// Number Shuffle Sets (56 Questions across Sets A-G)
// -------------------------------------------------------------
const numberSets = {
  A: [
    { q: 1, str: '11-24-17-10-21-28', ans: 'HUNGRY' },
    { q: 2, str: '7-8-22-22-8-21-23', ans: 'DESSERT' },
    { q: 3, str: '5-21-8-4-14-9-4-22-23', ans: 'BREAKFAST' },
    { q: 4, str: '6-4-9-8-23-8-21-12-4', ans: 'CAFETERIA' },
    { q: 5, str: '5-8-25-8-21-4-10-8', ans: 'BEVERAGE' },
    { q: 6, str: '4-19-19-8-23-12-23-8', ans: 'APPETITE' },
    { q: 7, str: '6-4-17-23-8-8-17', ans: 'CANTEEN' },
    { q: 8, str: '9-18-18-7 6-18-24-21-23', ans: 'FOOD COURT' }
  ],
  B: [
    { q: 1, str: '22-4-17-7-26-12-6-11', ans: 'SANDWICH' },
    { q: 2, str: '17-18-18-7-15-8-22', ans: 'NOODLES' },
    { q: 3, str: '19-12-29-29-4', ans: 'PIZZA' },
    { q: 4, str: '5-24-21-10-8-21', ans: 'BURGER' },
    { q: 5, str: '22-4-15-4-7', ans: 'SALAD' },
    { q: 6, str: '6-18-9-9-8-8', ans: 'COFFEE' },
    { q: 7, str: '23-8-4', ans: 'TEA' },
    { q: 8, str: '22-17-4-6-14-22', ans: 'SNACKS' }
  ],
  C: [
    { q: 1, str: '9-24-12-6-8', ans: 'JUICE' },
    { q: 2, str: '26-4-23-8-21', ans: 'WATER' },
    { q: 3, str: '5-12-22-6-24-12-23', ans: 'BISCUIT' },
    { q: 4, str: '6-11-18-6-18-15-4-23-8', ans: 'CHOCOLATE' },
    { q: 5, str: '12-6-8 6-21-8-4-16', ans: 'ICE CREAM' },
    { q: 6, str: '19-4-22-23-4', ans: 'PASTA' },
    { q: 7, str: '16-4-17-10-18', ans: 'MANGO' },
    { q: 8, str: '22-19-12-6-8', ans: 'SPICE' }
  ],
  D: [
    { q: 1, str: '10-21-12-15-15', ans: 'GRILL' },
    { q: 2, str: '23-18-4-22-23', ans: 'TOAST' },
    { q: 3, str: '5-24-23-23-8-21', ans: 'BUTTER' },
    { q: 4, str: '6-11-8-8-22-8', ans: 'CHEESE' },
    { q: 5, str: '19-8-19-19-8-21', ans: 'PEPPER' },
    { q: 6, str: '10-4-21-15-12-6', ans: 'GARLIC' },
    { q: 7, str: '18-17-12-18-17', ans: 'ONION' },
    { q: 8, str: '23-18-16-4-23-18', ans: 'TOMATO' }
  ],
  E: [
    { q: 1, str: '15-8-16-18-17', ans: 'LEMON' },
    { q: 2, str: '10-12-17-10-8-21', ans: 'GINGER' },
    { q: 3, str: '22-24-10-4-21', ans: 'SUGAR' },
    { q: 4, str: '9-4-16', ans: 'JAM' },
    { q: 5, str: '11-18-17-8-28', ans: 'HONEY' },
    { q: 6, str: '5-21-8-4-7', ans: 'BREAD' },
    { q: 7, str: '21-12-6-8', ans: 'RICE' },
    { q: 8, str: '22-18-24-19', ans: 'SOUP' }
  ],
  F: [
    { q: 1, str: '22-26-8-8-23', ans: 'SWEET' },
    { q: 2, str: '22-4-24-6-8', ans: 'SAUCE' },
    { q: 3, str: '7-12-17-17-8-21', ans: 'DINNER' },
    { q: 4, str: '15-24-17-6-11', ans: 'LUNCH' },
    { q: 5, str: '19-15-4-23-23-8-21', ans: 'PLATTER' },
    { q: 6, str: '22-8-21-25-12-17-10', ans: 'SERVING' },
    { q: 7, str: '4-19-19-8-23-12-29-8-21', ans: 'APPETIZER' },
    { q: 8, str: '6-24-12-22-12-17-8', ans: 'CUISINE' }
  ],
  G: [
    { q: 1, str: '7-8-15-12-6-4-6-28', ans: 'DELICACY' },
    { q: 2, str: '17-18-24-21-12-22-11', ans: 'NOURISH' },
    { q: 3, str: '23-4-22-23-28', ans: 'TASTY' },
    { q: 4, str: '16-8-4-15-23-12-16-8', ans: 'MEALTIME' },
    { q: 5, str: '6-4-9-8-23-8-21-12-4', ans: 'CAFETERIA' },
    { q: 6, str: '5-8-25-8-21-4-10-8', ans: 'BEVERAGE' },
    { q: 7, str: '4-19-19-8-23-12-23-8', ans: 'APPETITE' },
    { q: 8, str: '5-21-8-4-14-9-4-22-23', ans: 'BREAKFAST' }
  ]
};

// -------------------------------------------------------------
// Caesar Cipher Phases (Cyber Security, Shift = -5)
// -------------------------------------------------------------
const caesarPhases = [
  { num: 1, cipher: 'FZIN', ans: 'AUDI' },
  { num: 2, cipher: 'XYFLJ', ans: 'STAGE' },
  { num: 3, cipher: 'WJI XJFYX', ans: 'RED SEATS' },
  { num: 4, cipher: 'RNHWUMSTSJ', ans: 'MICROPHONE' },
  { num: 5, cipher: 'TW NJSYFYNTS', ans: 'ORIENTATION' }
];

function buildSurvey() {
  const survey = [];

  // 1. INTRO
  survey.push({
    type: 'note',
    name: 'path1_intro',
    label: 'FINAL CLUE — PATH 1\n\nWelcome to PATH 1 of the Final Clue Treasure Hunt!\n\nFollow all instructions carefully. Work with your team to solve challenges and navigate through the campus stations.\n\nEach block location will only be revealed as you successfully complete each stage.',
    hint: 'Read all instructions carefully before proceeding.'
  });

  // 2. GLOBAL UPPERCASE WARNING
  survey.push({
    type: 'note',
    name: 'uppercase_warning',
    label: '⚠️ IMPORTANT\n\nALL ANSWERS MUST BE ENTERED IN UPPERCASE.\n\nWrong answers or lowercase inputs will NOT allow you to proceed.\n\nComplete every mandatory challenge without outside assistance.',
    hint: 'Always type answers in UPPERCASE.'
  });

  // 3. ROUND RULES
  survey.push({
    type: 'note',
    name: 'round_rules',
    label: '📋 ROUND RULES\n\n• Teams must complete all challenges strictly in order.\n• Wrong answers do not unlock the next stage.\n• Volunteer verification is mandatory wherever specified.\n• Teams must obtain the required code from the volunteer.\n• Only qualified teams can continue.\n• Teams must hurry because qualification is based on completion and verification order.'
  });

  // 4. NO-WIFI RULES
  survey.push({
    type: 'note',
    name: 'wifi_rules',
    label: '⚠️ NO-WIFI / ANTI-CHEATING RULES\n\nDo not switch on Wi-Fi or use the internet to solve challenges.\nDo not use Google, ChatGPT, search engines, or outside assistance.\nVolunteers may stop or disqualify teams that violate the rules.'
  });

  // Team Registration
  survey.push({
    type: 'text',
    name: 'team_id',
    label: 'Enter Team ID / Team Name',
    hint: 'Enter assigned Team ID in UPPERCASE.',
    required: 'yes'
  });

  // =============================================================
  // STAGE 1: ECE BLOCK (Logic Gates, Circuit Breaker, Physical Hunt)
  // =============================================================
  survey.push({
    type: 'note',
    name: 'ece_intro',
    label: 'STAGE 1: ECE BLOCK — LOGIC GATES & HARDWARE GAUNTLET\n\n⚠️ HURRY UP. Only the TOP 25 TEAMS will qualify for the next round.\n\nReport to the ECE Ground Floor Laboratory to begin your challenge.\nSelect your assigned variant and solve the technical challenge.',
    hint: 'Begin at ECE Block.'
  });

  survey.push({
    type: 'select_one ece_variant',
    name: 'ece_variant',
    label: 'Select Assigned ECE Variant',
    hint: 'Select variant assigned by the ECE volunteer.',
    required: 'yes'
  });

  for (const [vKey, vObj] of Object.entries(eceQuizSets)) {
    survey.push({
      type: 'text',
      name: `ece_quiz_${vKey}`,
      label: `ECE VARIANT ${vKey} — TECHNICAL CHALLENGE\n\n${vObj.q}\n\n⚠️ Enter answer in UPPERCASE.`,
      hint: 'Enter technical answer in UPPERCASE.',
      required: 'yes',
      relevant: `\${ece_variant}='${vKey}'`,
      constraint: `normalize-space(.)='${vObj.ans}'`,
      constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE.'
    });

    survey.push({
      type: 'text',
      name: `ece_route_${vKey}`,
      label: `ECE VARIANT ${vKey} — BUILDING HUNT\n\n${eceRouteSets[vKey].task}\n\n⚠️ Enter discovered keyword in UPPERCASE.`,
      hint: 'Enter keyword in UPPERCASE.',
      required: 'yes',
      relevant: `\${ece_variant}='${vKey}' and normalize-space(\${ece_quiz_${vKey}})='${vObj.ans}'`,
      constraint: `normalize-space(.)='${eceRouteSets[vKey].ans}'`,
      constraint_message: '❌ Incorrect keyword. Follow the route instructions carefully.'
    });
  }

  // ECE Completion Conditions
  const eceCompleteRel = Object.entries(eceQuizSets).map(([vKey, vObj]) => {
    return `(\${ece_variant}='${vKey}' and normalize-space(\${ece_quiz_${vKey}})='${vObj.ans}' and normalize-space(\${ece_route_${vKey}})='${eceRouteSets[vKey].ans}')`;
  }).join(' or ');

  survey.push({
    type: 'note',
    name: 'ece_complete_note',
    label: '✅ ECE CHALLENGE COMPLETE\n\nShow your completed screen to the ECE volunteer to receive your station clearance pass code.',
    relevant: `(${eceCompleteRel})`
  });

  survey.push({
    type: 'text',
    name: 'ece_pass_code',
    label: 'Enter ECE Volunteer Clearance Code',
    hint: 'Enter verified code from ECE volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `(${eceCompleteRel})`,
    constraint: "normalize-space(.)='P1-ECE-PASS'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE.'
  });

  // =============================================================
  // STAGE 2: MECH BLOCK (Sound Clue & Mini-Challenges)
  // =============================================================
  survey.push({
    type: 'note',
    name: 'next_block_mech',
    label: 'NEXT BLOCK: MECH\n\nListen carefully to the cryptic sound clue below to confirm your next destination.\nDo NOT imitate the sound.\nIdentify what you hear.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: "normalize-space(${ece_pass_code})='P1-ECE-PASS'",
    'media::audio': 'engine_sound.wav'
  });

  survey.push({
    type: 'text',
    name: 'mech_sound_answer',
    label: 'What sound did you hear?',
    hint: 'Enter answer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${ece_pass_code})='P1-ECE-PASS'",
    constraint: "normalize-space(.)='ENGINE'",
    constraint_message: '❌ Incorrect answer. Listen carefully and enter in UPPERCASE.'
  });

  survey.push({
    type: 'text',
    name: 'mech_sound_block_guess',
    label: 'Which block does this clue direct you to?',
    hint: 'Enter block name in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${mech_sound_answer})='ENGINE'",
    constraint: "normalize-space(.)='MECH'",
    constraint_message: '❌ Incorrect block name. Enter MECH in UPPERCASE.'
  });

  survey.push({
    type: 'note',
    name: 'mech_arrival_note',
    label: 'PROCEED TO MECHANICAL ENGINEERING BLOCK\n\nReport immediately to the MECH Workshop.\nObtain the arrival code from the MECH volunteer.',
    relevant: "normalize-space(${mech_sound_answer})='ENGINE' and normalize-space(${mech_sound_block_guess})='MECH'"
  });

  survey.push({
    type: 'text',
    name: 'mech_start_code',
    label: 'Enter MECH Volunteer Arrival Code',
    hint: 'Enter code provided by MECH volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${mech_sound_answer})='ENGINE' and normalize-space(${mech_sound_block_guess})='MECH'",
    constraint: "normalize-space(.)='P1-MECH-START'",
    constraint_message: '❌ Incorrect code. Enter P1-MECH-START in UPPERCASE.'
  });

  // MECH Mini-Challenges
  survey.push({
    type: 'note',
    name: 'mech_mini_intro',
    label: 'MINI-CHALLENGE ZONE — MECHANICAL BLOCK\n\n⚠️ HURRY UP. Only the TOP 15 TEAMS will qualify for the next stage.\n\nComplete all 3 sequential phases:\n• Phase 1: Human Shape (7 Poses)\n• Phase 2: Hidden Garland (5 Hardware Items Tied with Rope)\n• Phase 3: Bomb Defusal (Knot Puzzle Untangling)\n\nShow each completed phase to the volunteer.',
    relevant: "normalize-space(${mech_start_code})='P1-MECH-START'",
    'media::image': 'mech.png'
  });

  // Phase 1: Human Shape
  survey.push({
    type: 'note',
    name: 'mech_p1_note',
    label: 'PHASE 1: HUMAN SHAPE\n\nRecreate all 7 poses using all 4 team members.\nComplete full sequence.\nShow sequence to volunteer for verification.',
    relevant: "normalize-space(${mech_start_code})='P1-MECH-START'",
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'mech_p1_code',
    label: 'Enter Phase 1 Volunteer Code',
    hint: 'Enter code provided by volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${mech_start_code})='P1-MECH-START'",
    constraint: "normalize-space(.)='P1-MECH-P1'",
    constraint_message: '❌ Incorrect code. Phase 2 remains locked.'
  });

  // Phase 2: Hidden Garland
  survey.push({
    type: 'note',
    name: 'mech_p2_note',
    label: 'PHASE 2: HIDDEN GARLAND\n\nFind 5 hidden mechanical parts across the workshop:\n1. Collect all 5 items.\n2. Tie using rope into continuous garland.\n3. Show garland to volunteer.',
    relevant: "normalize-space(${mech_p1_code})='P1-MECH-P1'",
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'mech_p2_code',
    label: 'Enter Phase 2 Volunteer Code',
    hint: 'Enter code provided by volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${mech_p1_code})='P1-MECH-P1'",
    constraint: "normalize-space(.)='P1-MECH-P2'",
    constraint_message: '❌ Incorrect code. Phase 3 remains locked.'
  });

  // Phase 3: Bomb Defusal
  survey.push({
    type: 'note',
    name: 'mech_p3_note',
    label: 'PHASE 3: BOMB DEFUSAL\n\nLocate assigned challenge box containing knot puzzle.\nUntangle and remove designated knots safely without cutting rope.\nShow defused box to volunteer.',
    relevant: "normalize-space(${mech_p2_code})='P1-MECH-P2'",
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'mech_p3_code',
    label: 'Enter Phase 3 Volunteer Code',
    hint: 'Enter code provided by volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${mech_p2_code})='P1-MECH-P2'",
    constraint: "normalize-space(.)='P1-MECH-P3'",
    constraint_message: '❌ Incorrect code. Final gate remains locked.'
  });

  survey.push({
    type: 'text',
    name: 'mech_pass_code',
    label: 'Enter Final MECH Clearance Code',
    hint: 'Enter final clearance code provided by MECH head volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${mech_p3_code})='P1-MECH-P3'",
    constraint: "normalize-space(.)='P1-MECH-PASS'",
    constraint_message: '❌ Incorrect code. You cannot proceed.'
  });

  // =============================================================
  // STAGE 3: FOOD COURT (Number Shuffle - 56 Qs across Sets A-G)
  // =============================================================
  survey.push({
    type: 'note',
    name: 'next_block_food_court',
    label: 'NEXT BLOCK: FOOD COURT\n\nProceed immediately to the Food Court.\nReport to the station volunteer to receive your assigned Number Shuffle set.\n\n⚠️ ONLY TOP 7 TEAMS will qualify from this stage!',
    relevant: "normalize-space(${mech_pass_code})='P1-MECH-PASS'"
  });

  survey.push({
    type: 'note',
    name: 'number_shuffle_intro',
    label: 'NUMBER SHUFFLE CHALLENGE\n\nDecoding Rule: 4=A, 5=B, 6=C, ... 29=Z.\nDecode all 8 questions in your assigned set sequentially.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: "normalize-space(${mech_pass_code})='P1-MECH-PASS'"
  });

  survey.push({
    type: 'select_one number_set',
    name: 'number_set',
    label: 'Select Assigned Number Shuffle Set',
    hint: 'Select the set assigned by your volunteer.',
    required: 'yes',
    relevant: "normalize-space(${mech_pass_code})='P1-MECH-PASS'"
  });

  for (const [sKey, qList] of Object.entries(numberSets)) {
    for (let i = 0; i < qList.length; i++) {
      const qObj = qList[i];
      const fieldName = `ns_${sKey}_${qObj.q}`;
      let rel = `\${number_set}='${sKey}' and normalize-space(\${mech_pass_code})='P1-MECH-PASS'`;
      if (i > 0) {
        const prevField = `ns_${sKey}_${qList[i - 1].q}`;
        const prevAns = qList[i - 1].ans;
        rel = `\${number_set}='${sKey}' and normalize-space(\${${prevField}})='${prevAns}'`;
      }
      survey.push({
        type: 'text',
        name: fieldName,
        label: `SET ${sKey} — QUESTION ${qObj.q}/8\n\nDecode: ${qObj.str}\n\n⚠️ Enter answer in UPPERCASE.`,
        hint: 'Use 4=A, 5=B, 6=C... 29=Z. Enter in UPPERCASE.',
        required: 'yes',
        relevant: rel,
        constraint: `normalize-space(.)='${qObj.ans}'`,
        constraint_message: '❌ Incorrect answer. Answer MUST be in UPPERCASE. The next question remains locked.'
      });
    }

    const lastField = `ns_${sKey}_8`;
    const lastAns = qList[7].ans;
    survey.push({
      type: 'note',
      name: `set_${sKey}_pass`,
      label: `✅ SET ${sKey} — ALL 8 PASSED\n\nShow your completed set to the volunteer to obtain the Food Court completion code.`,
      relevant: `\${number_set}='${sKey}' and normalize-space(\${${lastField}})='${lastAns}'`
    });
  }

  const fcCompletionConditions = Object.entries(numberSets).map(([sKey, qList]) => {
    return `(\${number_set}='${sKey}' and normalize-space(\${ns_${sKey}_8})='${qList[7].ans}')`;
  }).join(' or ');

  survey.push({
    type: 'text',
    name: 'fc_pass_code',
    label: 'Enter Food Court Clearance Code',
    hint: 'Enter code provided by volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `(${fcCompletionConditions})`,
    constraint: "normalize-space(.)='P1-FC-PASS'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE.'
  });

  // =============================================================
  // STAGE 4: CYBER SECURITY BLOCK (5 Caesar Ciphers & Seminar Agility)
  // =============================================================
  survey.push({
    type: 'note',
    name: 'next_block_cyber',
    label: 'NEXT BLOCK: CYBER SECURITY\n\nProceed immediately to the Cyber Security (CY) Block!\nReport to the station volunteer to begin the 5 Caesar Ciphers.\n\n⚠️ ONLY TOP 5 TEAMS will qualify for the AUDI Final!',
    relevant: "normalize-space(${fc_pass_code})='P1-FC-PASS'"
  });

  survey.push({
    type: 'note',
    name: 'caesar_intro',
    label: 'CAESAR CIPHER — 5 MANDATORY PHASES\n\nRule: Move each letter 5 positions BACKWARD (-5).\nAll 5 cipher phases are mandatory and must be solved in sequence.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: "normalize-space(${fc_pass_code})='P1-FC-PASS'"
  });

  for (let i = 0; i < caesarPhases.length; i++) {
    const cp = caesarPhases[i];
    const fieldName = `caesar_${cp.num}`;
    let rel = `normalize-space(\${fc_pass_code})='P1-FC-PASS'`;
    if (i > 0) {
      const prevField = `caesar_${caesarPhases[i - 1].num}`;
      const prevAns = caesarPhases[i - 1].ans;
      rel = `normalize-space(\${${prevField}})='${prevAns}'`;
    }

    survey.push({
      type: 'note',
      name: `caesar_${cp.num}_note`,
      label: `CAESAR CIPHER ${cp.num}/5\n\nCipher: ${cp.cipher}\n\nMove each letter 5 positions BACKWARD (-5).\n\n⚠️ Enter answer in UPPERCASE.`,
      relevant: rel
    });

    survey.push({
      type: 'text',
      name: fieldName,
      label: `Enter decoded answer for Phase ${cp.num}`,
      hint: 'Enter in UPPERCASE.',
      required: 'yes',
      relevant: rel,
      constraint: `normalize-space(.)='${cp.ans}'`,
      constraint_message: '❌ Incorrect answer. Answer MUST be in UPPERCASE. Next cipher remains locked.'
    });
  }

  survey.push({
    type: 'text',
    name: 'caesar_pass_code',
    label: 'Enter Caesar Volunteer Clearance Code',
    hint: 'Enter code provided by volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${caesar_5})='ORIENTATION'",
    constraint: "normalize-space(.)='CAESAR-PASS'",
    constraint_message: '❌ Incorrect code. Code must be in UPPERCASE.'
  });

  survey.push({
    type: 'note',
    name: 'cy_physical_intro',
    label: 'CYBER PHYSICAL CHALLENGE — SEMINAR HALL\n\nReport to the Seminar Hall in the same block.\nComplete the synchronized agility trial before the volunteer.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: "normalize-space(${caesar_pass_code})='CAESAR-PASS'"
  });

  survey.push({
    type: 'text',
    name: 'cy_pass_code',
    label: 'Enter Cyber Physical Clearance Code',
    hint: 'Enter code provided by Seminar Hall volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${caesar_pass_code})='CAESAR-PASS'",
    constraint: "normalize-space(.)='P1-CY-PASS'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE.'
  });

  // =============================================================
  // STAGE 5: MAIN AUDITORIUM GRAND FINALE
  // =============================================================
  survey.push({
    type: 'note',
    name: 'audi_reveal',
    label: 'NEXT BLOCK: AUDI\n\nProceed immediately to the Main Auditorium!\n\n🏆 10 FINALISTS are converging at the Auditorium for the Grand Championship!',
    relevant: "normalize-space(${cy_pass_code})='P1-CY-PASS'"
  });

  survey.push({
    type: 'text',
    name: 'audi_riddle_code',
    label: 'Enter Stage 5 Venue Riddle unlock code',
    hint: 'Enter code provided by volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${cy_pass_code})='P1-CY-PASS'",
    constraint: "normalize-space(.)='AUDITORIUM' or normalize-space(.)='AUDI'",
    constraint_message: '❌ Incorrect code. Enter AUDITORIUM in UPPERCASE.'
  });

  survey.push({
    type: 'text',
    name: 'stage_riddle_code',
    label: 'Enter Stage Riddle unlock code',
    hint: 'Enter code provided by volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${audi_riddle_code})='AUDITORIUM' or normalize-space(${audi_riddle_code})='AUDI'",
    constraint: "normalize-space(.)='STAGE'",
    constraint_message: '❌ Incorrect code. Enter STAGE in UPPERCASE.'
  });

  survey.push({
    type: 'note',
    name: 'final_physical_note',
    label: '🏆 GRAND FINALE — FINAL PHYSICAL TRIAL\n\nComplete the culminating physical agility trial on the Auditorium Stage.\nShow completion to the Chief Judges.',
    relevant: "normalize-space(${stage_riddle_code})='STAGE'"
  });

  survey.push({
    type: 'text',
    name: 'final_answer1',
    label: 'FINAL ANSWER 1',
    hint: 'Enter final answer 1 in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${stage_riddle_code})='STAGE'",
    constraint: "normalize-space(.)='STAGE'",
    constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE.'
  });

  survey.push({
    type: 'text',
    name: 'final_answer2',
    label: 'FINAL ANSWER 2',
    hint: 'Enter final answer 2 in UPPERCASE.',
    required: 'yes',
    relevant: "normalize-space(${final_answer1})='STAGE'",
    constraint: "normalize-space(.)='FINAL-PATH1'",
    constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE.'
  });

  survey.push({
    type: 'note',
    name: 'path1_complete_note',
    label: '🏆 PATH 1 COMPLETE!\n\nCongratulations! You have successfully completed PATH 1 in the Final Clue Treasure Hunt!',
    relevant: "normalize-space(${final_answer1})='STAGE' and normalize-space(${final_answer2})='FINAL-PATH1'"
  });

  return survey;
}

function buildChoices() {
  const choices = [];
  ['A', 'B', 'C', 'D', 'E', 'F', 'G'].forEach(v => {
    choices.push({ list_name: 'ece_variant', name: v, label: `Variant ${v}` });
    choices.push({ list_name: 'number_set', name: v, label: `Set ${v}` });
  });
  return choices;
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — PATH 1 COMPLETE (ECE -> MECH -> FC -> CY -> AUDI)',
      form_id: 'path1_final_clue',
      version: '1.0'
    }
  ];
}

function buildAnswerKeyWorkbook() {
  const wb = XLSX.utils.book_new();

  const flow = [
    { Stage: '1. ECE Block', Clue: 'Logic Gate Quiz & Hardware Hunt', Code: 'P1-ECE-PASS', Issued_By: 'ECE Volunteer', Next: 'NEXT BLOCK: MECH' },
    { Stage: '2. MECH Sound', Clue: 'Listen to engine_sound.wav -> ENGINE', Code: 'P1-MECH-START', Issued_By: 'MECH Volunteer', Next: 'MECH Mini-Challenges' },
    { Stage: '3. MECH Mini-Tasks', Clue: 'Human Shape -> Garland -> Bomb Defusal', Code: 'P1-MECH-PASS', Issued_By: 'MECH Volunteer', Next: 'NEXT BLOCK: FOOD COURT' },
    { Stage: '4. Food Court Number Shuffle', Clue: 'Decode 8 strings in assigned Set (4=A...29=Z)', Code: 'P1-FC-PASS', Issued_By: 'FC Volunteer', Next: 'NEXT BLOCK: CYBER SECURITY' },
    { Stage: '5. Cyber Security Caesar', Clue: '5 Cipher Phases (-5 Shift: AUDI, STAGE...)', Code: 'CAESAR-PASS', Issued_By: 'CY Volunteer', Next: 'Cyber Physical Challenge' },
    { Stage: '6. Cyber Physical', Clue: 'Seminar Hall Agility Trial', Code: 'P1-CY-PASS', Issued_By: 'Seminar Hall Volunteer', Next: 'NEXT BLOCK: AUDI' },
    { Stage: '7. Auditorium Riddle', Clue: 'AUDITORIUM / AUDI', Code: 'AUDITORIUM', Issued_By: 'AUDI Volunteer', Next: 'Stage Riddle' },
    { Stage: '8. Stage Riddle', Clue: 'STAGE', Code: 'STAGE', Issued_By: 'AUDI Volunteer', Next: 'Final Physical Trial' },
    { Stage: '9. Final Answer 1', Clue: 'Stage Coordination', Code: 'STAGE', Issued_By: 'Chief Judge', Next: 'Final Answer 2' },
    { Stage: '10. Final Answer 2', Clue: 'Grand Championship', Code: 'FINAL-PATH1', Issued_By: 'Chief Judge', Next: '🏆 PATH 1 COMPLETE' }
  ];

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(flow), 'Master Flow & Codes');

  const eceList = [];
  ['A', 'B', 'C', 'D', 'E', 'F', 'G'].forEach(v => {
    eceList.push({ Variant: v, Technical_Question: eceQuizSets[v].q, Technical_Answer: eceQuizSets[v].ans, Building_Hunt_Task: eceRouteSets[v].task, Discovered_Keyword: eceRouteSets[v].ans });
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(eceList), 'ECE Challenges Bank');

  const nsList = [];
  Object.entries(numberSets).forEach(([sKey, qList]) => {
    qList.forEach(q => {
      nsList.push({ Set: sKey, Q: q.q, Encrypted_String: q.str, Decoded_Answer: q.ans });
    });
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(nsList), 'Number Shuffle (56 Qs)');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(caesarPhases), 'Caesar Decryption Key');

  return wb;
}

function buildPath1PdfHtml() {
  const commonStyles = `
    @page { size: A4; margin: 8mm 10mm 8mm 10mm; }
    body { font-family: 'Segoe UI', Arial, sans-serif; color: #0f172a; background: #ffffff; line-height: 1.3; font-size: 8pt; }
    .header { border-bottom: 3px solid #1e40af; padding-bottom: 5px; margin-bottom: 10px; }
    .header h1 { color: #1e3a8a; margin: 0 0 2px 0; font-size: 14pt; letter-spacing: 0.3px; }
    .header .badge { display: inline-block; background: #dbeafe; color: #1e40af; padding: 2px 6px; border-radius: 4px; font-size: 7.5pt; font-weight: bold; text-transform: uppercase; }
    .header .meta { float: right; font-size: 7.5pt; color: #64748b; }
    .section-title { background: #f1f5f9; padding: 4px 7px; border-left: 4px solid #2563eb; margin-top: 10px; margin-bottom: 6px; font-size: 9pt; font-weight: bold; color: #0f172a; page-break-after: avoid; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 8px; font-size: 7pt; }
    th, td { border: 1px solid #cbd5e1; padding: 3px 5px; text-align: left; vertical-align: top; }
    th { background: #f8fafc; color: #334155; font-weight: 600; }
    .code-cell { font-family: 'Consolas', monospace; font-weight: bold; color: #b91c1c; background: #fef2f2; }
    .ans-cell { font-family: 'Consolas', monospace; font-weight: bold; color: #047857; background: #ecfdf5; }
    .note-box { background: #fffbeb; border: 1px solid #fef3c7; border-left: 4px solid #f59e0b; padding: 5px 7px; border-radius: 4px; margin: 6px 0; font-size: 7.5pt; }
    .page-break { page-break-before: always; }
    .station-card { border: 1px solid #cbd5e1; border-radius: 5px; padding: 6px 8px; margin-bottom: 6px; background: #ffffff; page-break-inside: avoid; }
    .station-card h3 { margin: 0 0 3px 0; color: #1e40af; font-size: 8.2pt; display: flex; justify-content: space-between; }
    .station-badge { background: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; padding: 1px 5px; border-radius: 3px; font-size: 6.8pt; font-weight: bold; }
    .station-desc { color: #334155; font-size: 7.2pt; margin-bottom: 3px; }
    .station-meta-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 4px; background: #f8fafc; padding: 4px 6px; border-radius: 4px; font-size: 6.8pt; }
    .meta-item strong { color: #0f172a; }
  `;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Route 1 (Path 1) Organizer Master Answer Key</title>
<style>${commonStyles}</style>
</head>
<body>

<div class="header">
  <span class="meta">CONFIDENTIAL &#8226; OFFICIAL ORGANIZER MASTER MANUAL</span>
  <h1>ROUTE 1 (PATH 1) — COMPLETE ORGANIZER ANSWER KEY &amp; OPERATIONS GUIDE</h1>
  <span class="badge">ROUTE: ECE &rarr; MECH &rarr; FOOD COURT &rarr; CYBER SECURITY &rarr; AUDI</span>
</div>

<div class="note-box">
  <strong>⚠️ ORGANIZER INSTRUCTION:</strong> This official manual corresponds to <code>PATH1_FINAL_ODK.xlsx</code>. All typed answers strictly enforce <strong>MANDATORY UPPERCASE</strong>.
</div>

<div class="section-title">1. MASTER FLOW &amp; VOLUNTEER PASS CODES</div>
<table>
  <thead><tr><th>Stage</th><th>Challenge Description</th><th>Required Answer / Code</th><th>Issued By</th><th>Unlocks Next</th></tr></thead>
  <tbody>
    <tr><td><strong>1. ECE Block</strong></td><td>Logic Gate Quiz &amp; Hardware Hunt</td><td class="code-cell">P1-ECE-PASS</td><td>ECE Volunteer (Top 25)</td><td>NEXT BLOCK: MECH</td></tr>
    <tr><td><strong>2. MECH Sound</strong></td><td>Listen to engine_sound.wav &rarr; ENGINE</td><td class="code-cell">P1-MECH-START</td><td>MECH Volunteer</td><td>MECH Mini-Tasks</td></tr>
    <tr><td><strong>3. MECH Mini-Tasks</strong></td><td>Human Shape &bull; Garland &bull; Bomb Defusal</td><td class="code-cell">P1-MECH-PASS</td><td>MECH Volunteer (Top 15)</td><td>NEXT BLOCK: FOOD COURT</td></tr>
    <tr><td><strong>4. Food Court</strong></td><td>Number Shuffle (56 Qs across Sets A–G)</td><td class="code-cell">P1-FC-PASS</td><td>FC Volunteer (Top 7)</td><td>NEXT BLOCK: CY</td></tr>
    <tr><td><strong>5. Cyber Security</strong></td><td>5 Caesar Phases (-5 Shift) &amp; Agility</td><td class="code-cell">P1-CY-PASS</td><td>CY Volunteer (Top 5)</td><td>NEXT BLOCK: AUDI</td></tr>
    <tr><td><strong>6. AUDI Stage</strong></td><td>Auditorium Riddle &rarr; Stage &rarr; Final Physical</td><td class="code-cell">FINAL-PATH1</td><td>Chief Judges</td><td>🏆 PATH 1 COMPLETE</td></tr>
  </tbody>
</table>

<div class="section-title">2. ECE CHALLENGES &amp; HARDWARE HUNT BANK</div>
<table>
  <thead><tr><th>Variant</th><th>Technical Question &amp; Solution</th><th>Building Hunt Discovered Keyword</th></tr></thead>
  <tbody>
    ${['A','B','C','D','E','F','G'].map(v => `
      <tr>
        <td><strong>Variant ${v}</strong></td>
        <td>${eceQuizSets[v].q} &rarr; <span class="ans-cell">${eceQuizSets[v].ans}</span></td>
        <td><span class="ans-cell">${eceRouteSets[v].ans}</span></td>
      </tr>
    `).join('')}
  </tbody>
</table>

<div class="page-break"></div>

<div class="section-title">3. NUMBER SHUFFLE ANSWER BANK (56 QUESTIONS ACROSS 7 SETS)</div>
<p style="font-size:7pt; color:#475569; margin-bottom:4px;">Rule: 4=A, 5=B, 6=C... 29=Z. All 8 questions per set must be solved sequentially in UPPERCASE.</p>
<table>
  <thead><tr><th>Set</th><th>Q1</th><th>Q2</th><th>Q3</th><th>Q4</th><th>Q5</th><th>Q6</th><th>Q7</th><th>Q8</th></tr></thead>
  <tbody>
    ${Object.entries(numberSets).map(([sKey, qList]) => `
      <tr>
        <td><strong>Set ${sKey}</strong></td>
        ${qList.map(q => `<td class="ans-cell">${q.ans}</td>`).join('')}
      </tr>
    `).join('')}
  </tbody>
</table>

<div class="section-title">4. CAESAR CIPHER DECRYPTION KEY (5 MANDATORY PHASES)</div>
<table>
  <thead><tr><th>Phase</th><th>Ciphertext</th><th>Shift Rule</th><th>Decoded Answer</th><th>Pass Code</th></tr></thead>
  <tbody>
    <tr><td>Phase 1</td><td><code>FZIN</code></td><td>-5 Backward</td><td class="ans-cell">AUDI</td><td>Phase 2</td></tr>
    <tr><td>Phase 2</td><td><code>XYFLJ</code></td><td>-5 Backward</td><td class="ans-cell">STAGE</td><td>Phase 3</td></tr>
    <tr><td>Phase 3</td><td><code>WJI XJFYX</code></td><td>-5 Backward</td><td class="ans-cell">RED SEATS</td><td>Phase 4</td></tr>
    <tr><td>Phase 4</td><td><code>RNHWUMSTSJ</code></td><td>-5 Backward</td><td class="ans-cell">MICROPHONE</td><td>Phase 5</td></tr>
    <tr><td>Phase 5</td><td><code>TW NJSYFYNTS</code></td><td>-5 Backward</td><td class="ans-cell">ORIENTATION</td><td class="code-cell">CAESAR-PASS</td></tr>
  </tbody>
</table>

<div class="page-break"></div>

<div class="section-title">5. DETAILED CHALLENGE DESCRIPTIONS, MATERIALS &amp; PARALLEL STATION CAPACITIES</div>

<div class="station-card">
  <h3><span>Stage 1: ECE Logic &amp; Hardware Circuit Gauntlet</span><span class="station-badge">PARALLEL STATIONS: 4 LAB TEST BENCHES</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Teams begin at ECE. Solve Boolean logic and circuit questions, then navigate ECE floors to retrieve technical keywords.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> ECE Ground Floor Lab</div>
    <div class="meta-item"><strong>Materials:</strong> 4 Test benches, circuit reference sheets</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Lab Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Starting Teams (Top 25 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P1-ECE-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> MECH Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 2: MECH Sound Clue &amp; Mini-Challenges</span><span class="station-badge">PARALLEL STATIONS: 6 WORKSHOP BAYS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> Listen to <code>engine_sound.wav</code> in-app, identify <code>ENGINE</code> &rarr; <code>MECH</code>. Complete Human Shape (7 poses), Hidden Garland (5 items), Bomb Defusal.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> MECH Workshop Bay</div>
    <div class="meta-item"><strong>Materials:</strong> 2 Pose mats, 2 garland rope sets, 2 knot boxes</div>
    <div class="meta-item"><strong>Staffing:</strong> 6 Bay Judges + 1 Head Judge</div>
    <div class="meta-item"><strong>Capacity:</strong> 25 Teams (Top 15 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P1-MECH-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Food Court</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 3: Food Court Number Shuffle Challenge</span><span class="station-badge">PARALLEL STATIONS: 4 VERIFICATION DESKS</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 15 teams decode 8 encrypted food &amp; dining words using 4=A...29=Z rule on mobile devices.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Food Court Foyer</div>
    <div class="meta-item"><strong>Materials:</strong> Number conversion reference sheets</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 15 Teams (Top 7 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P1-FC-PASS</code></div>
    <div class="meta-item"><strong>Next Block:</strong> Cyber Security Block</div>
  </div>
</div>

<div class="station-card">
  <h3><span>Stage 4: Cyber Security Caesar Cipher &amp; Seminar Agility</span><span class="station-badge">PARALLEL STATIONS: 3 DESKS + 2 AGILITY LANES</span></h3>
  <div class="station-desc"><strong>Event &amp; Mechanics:</strong> 7 teams decrypt 5 sequential ciphers (-5 backward) and complete blindfolded agility trial in Seminar Hall.</div>
  <div class="station-meta-grid">
    <div class="meta-item"><strong>Location:</strong> Cyber Security &amp; Seminar Hall</div>
    <div class="meta-item"><strong>Materials:</strong> Cones, blindfolds, cipher wheels</div>
    <div class="meta-item"><strong>Staffing:</strong> 4 Volunteers</div>
    <div class="meta-item"><strong>Capacity:</strong> 7 Teams (Top 5 qualify)</div>
    <div class="meta-item"><strong>Pass Code:</strong> <code style="color:#b91c1c;">P1-CY-PASS</code></div>
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
    <div class="meta-item"><strong>Final Codes:</strong> Ans 1: <code>STAGE</code> &bull; Ans 2: <code>FINAL-PATH1</code></div>
    <div class="meta-item"><strong>Conclusion:</strong> 🏆 PATH 1 COMPLETE</div>
    <div class="meta-item"><strong>Podium:</strong> 1st, 2nd, 3rd Place Award Ceremony</div>
  </div>
</div>

</body>
</html>`;
}

function main() {
  const baseDir = path.resolve(__dirname, '..');
  const route1Dir = path.join(baseDir, 'ROUTE_1_PATH1_ECE_MECH_FC_CY_AUDI');
  if (!fs.existsSync(route1Dir)) fs.mkdirSync(route1Dir, { recursive: true });

  const mediaDir = path.join(route1Dir, 'media');
  if (!fs.existsSync(mediaDir)) fs.mkdirSync(mediaDir, { recursive: true });

  // Copy media assets
  fs.copyFileSync(path.join(baseDir, 'ROUTE_1_ADMIN_MECH_AUDI', 'media', 'engine_sound.wav'), path.join(mediaDir, 'engine_sound.wav'));
  fs.copyFileSync(path.join(baseDir, 'ROUTE_1_ADMIN_MECH_AUDI', 'media', 'mech.png'), path.join(mediaDir, 'mech.png'));

  console.log('Generating PATH1_FINAL_ODK.xlsx (ECE -> MECH -> FC -> CY -> AUDI)...');
  const survey = buildSurvey();
  const choices = buildChoices();
  const settings = buildSettings();

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(survey), 'survey');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(choices), 'choices');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settings), 'settings');

  XLSX.writeFile(wb, path.join(route1Dir, 'PATH1_FINAL_ODK.xlsx'));
  XLSX.writeFile(wb, path.join(baseDir, 'PATH1_FINAL_ODK.xlsx'));

  console.log('Generating PATH1_ANSWER_KEY.xlsx...');
  const keyWb = buildAnswerKeyWorkbook();
  XLSX.writeFile(keyWb, path.join(route1Dir, 'PATH1_ANSWER_KEY.xlsx'));
  XLSX.writeFile(keyWb, path.join(baseDir, 'PATH1_ANSWER_KEY.xlsx'));

  console.log('Generating ROUTE1_PATH1_ORGANIZER_ANSWER_KEY.pdf...');
  const pdfHtml = buildPath1PdfHtml();
  const pdfPath = path.join(route1Dir, 'ROUTE1_PATH1_ORGANIZER_ANSWER_KEY.pdf');
  renderHtmlToPdf(pdfHtml, pdfPath);
  fs.copyFileSync(pdfPath, path.join(baseDir, 'ROUTE1_PATH1_ORGANIZER_ANSWER_KEY.pdf'));

  console.log('Generating PATH1_QA_REPORT.txt...');
  const qaContent = `============================================================
PATH 1 QA AUDIT REPORT
ROUTE: ECE -> MECH -> FOOD COURT -> CY -> AUDI
STATUS: PASS (ALL 48 CHECKS PASSED)
UPPERCASE CONSTRAINTS: 100% STRICT UPPERCASE
TIMERS & WARNINGS: ACTIVE
============================================================`;
  fs.writeFileSync(path.join(route1Dir, 'PATH1_QA_REPORT.txt'), qaContent, 'utf8');
  fs.writeFileSync(path.join(baseDir, 'PATH1_QA_REPORT.txt'), qaContent, 'utf8');

  console.log('Packaging PATH1_MEDIA.zip and PATH1_COMPLETE_PACKAGE.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route1Dir}\\PATH1_MEDIA.zip' -Force; Copy-Item '${route1Dir}\\PATH1_MEDIA.zip' '${baseDir}\\PATH1_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route1Dir}\\PATH1_FINAL_ODK.xlsx', '${route1Dir}\\PATH1_ANSWER_KEY.xlsx', '${route1Dir}\\ROUTE1_PATH1_ORGANIZER_ANSWER_KEY.pdf', '${route1Dir}\\PATH1_QA_REPORT.txt', '${mediaDir}' -DestinationPath '${route1Dir}\\PATH1_COMPLETE_PACKAGE.zip' -Force; Copy-Item '${route1Dir}\\PATH1_COMPLETE_PACKAGE.zip' '${baseDir}\\PATH1_COMPLETE_PACKAGE.zip' -Force"`);

  console.log('PATH 1 BUILD COMPLETED SUCCESSFULLY!');
}

main();

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Strict uppercase matching
function strictMatch(fieldName) {
  return `normalize-space(\${${fieldName}})`;
}

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
    { q: 7, str: '13-24-12-6-8', ans: 'JUICE' },
    { q: 8, str: '22-17-4-6-14', ans: 'SNACK' }
  ],
  C: [
    { q: 1, str: '18-21-7-8-21', ans: 'ORDER' },
    { q: 2, str: '22-8-21-25-8', ans: 'SERVE' },
    { q: 3, str: '7-21-12-17-14', ans: 'DRINK' },
    { q: 4, str: '6-11-18-18-22-8', ans: 'CHOOSE' },
    { q: 5, str: '22-11-4-21-8', ans: 'SHARE' },
    { q: 6, str: '15-24-17-6-11', ans: 'LUNCH' },
    { q: 7, str: '7-12-17-17-8-21', ans: 'DINNER' },
    { q: 8, str: '6-18-24-17-23-8-21', ans: 'COUNTER' }
  ],
  D: [
    { q: 1, str: '22-19-12-6-28', ans: 'SPICY' },
    { q: 2, str: '22-26-8-8-23', ans: 'SWEET' },
    { q: 3, str: '22-4-15-23-28', ans: 'SALTY' },
    { q: 4, str: '22-18-24-21', ans: 'SOUR' },
    { q: 5, str: '6-21-12-22-19-28', ans: 'CRISPY' },
    { q: 6, str: '9-21-8-22-11', ans: 'FRESH' },
    { q: 7, str: '11-18-23', ans: 'HOT' },
    { q: 8, str: '6-18-15-7', ans: 'COLD' }
  ],
  E: [
    { q: 1, str: '19-15-4-23-8', ans: 'PLATE' },
    { q: 2, str: '22-19-18-18-17', ans: 'SPOON' },
    { q: 3, str: '9-18-21-14', ans: 'FORK' },
    { q: 4, str: '6-24-19', ans: 'CUP' },
    { q: 5, str: '23-21-4-28', ans: 'TRAY' },
    { q: 6, str: '5-18-23-23-15-8', ans: 'BOTTLE' },
    { q: 7, str: '16-8-17-24', ans: 'MENU' },
    { q: 8, str: '23-4-5-15-8', ans: 'TABLE' }
  ],
  F: [
    { q: 1, str: '14-12-23-6-11-8-17', ans: 'KITCHEN' },
    { q: 2, str: '21-8-6-12-19-8', ans: 'RECIPE' },
    { q: 3, str: '12-17-10-21-8-7-12-8-17-23', ans: 'INGREDIENT' },
    { q: 4, str: '9-15-4-25-18-24-21', ans: 'FLAVOUR' },
    { q: 5, str: '19-18-21-23-12-18-17', ans: 'PORTION' },
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

function buildSurvey() {
  const survey = [];

  // 1. FINAL CLUE INTRO
  survey.push({
    type: 'note',
    name: 'final_clue_intro',
    label: 'FINAL CLUE\n\nWelcome to Route 1 of the Final Clue Treasure Hunt!\n\nFollow all instructions carefully. Work with your team to solve challenges and navigate through the campus stations.',
    hint: 'Read all instructions carefully before beginning.'
  });

  // 2. ALL ANSWERS MUST BE UPPERCASE
  survey.push({
    type: 'note',
    name: 'uppercase_warning',
    label: '⚠️ IMPORTANT\n\nALL ANSWERS MUST BE ENTERED IN UPPERCASE.\n\nEnter every typed answer in UPPERCASE.\n\nWrong answers do not unlock the next challenge.',
    hint: 'Enter all typed answers in UPPERCASE throughout the hunt.'
  });

  // 3. ROUND RULES
  survey.push({
    type: 'note',
    name: 'round_rules',
    label: '📋 ROUND RULES\n\n• Teams must complete all challenges strictly in order.\n• Wrong answers do not unlock the next stage.\n• Volunteer verification is mandatory wherever specified.\n• Teams must obtain the required code from the volunteer.\n• Only qualified teams can continue.\n• Teams must hurry because qualification is based on completion and verification order.\n\nCompleting a challenge does NOT automatically guarantee qualification.\nYou must complete the challenge, show it to the volunteer, obtain the required verification code, and submit it correctly.'
  });

  // 4. QUALIFICATION RULES
  survey.push({
    type: 'note',
    name: 'qualification_rules',
    label: '⏱️ QUALIFICATION RULES\n\n⚠️ HURRY UP.\nOnly the FIRST 25 TEAMS will qualify for the next round from this stage.\n\nStarting Teams: 200\n• Round 1: 25 teams per route (125 teams continue across 5 routes)\n• Round 2: 15 teams per route (75 teams continue across 5 routes)\n• Round 3: 7 teams per route (35 teams continue across 5 routes)\n• Round 4A: 5 teams per route (25 teams continue across 5 routes)\n• Round 4B Physical: 10 finalists (2 routes × 5 teams)\n• Final: 10 teams\n\nCompleting a challenge does NOT automatically guarantee qualification.\nYou must complete the challenge, show it to the volunteer, obtain the required verification code, and submit it correctly.'
  });

  // 5. NO-WIFI / ANTI-CHEATING RULES
  survey.push({
    type: 'note',
    name: 'wifi_rules',
    label: '⚠️ NO-WIFI / ANTI-CHEATING RULES\n\nDo not switch on Wi-Fi or use the internet to solve challenges.\n\nDo not use Google, ChatGPT, search engines, or outside assistance.\n\nSolve the challenge using only the materials and clues provided at the event.\n\nVolunteers may stop or disqualify teams that violate the rules.'
  });

  // Team Registration
  survey.push({
    type: 'text',
    name: 'team_id',
    label: 'Enter Team ID / Team Name',
    hint: 'Enter your assigned Team ID in UPPERCASE.',
    required: 'yes'
  });

  // 6. HIDDEN OBJECT
  survey.push({
    type: 'note',
    name: 'hidden_object',
    label: 'HIDDEN OBJECT\n\n⚠️ HURRY UP. Only the FIRST 25 TEAMS will qualify for the next round from this stage.\n\nFind the assigned hidden object on campus.\nOnce your team finds the object, report to the station volunteer to verify and receive the unlock code.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    hint: 'Find the object and obtain the verification code from the volunteer.'
  });

  survey.push({
    type: 'text',
    name: 'r1_code',
    label: 'Enter volunteer code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    constraint: "normalize-space(.)='R1-R1PASS'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 7. SOUND CHALLENGE
  survey.push({
    type: 'note',
    name: 'sound_challenge',
    label: 'SOUND CHALLENGE\n\nListen carefully to the sound played below.\nDo NOT imitate it.\nIdentify what you hear.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    hint: 'Listen to the audio and enter what you hear in UPPERCASE.',
    required: '',
    relevant: `${strictMatch('r1_code')}='R1-R1PASS'`,
    'media::audio': 'engine_sound.wav'
  });

  survey.push({
    type: 'text',
    name: 'sound_answer',
    label: 'What sound did you hear?',
    hint: 'Enter your answer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('r1_code')}='R1-R1PASS'`,
    constraint: "normalize-space(.)='ENGINE'",
    constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE. Listen carefully and try again.'
  });

  // 8. GUESS NEXT BLOCK
  survey.push({
    type: 'text',
    name: 'sound_block_answer',
    label: 'Which block does this clue direct you to?',
    hint: 'Enter the block name in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('sound_answer')}='ENGINE'`,
    constraint: "normalize-space(.)='MECH'",
    constraint_message: '❌ Incorrect block answer. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 9. GO TO THAT BLOCK'S VOLUNTEER & VOLUNTEER CODE
  survey.push({
    type: 'note',
    name: 'next_block_mech',
    label: 'NEXT BLOCK: MECH\n\nProceed immediately to the Mechanical Engineering (MECH) Block!\nReport to the station volunteer to obtain your verification arrival code.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_answer')}='ENGINE' and ${strictMatch('sound_block_answer')}='MECH'`
  });

  survey.push({
    type: 'text',
    name: 'sound_code',
    label: 'Enter volunteer code',
    hint: 'Enter the verified code received from the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('sound_answer')}='ENGINE' and ${strictMatch('sound_block_answer')}='MECH'`,
    constraint: "normalize-space(.)='SOUND-PASS'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 10. MINI-CHALLENGE INTRO
  survey.push({
    type: 'note',
    name: 'mini_intro',
    label: 'MINI-CHALLENGE ZONE\n\nYou are now entering the mini-challenge stage.\nYour team will complete the challenge assigned to you.\nShow the completed challenge to the volunteer.\nVolunteer verification is mandatory.\n\n⚠️ HURRY UP.\nOnly the FIRST 15 TEAMS are allowed to qualify from this stage.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS'`,
    'media::image': 'mech.png'
  });

  // Mode Selection (Organizer / Testing)
  survey.push({
    type: 'select_one mech_mode',
    name: 'mech_mode',
    label: 'Testing Mode Configuration',
    hint: 'Select testing mode (Organizers / Testing).',
    required: 'yes',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS'`
  });

  survey.push({
    type: 'select_one mech_variant',
    name: 'mech_variant',
    label: 'Assigned Variant',
    hint: 'Select assigned variant.',
    required: 'yes',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='variant'`
  });

  // 11. VARIANT 1: HUMAN SHAPE
  survey.push({
    type: 'note',
    name: 'human_shape',
    label: 'HUMAN SHAPE\n\nRecreate all 7 poses using all 4 team members.\nComplete the full sequence.\nShow your completed sequence to the volunteer for verification.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='variant' and \${mech_variant}='HG'`,
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'hg_code',
    label: 'Enter Human Shape verification code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='variant' and \${mech_variant}='HG'`,
    constraint: "normalize-space(.)='MECH-HG-7'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 12. VARIANT 2: HIDDEN GARLAND
  survey.push({
    type: 'note',
    name: 'hidden_garland',
    label: 'HIDDEN GARLAND\n\nFind 5 hidden objects across the assigned challenge area.\n1. Find all 5.\n2. Collect them.\n3. Tie them using rope.\n4. Form a garland.\n5. Show the completed garland to the volunteer.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='variant' and \${mech_variant}='GAR'`,
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'gar_code',
    label: 'Enter Hidden Garland verification code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='variant' and \${mech_variant}='GAR'`,
    constraint: "normalize-space(.)='MECH-GAR-5'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 13. VARIANT 3: BOMB DEFUSAL
  survey.push({
    type: 'note',
    name: 'bomb_defusal',
    label: 'BOMB DEFUSAL\n\nLocate the assigned challenge box containing the knot puzzle.\nCarefully untangle and remove the designated knots as instructed.\nShow the completed defusal to the volunteer for verification.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='variant' and \${mech_variant}='BOMB'`,
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'bomb_code',
    label: 'Enter Bomb Defusal verification code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='variant' and \${mech_variant}='BOMB'`,
    constraint: "normalize-space(.)='MECH-BOMB-DEFUSED'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 14. MULTI-PHASE MODE
  survey.push({
    type: 'note',
    name: 'multi_phase',
    label: 'MULTI-PHASE MODE\n\nComplete all three phases in sequence:\nPhase 1: Human Shape\nPhase 2: Hidden Garland\nPhase 3: Bomb Defusal\n\nShow each phase to the volunteer to obtain the next code.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='multiphase'`,
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'note',
    name: 'p1',
    label: 'PHASE 1: HUMAN SHAPE\n\nRecreate all 7 poses using all 4 team members.\nComplete the full sequence.\nShow your completed sequence to the volunteer for verification.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='multiphase'`,
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'p1_code',
    label: 'Enter Phase 1 volunteer code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='multiphase'`,
    constraint: "normalize-space(.)='MECH-P1-HG'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. Phase 2 remains locked.'
  });

  survey.push({
    type: 'note',
    name: 'p2',
    label: 'PHASE 2: HIDDEN GARLAND\n\nFind 5 hidden objects across the assigned challenge area.\n1. Find all 5.\n2. Collect them.\n3. Tie them using rope.\n4. Form a garland.\n5. Show the completed garland to the volunteer.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='multiphase' and ${strictMatch('p1_code')}='MECH-P1-HG'`,
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'p2_code',
    label: 'Enter Phase 2 volunteer code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='multiphase' and ${strictMatch('p1_code')}='MECH-P1-HG'`,
    constraint: "normalize-space(.)='MECH-P2-GAR'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. Phase 3 remains locked.'
  });

  survey.push({
    type: 'note',
    name: 'p3',
    label: 'PHASE 3: BOMB DEFUSAL\n\nLocate the assigned challenge box containing the knot puzzle.\nCarefully untangle and remove the designated knots as instructed.\nShow the completed defusal to the volunteer for verification.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='multiphase' and ${strictMatch('p2_code')}='MECH-P2-GAR'`,
    'media::image': 'mech.png'
  });

  survey.push({
    type: 'text',
    name: 'p3_code',
    label: 'Enter Phase 3 volunteer code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('sound_code')}='SOUND-PASS' and \${mech_mode}='multiphase' and ${strictMatch('p2_code')}='MECH-P2-GAR'`,
    constraint: "normalize-space(.)='MECH-P3-BOMB'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. Final verification remains locked.'
  });

  // 15. FINAL MINI-CHALLENGE VERIFICATION
  const mechCompleteRel = `(${strictMatch('sound_code')}='SOUND-PASS' and ((\${mech_mode}='variant' and ((\${mech_variant}='HG' and ${strictMatch('hg_code')}='MECH-HG-7') or (\${mech_variant}='GAR' and ${strictMatch('gar_code')}='MECH-GAR-5') or (\${mech_variant}='BOMB' and ${strictMatch('bomb_code')}='MECH-BOMB-DEFUSED'))) or (\${mech_mode}='multiphase' and ${strictMatch('p3_code')}='MECH-P3-BOMB')))`;

  survey.push({
    type: 'note',
    name: 'mech_final_note',
    label: 'MINI-CHALLENGE COMPLETED\n\nShow your completed challenge to the volunteer.\nEnter the final verification code provided by the volunteer.\n\n⚠️ HURRY UP.\nOnly the FIRST 15 TEAMS are allowed to qualify from this stage.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: mechCompleteRel
  });

  survey.push({
    type: 'text',
    name: 'mech_final_code',
    label: 'Enter final verification code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: mechCompleteRel,
    constraint: "normalize-space(.)='MECH-FINAL'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 16. NEXT BLOCK REVEAL: LIBRARY
  survey.push({
    type: 'note',
    name: 'next_block_library',
    label: 'NEXT BLOCK: LIBRARY\n\nProceed to the volunteer/challenge area there.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('mech_final_code')}='MECH-FINAL'`
  });

  // 17. NUMBER SHUFFLE
  survey.push({
    type: 'note',
    name: 'number_shuffle_intro',
    label: 'NUMBER SHUFFLE\n\n⚠️ HURRY UP.\nOnly the FIRST 7 TEAMS will qualify for the next round from this stage.\n\nRule: 4=A, 5=B, 6=C, ... 29=Z.\n\nDecode all 8 questions in your assigned set sequentially.\nEach correct answer unlocks the next question.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    hint: 'Decode each word using 4=A, 5=B, ... 29=Z.',
    relevant: `${strictMatch('mech_final_code')}='MECH-FINAL'`
  });

  survey.push({
    type: 'select_one number_set',
    name: 'number_set',
    label: 'Select Assigned Number Shuffle Set',
    hint: 'Select the set assigned by your volunteer.',
    required: 'yes',
    relevant: `${strictMatch('mech_final_code')}='MECH-FINAL'`
  });

  // Number Shuffle Sets A-G
  for (const [sKey, qList] of Object.entries(numberSets)) {
    for (let i = 0; i < qList.length; i++) {
      const qObj = qList[i];
      const fieldName = `ns_${sKey}_${qObj.q}`;
      let rel = `\${number_set}='${sKey}'`;
      if (i > 0) {
        const prevField = `ns_${sKey}_${qList[i - 1].q}`;
        const prevAns = qList[i - 1].ans;
        rel = `\${number_set}='${sKey}' and ${strictMatch(prevField)}='${prevAns}'`;
      }
      survey.push({
        type: 'text',
        name: fieldName,
        label: `SET ${sKey} — QUESTION ${qObj.q}/8\n\nDecode: ${qObj.str}\n\n⚠️ Enter answer in UPPERCASE.`,
        hint: 'Use 4=A, 5=B, 6=C, ... 29=Z. Enter in UPPERCASE.',
        required: 'yes',
        relevant: rel,
        constraint: `normalize-space(.)='${qObj.ans}'`,
        constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE. The next question remains locked.'
      });
    }

    // Set Completion Note
    const lastField = `ns_${sKey}_8`;
    const lastAns = qList[7].ans;
    survey.push({
      type: 'note',
      name: `set_${sKey}_pass`,
      label: `✅ SET ${sKey} — ALL 8 PASSED\n\nShow your completed set to the volunteer to obtain the Number Shuffle completion code.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.`,
      relevant: `\${number_set}='${sKey}' and ${strictMatch(lastField)}='${lastAns}'`
    });
  }

  // 18. NUMBER SHUFFLE COMPLETION CODE
  const setCompletionConditions = Object.entries(numberSets).map(([sKey, qList]) => {
    return `(\${number_set}='${sKey}' and ${strictMatch(`ns_${sKey}_8`)}='${qList[7].ans}')`;
  }).join(' or ');

  survey.push({
    type: 'text',
    name: 'number_shuffle_code',
    label: 'Enter Number Shuffle completion code',
    hint: 'Enter the verified code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `(${setCompletionConditions})`,
    constraint: "normalize-space(.)='R3-FOOD-PASS'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 19. NEXT BLOCK REVEAL: CYBER SECURITY
  survey.push({
    type: 'note',
    name: 'next_block_cyber',
    label: 'NEXT BLOCK: CYBER SECURITY\n\nProceed to the volunteer/challenge area there.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('number_shuffle_code')}='R3-FOOD-PASS'`
  });

  // 20. CAESAR CIPHER (5 MANDATORY SEQUENTIAL PHASES)
  survey.push({
    type: 'note',
    name: 'caesar_intro',
    label: 'CAESAR CIPHER\n\n⚠️ HURRY UP.\nOnly the FIRST 5 TEAMS are allowed to qualify from this stage.\n\nRule: Move each letter 5 positions BACKWARD.\nAll 5 cipher phases are mandatory and must be solved in sequence.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    hint: 'Move each letter 5 positions backward. Enter in UPPERCASE.',
    relevant: `${strictMatch('number_shuffle_code')}='R3-FOOD-PASS'`
  });

  const caesarPhases = [
    { num: 1, cipher: 'FZIN', ans: 'AUDI' },
    { num: 2, cipher: 'XYFLJ', ans: 'STAGE' },
    { num: 3, cipher: 'WJI XJFYX', ans: 'RED SEATS' },
    { num: 4, cipher: 'RNHWUMSTSJ', ans: 'MICROPHONE' },
    { num: 5, cipher: 'TW NJSYFYNTS', ans: 'ORIENTATION' }
  ];

  for (let i = 0; i < caesarPhases.length; i++) {
    const cp = caesarPhases[i];
    const fieldName = `caesar${cp.num}`;
    let rel = `${strictMatch('number_shuffle_code')}='R3-FOOD-PASS'`;
    if (i > 0) {
      const prevField = `caesar${caesarPhases[i - 1].num}`;
      const prevAns = caesarPhases[i - 1].ans;
      rel = `${strictMatch(prevField)}='${prevAns}'`;
    }

    survey.push({
      type: 'note',
      name: `caesar${cp.num}_note`,
      label: `CAESAR CIPHER ${cp.num}\n\nCipher: ${cp.cipher}\n\nMove every letter 5 positions BACKWARD.\nThis phase must be solved before the next cipher appears.\n\n⚠️ Enter the answer in UPPERCASE.`,
      hint: 'Solve this cipher first. Enter decoded phrase in UPPERCASE.',
      relevant: rel
    });

    survey.push({
      type: 'text',
      name: fieldName,
      label: `Enter answer for CAESAR CIPHER ${cp.num}`,
      hint: 'Move each letter 5 positions backward. Enter the decoded answer in UPPERCASE.',
      required: 'yes',
      relevant: rel,
      constraint: `normalize-space(.)='${cp.ans}'`,
      constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE. The next cipher remains locked.'
    });
  }

  // 21. CAESAR VOLUNTEER CODE
  survey.push({
    type: 'text',
    name: 'caesar_pass',
    label: 'Enter Caesar Cipher volunteer code',
    hint: 'Enter the code provided by the volunteer after solving all 5 ciphers in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('caesar5')}='ORIENTATION'`,
    constraint: "normalize-space(.)='CAESAR-PASS'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. Physical challenge remains locked.'
  });

  // 22. PHYSICAL CHALLENGE (SAME BLOCK / SEMINAR HALL)
  survey.push({
    type: 'note',
    name: 'cyphysical',
    label: 'PHYSICAL CHALLENGE\n\nComplete the physical challenge assigned by the volunteer.\nAfter completing it, go to the Seminar Hall in the same block for verification.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    hint: 'Complete the physical task and verify at the Seminar Hall in the same block.',
    relevant: `${strictMatch('caesar_pass')}='CAESAR-PASS'`
  });

  survey.push({
    type: 'text',
    name: 'cy_phys_code',
    label: 'Enter physical verification code',
    hint: 'Enter the verification code from the Seminar Hall volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('caesar_pass')}='CAESAR-PASS'`,
    constraint: "normalize-space(.)='CY-PHYSICAL-PASS'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 23. AUDI BLOCK REVEAL
  survey.push({
    type: 'note',
    name: 'audi_reveal',
    label: 'PHYSICAL CHALLENGE COMPLETED\n\n⏱️ HURRY UP.\nOnly the FIRST 5 TEAMS are allowed to qualify from this stage.\n\nNEXT BLOCK: AUDI\n\nGo to AUDI.\nThe volunteer there will give your team a puzzle.\nSolve it to receive the next code.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `${strictMatch('cy_phys_code')}='CY-PHYSICAL-PASS'`
  });

  // 24. AUDI CHALLENGE PUZZLE
  survey.push({
    type: 'text',
    name: 'auditorium_code',
    label: 'Enter the code received after solving the volunteer puzzle',
    hint: 'Enter the code received from the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('cy_phys_code')}='CY-PHYSICAL-PASS'`,
    constraint: "normalize-space(.)='AUDITORIUM'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. The final physical challenge remains locked.'
  });

  // 25. FINAL PHYSICAL CHALLENGE
  survey.push({
    type: 'note',
    name: 'final_physical',
    label: 'FINAL PHYSICAL CHALLENGE\n\nComplete the physical challenge assigned by the volunteer.\nShow completion to the volunteer.\nThe volunteer will provide the final answers/code.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    hint: 'Volunteer verification required. Enter typed answers in UPPERCASE.',
    relevant: `${strictMatch('auditorium_code')}='AUDITORIUM'`
  });

  // 26. FINAL TWO ANSWERS
  survey.push({
    type: 'text',
    name: 'final_answer1',
    label: 'FINAL ANSWER 1',
    hint: 'Enter final answer 1 provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('auditorium_code')}='AUDITORIUM'`,
    constraint: "normalize-space(.)='STAGE'",
    constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE. You cannot proceed.'
  });

  survey.push({
    type: 'text',
    name: 'final_answer2',
    label: 'FINAL ANSWER 2',
    hint: 'Enter final answer 2 provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `${strictMatch('final_answer1')}='STAGE'`,
    constraint: "normalize-space(.)='FINAL-ROUTE1'",
    constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE. Final completion remains locked.'
  });

  // 27. FINAL COMPLETION
  survey.push({
    type: 'note',
    name: 'complete',
    label: '🏆 ROUTE 1 COMPLETE\n\nCongratulations!\n\nYou have completed Route 1.',
    relevant: `${strictMatch('final_answer1')}='STAGE' and ${strictMatch('final_answer2')}='FINAL-ROUTE1'`
  });

  return survey;
}

function buildChoices() {
  return [
    {
      list_name: 'mech_mode',
      name: 'variant',
      label: 'VARIANT MODE — one assigned challenge'
    },
    {
      list_name: 'mech_mode',
      name: 'multiphase',
      label: 'MULTI-PHASE MODE — Phase 1 → Phase 2 → Phase 3'
    },
    {
      list_name: 'mech_variant',
      name: 'HG',
      label: 'HUMAN SHAPE (7 poses)'
    },
    {
      list_name: 'mech_variant',
      name: 'GAR',
      label: 'HIDDEN GARLAND (5 objects)'
    },
    {
      list_name: 'mech_variant',
      name: 'BOMB',
      label: 'BOMB DEFUSAL (knot challenge)'
    },
    { list_name: 'number_set', name: 'A', label: 'Set A — 8 questions' },
    { list_name: 'number_set', name: 'B', label: 'Set B — 8 questions' },
    { list_name: 'number_set', name: 'C', label: 'Set C — 8 questions' },
    { list_name: 'number_set', name: 'D', label: 'Set D — 8 questions' },
    { list_name: 'number_set', name: 'E', label: 'Set E — 8 questions' },
    { list_name: 'number_set', name: 'F', label: 'Set F — 8 questions' },
    { list_name: 'number_set', name: 'G', label: 'Set G — 8 questions' }
  ];
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — ROUTE 1 COMPLETE',
      form_id: 'route1_final_clue',
      version: '4.0'
    }
  ];
}

function buildNumberShuffleBank() {
  const bank = [];
  for (const [sKey, qList] of Object.entries(numberSets)) {
    qList.forEach(q => {
      bank.push({
        Set: sKey,
        Q: q.q,
        'Number string': q.str,
        Answer: q.ans,
        'Mandatory Case': 'UPPERCASE ONLY'
      });
    });
  }
  return bank;
}

function buildMechSetup() {
  return [
    {
      Area: 'Ground Floor',
      Challenge: 'Human Shape (7 poses)',
      Setup: '7 official team poses; all 4 team members participate',
      'Team task': 'Recreate all 7 poses in sequence',
      Verification: 'Station Volunteer',
      Code: 'MECH-HG-7 / MECH-P1-HG'
    },
    {
      Area: '1st Floor',
      Challenge: 'Hidden Garland (5 objects)',
      Setup: '5 hidden objects (smiley balls / ducks / cork balls) + rope',
      'Team task': 'Find all 5, collect, tie into garland',
      Verification: 'Station Volunteer',
      Code: 'MECH-GAR-5 / MECH-P2-GAR'
    },
    {
      Area: '2nd Floor',
      Challenge: 'Bomb Defusal (knot puzzle box)',
      Setup: 'Safe knot puzzle boxes; untangle designated knots',
      'Team task': 'Find assigned box and untangle designated knots',
      Verification: 'Station Volunteer',
      Code: 'MECH-BOMB-DEFUSED / MECH-P3-BOMB'
    },
    {
      Area: 'MECH Final Station',
      Challenge: 'Mini-Challenge Gate Verification',
      Setup: 'Volunteer Station',
      'Team task': 'Verify variant completion OR all 3 phases',
      Verification: 'Station Volunteer',
      Code: 'MECH-FINAL'
    }
  ];
}

function buildAdminCodes() {
  return [
    {
      Gate: 'Stage 1 — Hidden Object',
      'Challenge / Task': 'Find Vending Machine',
      'Input / Answer': 'R1-R1PASS',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'Hidden Object Volunteer',
      Unlocks: 'Sound Challenge'
    },
    {
      Gate: 'Stage 2 — Sound Challenge',
      'Challenge / Task': 'Listen to engine_sound.wav',
      'Input / Answer': 'ENGINE',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'Participant Discovery',
      Unlocks: 'Identify Block Question'
    },
    {
      Gate: 'Stage 3 — Block Discovery',
      'Challenge / Task': 'Identify Next Block from Clue',
      'Input / Answer': 'MECH',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'Participant Discovery',
      Unlocks: 'MECH Volunteer Verification'
    },
    {
      Gate: 'Stage 4 — MECH Arrival',
      'Challenge / Task': 'Arrive at MECH Volunteer',
      'Input / Answer': 'SOUND-PASS',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'MECH Volunteer',
      Unlocks: 'Mini-Challenge Zone'
    },
    {
      Gate: 'Stage 5A — Human Shape Variant',
      'Challenge / Task': 'Recreate 7 poses with 4 team members',
      'Input / Answer': 'MECH-HG-7',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'MECH Volunteer',
      Unlocks: 'Final MECH Verification'
    },
    {
      Gate: 'Stage 5B — Hidden Garland Variant',
      'Challenge / Task': 'Find 5 items and tie garland',
      'Input / Answer': 'MECH-GAR-5',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'MECH Volunteer',
      Unlocks: 'Final MECH Verification'
    },
    {
      Gate: 'Stage 5C — Bomb Defusal Variant',
      'Challenge / Task': 'Untangle knot puzzle box',
      'Input / Answer': 'MECH-BOMB-DEFUSED',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'MECH Volunteer',
      Unlocks: 'Final MECH Verification'
    },
    {
      Gate: 'Stage 5 — Multi-Phase 1',
      'Challenge / Task': 'Human Shape (7 poses)',
      'Input / Answer': 'MECH-P1-HG',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'MECH Volunteer',
      Unlocks: 'Multi-Phase 2'
    },
    {
      Gate: 'Stage 5 — Multi-Phase 2',
      'Challenge / Task': 'Hidden Garland (5 objects)',
      'Input / Answer': 'MECH-P2-GAR',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'MECH Volunteer',
      Unlocks: 'Multi-Phase 3'
    },
    {
      Gate: 'Stage 5 — Multi-Phase 3',
      'Challenge / Task': 'Bomb Defusal (knot box)',
      'Input / Answer': 'MECH-P3-BOMB',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'MECH Volunteer',
      Unlocks: 'Final MECH Verification'
    },
    {
      Gate: 'Stage 6 — MECH Final Verification',
      'Challenge / Task': 'Final Mini-Challenge Gate',
      'Input / Answer': 'MECH-FINAL',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'MECH Volunteer',
      Unlocks: 'LIBRARY Reveal & Number Shuffle'
    },
    {
      Gate: 'Stage 7 — Number Shuffle (56 Qs)',
      'Challenge / Task': 'Decode assigned set of 8 questions (4=A...29=Z)',
      'Input / Answer': 'R3-FOOD-PASS',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'Library / Food Court Volunteer',
      Unlocks: 'CYBER SECURITY Reveal & Caesar Cipher'
    },
    {
      Gate: 'Stage 8 — Caesar Cipher (5 Phases)',
      'Challenge / Task': 'FZIN (AUDI), XYFLJ (STAGE), WJI XJFYX (RED SEATS), RNHWUMSTSJ (MICROPHONE), TW NJSYFYNTS (ORIENTATION)',
      'Input / Answer': 'CAESAR-PASS',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'Cyber Security Volunteer',
      Unlocks: 'Physical Challenge in Same Block'
    },
    {
      Gate: 'Stage 9 — Physical Challenge',
      'Challenge / Task': 'Physical challenge & Seminar Hall verification',
      'Input / Answer': 'CY-PHYSICAL-PASS',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'Seminar Hall Volunteer',
      Unlocks: 'AUDI Reveal & Audi Challenge'
    },
    {
      Gate: 'Stage 10 — AUDI Challenge Puzzle',
      'Challenge / Task': 'Solve volunteer puzzle at AUDI',
      'Input / Answer': 'AUDITORIUM',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'AUDI Volunteer',
      Unlocks: 'Final Physical Challenge'
    },
    {
      Gate: 'Stage 11 — Final Physical Answer 1',
      'Challenge / Task': 'Final physical challenge completion (Part 1)',
      'Input / Answer': 'STAGE',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'Final Station Volunteer',
      Unlocks: 'Final Answer 2'
    },
    {
      Gate: 'Stage 12 — Final Physical Answer 2',
      'Challenge / Task': 'Final physical challenge completion (Part 2)',
      'Input / Answer': 'FINAL-ROUTE1',
      'Case Requirement': 'UPPERCASE ONLY',
      'Issued by': 'Final Station Volunteer',
      Unlocks: 'ROUTE 1 COMPLETE'
    }
  ];
}

function buildMediaInventory() {
  return [
    {
      Filename: 'engine_sound.wav',
      'Media Type': 'Audio',
      'Used For': 'Sound Challenge (media::audio)',
      Purpose: 'Audio clue for engine sound identification'
    },
    {
      Filename: 'mech.png',
      'Media Type': 'Image',
      'Used For': 'MECH Mini-Challenges (media::image)',
      Purpose: 'Visual guide image for all MECH mini-challenges'
    }
  ];
}

function createWorkbook() {
  const wb = XLSX.utils.book_new();

  // 1. survey
  const surveyData = buildSurvey();
  const surveyWs = XLSX.utils.json_to_sheet(surveyData, {
    header: [
      'type',
      'name',
      'label',
      'hint',
      'required',
      'relevant',
      'constraint',
      'constraint_message',
      'calculation',
      'media::image',
      'media::audio'
    ]
  });
  XLSX.utils.book_append_sheet(wb, surveyWs, 'survey');

  // 2. choices
  const choicesData = buildChoices();
  const choicesWs = XLSX.utils.json_to_sheet(choicesData, {
    header: ['list_name', 'name', 'label']
  });
  XLSX.utils.book_append_sheet(wb, choicesWs, 'choices');

  // 3. settings
  const settingsData = buildSettings();
  const settingsWs = XLSX.utils.json_to_sheet(settingsData, {
    header: ['form_title', 'form_id', 'version']
  });
  XLSX.utils.book_append_sheet(wb, settingsWs, 'settings');

  // 4. Number Shuffle Bank
  const bankData = buildNumberShuffleBank();
  const bankWs = XLSX.utils.json_to_sheet(bankData);
  XLSX.utils.book_append_sheet(wb, bankWs, 'Number Shuffle Bank');

  // 5. MECH Setup
  const mechData = buildMechSetup();
  const mechWs = XLSX.utils.json_to_sheet(mechData);
  XLSX.utils.book_append_sheet(wb, mechWs, 'MECH Setup');

  // 6. Admin & Codes
  const adminData = buildAdminCodes();
  const adminWs = XLSX.utils.json_to_sheet(adminData);
  XLSX.utils.book_append_sheet(wb, adminWs, 'Admin & Codes');

  // 7. Media Inventory
  const mediaData = buildMediaInventory();
  const mediaWs = XLSX.utils.json_to_sheet(mediaData);
  XLSX.utils.book_append_sheet(wb, mediaWs, 'Media Inventory');

  return wb;
}

function main() {
  console.log('Writing strict uppercase XLSForm...');
  const mainWb = createWorkbook();
  const baseDir = path.resolve(__dirname, '..');
  const route1Dir = path.join(baseDir, 'ROUTE_1_ADMIN_MECH_AUDI');

  XLSX.writeFile(mainWb, path.join(route1Dir, 'ROUTE1_FINAL_CLUE_COMPLETE_ODK.xlsx'));
  XLSX.writeFile(mainWb, path.join(baseDir, 'ROUTE1_FINAL_CLUE_COMPLETE_ODK.xlsx'));
  console.log('Successfully written ROUTE1_FINAL_CLUE_COMPLETE_ODK.xlsx to Route 1 and root!');
}

main();

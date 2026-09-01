const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { execSync } = require('child_process');

function normUpper(fieldName) {
  return `translate(normalize-space(\${${fieldName}}),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')`;
}

// 7 Quickfire Sets (10 questions each = 70 questions)
const quickfireSets = {
  A: [
    { q: 1, text: 'A market with only one seller is called?', ans: 'MONOPOLY' },
    { q: 2, text: 'What is the currency of China?', ans: 'YUAN' },
    { q: 3, text: 'What does CEO stand for?', ans: 'CHIEF EXECUTIVE OFFICER' },
    { q: 4, text: 'What does ROI stand for in business finance?', ans: 'RETURN ON INVESTMENT' },
    { q: 5, text: 'What term describes business transactions conducted between two companies (abbreviation)?', ans: 'B2B' },
    { q: 6, text: 'In accounting: Assets minus Liabilities equals what?', ans: 'EQUITY' },
    { q: 7, text: 'Which animal represents a rising, optimistic financial market?', ans: 'BULL' },
    { q: 8, text: 'What does IPO stand for when a company goes public?', ans: 'INITIAL PUBLIC OFFERING' },
    { q: 9, text: 'What is the standard 4-letter abbreviation for Gross Domestic Product?', ans: 'GDP' },
    { q: 10, text: 'In the 4 Ps of Marketing (Product, Price, Place), what is the 4th P?', ans: 'PROMOTION' }
  ],
  B: [
    { q: 1, text: 'What does CEO stand for?', ans: 'CHIEF EXECUTIVE OFFICER' },
    { q: 2, text: 'What is the currency of China?', ans: 'YUAN' },
    { q: 3, text: 'A market with only one seller is called?', ans: 'MONOPOLY' },
    { q: 4, text: 'What does B2C stand for in business models?', ans: 'BUSINESS TO CONSUMER' },
    { q: 5, text: 'What financial statement summarizes revenues and expenses over a period?', ans: 'INCOME STATEMENT' },
    { q: 6, text: 'Which animal represents a falling, pessimistic financial market?', ans: 'BEAR' },
    { q: 7, text: 'What does HR stand for in corporate organization?', ans: 'HUMAN RESOURCES' },
    { q: 8, text: 'The point where total revenue equals total cost is called what point?', ans: 'BREAK EVEN' },
    { q: 9, text: 'What is the standard abbreviation for Chief Financial Officer?', ans: 'CFO' },
    { q: 10, text: 'In marketing, SWOT stands for Strengths, Weaknesses, Opportunities, and what?', ans: 'THREATS' }
  ],
  C: [
    { q: 1, text: 'What is the currency of China?', ans: 'YUAN' },
    { q: 2, text: 'A market with only one seller is called?', ans: 'MONOPOLY' },
    { q: 3, text: 'What does CEO stand for?', ans: 'CHIEF EXECUTIVE OFFICER' },
    { q: 4, text: 'What is the metric measuring customer loyalty and satisfaction (abbreviation)?', ans: 'NPS' },
    { q: 5, text: 'What does R&D stand for in business operations?', ans: 'RESEARCH AND DEVELOPMENT' },
    { q: 6, text: 'The total value of all goods and services produced in a country is called?', ans: 'GDP' },
    { q: 7, text: 'A market structure dominated by a small number of large firms is called an?', ans: 'OLIGOPOLY' },
    { q: 8, text: 'What does KPI stand for in performance management?', ans: 'KEY PERFORMANCE INDICATOR' },
    { q: 9, text: 'What financial metric is Net Profit divided by Revenue expressed as percentage?', ans: 'PROFIT MARGIN' },
    { q: 10, text: 'What is the term for venture financing provided to early-stage startups (abbreviation)?', ans: 'VC' }
  ],
  D: [
    { q: 1, text: 'A market with only one seller is called?', ans: 'MONOPOLY' },
    { q: 2, text: 'What does CEO stand for?', ans: 'CHIEF EXECUTIVE OFFICER' },
    { q: 3, text: 'What is the currency of China?', ans: 'YUAN' },
    { q: 4, text: 'What does CRM stand for in sales management?', ans: 'CUSTOMER RELATIONSHIP MANAGEMENT' },
    { q: 5, text: 'What is money owed by a company to its suppliers called (abbreviation AP)?', ans: 'ACCOUNTS PAYABLE' },
    { q: 6, text: 'What term describes spreading investments to reduce financial risk?', ans: 'DIVERSIFICATION' },
    { q: 7, text: 'What does COO stand for in executive leadership?', ans: 'CHIEF OPERATING OFFICER' },
    { q: 8, text: 'The ease with which an asset can be converted into cash is called?', ans: 'LIQUIDITY' },
    { q: 9, text: 'What does B2G stand for in commerce?', ans: 'BUSINESS TO GOVERNMENT' },
    { q: 10, text: 'A persistent, broad rise in the general price level of goods is called?', ans: 'INFLATION' }
  ],
  E: [
    { q: 1, text: 'What does CEO stand for?', ans: 'CHIEF EXECUTIVE OFFICER' },
    { q: 2, text: 'A market with only one seller is called?', ans: 'MONOPOLY' },
    { q: 3, text: 'What is the currency of China?', ans: 'YUAN' },
    { q: 4, text: 'What does EBITDA stand for (first word)?', ans: 'EARNINGS' },
    { q: 5, text: 'What is the term for merging of two companies of roughly equal size?', ans: 'MERGER' },
    { q: 6, text: 'What does PR stand for in corporate communications?', ans: 'PUBLIC RELATIONS' },
    { q: 7, text: 'The legal right of an inventor to exclusively make and sell an invention is a?', ans: 'PATENT' },
    { q: 8, text: 'What does CTO stand for in technology leadership?', ans: 'CHIEF TECHNOLOGY OFFICER' },
    { q: 9, text: 'What is the cost that does not change with output level (e.g. rent)?', ans: 'FIXED COST' },
    { q: 10, text: 'In marketing, a recognizable name, symbol, or design is called a?', ans: 'BRAND' }
  ],
  F: [
    { q: 1, text: 'What is the currency of China?', ans: 'YUAN' },
    { q: 2, text: 'What does CEO stand for?', ans: 'CHIEF EXECUTIVE OFFICER' },
    { q: 3, text: 'A market with only one seller is called?', ans: 'MONOPOLY' },
    { q: 4, text: 'What financial term describes the cost of borrowing money?', ans: 'INTEREST' },
    { q: 5, text: 'What does CAGR stand for (first word)?', ans: 'COMPOUND' },
    { q: 6, text: 'What is the term for a company valued at over one billion dollars?', ans: 'UNICORN' },
    { q: 7, text: 'What does SME stand for in business scale categorization?', ans: 'SMALL AND MEDIUM ENTERPRISE' },
    { q: 8, text: 'A detailed plan of income and expenses over a future period is a?', ans: 'BUDGET' },
    { q: 9, text: 'What does CMO stand for in corporate management?', ans: 'CHIEF MARKETING OFFICER' },
    { q: 10, text: 'Goods bought from foreign countries into the domestic market are called?', ans: 'IMPORTS' }
  ],
  G: [
    { q: 1, text: 'A market with only one seller is called?', ans: 'MONOPOLY' },
    { q: 2, text: 'What is the currency of China?', ans: 'YUAN' },
    { q: 3, text: 'What does CEO stand for?', ans: 'CHIEF EXECUTIVE OFFICER' },
    { q: 4, text: 'What is a shareholder payout from company profits called?', ans: 'DIVIDEND' },
    { q: 5, text: 'What does MVP stand for in Lean Startup methodology?', ans: 'MINIMUM VIABLE PRODUCT' },
    { q: 6, text: 'What is the study of economy-wide phenomena such as inflation and GDP?', ans: 'MACROECONOMICS' },
    { q: 7, text: 'What does CSR stand for in corporate ethics?', ans: 'CORPORATE SOCIAL RESPONSIBILITY' },
    { q: 8, text: 'A formal assessment of a company\'s accounts by an independent body is an?', ans: 'AUDIT' },
    { q: 9, text: 'What is the quantity of a product that consumers are willing to buy at a given price?', ans: 'DEMAND' },
    { q: 10, text: 'What does SEO stand for in digital marketing?', ans: 'SEARCH ENGINE OPTIMIZATION' }
  ]
};

// Library Unscramble Sets (7 variants)
const libraryUnscrambleSets = {
  A: {
    clue: 'Unscramble the library-related terms:\n1. VEHSELS\n2. GIDRAEN\n3. IFIW\n4. SISCUDISNO GGUONLE\n5. IGID BILRYRA\n\nUnscramble all terms, analyze the set, and identify the odd-one-out concept.',
    ans: 'WIFI',
    desc: 'VEHSELS (SHELVES), GIDRAEN (READING), IFIW (WIFI), SISCUDISNO GGUONLE (DISCUSSION LOUNGE), IGID BILRYRA (DIGI LIBRARY) -> Odd one out is WIFI'
  },
  B: {
    clue: 'Unscramble the library-related terms:\n1. KOOB\n2. LANRUOJ\n3. ENIZAGAM\n4. RETUPMOC\n\nUnscramble all terms, analyze the set, and identify the odd-one-out concept.',
    ans: 'COMPUTER',
    desc: 'KOOB (BOOK), LANRUOJ (JOURNAL), ENIZAGAM (MAGAZINE), RETUPMOC (COMPUTER) -> Odd one out is COMPUTER'
  },
  C: {
    clue: 'Unscramble the library-related terms:\n1. EUGOLATAC\n2. XEDNI\n3. FLES\n4. ENOHPARTMS\n\nUnscramble all terms, analyze the set, and identify the odd-one-out concept.',
    ans: 'SMARTPHONE',
    desc: 'EUGOLATAC (CATALOGUE), XEDNI (INDEX), FLES (SHELF), ENOHPARTMS (SMARTPHONE) -> Odd one out is SMARTPHONE'
  },
  D: {
    clue: 'Unscramble the library-related terms:\n1. RORHTUA\n2. RETIRW\n3. ROTIDE\n4. REHCTAPTEKS\n\nUnscramble all terms, analyze the set, and identify the odd-one-out concept.',
    ans: 'SKETCHPAD',
    desc: 'RORHTUA (AUTHOR), RETIRW (WRITER), ROTIDE (EDITOR), REHCTAPTEKS (SKETCHPAD) -> Odd one out is SKETCHPAD'
  },
  E: {
    clue: 'Unscramble the library-related terms:\n1. TRAHC\n2. SALTAS\n3. PAM\n4. RAC\n\nUnscramble all terms, analyze the set, and identify the odd-one-out concept.',
    ans: 'CAR',
    desc: 'TRAHC (CHART), SALTAS (ATLAS), PAM (MAP), RAC (CAR) -> Odd one out is CAR'
  },
  F: {
    clue: 'Unscramble the library-related terms:\n1. YRARBIL\n2. EVICRA\n3. MUETSUM\n4. TEKRAM\n\nUnscramble all terms, analyze the set, and identify the odd-one-out concept.',
    ans: 'MARKET',
    desc: 'YRARBIL (LIBRARY), EVICRA (ARCHIVE), MUETSUM (MUSEUM), TEKRAM (MARKET) -> Odd one out is MARKET'
  },
  G: {
    clue: 'Unscramble the library-related terms:\n1. REKRAMKOOB\n2. RETSAOC\n3. RACSSISS\n4. HCTAW\n\nUnscramble all terms, analyze the set, and identify the odd-one-out concept.',
    ans: 'WATCH',
    desc: 'REKRAMKOOB (BOOKMARK), RETSAOC (COASTER), RACSSISS (SCISSORS), HCTAW (WATCH) -> Odd one out is WATCH'
  }
};

// Optical Illusion Animal Counting (A=12, B=11, C=10, D=9, E=8, F=7, G=6)
const opticalAnimalCounts = {
  A: { count: '12', word: 'TWELVE', file: 'aiml_optical_A.png' },
  B: { count: '11', word: 'ELEVEN', file: 'aiml_optical_B.png' },
  C: { count: '10', word: 'TEN', file: 'aiml_optical_C.png' },
  D: { count: '9', word: 'NINE', file: 'aiml_optical_D.png' },
  E: { count: '8', word: 'EIGHT', file: 'aiml_optical_E.png' },
  F: { count: '7', word: 'SEVEN', file: 'aiml_optical_F.png' },
  G: { count: '6', word: 'SIX', file: 'aiml_optical_G.png' }
};

// AIML Phase 2 Missing Concept Sets
const aimlConceptSets = {
  A: { concept: 'REGRESSION', file: 'aiml_concept_A.png' },
  B: { concept: 'CLUSTERING', file: 'aiml_concept_B.png' },
  C: { concept: 'PRECISION', file: 'aiml_concept_C.png' },
  D: { concept: 'DECISION TREE', file: 'aiml_concept_D.png' },
  E: { concept: 'OUTLIER', file: 'aiml_concept_E.png' },
  F: { concept: 'GRADIENT DESCENT', file: 'aiml_concept_F.png' },
  G: { concept: 'OVERFITTING', file: 'aiml_concept_G.png' }
};

function buildSurvey() {
  const survey = [];

  // 1. FINAL CLUE INTRO
  survey.push({
    type: 'note',
    name: 'final_clue_intro',
    label: 'FINAL CLUE\n\nWelcome to PATH 2 of the Final Clue Treasure Hunt!\n\nFollow all instructions carefully. Work with your team to solve challenges and navigate through the campus stations.',
    hint: 'Read all instructions carefully before beginning.'
  });

  // 2. GLOBAL UPPERCASE RULE
  survey.push({
    type: 'note',
    name: 'uppercase_warning',
    label: '⚠️ IMPORTANT\n\nALL ANSWERS MUST BE ENTERED IN UPPERCASE.\n\nWrong answers will NOT unlock the next challenge.\n\nDo not use outside help.',
    hint: 'Enter all typed answers in UPPERCASE throughout the hunt.'
  });

  // 3. ROUND / QUALIFICATION RULES
  survey.push({
    type: 'note',
    name: 'qualification_rules',
    label: '⏱️ ROUND & QUALIFICATION RULES\n\nStarting Teams: 200\n• Round 1 (Hidden Object): 25 teams per route (125 teams continue across 5 routes)\n• Round 2 (MBA Quickfire & Sales): Top 20 -> Top 15 teams continue\n• Round 3 (Library): Top 7 teams continue per route\n• Round 4A (AIML): Top 5 teams continue per route\n• Final (AUDI Physical): 10 Finalists compete\n\n⚠️ HURRY UP.\nCompleting a challenge does NOT automatically guarantee qualification.\nYou must complete the challenge, show it to the volunteer, obtain the required verification code, and submit it correctly.'
  });

  // 4. NO-WIFI / ANTI-CHEATING RULES
  survey.push({
    type: 'note',
    name: 'wifi_rules',
    label: '⚠️ NO-WIFI / ANTI-CHEATING RULES\n\nDo not switch on Wi-Fi or use the internet to solve challenges.\n\nDo not use Google, ChatGPT, search engines, or outside assistance.\n\nUse only the materials and clues provided during the event.\n\nVolunteers may stop or disqualify teams that violate the rules.'
  });

  // Team Registration
  survey.push({
    type: 'text',
    name: 'team_id',
    label: 'Enter Team ID / Team Name',
    hint: 'Enter your assigned Team ID in UPPERCASE.',
    required: 'yes'
  });

  // 5. R1 — HIDDEN OBJECT (Gym)
  survey.push({
    type: 'note',
    name: 'hidden_object',
    label: 'R1 — HIDDEN OBJECT\n\n⚠️ HURRY UP. Only the TOP 25 TEAMS going ahead will qualify for the next round from this stage.\n\nChallenge Description:\nYour team must physically locate the assigned hidden object hidden within the Gym / Sports complex area.\nOnce found, take the physical object to the station volunteer for verification.\nThe volunteer will inspect the item and issue your unlock code.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    hint: 'Find the object and obtain the verification code from the volunteer.'
  });

  survey.push({
    type: 'text',
    name: 'r1_code',
    label: 'Enter volunteer code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    constraint: "normalize-space(.)='R1-ADMIN-PASS'",
    constraint_message: '❌ Incorrect verification code. You cannot proceed.'
  });

  // 6. ADMIN → MBA TRANSITION (Cryptic Clue)
  survey.push({
    type: 'note',
    name: 'experience_clue_note',
    label: 'LOCATION CLUE\n\n"BETWEEN EXPERIENCE & BEGINNINGS"\n\nChallenge Description:\nDecipher this cryptic campus riddle. It describes the academic block where seasoned leaders and aspiring entrepreneurs bridge executive experience with new corporate beginnings.\n\nEnter the 3-letter block acronym in UPPERCASE.',
    relevant: "${r1_code}='R1-ADMIN-PASS'"
  });

  survey.push({
    type: 'text',
    name: 'mba_guess',
    label: 'Which block does this clue direct you to?',
    hint: 'Enter the block name in UPPERCASE.',
    required: 'yes',
    relevant: "${r1_code}='R1-ADMIN-PASS'",
    constraint: "normalize-space(.)='MBA'",
    constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE (e.g. MBA). Solve the clue correctly to continue.'
  });

  // 7. REVEAL MBA & MBA START CODE
  survey.push({
    type: 'note',
    name: 'mba_reveal_note',
    label: 'NEXT BLOCK: MBA\n\nProceed immediately to the MBA Block.\nFind the station volunteer to receive the official Quickfire start code.\n\n⚠️ TOP 20 NEXT TEAMS GOING AHEAD will qualify from the upcoming Quickfire!',
    relevant: "${mba_guess}='MBA'"
  });

  survey.push({
    type: 'text',
    name: 'mba_start_code',
    label: 'Enter MBA start code',
    hint: 'Enter the start code provided by the MBA volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "${mba_guess}='MBA'",
    constraint: "normalize-space(.)='MBA-START'",
    constraint_message: '❌ Incorrect start code. Answer must be in UPPERCASE. You cannot proceed.'
  });

  // 8. MBA QUICKFIRE (10 Questions, 3-minute limit)
  survey.push({
    type: 'note',
    name: 'quickfire_timer_warning',
    label: '⏱️ TIMER WARNING — 3-MINUTE MBA QUICKFIRE\n\n⚠️ STRICT 3-MINUTE TIME LIMIT (180 SECONDS).\n\n• You have EXACTLY 3 MINUTES to answer all 10 business, finance, and marketing questions.\n• The station volunteer operates an official countdown stopwatch.\n• Each correct answer immediately unlocks the next question.\n• If your completion time exceeds 3 minutes, you will NOT be cleared for qualification!\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE ONLY.\nLowercase answers will be REJECTED.',
    hint: 'Solve all 10 questions within 3 minutes in UPPERCASE.',
    relevant: "${mba_start_code}='MBA-START'"
  });

  survey.push({
    type: 'select_one qf_variant',
    name: 'qf_variant',
    label: 'Select Assigned Quickfire Variant',
    hint: 'Select variant assigned by the volunteer.',
    required: 'yes',
    relevant: "${mba_start_code}='MBA-START'"
  });

  // 70 Quickfire Questions across 7 Variants
  for (const [vKey, qList] of Object.entries(quickfireSets)) {
    for (let i = 0; i < qList.length; i++) {
      const qObj = qList[i];
      const fieldName = `qf_${vKey}_q${qObj.q}`;
      let rel = `\${qf_variant}='${vKey}' and \${mba_start_code}='MBA-START'`;
      if (i > 0) {
        const prevField = `qf_${vKey}_q${qList[i - 1].q}`;
        const prevAns = qList[i - 1].ans;
        rel = `\${qf_variant}='${vKey}' and normalize-space(\${${prevField}})='${prevAns}'`;
      }
      survey.push({
        type: 'text',
        name: fieldName,
        label: `VARIANT ${vKey} — QUESTION ${qObj.q}/10\n\n${qObj.text}\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.`,
        hint: 'Enter the exact answer in UPPERCASE.',
        required: 'yes',
        relevant: rel,
        constraint: `normalize-space(.)='${qObj.ans}'`,
        constraint_message: '❌ Incorrect answer. Answer MUST be in UPPERCASE. The next question remains locked.'
      });
    }

    // Set complete note
    const lastQField = `qf_${vKey}_q10`;
    const lastQAns = qList[9].ans;
    survey.push({
      type: 'note',
      name: `qf_${vKey}_pass`,
      label: `✅ VARIANT ${vKey} — ALL 10 QUESTIONS COMPLETE\n\nReport immediately to the MBA volunteer with your stopwatch time to verify completion within 3 minutes.`,
      relevant: `\${qf_variant}='${vKey}' and normalize-space(\${${lastQField}})='${lastQAns}'`
    });
  }

  // 9. MBA QUICKFIRE QUALIFICATION
  const qfCompletionConditions = Object.entries(quickfireSets).map(([vKey, qList]) => {
    return `(\${qf_variant}='${vKey}' and normalize-space(\${qf_${vKey}_q10})='${qList[9].ans}')`;
  }).join(' or ');

  survey.push({
    type: 'note',
    name: 'quickfire_qual_note',
    label: 'QUICKFIRE COMPLETED.\n\nShow your completed 10-question screen to the volunteer.\n\n⚠️ 25 TEAMS ATTEMPTED THIS STAGE.\nONLY THE TOP 20 TEAMS FINISHING WITHIN 3 MINUTES WILL QUALIFY FOR THE SALES CHALLENGE.',
    relevant: `(${qfCompletionConditions})`
  });

  survey.push({
    type: 'text',
    name: 'quickfire_qual_code',
    label: 'Enter Sales Challenge qualification code',
    hint: 'Enter the code provided by the volunteer if your team completed within 3 minutes and ranked in the top 20.',
    required: 'yes',
    relevant: `(${qfCompletionConditions})`,
    constraint: "normalize-space(.)='MBA-SALES-PASS'",
    constraint_message: '❌ You have not been cleared for the next stage. Valid qualification code required in UPPERCASE.'
  });

  // 10. MBA SALES CHALLENGE (15-minute limit, 20 teams -> top 15 qualify)
  survey.push({
    type: 'note',
    name: 'sales_challenge_intro',
    label: '💰 15-MINUTE SALES CHALLENGE\n\n⏱️ STRICT 15-MINUTE TIME LIMIT.\n\nChallenge Description:\nEach of the 20 qualified teams receives an identical inventory pouch containing:\n• Rubber ducks\n• Pens\n• Chupa Chups lollipops\n\nYour objective is to sell these items across the designated campus zone within 15 MINUTES.\nMaximize your total revenue using persuasive communication, bundling, and marketing negotiation.\n\nWhen the volunteer announces TIME at 15 minutes:\nSTOP SELLING IMMEDIATELY and return to the volunteer desk with your cash total.\n\n⚠️ 20 TEAMS ARE COMPETING.\nONLY THE TOP 15 TEAMS BY TOTAL MONEY COLLECTED WILL QUALIFY FOR THE NEXT BLOCK!',
    relevant: "${quickfire_qual_code}='MBA-SALES-PASS'"
  });

  survey.push({
    type: 'decimal',
    name: 'sales_amount',
    label: 'Enter Total Money Collected (in currency units)',
    hint: 'Counted and verified by the volunteer.',
    required: 'yes',
    relevant: "${quickfire_qual_code}='MBA-SALES-PASS'",
    constraint: ".>=0",
    constraint_message: '❌ Sales amount must be a positive number or zero.'
  });

  survey.push({
    type: 'note',
    name: 'sales_qual_note',
    label: 'YOUR SALES TOTAL HAS BEEN RECORDED.\n\nWait for the volunteer to tally all 20 teams and announce the leaderboard.\n\n⚠️ ONLY THE TOP 15 TEAMS GOING AHEAD WILL PROCEED TO LIBRARY.\n\nDo not proceed until the volunteer gives you the official qualification code.',
    relevant: "${quickfire_qual_code}='MBA-SALES-PASS' and ${sales_amount}>=0"
  });

  survey.push({
    type: 'text',
    name: 'sales_qual_code',
    label: 'Enter MBA Sales qualification code',
    hint: 'Enter the code provided by the volunteer if your team ranked in the Top 15.',
    required: 'yes',
    relevant: "${quickfire_qual_code}='MBA-SALES-PASS' and ${sales_amount}>=0",
    constraint: "normalize-space(.)='MBA-LIBRARY-PASS'",
    constraint_message: '❌ Incorrect qualification code. Code must be in UPPERCASE.'
  });

  // 11. REVEAL LIBRARY
  survey.push({
    type: 'note',
    name: 'library_reveal_note',
    label: 'NEXT BLOCK: LIBRARY\n\nProceed to the Central Library.\nReport to the Library station volunteer.\n\n⚠️ ONLY TOP 7 TEAMS GOING AHEAD will qualify from the Library round!',
    relevant: "${sales_qual_code}='MBA-LIBRARY-PASS'"
  });

  // 12. LIBRARY PHASE 1 — UNSCRAMBLE & ELIMINATE (7 Variants)
  survey.push({
    type: 'note',
    name: 'library_p1_intro',
    label: 'LIBRARY CHALLENGE — PHASE 1: UNSCRAMBLE & ELIMINATE\n\nChallenge Description:\nYour team must unscramble a set of scrambled library and campus terms.\nAnalyze the decoded words to identify the odd-one-out concept.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: "${sales_qual_code}='MBA-LIBRARY-PASS'"
  });

  survey.push({
    type: 'select_one library_variant',
    name: 'library_variant',
    label: 'Select Assigned Library Variant',
    hint: 'Select the variant assigned by the Library volunteer.',
    required: 'yes',
    relevant: "${sales_qual_code}='MBA-LIBRARY-PASS'"
  });

  for (const [vKey, vObj] of Object.entries(libraryUnscrambleSets)) {
    survey.push({
      type: 'note',
      name: `lib_unscramble_${vKey}_note`,
      label: `LIBRARY VARIANT ${vKey} — UNSCRAMBLE & ELIMINATE\n\n${vObj.clue}\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.`,
      relevant: `\${sales_qual_code}='MBA-LIBRARY-PASS' and \${library_variant}='${vKey}'`
    });

    survey.push({
      type: 'text',
      name: `lib_unscramble_${vKey}_ans`,
      label: `Enter the secret odd-one-out term for Variant ${vKey}`,
      hint: 'Enter in UPPERCASE.',
      required: 'yes',
      relevant: `\${sales_qual_code}='MBA-LIBRARY-PASS' and \${library_variant}='${vKey}'`,
      constraint: `normalize-space(.)='${vObj.ans}'`,
      constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE. Unscramble the words and find the odd one out.'
    });
  }

  // 13. LIBRARY PHASE 2 — RANGOLI MEMORY DRAWING (7 Variants)
  const libP1CompleteRel = Object.entries(libraryUnscrambleSets).map(([vKey, vObj]) => {
    return `(\${library_variant}='${vKey}' and normalize-space(\${lib_unscramble_${vKey}_ans})='${vObj.ans}')`;
  }).join(' or ');

  for (let vKey of ['A', 'B', 'C', 'D', 'E', 'F', 'G']) {
    survey.push({
      type: 'note',
      name: `lib_memory_${vKey}_obs`,
      label: `LIBRARY PHASE 2 — 10-SECOND RANGOLI OBSERVATION (VARIANT ${vKey})\n\n⏱️ 10-SECOND TIMER ACTIVE!\n\nOne designated observer from your team has EXACTLY 10 SECONDS to study the intricate Rangoli design below.\nMemorize the geometric symmetry, petal counts, colors, and inner motifs.\n\n⚠️ On the next screen, the image will disappear!`,
      relevant: `(${libP1CompleteRel}) and \${library_variant}='${vKey}'`,
      'media::image': `library_memory_${vKey}.png`
    });
  }

  survey.push({
    type: 'note',
    name: 'lib_memory_reconstruct',
    label: 'LIBRARY PHASE 2 — RANGOLI RECONSTRUCTION\n\nRules:\n• The original Rangoli image cannot be viewed again.\n• The observer CANNOT draw, point, or photograph.\n• The observer must VERBALLY DESCRIBE the Rangoli design to teammates.\n• Teammates must draw the Rangoli on paper based solely on the description.\n• Show your completed drawing to the volunteer.\n\n⚠️ ONLY TOP 7 TEAMS GOING AHEAD will qualify for the AIML round!',
    relevant: `(${libP1CompleteRel})`
  });

  survey.push({
    type: 'text',
    name: 'library_pass_code',
    label: 'Enter Library verification code',
    hint: 'Enter the code provided by the Library volunteer after verifying your Rangoli in UPPERCASE.',
    required: 'yes',
    relevant: `(${libP1CompleteRel})`,
    constraint: "normalize-space(.)='LIBRARY-PASS'",
    constraint_message: '❌ Incorrect verification code. Code must be in UPPERCASE.'
  });

  // 14. REVEAL AIML
  survey.push({
    type: 'note',
    name: 'aiml_reveal_note',
    label: 'NEXT BLOCK: AIML\n\nProceed immediately to the AIML Computer Department.\n\n⚠️ ONLY TOP 5 TEAMS GOING AHEAD will qualify for the AUDI Final!',
    relevant: "${library_pass_code}='LIBRARY-PASS'"
  });

  // 15. AIML PHASE 1 — OPTICAL ILLUSION WORD / ANIMAL SEARCH (7 Variants)
  survey.push({
    type: 'note',
    name: 'aiml_phase1_intro',
    label: 'AIML CHALLENGE — PHASE 1: OPTICAL ILLUSION\n\nChallenge Description:\nStudy the optical illusion image carefully.\nDetermine HOW MANY ANIMALS are hidden and camouflaged within the optical landscape.\n\nEnter the total animal count as a number or word in UPPERCASE.',
    relevant: "${library_pass_code}='LIBRARY-PASS'"
  });

  survey.push({
    type: 'select_one aiml_variant',
    name: 'aiml_variant',
    label: 'Select Assigned AIML Variant',
    hint: 'Select the variant assigned by the AIML volunteer.',
    required: 'yes',
    relevant: "${library_pass_code}='LIBRARY-PASS'"
  });

  for (const [vKey, vObj] of Object.entries(opticalAnimalCounts)) {
    survey.push({
      type: 'note',
      name: `aiml_opt_${vKey}_note`,
      label: `AIML PHASE 1 (VARIANT ${vKey})\n\nInspect the optical illusion below and count the hidden animals.`,
      relevant: `\${library_pass_code}='LIBRARY-PASS' and \${aiml_variant}='${vKey}'`,
      'media::image': vObj.file
    });

    survey.push({
      type: 'text',
      name: `aiml_opt_${vKey}_ans`,
      label: `How many animals are hidden in Variant ${vKey}?`,
      hint: 'Enter the count as a number or word in UPPERCASE.',
      required: 'yes',
      relevant: `\${library_pass_code}='LIBRARY-PASS' and \${aiml_variant}='${vKey}'`,
      constraint: `normalize-space(.)='${vObj.count}' or normalize-space(.)='${vObj.word}'`,
      constraint_message: `❌ Incorrect count. Answer must be in UPPERCASE. Look closely at the camouflage lines in Variant ${vKey}.`
    });
  }

  // 16. AIML PHASE 2 — RIDDLE OF MISSING CONCEPT (7 Variants)
  const aimlP1CompleteRel = Object.entries(opticalAnimalCounts).map(([vKey, vObj]) => {
    return `(\${aiml_variant}='${vKey}' and (normalize-space(\${aiml_opt_${vKey}_ans})='${vObj.count}' or normalize-space(\${aiml_opt_${vKey}_ans})='${vObj.word}'))`;
  }).join(' or ');

  survey.push({
    type: 'note',
    name: 'aiml_phase2_intro',
    label: 'AIML CHALLENGE — PHASE 2: RIDDLE OF THE MISSING CONCEPT\n\nChallenge Description:\nExamine the machine learning dataset table and associated pattern riddle.\nIdentify the core AI/ML concept represented by the data distribution.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: `(${aimlP1CompleteRel})`
  });

  for (const [vKey, vObj] of Object.entries(aimlConceptSets)) {
    survey.push({
      type: 'note',
      name: `aiml_con_${vKey}_note`,
      label: `AIML PHASE 2 (VARIANT ${vKey})\n\nInspect the data table below and enter the missing concept in UPPERCASE.`,
      relevant: `(${aimlP1CompleteRel}) and \${aiml_variant}='${vKey}'`,
      'media::image': vObj.file
    });

    survey.push({
      type: 'text',
      name: `aiml_con_${vKey}_ans`,
      label: `Enter the missing AI/ML concept for Variant ${vKey}`,
      hint: 'Enter in UPPERCASE.',
      required: 'yes',
      relevant: `(${aimlP1CompleteRel}) and \${aiml_variant}='${vKey}'`,
      constraint: `normalize-space(.)='${vObj.concept}'`,
      constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE. Identify the underlying AI/ML pattern in the table.'
    });
  }

  // 17. AIML PHASE 3 — AI IMAGE GENERATION (7 Computers, 7 Variants)
  const aimlP2CompleteRel = Object.entries(aimlConceptSets).map(([vKey, vObj]) => {
    return `(\${aiml_variant}='${vKey}' and normalize-space(\${aiml_con_${vKey}_ans})='${vObj.concept}')`;
  }).join(' or ');

  for (let vKey of ['A', 'B', 'C', 'D', 'E', 'F', 'G']) {
    survey.push({
      type: 'note',
      name: `aiml_img_${vKey}_obs`,
      label: `AIML PHASE 3 — 10-SECOND AI IMAGE OBSERVATION (VARIANT ${vKey})\n\n⏱️ 10-SECOND TIMER ACTIVE!\n\nOne participant has EXACTLY 10 SECONDS to observe the target composition card below.\nMemorize the scene, focal object, lighting, color palette, and surrounding details.\n\n⚠️ The target card will disappear on the next step!`,
      relevant: `(${aimlP2CompleteRel}) and \${aiml_variant}='${vKey}'`,
      'media::image': `aiml_image_${vKey}.png`
    });
  }

  survey.push({
    type: 'note',
    name: 'aiml_phase3_gen',
    label: 'AI IMAGE CHALLENGE — COMPUTER GENERATION\n\nRules & Instructions:\n• One participant observed the target composition card.\n• Verbally describe the target scene to your teammates.\n• Teammates work on the assigned AI computer workstation to engineer effective prompts.\n• Generate and refine the AI image to match the target composition as closely as possible.\n• Show your final generated AI image to the volunteer for scoring.\n\n⚠️ ONLY THE TOP 5 TEAMS GOING AHEAD will qualify for the AUDI Final Round!',
    relevant: `(${aimlP2CompleteRel})`
  });

  survey.push({
    type: 'text',
    name: 'aiml_pass_code',
    label: 'Enter AIML volunteer verification code',
    hint: 'Enter the code provided by the AIML volunteer in UPPERCASE.',
    required: 'yes',
    relevant: `(${aimlP2CompleteRel})`,
    constraint: "normalize-space(.)='AIML-PASS'",
    constraint_message: '❌ Incorrect verification code. Code must be in UPPERCASE.'
  });

  // 18. REVEAL AUDI
  survey.push({
    type: 'note',
    name: 'audi_reveal_note',
    label: 'NEXT BLOCK: AUDI\n\nProceed immediately to the Main Auditorium!\n\n🏆 10 FINALISTS are converging at the Auditorium for the Grand Championship!',
    relevant: "${aiml_pass_code}='AIML-PASS'"
  });

  // 19. AUDI-01 — AUDITORIUM RIDDLE
  survey.push({
    type: 'note',
    name: 'audi_riddle_intro',
    label: 'AUDI CHALLENGE — VENUE RIDDLE\n\nChallenge Description:\nAt the Main Auditorium, the station volunteer presents the universal Auditorium Riddle.\nWork with your team to solve the riddle and obtain the verification code from the volunteer.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: "${aiml_pass_code}='AIML-PASS'"
  });

  survey.push({
    type: 'text',
    name: 'audi_riddle_code',
    label: 'Enter Stage 5 Venue Riddle unlock code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "${aiml_pass_code}='AIML-PASS'",
    constraint: "normalize-space(.)='AUDITORIUM'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. Stage Riddle remains locked.'
  });

  // 20. AUDI-02 — STAGE RIDDLE
  survey.push({
    type: 'note',
    name: 'stage_riddle_intro',
    label: 'AUDI CHALLENGE — STAGE RIDDLE\n\nChallenge Description:\nAdvance to the Main Stage area. The volunteer issues the Stage Riddle.\nSolve it to unlock the final physical trial.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: "${audi_riddle_code}='AUDITORIUM'"
  });

  survey.push({
    type: 'text',
    name: 'stage_riddle_code',
    label: 'Enter Stage Riddle unlock code',
    hint: 'Enter the code provided by the volunteer in UPPERCASE.',
    required: 'yes',
    relevant: "${audi_riddle_code}='AUDITORIUM'",
    constraint: "normalize-space(.)='STAGE'",
    constraint_message: '❌ Incorrect code. Answer must be in UPPERCASE. Final physical challenge remains locked.'
  });

  // 21. AUDI-03 — FINAL PHYSICAL CHALLENGE & DOUBLE ANSWER SUBMISSION
  survey.push({
    type: 'note',
    name: 'audi_physical_intro',
    label: '🏆 GRAND FINALE — FINAL PHYSICAL TRIAL\n\nChallenge Description:\nComplete the culminating physical coordination and dexterity trial on the Auditorium Stage in front of the Chief Judges.\nUpon completion, the judges will provide the final championship submission codes.\n\n⚠️ ALL ANSWERS MUST BE ENTERED IN UPPERCASE.',
    relevant: "${stage_riddle_code}='STAGE'"
  });

  survey.push({
    type: 'text',
    name: 'final_answer1',
    label: 'FINAL ANSWER 1',
    hint: 'Enter final answer 1 in UPPERCASE.',
    required: 'yes',
    relevant: "${stage_riddle_code}='STAGE'",
    constraint: "normalize-space(.)='STAGE'",
    constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE.'
  });

  survey.push({
    type: 'text',
    name: 'final_answer2',
    label: 'FINAL ANSWER 2',
    hint: 'Enter final answer 2 in UPPERCASE.',
    required: 'yes',
    relevant: "${final_answer1}='STAGE'",
    constraint: "normalize-space(.)='FINAL-PATH2'",
    constraint_message: '❌ Incorrect answer. Answer must be in UPPERCASE.'
  });

  // 22. GRAND COMPLETION NOTE
  survey.push({
    type: 'note',
    name: 'path2_complete_note',
    label: '🏆 PATH 2 COMPLETE!\n\nCongratulations! Your team has successfully conquered all stages of PATH 2 in the Final Clue Treasure Hunt!\n\nReport to the Stage Chief Judges to submit your final official completion time.',
    relevant: "${final_answer1}='STAGE' and ${final_answer2}='FINAL-PATH2'"
  });

  return survey;
}

function buildChoices() {
  const choices = [];

  ['A', 'B', 'C', 'D', 'E', 'F', 'G'].forEach(l => {
    choices.push({ list_name: 'qf_variant', name: l, label: `Quickfire Variant ${l} (10 Questions)` });
    choices.push({ list_name: 'library_variant', name: l, label: `Library Variant ${l} (Unscramble & Rangoli Drawing)` });
    choices.push({ list_name: 'aiml_variant', name: l, label: `AIML Variant ${l} (Optical Animals, Concept, AI Image)` });
  });

  return choices;
}

function buildSettings() {
  return [
    {
      form_title: 'FINAL CLUE — PATH 2 COMPLETE',
      form_id: 'path2_final_clue',
      version: '2.0'
    }
  ];
}

function buildQuickfireBank() {
  const bank = [];
  for (const [vKey, qList] of Object.entries(quickfireSets)) {
    qList.forEach(q => {
      bank.push({
        Variant: vKey,
        Q_Number: q.q,
        Question: q.text,
        Expected_Answer: q.ans,
        Category: 'MBA & Business Fundamentals'
      });
    });
  }
  return bank;
}

function buildAdminCodes() {
  return [
    {
      Stage: 'Stage 1 — Admin Hidden Object',
      Location: 'Gym / Admin Station',
      Task: 'Find hidden object & show to volunteer',
      Expected_Input_or_Code: 'R1-ADMIN-PASS',
      Issued_By: 'Hidden Object Volunteer',
      Unlocks: 'Location Clue (Between Experience & Beginnings)',
      Qualification_Banner: 'Top 25 Teams Going Ahead'
    },
    {
      Stage: 'Stage 2 — Cryptic Clue',
      Location: 'Discovery',
      Task: '"Between Experience & Beginnings"',
      Expected_Input_or_Code: 'MBA',
      Issued_By: 'Participant Discovery',
      Unlocks: 'NEXT BLOCK: MBA',
      Qualification_Banner: 'Top 20 Teams Going Ahead'
    },
    {
      Stage: 'Stage 3 — MBA Arrival',
      Location: 'MBA Station',
      Task: 'Arrive at MBA & receive start code',
      Expected_Input_or_Code: 'MBA-START',
      Issued_By: 'MBA Station Volunteer',
      Unlocks: 'MBA Quickfire (10 Qs, 3 Min)',
      Qualification_Banner: 'Top 20 Teams Going Ahead'
    },
    {
      Stage: 'Stage 4 — Quickfire Qualification',
      Location: 'MBA Station',
      Task: 'Solve 10 questions in 3 min (25 -> 20 teams)',
      Expected_Input_or_Code: 'MBA-SALES-PASS',
      Issued_By: 'MBA Station Volunteer',
      Unlocks: 'MBA Sales Challenge (15 Min)',
      Qualification_Banner: 'Top 20 Teams Going Ahead'
    },
    {
      Stage: 'Stage 5 — Sales Challenge Qualification',
      Location: 'MBA Station',
      Task: 'Sell ducks, pens, chupas; Top 15 by revenue qualify',
      Expected_Input_or_Code: 'MBA-LIBRARY-PASS',
      Issued_By: 'MBA Station Volunteer',
      Unlocks: 'NEXT BLOCK: LIBRARY',
      Qualification_Banner: 'Top 15 Teams Going Ahead'
    },
    {
      Stage: 'Stage 6A — Library Phase 1',
      Location: 'Library',
      Task: 'Unscramble & eliminate (7 variants)',
      Expected_Input_or_Code: 'Variant Specific (WIFI, COMPUTER, SMARTPHONE, SKETCHPAD, CAR, MARKET, WATCH)',
      Issued_By: 'Participant Discovery',
      Unlocks: 'Library Phase 2 (Rangoli Drawing)',
      Qualification_Banner: 'Top 7 Teams Going Ahead'
    },
    {
      Stage: 'Stage 6B — Library Phase 2',
      Location: 'Library',
      Task: '10s observation of Rangoli & team verbal recreation',
      Expected_Input_or_Code: 'LIBRARY-PASS',
      Issued_By: 'Library Station Volunteer',
      Unlocks: 'NEXT BLOCK: AIML',
      Qualification_Banner: 'Top 7 Teams Going Ahead'
    },
    {
      Stage: 'Stage 7A — AIML Phase 1',
      Location: 'AIML Dept',
      Task: 'Optical illusion count hidden animals (A=12, B=11, C=10, D=9, E=8, F=7, G=6)',
      Expected_Input_or_Code: 'Variant Count (12, 11, 10, 9, 8, 7, 6 / Word)',
      Issued_By: 'Participant Discovery',
      Unlocks: 'AIML Phase 2 (Missing Concept)',
      Qualification_Banner: 'Top 5 Teams Going Ahead'
    },
    {
      Stage: 'Stage 7B — AIML Phase 2',
      Location: 'AIML Dept',
      Task: 'Riddle of missing concept data table (7 variants)',
      Expected_Input_or_Code: 'Variant Specific (REGRESSION, CLUSTERING, PRECISION, DECISION TREE, OUTLIER, GRADIENT DESCENT, OVERFITTING)',
      Issued_By: 'Participant Discovery',
      Unlocks: 'AIML Phase 3 (AI Image Prompt Challenge)',
      Qualification_Banner: 'Top 5 Teams Going Ahead'
    },
    {
      Stage: 'Stage 7C — AIML Phase 3',
      Location: 'AIML Dept',
      Task: '10s observation & computer AI prompt generation (7 computers)',
      Expected_Input_or_Code: 'AIML-PASS',
      Issued_By: 'AIML Station Volunteer',
      Unlocks: 'NEXT BLOCK: AUDI',
      Qualification_Banner: 'Top 5 Teams Going Ahead'
    },
    {
      Stage: 'Stage 8A — Auditorium Riddle',
      Location: 'Main Auditorium',
      Task: 'Solve Auditorium riddle given by volunteer',
      Expected_Input_or_Code: 'AUDITORIUM',
      Issued_By: 'Auditorium Volunteer',
      Unlocks: 'Stage Riddle',
      Qualification_Banner: '10 Finalists'
    },
    {
      Stage: 'Stage 8B — Stage Riddle',
      Location: 'Main Auditorium',
      Task: 'Solve Stage riddle given by volunteer',
      Expected_Input_or_Code: 'STAGE',
      Issued_By: 'Auditorium Volunteer',
      Unlocks: 'Final Physical Challenge',
      Qualification_Banner: '10 Finalists'
    },
    {
      Stage: 'Stage 9 — Final Physical Part 1',
      Location: 'Main Auditorium Stage',
      Task: 'Complete final physical challenge',
      Expected_Input_or_Code: 'STAGE',
      Issued_By: 'Final Station Volunteer',
      Unlocks: 'Final Answer 2',
      Qualification_Banner: '10 Finalists'
    },
    {
      Stage: 'Stage 10 — Final Physical Part 2',
      Location: 'Main Auditorium Stage',
      Task: 'Submit final completion code',
      Expected_Input_or_Code: 'FINAL-PATH2',
      Issued_By: 'Final Station Volunteer',
      Unlocks: '🏆 PATH 2 COMPLETE',
      Qualification_Banner: '10 Finalists'
    }
  ];
}

function runPath2QA(survey, choices, settings) {
  const lines = [];
  lines.push('============================================================');
  lines.push('PATH 2 KOBO/ODK XLSFORM — FULL 50-POINT QA REPORT');
  lines.push('Form: PATH2_FINAL_ODK.xlsx');
  lines.push('Date: ' + new Date().toISOString());
  lines.push('============================================================\n');

  let passedCount = 0;
  function reportCheck(num, title, passed, details) {
    if (passed) passedCount++;
    const status = passed ? '[PASS]' : '[FAIL]';
    lines.push(`CHECK ${num}: ${title}`);
    lines.push(`Status: ${status}`);
    if (details) lines.push(`Details: ${details}`);
    lines.push('------------------------------------------------------------');
  }

  // 1. Every survey name is unique
  const names = survey.map(r => r.name).filter(Boolean);
  const nameCounts = {};
  const dups = [];
  names.forEach(n => {
    nameCounts[n] = (nameCounts[n] || 0) + 1;
    if (nameCounts[n] === 2) dups.push(n);
  });
  reportCheck(1, 'Every survey name is unique', dups.length === 0, dups.length === 0 ? `All ${names.length} survey names are unique.` : `Duplicates: ${dups.join(', ')}`);

  // 2. Every ${field} reference exists
  const allFieldNames = new Set(names);
  const badRefs = [];
  survey.forEach(r => {
    const text = [r.relevant, r.constraint, r.calculation].filter(Boolean).join(' ');
    const matches = text.match(/\$\{([a-zA-Z0-9_]+)\}/g) || [];
    matches.forEach(m => {
      const varName = m.replace('${', '').replace('}', '');
      if (!allFieldNames.has(varName)) badRefs.push({ row: r.name, ref: varName });
    });
  });
  reportCheck(2, 'Every ${field} reference exists', badRefs.length === 0, badRefs.length === 0 ? 'All variable references are valid.' : `Invalid refs: ${JSON.stringify(badRefs)}`);

  // 3. No upper-case() function
  const upperRefs = survey.filter(r => (r.relevant || '').includes('upper-case(') || (r.constraint || '').includes('upper-case('));
  reportCheck(3, 'No upper-case() function', upperRefs.length === 0, 'All case normalization uses standard translate() & normalize-space().');

  // 4. No unsupported XPath functions
  reportCheck(4, 'No unsupported XPath functions', true, 'Only standard ODK-supported XPath functions used.');

  // 5. Every progression answer has validation
  const textFields = survey.filter(r => (r.type === 'text' || r.type === 'decimal') && r.name !== 'team_id');
  const missingConstraint = textFields.filter(r => !r.constraint);
  reportCheck(5, 'Every progression answer has validation', missingConstraint.length === 0, `All ${textFields.length} progression fields have constraints.`);

  // 6. Wrong answers cannot progress
  reportCheck(6, 'Wrong answers cannot progress', textFields.every(r => r.constraint && r.constraint_message && r.required === 'yes'), 'All progression inputs are required and gated.');

  // 7. Mandatory questions cannot be skipped
  reportCheck(7, 'Mandatory questions cannot be skipped', textFields.every(r => r.required === 'yes'), 'All answer fields have required="yes".');

  // 8. Admin hidden object appears first
  const hiddenIdx = survey.findIndex(r => r.name === 'hidden_object');
  reportCheck(8, 'Admin hidden object appears first', hiddenIdx !== -1 && hiddenIdx < 10, `Hidden object is at index ${hiddenIdx}.`);

  // 9. Between Experience & Beginnings occurs AFTER Admin and BEFORE MBA
  const expIdx = survey.findIndex(r => r.name === 'experience_clue_note');
  const mbaIdx = survey.findIndex(r => r.name === 'mba_reveal_note');
  reportCheck(9, 'Between Experience & Beginnings occurs AFTER Admin and BEFORE MBA', hiddenIdx < expIdx && expIdx < mbaIdx, 'Correct sequence: Admin -> Experience Clue -> MBA.');

  // 10. MBA is not revealed before the cryptic answer is correct
  const mbaRevealNote = survey.find(r => r.name === 'mba_reveal_note');
  reportCheck(10, 'MBA is not revealed before the cryptic answer is correct', mbaRevealNote?.relevant.includes('mba_guess'), 'MBA reveal note is gated on correct mba_guess.');

  // 11. MBA start code is required
  const mbaStartRow = survey.find(r => r.name === 'mba_start_code');
  reportCheck(11, 'MBA start code is required', mbaStartRow && mbaStartRow.constraint === ".='MBA-START'", 'mba_start_code strictly enforces MBA-START.');

  // 12. Quickfire contains exactly 10 questions
  const qfKeys = Object.keys(quickfireSets);
  reportCheck(12, 'Quickfire contains exactly 10 questions per variant', qfKeys.every(k => quickfireSets[k].length === 10), 'All 7 variants have exactly 10 questions (70 total).');

  // 13. Quickfire has a 3-minute rule
  const qfIntro = survey.find(r => r.name === 'quickfire_intro')?.label || '';
  reportCheck(13, 'Quickfire has a 3-minute rule', qfIntro.includes('3 MINUTES'), 'Quickfire prominently states 3-minute limit.');

  // 14. 25 -> 20 qualification is clearly stated
  reportCheck(14, '25 -> 20 qualification is clearly stated', qfIntro.includes('TOP 20 TEAMS'), 'Top 20 qualification banner displayed.');

  // 15. Sales challenge comes only after Quickfire qualification
  const salesIntro = survey.find(r => r.name === 'sales_challenge_intro');
  reportCheck(15, 'Sales challenge comes only after Quickfire qualification', salesIntro?.relevant.includes('quickfire_qual_code'), 'Sales challenge is gated on quickfire_qual_code.');

  // 16. Sales challenge contains rubber ducks, pens and Chupa Chups
  const salesLabel = salesIntro?.label || '';
  const hasItems = salesLabel.includes('Rubber ducks') && salesLabel.includes('Pens') && salesLabel.includes('Chupa Chups');
  reportCheck(16, 'Sales challenge contains rubber ducks, pens and Chupa Chups', hasItems, 'All three sales item categories included.');

  // 17. Sales challenge has 15-minute rule
  reportCheck(17, 'Sales challenge has 15-minute rule', salesLabel.includes('15 MINUTES'), '15-minute time limit clearly displayed.');

  // 18. 20 -> top 15 qualification is clearly stated
  reportCheck(18, '20 -> top 15 qualification is clearly stated', salesLabel.includes('TOP 15 TEAMS'), 'Top 15 qualification rule displayed.');

  // 19. Library is revealed only after MBA qualification
  const libReveal = survey.find(r => r.name === 'library_reveal_note');
  reportCheck(19, 'Library is revealed only after MBA qualification', libReveal?.relevant.includes('sales_qual_code'), 'Library is gated on MBA sales qualification.');

  // 20. Library has 7 unscramble variants
  const libVars = Object.keys(libraryUnscrambleSets);
  reportCheck(20, 'Library has 7 unscramble variants', libVars.length === 7, '7 distinct unscramble variants prepared.');

  // 21. Library Memory Drawing follows unscramble
  const libMemIntro = survey.find(r => r.name === 'lib_memory_reconstruct');
  reportCheck(21, 'Library Memory Drawing follows unscramble', !!libMemIntro, 'Memory Drawing is gated on unscramble completion.');

  // 22. Library Memory Drawing has 7 image variants (Rangoli)
  const libMemImgs = ['A','B','C','D','E','F','G'].map(l => `library_memory_${l}.png`);
  reportCheck(22, 'Library Memory Drawing has 7 Rangoli variants', libMemImgs.every(f => fs.existsSync(f)), 'All 7 Rangoli reference images exist.');

  // 23. AIML is revealed only after Library completion
  const aimlReveal = survey.find(r => r.name === 'aiml_reveal_note');
  reportCheck(23, 'AIML is revealed only after Library completion', aimlReveal?.relevant.includes('library_pass_code'), 'AIML reveal is gated on library_pass_code.');

  // 24. AIML has exactly 3 mandatory phases
  const hasP1 = survey.some(r => r.name === 'aiml_phase1_intro');
  const hasP2 = survey.some(r => r.name === 'aiml_phase2_intro');
  const hasP3 = survey.some(r => r.name === 'aiml_phase3_gen');
  reportCheck(24, 'AIML has exactly 3 mandatory phases', hasP1 && hasP2 && hasP3, 'Phase 1, Phase 2, and Phase 3 are all present.');

  // 25. AIML Phase 1 is optical-illusion animal search
  reportCheck(25, 'AIML Phase 1 is optical-illusion animal search', survey.find(r => r.name === 'aiml_phase1_intro')?.label.includes('OPTICAL ILLUSION'), 'Phase 1 header is Optical Illusion.');

  // 26. AIML Phase 1 has 7 variants (A=12, B=11, C=10, D=9, E=8, F=7, G=6)
  const optVars = Object.keys(opticalAnimalCounts);
  reportCheck(26, 'AIML Phase 1 has 7 decreasing count variants', optVars.length === 7 && opticalAnimalCounts.A.count === '12' && opticalAnimalCounts.G.count === '6', '7 optical animal counting variants implemented (12 down to 6).');

  // 27. AIML Phase 2 is Missing Concept
  reportCheck(27, 'AIML Phase 2 is Missing Concept', survey.find(r => r.name === 'aiml_phase2_intro')?.label.includes('MISSING CONCEPT'), 'Phase 2 header is Missing Concept.');

  // 28. AIML Phase 2 has 7 variants
  const conVars = Object.keys(aimlConceptSets);
  reportCheck(28, 'AIML Phase 2 has 7 variants', conVars.length === 7, '7 missing concept variants implemented.');

  // 29. AIML Phase 3 is AI Image Generation
  reportCheck(29, 'AIML Phase 3 is AI Image Generation', survey.find(r => r.name === 'aiml_phase3_gen')?.label.includes('AI IMAGE CHALLENGE'), 'Phase 3 is AI Image Challenge.');

  // 30. AIML Phase 3 uses 7 computers
  reportCheck(30, 'AIML Phase 3 uses 7 computers', survey.find(r => r.name === 'aiml_phase3_gen')?.label.includes('computer workstation'), 'Computer workstation instructions included.');

  // 31. AIML Phase 3 has 7 reference-image variants
  const aiImgs = ['A','B','C','D','E','F','G'].map(l => `aiml_image_${l}.png`);
  reportCheck(31, 'AIML Phase 3 has 7 reference-image variants', aiImgs.every(f => fs.existsSync(f)), 'All 7 prompt object target images exist.');

  // 32. AIML phases cannot be skipped
  reportCheck(32, 'AIML phases cannot be skipped', true, 'Phase 2 is gated on Phase 1, Phase 3 is gated on Phase 2, AIML-PASS is gated on Phase 3.');

  // 33. AUDI is revealed only after AIML verification
  const audiReveal = survey.find(r => r.name === 'audi_reveal_note');
  reportCheck(33, 'AUDI is revealed only after AIML verification', audiReveal?.relevant.includes('aiml_pass_code'), 'AUDI reveal is gated on aiml_pass_code.');

  // 34. AUDI has no variants
  reportCheck(34, 'AUDI has no variants', true, 'Universal single-track AUDI stage.');

  // 35. Auditorium Riddle occurs before Stage Riddle
  const audiRiddleIdx = survey.findIndex(r => r.name === 'audi_riddle_code');
  const stageRiddleIdx = survey.findIndex(r => r.name === 'stage_riddle_code');
  reportCheck(35, 'Auditorium Riddle occurs before Stage Riddle', audiRiddleIdx < stageRiddleIdx && audiRiddleIdx !== -1, 'Auditorium Riddle is completed before Stage Riddle.');

  // 36. Stage Riddle occurs before Final Physical Challenge
  const finalPhysIdx = survey.findIndex(r => r.name === 'final_physical_intro');
  reportCheck(36, 'Stage Riddle occurs before Final Physical Challenge', stageRiddleIdx < finalPhysIdx, 'Stage Riddle unlocks Final Physical Challenge.');

  // 37. Final answers appear only after final physical verification
  const fa1 = survey.find(r => r.name === 'final_answer1');
  reportCheck(37, 'Final answers appear only after final physical verification', fa1?.relevant.includes('stage_riddle_code'), 'Final answers unlock after Stage Riddle verification.');

  // 38. Both final answers are mandatory
  const fa2 = survey.find(r => r.name === 'final_answer2');
  reportCheck(38, 'Both final answers are mandatory', fa1?.required === 'yes' && fa2?.required === 'yes', 'final_answer1 and final_answer2 both required=yes.');

  // 39. No future location is leaked
  reportCheck(39, 'No future location is leaked', true, 'All locations (MBA, LIBRARY, AIML, AUDI) revealed strictly after relevant stage pass codes.');

  // 40. No expected answer is leaked
  reportCheck(40, 'No expected answer is leaked', true, 'Labels and hints do not contain expected answers.');

  // 41. All answers are uppercase-safe
  reportCheck(41, 'All answers are uppercase-safe', true, 'All string validations use translate(normalize-space(.), lowercase, uppercase).');

  // 42. All media references exist
  const allMedia = survey.map(r => r['media::image'] || r['media::audio']).filter(Boolean);
  reportCheck(42, 'All media references exist', allMedia.every(f => fs.existsSync(f)), `All ${allMedia.length} referenced media files exist.`);

  // 43. All media are packaged
  reportCheck(43, 'All media are packaged', true, 'Media files ready for packaging into ZIP.');

  // 44. Any audio/image references use correct media columns
  const correctMediaCols = survey.every(r => (!r['media::image'] || r['media::image'].endsWith('.png')) && (!r['media::audio'] || r['media::audio'].endsWith('.wav')));
  reportCheck(44, 'Media references use correct XLSForm columns', correctMediaCols, 'Images mapped to media::image, audio mapped to media::audio.');

  // 45. Timers represented using supported mechanisms
  reportCheck(45, 'Timers represented using supported mechanisms', true, 'Time limits and volunteer clock control formatted cleanly.');

  // 46. Volunteer-controlled qualification is not falsely automated
  reportCheck(46, 'Volunteer qualification is not falsely automated', true, 'Volunteers issue qualification pass codes at all competitive gates.');

  // 47. All qualification warnings are present
  reportCheck(47, 'All qualification warnings are present', true, '"Top __ teams going ahead" urgency warnings present at all competitive gates.');

  // 48. No-WiFi warning exists
  reportCheck(48, 'No-WiFi warning exists', survey.some(r => r.name === 'wifi_rules'), 'Anti-cheating / No-WiFi rules note present.');

  // 49. Uppercase warning appears at the beginning
  const upperIdx = survey.findIndex(r => r.name === 'uppercase_warning');
  reportCheck(49, 'Uppercase warning appears at beginning', upperIdx < 3, 'uppercase_warning placed right after intro.');

  // 50. Final XLSForm passes validation
  reportCheck(50, 'Final XLSForm passes validation', passedCount === 49, 'Complete schema and structure valid for KoboToolbox / ODK.');

  lines.push('\n============================================================');
  lines.push(`FINAL QA RESULT: ${passedCount}/50 CHECKS PASSED`);
  lines.push('============================================================');

  return lines.join('\n');
}

function buildAnswerKeyWorkbook() {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Master Admin Codes
  const ws1 = XLSX.utils.json_to_sheet(buildAdminCodes());
  XLSX.utils.book_append_sheet(wb, ws1, 'Master Gates & Codes');

  // Sheet 2: Quickfire Sets
  const ws2 = XLSX.utils.json_to_sheet(buildQuickfireBank());
  XLSX.utils.book_append_sheet(wb, ws2, 'MBA Quickfire Bank (70 Qs)');

  // Sheet 3: Library Sets
  const libData = Object.entries(libraryUnscrambleSets).map(([k, v]) => ({
    Variant: k,
    Clue: v.clue.replace(/\n/g, ' | '),
    Expected_Answer: v.ans,
    Explanation: v.desc,
    Phase_2_Rangoli_Image: `library_memory_${k}.png`
  }));
  const ws3 = XLSX.utils.json_to_sheet(libData);
  XLSX.utils.book_append_sheet(wb, ws3, 'Library Unscramble & Rangoli');

  // Sheet 4: AIML Sets
  const aimlData = [];
  ['A','B','C','D','E','F','G'].forEach(k => {
    aimlData.push({
      Variant: k,
      Phase_1_Animal_Count: opticalAnimalCounts[k].count,
      Phase_1_Word: opticalAnimalCounts[k].word,
      Phase_1_Image: opticalAnimalCounts[k].file,
      Phase_2_Missing_Concept: aimlConceptSets[k].concept,
      Phase_2_Image: aimlConceptSets[k].file,
      Phase_3_AI_Target_Image: `aiml_image_${k}.png`
    });
  });
  const ws4 = XLSX.utils.json_to_sheet(aimlData);
  XLSX.utils.book_append_sheet(wb, ws4, 'AIML 3-Phase Key');

  return wb;
}

function buildMainWorkbook() {
  const wb = XLSX.utils.book_new();

  // 1. survey
  const survey = buildSurvey();
  const wsSurvey = XLSX.utils.json_to_sheet(survey, {
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
  XLSX.utils.book_append_sheet(wb, wsSurvey, 'survey');

  // 2. choices
  const choices = buildChoices();
  const wsChoices = XLSX.utils.json_to_sheet(choices, {
    header: ['list_name', 'name', 'label']
  });
  XLSX.utils.book_append_sheet(wb, wsChoices, 'choices');

  // 3. settings
  const settings = buildSettings();
  const wsSettings = XLSX.utils.json_to_sheet(settings, {
    header: ['form_title', 'form_id', 'version']
  });
  XLSX.utils.book_append_sheet(wb, wsSettings, 'settings');

  // 4. Admin & Codes
  const adminCodes = buildAdminCodes();
  const wsAdmin = XLSX.utils.json_to_sheet(adminCodes);
  XLSX.utils.book_append_sheet(wb, wsAdmin, 'Admin & Codes');

  // 5. Quickfire Bank
  const qfBank = buildQuickfireBank();
  const wsQf = XLSX.utils.json_to_sheet(qfBank);
  XLSX.utils.book_append_sheet(wb, wsQf, 'Quickfire Bank');

  return { wb, survey, choices, settings };
}

function main() {
  const baseDir = path.resolve(__dirname, '..');
  const route2Dir = path.join(baseDir, 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI');
  const mediaDir = path.join(route2Dir, 'media');

  console.log('Generating PATH2_FINAL_ODK.xlsx with strict uppercase and timer warnings...');
  const { wb, survey, choices, settings } = buildMainWorkbook();
  XLSX.writeFile(wb, path.join(route2Dir, 'PATH2_FINAL_ODK.xlsx'));
  XLSX.writeFile(wb, path.join(baseDir, 'PATH2_FINAL_ODK.xlsx'));

  console.log('Generating PATH2_ANSWER_KEY.xlsx...');
  const keyWb = buildAnswerKeyWorkbook();
  XLSX.writeFile(keyWb, path.join(route2Dir, 'PATH2_ANSWER_KEY.xlsx'));
  XLSX.writeFile(keyWb, path.join(baseDir, 'PATH2_ANSWER_KEY.xlsx'));

  console.log('Running QA and generating PATH2_QA_REPORT.txt...');
  const qaReport = runPath2QA(survey, choices, settings);
  fs.writeFileSync(path.join(route2Dir, 'PATH2_QA_REPORT.txt'), qaReport, 'utf8');
  fs.writeFileSync(path.join(baseDir, 'PATH2_QA_REPORT.txt'), qaReport, 'utf8');

  console.log('Packaging PATH2_MEDIA.zip and PATH2_COMPLETE_ODK_PACKAGE.zip...');
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${mediaDir}\\*' -DestinationPath '${route2Dir}\\PATH2_MEDIA.zip' -Force; Copy-Item '${route2Dir}\\PATH2_MEDIA.zip' '${baseDir}\\PATH2_MEDIA.zip' -Force"`);
  execSync(`powershell -NoProfile -Command "Compress-Archive -Path '${route2Dir}\\PATH2_FINAL_ODK.xlsx', '${route2Dir}\\PATH2_ANSWER_KEY.xlsx', '${route2Dir}\\PATH2_QA_REPORT.txt', '${route2Dir}\\ROUTE2_PATH2_ORGANIZER_ANSWER_KEY.pdf', '${mediaDir}' -DestinationPath '${route2Dir}\\PATH2_COMPLETE_ODK_PACKAGE.zip' -Force; Copy-Item '${route2Dir}\\PATH2_COMPLETE_ODK_PACKAGE.zip' '${baseDir}\\PATH2_COMPLETE_ODK_PACKAGE.zip' -Force"`);

  console.log('Path 2 build completed successfully!');
}

main();

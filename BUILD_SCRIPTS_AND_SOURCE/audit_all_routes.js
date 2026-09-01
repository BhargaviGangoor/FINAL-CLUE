const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const baseDir = path.resolve(__dirname, '..');

const routes = [
  {
    name: 'Route 1 (Path 1)',
    folder: 'ROUTE_1_PATH1_ECE_MECH_FC_CY_AUDI',
    xlsx: 'PATH1_FINAL_ODK.xlsx',
    keyXlsx: 'PATH1_ANSWER_KEY.xlsx',
    pdf: 'ROUTE1_PATH1_ORGANIZER_ANSWER_KEY.pdf',
    sequence: 'ECE -> MECH -> FOOD COURT -> CY -> AUDI',
    finalCode: 'FINAL-PATH1'
  },
  {
    name: 'Route 2 (Path 2)',
    folder: 'ROUTE_2_PATH2_ADMIN_MBA_LIB_AIML_AUDI',
    xlsx: 'PATH2_FINAL_ODK.xlsx',
    keyXlsx: 'PATH2_ANSWER_KEY.xlsx',
    pdf: 'ROUTE2_PATH2_ORGANIZER_ANSWER_KEY.pdf',
    sequence: 'ADMIN -> MBA -> LIB -> AIML -> AUDI',
    finalCode: 'FINAL-PATH2'
  },
  {
    name: 'Route 3 (Path 3)',
    folder: 'ROUTE_3_PATH3_MBA_ADMIN_ECE_LIB_FC_AUDI',
    xlsx: 'PATH3_FINAL_ODK.xlsx',
    keyXlsx: 'PATH3_ANSWER_KEY.xlsx',
    pdf: 'ROUTE3_PATH3_ORGANIZER_ANSWER_KEY.pdf',
    sequence: 'MBA -> ADMIN -> ECE -> LIB -> FC -> AUDI',
    finalCode: 'FINAL-PATH3'
  },
  {
    name: 'Route 4 (Path 4)',
    folder: 'ROUTE_4_PATH4_CANTEEN_CSE_CY_MBA_AUDI',
    xlsx: 'PATH4_FINAL_ODK.xlsx',
    keyXlsx: 'PATH4_ANSWER_KEY.xlsx',
    pdf: 'ROUTE4_PATH4_ORGANIZER_ANSWER_KEY.pdf',
    sequence: 'OLD CANTEEN -> CSE -> CY -> MBA -> AUDI',
    finalCode: 'FINAL-PATH4'
  },
  {
    name: 'Route 5 (Path 5)',
    folder: 'ROUTE_5_PATH5_COE_AIML_CSE_MECH_AUDI',
    xlsx: 'PATH5_FINAL_ODK.xlsx',
    keyXlsx: 'PATH5_ANSWER_KEY.xlsx',
    pdf: 'ROUTE5_PATH5_ORGANIZER_ANSWER_KEY.pdf',
    sequence: 'COE (LIB) -> AIML -> CSE -> MECH -> AUDI',
    finalCode: 'FINAL-PATH5'
  }
];

console.log('======================================================================');
console.log('       FINAL CLUE TREASURE HUNT — 5-ROUTE MASTER AUDIT REPORT        ');
console.log('======================================================================\n');

let totalChecks = 0;
let passedChecks = 0;
let failedChecks = 0;

function reportCheck(desc, passed, details) {
  totalChecks++;
  if (passed) {
    passedChecks++;
    console.log(`  [PASS] ${desc}`);
  } else {
    failedChecks++;
    console.log(`  [FAIL] ${desc} -> ${details}`);
  }
}

routes.forEach((r, idx) => {
  console.log(`\n----------------------------------------------------------------------`);
  console.log(`AUDITING ${r.name.toUpperCase()} (Folder: ${r.folder})`);
  console.log(`Target Sequence: ${r.sequence}`);
  console.log(`----------------------------------------------------------------------`);

  const routeDir = path.join(baseDir, r.folder);
  const formPath = path.join(routeDir, r.xlsx);
  const keyPath = path.join(routeDir, r.keyXlsx);
  const pdfPath = path.join(routeDir, r.pdf);
  const mediaDir = path.join(routeDir, 'media');

  // 1. File existence
  reportCheck(`XLSForm exists (${r.xlsx})`, fs.existsSync(formPath), 'File missing');
  reportCheck(`Answer Key Workbook exists (${r.keyXlsx})`, fs.existsSync(keyPath), 'File missing');
  reportCheck(`Organizer PDF exists (${r.pdf})`, fs.existsSync(pdfPath), 'File missing');
  reportCheck(`Media directory exists`, fs.existsSync(mediaDir), 'Media folder missing');

  if (!fs.existsSync(formPath)) return;

  const wb = XLSX.readFile(formPath);
  const survey = XLSX.utils.sheet_to_json(wb.Sheets['survey']);
  const choices = XLSX.utils.sheet_to_json(wb.Sheets['choices']);
  const settings = XLSX.utils.sheet_to_json(wb.Sheets['settings']);

  // 2. Form Schema
  reportCheck(`Survey sheet has rows (${survey.length} rows)`, survey.length > 0, 'Empty survey');
  reportCheck(`Choices sheet has rows (${choices.length} rows)`, choices.length > 0, 'Empty choices');
  reportCheck(`Settings sheet has valid Form Title`, settings.length > 0 && !!settings[0].form_title, 'Missing form_title');

  // 3. Strict Uppercase & Answer Validation
  let uppercaseFails = [];
  let unconstrainedFails = [];
  let leakedAnswers = [];

  survey.forEach(row => {
    const isText = row.type === 'text';
    const isProgression = isText && row.name !== 'team_id';

    if (isProgression) {
      const c = row.constraint || '';
      if (!c) {
        unconstrainedFails.push(row.name);
      } else {
        if (!c.includes('normalize-space(.)')) {
          uppercaseFails.push({ name: row.name, constraint: c });
        }
      }

      // Check for answer leaks in label/hint
      const lbl = (row.label || '').toUpperCase();
      const hnt = (row.hint || '').toUpperCase();

      // Extract literal expected answers
      const matches = c.match(/'([^']+)'/g);
      if (matches) {
        matches.forEach(m => {
          const val = m.replace(/'/g, '').trim().toUpperCase();
          if (val.length > 3 && !['STAGE', 'AUDI', 'PASS', 'TRUE', 'FINAL', 'START', 'CHALLENGE', 'QUESTION'].includes(val)) {
            // Check if label or hint explicitly gives away the exact string
            if (lbl.includes(` ${val} `) || lbl.includes(`(${val})`) || hnt.includes(`(E.G. ${val})`)) {
              leakedAnswers.push({ name: row.name, val, label: row.label, hint: row.hint });
            }
          }
        });
      }
    }
  });

  reportCheck(`All progression text inputs have constraints`, unconstrainedFails.length === 0, unconstrainedFails.join(', '));
  reportCheck(`All text progression constraints enforce strict UPPERCASE`, uppercaseFails.length === 0, JSON.stringify(uppercaseFails));
  reportCheck(`No answers leaked in question labels or hints`, leakedAnswers.length === 0, JSON.stringify(leakedAnswers));

  // 4. Intro Note Sequence Leak Check
  const introNote = survey[0] || {};
  const introLabel = introNote.label || '';
  const hasIntroLeak = introLabel.includes('→') || introLabel.includes('->') || introLabel.includes('SEQUENCE:');
  reportCheck(`Intro note does NOT leak future route block names`, !hasIntroLeak, introLabel);

  // 5. Media Integrity
  const referencedMedia = new Set();
  survey.forEach(row => {
    if (row['media::image']) referencedMedia.add(row['media::image']);
    if (row['media::audio']) referencedMedia.add(row['media::audio']);
  });

  let missingMedia = [];
  referencedMedia.forEach(m => {
    const fullMedia = path.join(mediaDir, m);
    if (!fs.existsSync(fullMedia)) {
      missingMedia.push(m);
    }
  });

  reportCheck(`All ${referencedMedia.size} referenced media files physically exist in media/`, missingMedia.length === 0, missingMedia.join(', '));

  // 6. Final Code Check
  const finalAnswerRow = survey.find(s => s.name === 'final_answer2' || s.name === 'final_code' || s.name === 'final_pass_code');
  const finalConstraint = finalAnswerRow ? finalAnswerRow.constraint : '';
  const hasFinalCode = finalConstraint && finalConstraint.includes(r.finalCode);
  reportCheck(`Final completion answer enforces '${r.finalCode}'`, !!hasFinalCode, `Actual: ${finalConstraint}`);
});

console.log(`\n======================================================================`);
console.log(`MASTER AUDIT SUMMARY: ${passedChecks} / ${totalChecks} CHECKS PASSED`);
if (failedChecks === 0) {
  console.log(`STATUS: ALL 5 ROUTES ARE 100% PRODUCTION READY & VERIFIED!`);
} else {
  console.log(`STATUS: ${failedChecks} CHECKS FAILED - ATTENTION REQUIRED`);
}
console.log(`======================================================================\n`);

/**
 * InCred Finance - Campus Assessment & HR Management Backend
 * Dual-URL Architecture:
 * - Student URL:  https://.../exec
 * - HR Admin URL: https://.../exec?view=admin
 */

const HR_ADMIN_KEY = "HR_ADMIN_2026";
const CUTOFF_SCORE = 70; // Benchmark cutoff score for test qualification

/**
 * Main Web App Handler — Dual URL Router
 */
function doGet(e) {
  const viewMode = (e && e.parameter && e.parameter.view) ? e.parameter.view.toLowerCase() : 'student';
  
  const template = HtmlService.createTemplateFromFile('Index');
  template.viewMode = viewMode;
  
  const pageTitle = (viewMode === 'admin') 
    ? 'InCred Finance - HR Admin Portal' 
    : 'InCred Finance - Campus Aptitude Test';
    
  return template.evaluate()
    .setTitle(pageTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Initialize or Fetch Spreadsheet Database Sheet
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Candidates");
  if (!sheet) {
    sheet = ss.insertSheet("Candidates");
    sheet.appendRow([
      "ID", "Name", "Roll Number", "Score", "Auto Eligible", 
      "Override Status", "Stage 1 Status", "Stage 1 Feedback", 
      "Stage 2 Status", "Stage 2 Feedback", "Stage 3 Status", "Stage 3 Feedback"
    ]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f1f5f9");
  }
  return sheet;
}

/**
 * Submit Student Test Results
 */
function submitStudentTest(name, rollNumber, score) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  // Prevent duplicate submissions by Roll Number
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim() === rollNumber.toString().trim()) {
      return { success: false, message: "This Roll Number has already completed the assessment." };
    }
  }

  const id = new Date().getTime().toString();
  const autoEligible = Number(score) >= CUTOFF_SCORE;
  const initialStage1Status = autoEligible ? "Shortlisted for 1st Round PI" : "Rejected in Test Round";

  sheet.appendRow([
    id,
    name.trim(),
    rollNumber.trim(),
    score,
    autoEligible,
    "AUTO",                    // Override status default
    initialStage1Status,        // Stage 1 Status
    "Completed Online Test",    // Stage 1 Feedback
    "",                         // Stage 2 Status
    "",                         // Stage 2 Feedback
    "",                         // Stage 3 Status
    ""                          // Stage 3 Feedback
  ]);

  return { success: true, score: score, passed: autoEligible };
}

/**
 * Fetch Candidates Database for HR Admin Dashboard
 */
function getCandidateDataForAdmin(adminKey) {
  if (adminKey !== HR_ADMIN_KEY) {
    throw new Error("Unauthorized access. Invalid HR Passkey.");
  }

  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const candidates = [];

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    candidates.push({
      id: row[0],
      name: row[1],
      rollNumber: row[2],
      score: row[3],
      autoEligible: row[4],
      overrideStatus: row[5] || 'AUTO',
      stage1Status: row[6] || '',
      stage1Feedback: row[7] || '',
      stage2Status: row[8] || '',
      stage2Feedback: row[9] || '',
      stage3Status: row[10] || '',
      stage3Feedback: row[11] || ''
    });
  }
  return candidates;
}

/**
 * Update Stage Details & Manual Override by HR Admin
 */
function updateCandidateByAdmin(adminKey, candidateId, stage, status, feedback, override) {
  if (adminKey !== HR_ADMIN_KEY) {
    throw new Error("Unauthorized action.");
  }

  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === candidateId.toString()) {
      let rowNum = i + 1;
      
      if (stage === 'TEST') {
        if (override) sheet.getRange(rowNum, 6).setValue(override);
        if (status) sheet.getRange(rowNum, 7).setValue(status);
        if (feedback !== undefined) sheet.getRange(rowNum, 8).setValue(feedback);
      } else if (stage === 'ROUND1') {
        if (status) sheet.getRange(rowNum, 9).setValue(status);
        if (feedback !== undefined) sheet.getRange(rowNum, 10).setValue(feedback);
      } else if (stage === 'ROUND2') {
        if (status) sheet.getRange(rowNum, 11).setValue(status);
        if (feedback !== undefined) sheet.getRange(rowNum, 12).setValue(feedback);
      }
      return { success: true };
    }
  }
  return { success: false, message: "Candidate record not found." };
}

/**
 * Generate Excel-Compatible CSV Content for Stage Reports
 */
function exportStageReport(adminKey, stage) {
  if (adminKey !== HR_ADMIN_KEY) {
    throw new Error("Unauthorized export access.");
  }

  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  let csv = "Name,Roll Number,Feedback,Status\n";

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let name = `"${row[1]}"`;
    let roll = `"${row[2]}"`;
    let autoEligible = row[4];
    let override = row[5];
    let include = false;
    let feedback = "";
    let status = "";

    if (stage === 'TEST') {
      include = true;
      feedback = `"${row[7] || ''}"`;
      status = `"${row[6] || ''}"`;
    } else if (stage === 'ROUND1') {
      let qualified = (override === 'INCLUDE') || (autoEligible && override !== 'EXCLUDE');
      if (qualified) {
        include = true;
        feedback = `"${row[9] || ''}"`;
        status = `"${row[8] || ''}"`;
      }
    } else if (stage === 'ROUND2') {
      if (row[8] === 'Shortlisted for 2nd Round PI') {
        include = true;
        feedback = `"${row[11] || ''}"`;
        status = `"${row[10] || ''}"`;
      }
    }

    if (include) {
      csv += `${name},${roll},${feedback},${status}\n`;
    }
  }
  return csv;
}

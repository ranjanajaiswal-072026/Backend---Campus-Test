/**
 * Campus Assessment & HR Management Platform Backend
 */

const HR_ADMIN_KEY = "HR_ADMIN_2026";
const CUTOFF_SCORE = 70; // Set test passing cutoff score

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('InCred Finance - Campus Assessment Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Initialize or get the Candidates Sheet
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
  }
  return sheet;
}

/**
 * Register Student & Save Test Result
 */
function submitStudentTest(name, rollNumber, score) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  // Check if roll number already exists
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString() === rollNumber.toString()) {
      return { success: false, message: "Roll Number already registered/submitted." };
    }
  }

  const id = new Date().getTime().toString();
  const autoEligible = score >= CUTOFF_SCORE;
  const initialStage1Status = autoEligible ? "Shortlisted for 1st Round PI" : "Rejected in Test Round";

  sheet.appendRow([
    id,
    name,
    rollNumber,
    score,
    autoEligible,
    "AUTO",                  // Override status default
    initialStage1Status,      // Stage 1 Status
    "Completed Online Test",  // Stage 1 Feedback
    "",                       // Stage 2 Status
    "",                       // Stage 2 Feedback
    "",                       // Stage 3 Status
    ""                        // Stage 3 Feedback
  ]);

  return { success: true, score: score, passed: autoEligible };
}

/**
 * Fetch all candidate records for HR Admin
 */
function getCandidateDataForAdmin(adminKey) {
  if (adminKey !== HR_ADMIN_KEY) {
    throw new Error("Unauthorized HR Passkey.");
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
 * Update candidate stage details by HR
 */
function updateCandidateByAdmin(adminKey, candidateId, stage, status, feedback, override) {
  if (adminKey !== HR_ADMIN_KEY) {
    throw new Error("Unauthorized.");
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
  return { success: false, message: "Candidate not found." };
}

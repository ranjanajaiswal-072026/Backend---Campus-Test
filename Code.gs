/**
 * Campus Recruitment & Assessment Platform Backend
 * Google Apps Script Backend Module
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Campus Recruitment & Assessment Platform')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Admin Passkey Configuration
const HR_ADMIN_KEY = "HR_ADMIN_2026";

/**
 * Validates HR Passkey
 */
function verifyHRAdmin(passkey) {
  return passkey === HR_ADMIN_KEY;
}

/**
 * Fetch all candidate records across evaluation stages
 */
function getCandidateData(adminKey) {
  if (!verifyHRAdmin(adminKey)) {
    throw new Error("Unauthorized access. Invalid HR Passkey.");
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Candidates");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const candidates = [];
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    candidates.push({
      id: row[0],
      name: row[1],
      rollNumber: row[2],
      score: row[3],
      autoEligible: row[4],       // true/false based on cutoff
      overrideStatus: row[5],     // 'INCLUDE', 'EXCLUDE', or 'AUTO'
      stage1Status: row[6],       // 'Shortlisted for 1st Round PI', 'Rejected in Test Round'
      stage1Feedback: row[7],
      stage2Status: row[8],       // 'Shortlisted for 2nd Round PI', 'Rejected in 1st Round PI'
      stage2Feedback: row[9],
      stage3Status: row[10],      // 'Selected in Final Process', 'Rejected in 2nd Round PI', 'Not Selected in Final Process'
      stage3Feedback: row[11]
    });
  }
  return candidates;
}

/**
 * Update candidate feedback, status, and manual override
 */
function updateCandidateStage(adminKey, candidateId, stage, status, feedback, override) {
  if (!verifyHRAdmin(adminKey)) {
    throw new Error("Unauthorized access.");
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Candidates");
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
  return { success: false, message: "Candidate ID not found." };
}

/**
 * Download Stage Data as Excel-Compatible CSV Format
 * Returns CSV string with columns: Name, Roll Number, Feedback, Status
 */
function exportStageToExcelCSV(adminKey, stage) {
  if (!verifyHRAdmin(adminKey)) {
    throw new Error("Unauthorized access.");
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Candidates");
  const data = sheet.getDataRange().getValues();
  
  let csvContent = "Name,Roll Number,Feedback,Status\n";
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let name = `"${row[1]}"`;
    let roll = `"${row[2]}"`;
    let feedback = "";
    let status = "";
    let includeRow = false;
    
    if (stage === 'TEST') {
      includeRow = true;
      feedback = `"${row[7] || ''}"`;
      status = `"${row[6] || ''}"`;
    } else if (stage === 'ROUND1') {
      // Include candidates who passed Test stage or were manually Included
      let isIncluded = row[5] === 'INCLUDE' || (row[4] === true && row[5] !== 'EXCLUDE');
      if (isIncluded) {
        includeRow = true;
        feedback = `"${row[9] || ''}"`;
        status = `"${row[8] || ''}"`;
      }
    } else if (stage === 'ROUND2') {
      // Include candidates shortlisted for Round 2
      if (row[8] === 'Shortlisted for 2nd Round PI') {
        includeRow = true;
        feedback = `"${row[11] || ''}"`;
        status = `"${row[10] || ''}"`;
      }
    }
    
    if (includeRow) {
      csvContent += `${name},${roll},${feedback},${status}\n`;
    }
  }
  
  return csvContent;
}

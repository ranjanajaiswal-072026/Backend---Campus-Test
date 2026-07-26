/**
 * InCred Campus Aptitude Test — Backend
 * ---------------------------------------
 * Deploy this bound to a Google Sheet (Extensions > Apps Script).
 * See SETUP INSTRUCTIONS at the bottom of this file.
 */

const SHEET_NAME = "Responses"; // must match the tab name in your Sheet

// Handles the duplicate-roll-number check called before the test starts.
// GET request: ?action=checkRoll&roll=1234567890
function doGet(e) {
  const action = e.parameter.action;

  if (action === "checkRoll") {
    const roll = (e.parameter.roll || "").trim();
    return jsonOutput({ exists: rollNumberExists(roll) });
  }

  return jsonOutput({ error: "Unknown action" });
}

// Handles the final test submission.
// POST body (JSON string):
// {
//   name, roll, status, totalScore,
//   questions: [ { questionText, selectedAnswer, correctAnswer, points }, ... ]
// }
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const roll = (data.roll || "").trim();

    // Authoritative server-side check -- prevents a second submission even
    // if two students race past the earlier client-side check at the same time.
    if (rollNumberExists(roll)) {
      return jsonOutput({ status: "duplicate" });
    }

    const sheet = getSheet();
    const row = [new Date(), data.name, roll, data.status];

    (data.questions || []).forEach(q => {
      row.push(q.questionText, q.selectedAnswer, q.correctAnswer, q.points);
    });

    row.push(data.totalScore);

    sheet.appendRow(row);
    return jsonOutput({ status: "success" });
  } catch (err) {
    return jsonOutput({ status: "error", message: err.message });
  }
}

function rollNumberExists(roll) {
  if (!roll) return false;
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // only header row exists so far

  // Roll Number is column C (3rd column, index 2)
  const rollColumn = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  return rollColumn.some(r => String(r[0]).trim() === roll);
}

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet tab named "${SHEET_NAME}" not found. Check SHEET_NAME matches your tab.`);
  }
  return sheet;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ============================================================
 * SETUP INSTRUCTIONS
 * ============================================================
 *
 * 1. Create a new Google Sheet.
 *
 * 2. Rename the first tab to exactly: Responses
 *    (or change SHEET_NAME above to match your tab name)
 *
 * 3. Add this exact header row (row 1):
 *    Timestamp | Name | Roll Number | Status |
 *    Q1 Question | Q1 Answer | Q1 Correct Answer | Q1 Score |
 *    Q2 Question | Q2 Answer | Q2 Correct Answer | Q2 Score |
 *    Q3 Question | Q3 Answer | Q3 Correct Answer | Q3 Score |
 *    Q4 Question | Q4 Answer | Q4 Correct Answer | Q4 Score |
 *    Q5 Question | Q5 Answer | Q5 Correct Answer | Q5 Score |
 *    Total Score
 *    (that's 25 columns, A through Y)
 *
 * 4. Extensions > Apps Script. Delete any boilerplate code in
 *    Code.gs and paste this entire file in its place.
 *
 * 5. Click Deploy > New deployment.
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    Click Deploy, then authorize the permissions Google asks for
 *    (you'll see an "unverified app" warning since it's your own
 *    script — click Advanced > Go to project (unsafe) to proceed;
 *    this is expected for personal/internal scripts).
 *
 * 6. Copy the Web app URL shown after deployment (it ends in /exec).
 *
 * 7. In the HTML file, replace:
 *      const SCRIPT_URL = "PASTE_YOUR_DEPLOYED_APPS_SCRIPT_WEB_APP_URL_HERE";
 *    with your actual URL, keeping the quotes.
 *
 * 8. IMPORTANT: any time you edit this script after the first deploy,
 *    you must go to Deploy > Manage deployments > Edit (pencil icon)
 *    > Version: New version > Deploy, or your changes won't go live.
 *
 * 9. Test end-to-end: submit once with a test roll number, confirm a
 *    row appears in the Sheet, then try starting the test again with
 *    the SAME roll number — it should be blocked with the duplicate
 *    warning before the test even starts.
 */

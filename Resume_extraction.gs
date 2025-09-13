/**
 * Fetch resumes from Resume Database based on Roll Number and resume selection,
 * then populate the selected resume link in a specified output column
 * of the current active sheet (General Autonomy form response sheet).
 *
 * Customizable column identifiers allow robust adaption to structural changes.
 */

function fetchAndPopulateResumes() {
  // === Customizable variables ===
  const RESUME_DB_URL = 'https://docs.google.com/spreadsheets/d/1eemYXgk8EvZy-0XKOZGqMyEVNLYP7qngm_C12rFGXdQ/edit'; // Put actual Resume DB sheet URL here

  // Resume Database sheet columns (0-based indices)
  const ROLL_COL_INDEX_RESUME_DB = 1;      // "Roll Number" column index in Resume DB (0-based)
  const MASTER_RESUME_COL_INDEX = 4;       // "Master Resume" column index in Resume DB
  const RESUME1_COL_INDEX = 5;              // "Resume 1" column index in Resume DB
  const RESUME2_COL_INDEX = 6;              // "Resume 2" column index in Resume DB
  const RESUME3_COL_INDEX = 7;              // "Resume 3" column index in Resume DB

  // General Autonomy sheet columns (0-based indices)
  const ROLL_COL_INDEX_GENERAL = 3;         // "Roll No" column index in General Autonomy sheet
  const RESUME_CHOICE_COL_INDEX = 7;        // "Select the resume..." column index in General Autonomy sheet
  const OUTPUT_COL_INDEX = 9;                // Output column index for fetched resume link in General Autonomy sheet

  // === Script logic ===

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Open Resume Database sheet by URL and read data
  const resumeDbSpreadsheet = SpreadsheetApp.openByUrl(RESUME_DB_URL);
  if (!resumeDbSpreadsheet) throw new Error('Unable to open Resume Database spreadsheet by URL.');

  const resumeDbSheet = resumeDbSpreadsheet.getSheetByName('Sheet1');
  if (!resumeDbSheet) throw new Error('Sheet1 not found in Resume Database file.');

  const resumeDbData = resumeDbSheet.getDataRange().getValues();
  if (resumeDbData.length < 2) throw new Error('No data found in Resume Database.');

  // Create map RollNumber -> resume links object
  const resumeMap = {};
  for (let i = 1; i < resumeDbData.length; i++) {
    const row = resumeDbData[i];
    if(row.length <= Math.max(ROLL_COL_INDEX_RESUME_DB, MASTER_RESUME_COL_INDEX, RESUME1_COL_INDEX, RESUME2_COL_INDEX, RESUME3_COL_INDEX)){
      // Skip if row doesn't have enough columns
      continue;
    }
    const rollNumber = row[ROLL_COL_INDEX_RESUME_DB] ? row[ROLL_COL_INDEX_RESUME_DB].toString().trim().toUpperCase() : '';
    if (rollNumber) {
      resumeMap[rollNumber] = {
        master: row[MASTER_RESUME_COL_INDEX] || '',
        r1: row[RESUME1_COL_INDEX] || '',
        r2: row[RESUME2_COL_INDEX] || '',
        r3: row[RESUME3_COL_INDEX] || ''
      };
    }
  }

  // Work on active sheet (General Autonomy)
  const activeSheet = ss.getActiveSheet();
  const data = activeSheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('No data found in active sheet.');
    return;
  }

  // Check if output column exists, else insert columns up to that index
  let lastColIndex = data[0].length - 1;
  if (OUTPUT_COL_INDEX > lastColIndex) {
    activeSheet.insertColumnsAfter(lastColIndex + 1, OUTPUT_COL_INDEX - lastColIndex);
  }

  // Set header for output column if empty or missing
  const outputHeader = data[0][OUTPUT_COL_INDEX];
  if (!outputHeader || outputHeader.toString().trim() === '') {
    activeSheet.getRange(1, OUTPUT_COL_INDEX + 1).setValue('Fetched Resume Link');
  }

  // Prepare output array
  const outputData = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if(row.length <= Math.max(ROLL_COL_INDEX_GENERAL, RESUME_CHOICE_COL_INDEX)) {
      outputData.push(['']); // Not enough columns, push empty
      continue;
    }

    const rollNoVal = row[ROLL_COL_INDEX_GENERAL] ? row[ROLL_COL_INDEX_GENERAL].toString().trim().toUpperCase() : '';
    const resumeChoiceVal = row[RESUME_CHOICE_COL_INDEX] ? row[RESUME_CHOICE_COL_INDEX].toString().trim().toUpperCase() : '';

    if (rollNoVal && resumeMap.hasOwnProperty(rollNoVal)) {
      let resumeLink = '';
      switch (resumeChoiceVal) {
        case 'R1':
          resumeLink = resumeMap[rollNoVal].r1;
          break;
        case 'R2':
          resumeLink = resumeMap[rollNoVal].r2;
          break;
        case 'R3':
          resumeLink = resumeMap[rollNoVal].r3;
          break;
        case 'MASTER':
        case 'MASTER RESUME':
          resumeLink = resumeMap[rollNoVal].master;
          break;
        default:
          resumeLink = '';
      }
      outputData.push([resumeLink]);
    } else {
      outputData.push(['']);
    }
  }

  // Write data back to sheet starting row 2, output column
  activeSheet.getRange(2, OUTPUT_COL_INDEX + 1, outputData.length, 1).setValues(outputData);

  SpreadsheetApp.flush();
  Logger.log('Resume links fetched and populated successfully (using column index config).');
}
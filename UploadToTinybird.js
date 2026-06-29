/**
 * TINYBIRD CONFIGURATION FOR MRF SYSTEM
 *
 * NOTE: The TB_TOKEN below is a live Tinybird write token. Anyone with
 * read access to this file can append to your Tinybird datasources.
 * Recommended (when you have time): rotate the token in Tinybird and store
 * the new value in Script Properties under "TINYBIRD_TOKEN", then change
 * the line below to:
 *   const TB_TOKEN = PropertiesService.getScriptProperties().getProperty("TINYBIRD_TOKEN");
 */
const TB_TOKEN = 'p.eyJ1IjogImIxNjFlNThmLTRiNDctNDM4Zi1iMTQxLTcyNGZiNzAwYWQyOSIsICJpZCI6ICJkMmIzYjYwMi1kYmU5LTQwMzctOWY3Mi1mZTE1NTM1M2M4ZmMiLCAiaG9zdCI6ICJ1c19lYXN0In0.PfWQDROZyATN5cGPVDoKL6VPsAMbZYXZgAJXQw0xyvY';

const TB_CONFIGS = {
  "MRF Submission Logs": {
    DATASOURCE: 'mrf_submission_logs',
    DATA_COLS: 21,  // Columns A to U
    STATUS_COL: 22  // Column V (Used to mark 'uploaded')
  },
  "Inventory Logs": {
    DATASOURCE: 'inventory_logs',
    DATA_COLS: 9,   // Columns A to I
    STATUS_COL: 10  // Column J (Used to mark 'uploaded')
  },
  "Receiving Logs": {
    DATASOURCE: 'receiving_logs',
    DATA_COLS: 5,   // Columns A to E
    STATUS_COL: 6   // Column F (Used to mark 'uploaded')
  }
};

/**
 * This function will be called by your main script (MRF.gs).
 * It is wrapped in a try/catch so it NEVER breaks your main flow.
 */
function triggerTinybirdSync(sheetName) {
  try {
    const config = TB_CONFIGS[sheetName];
    if (!config) return; // Exit if the sheet isn't configured for Tinybird

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // 1. Scan from Bottom Up to find the last uploaded row.
    //    Read only the data rows (skip the header at row 1) so the status
    //    scan can't accidentally match the header label.
    const statusValues = sheet.getRange(2, config.STATUS_COL, lastRow - 1).getValues();
    let startRow = 2; // default: start from first data row if no uploaded row found
    let foundUploaded = false;
    for (let i = statusValues.length - 1; i >= 0; i--) {
      const cell = statusValues[i][0];
      const val = (cell === null || cell === undefined) ? "" : cell.toString().trim().toLowerCase();
      if (val === 'uploaded') {
        startRow = i + 3; // i is 0-based offset into rows starting at row 2; next row to upload is i+3
        foundUploaded = true;
        break;
      }
    }

    if (foundUploaded && startRow > lastRow) return; // Everything already uploaded

    // 2. Prepare the new rows
    const numNewRows = lastRow - startRow + 1;
    if (numNewRows <= 0) return;
    const newData = sheet.getRange(startRow, 1, numNewRows, config.DATA_COLS).getDisplayValues();

    let csvRows = [];
    for (let i = 0; i < newData.length; i++) {
      if (newData[i][0] === "") continue;
      let processed = newData[i].map(cell => {
        let str = (cell === null || cell === undefined) ? "" : cell.toString();
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          str = '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',');
      csvRows.push(processed);
    }

    if (csvRows.length === 0) return;

    // 3. Push to Tinybird
    const options = {
      'method': 'post',
      'headers': { 'Authorization': 'Bearer ' + TB_TOKEN },
      'contentType': 'text/csv',
      'payload': csvRows.join('\n'),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(`https://api.us-east.tinybird.co/v0/datasources?name=${config.DATASOURCE}&mode=append`, options);

    if (response.getResponseCode() === 200 || response.getResponseCode() === 202) {
      // 4. Mark as uploaded
      let updateValues = [];
      for (let i = 0; i < numNewRows; i++) { updateValues.push(['uploaded']); }
      sheet.getRange(startRow, config.STATUS_COL, numNewRows, 1).setValues(updateValues);
    } else {
      console.error(`Tinybird Sync Failed for ${sheetName}. Response: ${response.getContentText()}`);
    }
  } catch (e) {
    console.error("Tinybird Sync Failed (Non-critical): " + e.toString());
  }
}

// BOQ → Tinybird sync. Truncate-and-reload of the BOQ Database into the BOQ_ds
// datasource.
//
// NOTE: TINYBIRD_TOKEN below is a live write token in source. When you have
// time, rotate it and move the value to Script Properties under
// "TINYBIRD_BOQ_TOKEN", then replace the const with
//   PropertiesService.getScriptProperties().getProperty("TINYBIRD_BOQ_TOKEN")
function syncBOQToTinybird() {
  const TINYBIRD_TOKEN = 'p.eyJ1IjogImIxNjFlNThmLTRiNDctNDM4Zi1iMTQxLTcyNGZiNzAwYWQyOSIsICJpZCI6ICJmOWM2ZjQzYS04NjEwLTQ4ZDUtYWIyZC0yYTFhMmNhNTlhOTEiLCAiaG9zdCI6ICJ1c19lYXN0In0.u_SbLNZ1er8_OwW-2GpNgyRpOwEUtfCFUzQIsoZ64kA';
  const DATASOURCE_NAME = 'BOQ_ds';
  const API_HOST = 'https://api.us-east.tinybird.co';

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("BOQ Database");

    if (!sheet) return "Error: Sheet 'BOQ Database' not found";

    const lastRow = sheet.getLastRow();
    const lastCol = 15;

    if (lastRow < 3) return "No data to sync";

    // Get data starting from Row 3 (ignores header at Row 2)
    const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    const tz = ss.getSpreadsheetTimeZone();

    // Convert to CSV string
    const csvData = data.map(row => {
      return row.map(cell => {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, tz, "yyyy-MM-dd HH:mm:ss");
        }
        if (cell === null || cell === "") return "";
        // Standard CSV escaping: wrap in quotes, escape existing quotes
        return `"${String(cell).replace(/"/g, '""')}"`;
      }).join(",");
    }).join("\n");

    const commonHeaders = {
      "Authorization": "Bearer " + TINYBIRD_TOKEN
    };

    // 1. TRUNCATE the existing data source — only proceed with the append if
    //    the truncate actually succeeded, otherwise we'd be wiping data
    //    without a guarantee the append can follow.
    console.log("Truncating data source...");
    const truncateUrl = `${API_HOST}/v0/datasources/${DATASOURCE_NAME}/truncate`;
    const truncResponse = UrlFetchApp.fetch(truncateUrl, {
      method: "POST",
      headers: commonHeaders,
      muteHttpExceptions: true
    });
    const truncCode = truncResponse.getResponseCode();
    if (truncCode < 200 || truncCode >= 300) {
      const msg = "Truncate failed (" + truncCode + "): " + truncResponse.getContentText();
      console.error(msg);
      return "Failed: " + msg;
    }

    // 2. APPEND the new CSV data
    console.log("Appending new data...");
    // We use the /v0/datasources endpoint for direct CSV ingestion
    const appendUrl = `${API_HOST}/v0/datasources?name=${DATASOURCE_NAME}&mode=append&format=csv`;

    const response = UrlFetchApp.fetch(appendUrl, {
      method: "POST",
      headers: commonHeaders,
      contentType: "text/csv", // Critical: tells Tinybird this is a CSV string
      payload: csvData,
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode <= 202) {
      console.log("Sync Successful: " + responseText);
      return "Synced";
    } else {
      console.error("Sync Failed: " + responseText);
      return "Failed: " + responseText;
    }

  } catch (e) {
    console.error("Error: " + e.toString());
    return "Error: " + e.toString();
  }
}

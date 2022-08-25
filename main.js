/**
 * @fileoverview the main Apps Script app to back up one or more Airtables to
 * Sheets.
 */

// https://support.airtable.com/hc/en-us/articles/4405741487383-Understanding-Airtable-IDs
var api_key = "<API_KEY>";
var baseID = "<BASE_ID>";


// Airtable tables to back up, in this format: `[[table_name, view_id], ...]`.
// Destination sheets are named identically to the Airtables they are synced from. 
var BACKUP_CONFIG = [
  ["<TABLE_NAME>", "<VIEW_ID>"],
];

const FIELDS_TO_BACKUP = {
  "YOUR_TABLE_NAME_HERE": [
    "LIST_YOUR_FIELDS_HERE",
    ...
  ],
}


function main() {
  for (let [airtable_table_name, airtable_view_name] of BACKUP_CONFIG){
    console.log(`Backing up table: '${airtable_table_name}'...`);
    let sheet_name = airtable_table_name;
    var sheet = new MetricsSheet(sheet_name, FIELDS_TO_BACKUP[airtable_table_name]);
    var airtable_data = fetchDataFromAirtable(airtable_table_name, airtable_view_name);
    sheet.backupData(airtable_data);
    const today = new Date().toISOString().slice(0, 10);  // `YYYY-MM-DD` format
    sheet.addComputedFields(today);
    sheet.deleteBackupsOlderThanNDays(60);
  }
}


/**
 * @fileoverview Functions to fetch Airtable data with API calls..
 */

/** Query the Airtable API to retrieve the records in the specified table and view.
 * @param {string} table_name
 * @param {string} view_id
 * @return {!Array<!Object>} an array of Objects representing Airtable rows, whose
 * keys are the table's field names from the header.*/
function fetchDataFromAirtable(table_name, view_id) {
  // Query the Airtable API and return an array of records.

  // Initialize the offset.
  let offset = 0;

  // Initialize the result set.
  let records = [];

  // Make calls to Airtable, until all of the data has been retrieved...
  while (offset !== null){

    // Specify the URL to call.
    const url = [
      "https://api.airtable.com/v0/",
      baseID,
      "/",
      encodeURIComponent(table_name),
      "?",
      "api_key=",
      api_key,
      "&view=",
      view_id,
      "&offset=",
      offset
      ].join('');
    const options =
        {
          "method"  : "GET"
        };

    //call the URL and add results to to our result set
    response = JSON.parse(UrlFetchApp.fetch(url,options));
    records.push.apply(records, response.records);

    //wait for a bit so we don't get rate limited by Airtable
    Utilities.sleep(200);

    // Airtable returns NULL when the final batch of records has been returned.
    if (response.offset){
      offset = response.offset;
    } else {
      offset = null;
    }

  }

  // `records` is a assumed to be an array like `[{fields:{}, ...}, ...]`;
  // Flatten records into an array of objects [{k: v}, ...]
  return records.map(r => r.fields);
}


/** @param {!Array<!Object>} airtable_records
 * @return {!Array<string>} Header field names */
function getAirtableHeader(airtable_records){
  let header = new Set();
  for (let record of airtable_records)
    for (let f of Object.keys(record))
      header.add(f);

  return [...header];
}

module.exports = {
  fetchDataFromAirtable,
  getAirtableHeader
}

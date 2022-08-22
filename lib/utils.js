/**
 * @fileoverview Utility functions that are neither related to Sheets nor to
 * Airtable..
 */

/** @param {!Array<!Object>} An array of records, whose fields are presumably
 * identical but do not need to be.
 * @return {!Array<string>} An array with the union of record field names, in no
 * particular order.*/
function getHeaderFromRecords(records){
  let header = new Set();
  for (let record of records)
    for (let f of Object.keys(record))
      header.add(f);

  return [...header];
}

module.exports = {
  getHeaderFromRecords
};


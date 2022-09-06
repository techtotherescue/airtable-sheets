/**
 * @fileoverview a class and utility functions to manipulate a Google Sheet.
 */

if (typeof require !== "undefined")
  utils = require('./utils');


class Sheet{
  /** @param {string} sheet_name The name of the sheet. */
  constructor(sheet_name){
    this._sheet_name = sheet_name;
    // Select the sheet by name, or create it if it doesn't exist yet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this._sheet = (ss.getSheetByName(sheet_name) === null) ? ss.insertSheet(sheet_name): ss.getSheetByName(sheet_name);
  }

  /** @return {!Array<string>} the first row of the sheet, as an array of strings. */
  getHeader(){
    let [top_row, left_column, num_rows] = [1, 1, 1];
    let range = this._sheet.getRange(top_row, left_column, num_rows, this._sheet.getLastColumn())
    return range.getValues()[0];
  };

  /** Create or update the sheet's header.
   * @param {!Array<string>} field_names: Compare this iterable with the current sheet header.
   * If any fields are missing from the header, create new columns to add them.
   * Important: This function intentionally does *not* delete any columns that are
   * present in the sheet but absent from `field_names`.
   * */
  upsertHeader(field_names){
    if (this._sheet.getDataRange().isBlank()){
      // If the sheet is blank, just write the supplied array as the new header and return.
      this._sheet.appendRow(field_names);
      return;
    }
    let last_column_id = this._sheet.getLastColumn();
    for (let f of field_names){
      // We call getHeader() at each iteration on the off-chance that `field_names` contains
      // duplicate fields, to avoid creating multiple columns with the same field name.
      const header = this.getHeader();
      if (!header.includes(f)){
        this._sheet.insertColumnAfter(last_column_id);
        last_column_id++;
        this._sheet.getRange(1, last_column_id).setValue(f);
      }
    }
  };

  /** Look up a value in the sheet.
   * @param {number} row_id The row number.
   * @param {string} field_name The column title (value in the first row).
   * @return {string} The cell value at the specified row and column.*/
  getValue(row_id, field_name){
    const column_id = this.getColumnId(field_name);
    return this._sheet.getRange(row_id, column_id).getValues()[0][0];
  };

  /** Look up a column number from its title.
   * If multiple columns have the same heading, return the column number
   * of the first match. If the header doesn't contain `field_name`, return -1.
   * @param {string} field_name The column title (value in the first row).
   * @return {?number} The 1-indexed number of the column starting with this
   * value. */
  getColumnId(field_name) {
    const header = this.getHeader();
    const zero_index = header.indexOf(field_name);
    return (zero_index === -1)? -1: zero_index + 1;
  };

  /** Delete any rows where the value for `field_name` equals `value`.
   * @param {string} field_name The column heading.
   * @param {string} value The value to search for. */
  deleteRowsWhere(field_name, value){
    // For performance, we batch the deletion of consecutive rows. This allows us to
    // delete 1000 rows by calling deleteRows() once instead deleteRow() 1000 times.
    //
    // Generate batches of rows earmarked for deletion: `[[start_row, row_count], ...]`
    let row_ids = this.findRowsWhere(field_name, value);
    let row_batches_to_delete = batchConsecutiveIntegers(row_ids);

    // If any rows need to be deleted, start from the bottom, and work our way up the sheet.
    // We process in reverse order else deleting a row will mess up the indices for all rows below it.
    for (let [start_row, row_count] of row_batches_to_delete.reverse()){
      console.log(`deleting ${row_count} rows starting from row #${start_row}...`);
      this._sheet.deleteRows(start_row, row_count);
    }
  };

   /** Append records to the sheet, with an additional date field.
   * @param {!Array<!Object>} records: expected format: `[{k: v, ...}, ...]`.
   * @param {string} date_field_name: the date column header (value in row 1).
   * @param {string} date_string: the date value to write on each appended row, in "YYYY-MM-DD" format. */
  appendWithDate(records, date_field_name, date_string){
    for (let record of records)
      record[date_field_name] = date_string;
    this.appendRows(records);
  };

  /** Clear the sheet then write the supplied records along with a header.
   * @param {!Array<!Object>} records: expected format: `[{k: v, ...}, ...]`. */
  overwrite(records){
    this._sheet.clear();
    this.appendRows(records);
  };

  /** @param {!Array<!Object>} data Rows to append. */
  appendRows(records){
    // Check the header, and add any missing fields to the sheet's header.
    this.upsertHeader(utils.getHeaderFromRecords(records));
    let header = this.getHeader();
    // Create an array of row arrays, that contain values in the same order as the header.
    let data = [];
    for (let record of records)
      data.push(header.map(f => record[f]));
    const start_row = this._sheet.getLastRow() + 1;
    let range = this._sheet.getRange(start_row, 1, data.length, data[0].length);
    range.setValues(data);
  }

  /** @param {string} field_name The name of the column to search.
   * @param {string} value Find rows where `field_name === value`.
   * @return {Array<number>} Array of 1-indexed row_ids where the value is found. */
  findRowsWhere(field_name, value){
    // Return early if the column with the header `field_name` doesn't exist.
    const field_id = this.getColumnId(field_name);
    if(field_id === -1) return [];

    const text_finder = this._sheet.createTextFinder(value).matchEntireCell(true);
    const cells = text_finder.findAll();
    return cells.filter(c => c.getColumn() === field_id).map(c => c.getRow());
  }

  /** @return {!Array<!Object>} Array of row objects with header as keys. */
  getSheetRows(){
    let header = this.getHeader();
    let rows = this._sheet.getDataRange().getValues().slice(1);  // Omit header.
    let result = rows.map(function(row){
      return Object.fromEntries(header.map((k, i) => [k, row[i]]));
    });
    return result;
  }

  /** @param {string} field_name The name of the field to filter by (on row 1).
   * @param {string} value Keep only rows where field_name === value.
   * @param {string} group_by_key Group by values or this field.
   * @return {!Object} An object whose keys are <group_by_key> values, and whose
   * values are the number of rows where field_name === value for each
   * `group_by_key` value. */
  countRowsWhere(field_name, values, group_by_key){
    const rows = this.getSheetRows().filter(r => values.includes(r[field_name]));
    let counts = {}
    rows.forEach(r => counts[r[group_by_key]] = (counts[r[group_by_key]] || 0) + 1);
    return counts;
  }

  /**
   * Retrieve a set of keys from `key_column_id`, in the `n_rows` rows starting
   * from `start_row`, then return a vertical array with the results of looking
   * up these keys in `values`.
   * @param {number} start_row 1-indexed row number to start reading keys from.
   * @param {number} n_rows Number of rows to read.
   * @param {number} key_column_id 1-indexed column number to read the keys in.
   * @param {!Object} values Object to read the return values from. Its keys
   * should exist in column `key_column_id`.
   * @param {any=} default_value Return this value for any rows whose key is
   * absent from `values`.
   * @return {!Array<!Array<string>>} A vertical array (array of single-value arrays)
   */
  lookupValues(start_row, n_rows, key_column_id, values, default_value = 0){
    let result = []
    for (let i of [...Array(n_rows).keys()]){
      const row_id = i + start_row;
      const key = this._sheet.getRange(row_id, key_column_id).getValue();
      const value = (values[key] === undefined)? default_value : values[key]
      result.push([value]);
    }
    return result;
  }
}


/** A class to represent a Backup worksheet, subclass of Sheet.  */
class BackupSheet extends Sheet {
  /**
   * @param {string} sheet_name The name of the Sheet storing backup data.
   * @param {!Array<string>}fields_to_backup The set of fields to back up.
   * @param {string=} backup_date_field_name A column in the sheet which
   * contains the backup date in YYYY-MM-DD format. */
  constructor(sheet_name,
              fields_to_backup,
              backup_date_field_name = "Backup Date") {
      super(sheet_name);
      this._FIELDS_TO_BACKUP = fields_to_backup;
      this._BACKUP_DATE_FIELD_NAME = backup_date_field_name;
  }

  /**
   * @param {number} n_days Delete any entries strictly older than `n_days`.
   */
  deleteBackupsOlderThanNDays(n_days){
    const today = new Date();
    let from_date = new Date();
    from_date.setDate(today.getDate() - n_days);
    from_date = from_date.toISOString().slice(0, 10);

    let backup_column_id = this.getColumnId(this._BACKUP_DATE_FIELD_NAME);
    let num_rows = this._sheet.getLastRow() - 1;
    let date_range = this._sheet.getRange(2, backup_column_id, num_rows);
    let unique_dates = [...new Set(date_range.getValues().map(r => r[0].toISOString().slice(0, 10)))];
    // Note: If Date.parse fails to parse a one of the values, it will return
    // NaN, which in turn will be excluded from the filter operation. This
    // allows us to conservatively keep any unexpected rows.
    let filtered = unique_dates.filter(e => e < from_date);
    for (let value of filtered)
      this.deleteRowsWhere(this._BACKUP_DATE_FIELD_NAME, value);
  }

  /**
   * @param {!Array<!Array<string>>} data Array of rows to append to the sheet.
   * @param {boolean=} overwrite_today Before writing the data, start by deleting
   * any rows with today's date. This feature is used for idempotence, to avoid
   * any duplication should this function be run multiple times on a given day.
   */
  backupData(data, overwrite_today = true){
    const today = new Date().toISOString().slice(0, 10);

    if (overwrite_today)
      this.deleteRowsWhere(this._BACKUP_DATE_FIELD_NAME, today);

    let filtered_data = [];
    for (let record of data){
      const new_record = {};
      for (let f of this._FIELDS_TO_BACKUP)
        new_record[f] = record[f];
      filtered_data.push(new_record);
    }

    // Write the data to the sheet along with today's date
    this.appendWithDate(filtered_data, this._BACKUP_DATE_FIELD_NAME, today);
  };
}


/** Batch consecutive integers in the supplied array.
 * @param {number} integers: A list of integers.
 * @return {!Array<!Array<number>>}: An array of 2-element arrays. In each nested array,
 * The first element is the first_integer in a sequence of consecutive integers, and
 * the second element is the sequence length. */
function batchConsecutiveIntegers(integers){
  let stencil = [true, ...integers.slice(1).map((v, i) => v !== integers[i] + 1)];
  let keys = [...Array(integers.length).keys()].filter((_v, i) => stencil[i]);
  let row_counts = [...keys, integers.length].slice(1).map((v, i) => v - keys[i]);
  const result = keys.map((v, i) => [integers[v], row_counts[i]]);
  return result;
}

/** Find the column number of `value` relative to the range's leftmost column.
 * @param {!Object} range An Apps Script range.
 * @param {string} value The value to look for.
 * @return {number} The column ID. */
function findColumn(range, value){
  const result_range = range.createTextFinder(value).matchEntireCell(true).findNext();
  if (result_range !== null)
    return result_range.getColumn();
}


if (typeof require !== "undefined")
  module.exports = {
    Sheet,
    batchConsecutiveIntegers
  };


/**
 * @fileoverview a class and utility functions to manipulate a Google Sheet.
 */

class Sheet{
  /** @param {string} sheet_name The name of the sheet. */
  constructor(sheet_name){
    this._sheet_name = sheet_name;
    // Select the sheet by name, or create it if it doesn't exist yet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this._sheet = (ss.getSheetByName(sheet_name) === null) ? ss.insertSheet(sheet_name): ss.getSheetByName(sheet_name);
  }

  /** @return {!Range} the first row of the sheet, as a range. */
  getHeader(){
    let [top_row, left_column, num_rows] = [1, 1, 1];
    return this._sheet.getRange(top_row, left_column, num_rows, this._sheet.getLastColumn());
  };

  /** Add fields to the sheet's header.
   * @param {!Array<string>} field_names: Compare this iterable with the current sheet header.
   * If any fields are missing from the header, create new columns to add them.
   * Important: This function intentionally does *not* delete any columns that are
   * present in the sheet but absent from `field_names`.
   * */
  addMissingHeadings(field_names){
    let last_column_id = this._sheet.getLastColumn();
    for (let f of field_names){
      // We call getHeader() at each iteration on the off-chance that `field_names` contains
      // duplicate fields, to avoid creating multiple columns with the same field name.
      const header = this.getHeader().getValues()[0];  // FIXME: getHeader() fails when the sheet is blank.
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
   * If multiple columns have the same heading, return the first one.
   * If the first row doesn't contain the supplied value, return `null`.
   * @param {string} field_name The column title (value in the first row).
   * @return {?number} The the number of the column starting with this value. */
  getColumnId(field_name) {
    let header = this.getHeader();
    return findColumn(header, field_name);
  };

  /** Delete any rows where the value for `field_name` equals `value`.
   * @param {string} field_name The column heading.
   * @param {string} value The value to search for. */
  deleteRowsWhere(field_name, value){
    // Return early if the column with the header `field_name` doesn't exist.
    const field_id = this.getColumnId(field_name);
    if(field_id === null) return;

    const text_finder = this._sheet.createTextFinder(value).matchEntireCell(true);
    let results = text_finder.findAll();

    // For performance, we batch the deletion of consecutive rows. This allows us to
    // delete 1000 rows by calling deleteRows() once instead deleteRow() 1000 times.

    // This next bit is slightly complicated, but we're just generating batches
    // of rows earmarked for deletion: `[[start_row, row_count], ...]`
    let rows_to_delete = results.filter(cell => cell.getColumn() === field_id);
    let row_ids = rows_to_delete.map(cell => cell.getRow());
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
    let header = this.getHeader().getValues()[0];
    for (let record of records){
      record[date_field_name] = date_string;
      // Create an array with the values
      let new_row = header.map(f => record[f]);
      this._sheet.appendRow(new_row);
    }
  };

  /** Clear the sheet and write the supplied records along with a header.
   * @param {!Array<!Object>} records: expected format: `[{k: v, ...}, ...]`. */
  overwrite(records){
    this._sheet.clear();
    let header = getAirtableHeader(records);
    this._sheet.appendRow(header);
    for (let record of records){
      let new_row = header.map(f => record[f]);
      this._sheet.appendRow(new_row);
    }
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

module.exports = {
  Sheet,
  batchConsecutiveIntegers
};


/**
 * @fileoverview This file defines the MetricsSheet class.
 * It is kept separate from the other classes in sheet.js because it contains
 * business logic and is therefore specific to TTTR.
 */

/** A class to represent sheets where periodic snapshots of fields and metrics
 * of interest are stored. */
class MetricsSheet extends BackupSheet {
  /** Add computed fields to the sheet. This function contains business logic.
   * @param {Date} backup_date The backup date to compute these metrics for. */
  addComputedFields(backup_date){
    // TODO: Switch to a dynamic dispatch pattern if/when we add new metrics.
    if (this._sheet_name === "Opportunities"){
      const counts = this.countRowsWhere("Stage", "Project verification", "id");

      this.upsertHeader(["Days in Verification"])
      const metric_column_id = this.getColumnId("Days in Verification");
      const key_column_id = this.getColumnId("id");

      const row_ids = this.findRowsWhere(this._BACKUP_DATE_FIELD_NAME, backup_date);
      const row_batches = batchConsecutiveIntegers(row_ids);
      for (let [start_row, n_rows] of row_batches){
        const values = this.lookupValues(start_row, n_rows, key_column_id, counts);
        this._sheet.getRange(start_row, metric_column_id, n_rows).setValues(values);
      }
    }
  }
}


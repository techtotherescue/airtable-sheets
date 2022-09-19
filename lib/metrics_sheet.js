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
      this.addComputedCountsField("Days waiting", "Stage", ["Project form submitted"], backup_date);
      this.addComputedCountsField("Days in Verification", "Stage", ["Project verification"], backup_date);
      const values = ["Project form submitted", "Project verification"];
      this.addComputedCountsField("Days waiting or in verification", "Stage", values, backup_date);
    }
    if (this._sheet_name === "Organizations"){
      this.addComputedCountsField("Days waiting for verification", "Stage", ["Waiting for verification"], backup_date);
    }
  }

  addComputedCountsField(computed_field_name, field_name, values, backup_date){
      const group_by_key = "id";
      const counts = this.countRowsWhere(field_name, values, group_by_key, backup_date);

      this.upsertHeader([computed_field_name])
      const metric_column_id = this.getColumnId(computed_field_name);
      const key_column_id = this.getColumnId(group_by_key);

      const row_ids = this.findRowsWhere(this._BACKUP_DATE_FIELD_NAME, backup_date);
      const row_batches = batchConsecutiveIntegers(row_ids);
      for (let [start_row, n_rows] of row_batches){
        const values = this.lookupValues(start_row, n_rows, key_column_id, counts);
        this._sheet.getRange(start_row, metric_column_id, n_rows).setValues(values);
      }
  }
}


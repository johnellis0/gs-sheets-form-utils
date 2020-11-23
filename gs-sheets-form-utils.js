/* Google Scripts Sheets Form Utilities
 * By John Ellis
 * https://github.com/johnellis0/gs-sheets-form-utils
 *
 * Provides utilities to merge form responses into another sheet, check & remove duplicates, automatically rotate
 * form submission sheet monthly/yearly
 */

/**
 * Moves range to sheet (destructive)
 * @param range
 * @param sheet
 */
function moveSubmissionToSheet(range, sheet){
    copySubmissionToSheet(range, sheet);
    range.getSheet().deleteRow(range.getRow());
}

/**
 * Copies range to sheet
 * @param range
 * @param sheet
 */
function copySubmissionToSheet(range, sheet){
    let destination = getFirstEmptyRange(sheet);
    range.copyTo(destination, {contentsOnly:true});
}

/**
 * Returns a 1x1 range at the beginning of the first empty row in the given Sheet
 * @param sheet
 * @returns {*}
 */
function getFirstEmptyRange(sheet){
    var first_empty_row = sheet.getLastRow() + 1;
    sheet.insertRowBefore(first_empty_row);

    return sheet.getRange(first_empty_row, 1);
}

/* Google Scripts Sheets Form Utilities
 * By John Ellis
 * https://github.com/johnellis0/gs-sheets-form-utils
 * Released under the MIT License
 *
 * Google Apps Script for Google Sheets that adds utility functions for working with Google Form submissions within
 * a Google Sheets file
 */

/**
 * Moves range to first empty row in sheet
 * @param {Range} range Range to move
 * @param {Sheet} sheet Sheet to insert range into
 * @param {boolean} digest Whether to add digest to destination range
 * @param {function} duplicateCallback Callback for if the range is detected as a duplicate. Will use digest duplicate
 * detection if `digest` is set to true. Callback will be called with the destination range as the first parameter.
 * @param {boolean} useLock Whether to use a Lock to prevent concurrent excecutions
 * @returns {Range} New range
 * @example
function onFormSubmit(e){ // Move all form submissions to sheet "Responses"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    moveToFirstEmptyRow(e.range, sheet);
}
 * @example
function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    moveToFirstEmptyRow(e.range, sheet, true, (range) => { // Move range to sheet "Responses" with duplicate callback
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
 */
function moveToFirstEmptyRow(range, sheet, digest=false, duplicateCallback=null, useLock=true){
    if(useLock){
        var lock = LockService.getScriptLock();
        lock.waitLock(300000); //Wait to get lock
    }

    let destination = copyToFirstEmptyRow(range, sheet, digest, duplicateCallback, false);
    range.getSheet().deleteRow(range.getRow());

    if(useLock)
        lock.releaseLock();

    return destination
}

/**
 * Copies range to first empty row in sheet
 * @param {Range} range Range to copy
 * @param {Sheet} sheet Sheet to insert range into
 * @param {boolean} digest Whether to add digest to destination range
 * @param {function} duplicateCallback Callback for if the range is detected as a duplicate. Will use digest duplicate
 * detection if `digest` is set to true. Callback will be called with the destination range as the first parameter.
 * @param {boolean} useLock Whether to use a Lock to prevent concurrent excecutions
 * @returns {Range} New range
 * @example
 function onFormSubmit(e){ // Copy all form submissions to sheet "Responses"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    copyToFirstEmptyRow(e.range, sheet);
}
 * @example
function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    copyToFirstEmptyRow(e.range, sheet, true, (range) => { // Copy range to sheet "Responses" with duplicate callback
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
 */
function copyToFirstEmptyRow(range, sheet, digest=false, duplicateCallback=null, useLock=true){
    if(useLock){
        var lock = LockService.getScriptLock();
        lock.waitLock(300000); //Wait to get lock
    }

    let duplicate = duplicateCallback !== null && isDuplicate(range, sheet, digest);
    let firstEmpty = getFirstEmptyRow(sheet);
    let destination = sheet.getRange(firstEmpty.getRow(), firstEmpty.getColumn(), 1, range.getNumColumns());

    range.copyTo(destination, {contentsOnly:true});

    if(digest)
        destination = addDigest(destination);
    if(duplicate)
        duplicateCallback(destination);

    if(useLock)
        lock.releaseLock();

    return destination;
}

/**
 * Returns current periodic sheet.
 *
 * Sheet will be created if it does not exist - a named template sheet can be supplied for this.
 *
 * @param {String} period Sheet period, values from: "month", "year"
 * @param {boolean} abbreviated Whether to use abbreviated names (eg. AUG / August)
 * @param {number} shift Time periods to shift by (+ or -)
 * @param {String} template Name of template sheet for sheet creation
 * @returns Sheet
 * @example
 // For example if the date were 01/01/2020 it would give the following sheet names:
 getPeriodicSheet("month"); // JAN20
 getPeriodicSheet("year"); // 2020
 getPeriodicSheet("month", false); // January 2020
 getPeriodicSheet("month", true, 1); // MAR20
 getPeriodicSheet("month", true, -1); // DEC19

 var templateName = "Template";
 getPeriodicSheet("month", true, 0, templateName); // Will make a copy of "Template" called 'JAN20'
 */
function getPeriodicSheet(period="month", abbreviated=true, shift=0, template=null){
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheetName;
    switch(period){
        case "month":
            sheetName = getMonthlySheetName(abbreviated, shift);
            break;
        case "year":
            sheetName = getYearlySheetName(shift);
            break;
        default:
            throw "Period not found: " + period;
    }

    var sheet = ss.getSheetByName(sheetName);

    if(!sheet){
        if(template){
            sheet = getNewSheetFromTemplate(template);
            sheet.setName(sheetName);
        }else{
            sheet = ss.insertSheet(sheetName);
        }
    }

    return sheet;
}

/**
 * Checks sheet for a duplicate of range
 * @param {Range} range Range to check for a duplicate of
 * @param {Sheet} sheet Sheet to check for duplicates
 * @param {boolean} useDigest Whether to use a digest column or if to check row values individually
 * @param {number} last Last row to check
 * @param {number} skip Columns to skip when calculating digest
 * @returns {boolean}
 */
function isDuplicate(range, sheet, useDigest=true, last=null, skip=1){
    last = last !== null ? last : sheet.getLastRow();

    Logger.log(last);

    if(last === 0)
        return false;

    if(useDigest){
        var digest = getDigest(range, skip);
        var col = range.getNumColumns() + 1;

        var data = sheet.getRange(1, col, last, 1).getValues();
        return data.flat().includes(digest);
    }else{
        // search all other rows
    }
}

/**
 * Appends cell to end of range containing digest of range values.
 *
 * Can be used to create a digest column and used with {@link isDuplicate} to check for duplicates more efficiently.
 *
 * @param {Range} range Range to add digest to
 * @param {number} skip Columns to skip from start of range (eg. to avoid timestamp)
 * @param {String} digest Calculated digest (will be calculated if not provided)
 * @returns {Range} Range with digest cell appended
 */
function addDigest(range, skip=1, digest=null){
    // Get digest
    digest = digest === null ? getDigest(range, skip) : digest;

    // Extend range to append a cell
    range = range.getSheet().getRange(range.getRow(), range.getColumn(), 1, range.getNumColumns() + 1);

    // Add digest to newly appended cell
    range.getCell(1, range.getNumColumns()).setValue(digest);

    return range; // Return extended range
}

/**
 *
 * @param sheetFrom
 * @param sheetTo
 * @param deleteFromSource
 */
function sweep(sheetFrom, sheetTo, deleteFromSource=true){
    if(deleteFromSource){
        // Get ranges & move them across (in reverse so deleting a row does not shift ranges below it)
        getRangesInUse(sheetFrom).reverse().forEach((range) => {
            moveToFirstEmptyRow(range, sheetTo);
        });
    }else{
        getRangesInUse(sheetFrom).forEach((range) => {
            copyToFirstEmptyRow(range, sheetTo);
        });
    }
}

/**
 * Returns range with specified amount of columns removed from the end
 * @param {Range} range Range to trim
 * @param {number} amount How many columns to remove from range
 * @returns {Range} Shortened range
 */
function trimRowRange(range, amount){
    return range.getSheet().getRange(range.getRow(), range.getColumn(), 1, range.getNumColumns() - amount);
}

/**
 *
 * @param {Sheet} sheet
 * @returns {Range}
 */
function getFirstEmptyRow(sheet){
    var first_empty_row = sheet.getLastRow() + 1;
    sheet.insertRowBefore(first_empty_row);

    return sheet.getRange(first_empty_row, 1);
}

/**
 * Returns full/abbreviated sheet name for current month (or shifted by +/- x months)
 *
 * See {@link getPeriodicSheet} for examples
 *
 * @param abbreviated - Whether to use abbreviated names
 * @param shift - Time periods to shift by (+/-)
 */
function getMonthlySheetName(abbreviated=true, shift=0){
    const monthNames = ["January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"];
    var date = new Date();

    var monthNum = date.getMonth() + shift;
    var year = date.getFullYear();

    if(monthNum < 0){
        monthNum = 11;
        year--;
    }else if(monthNum > 11){
        monthNum = 0;
        year++;
    }

    var month = monthNames[monthNum];

    if(abbreviated){
        return month.substr(0,3).toUpperCase() + year.toString().slice(-2);
    }else{
        return month + " " + year;
    }
}

/**
 * Return sheet name for current year (or shifted by +/- x years)
 *
 * @param shift
 * @returns {number}
 */
function getYearlySheetName(shift=0){
    var date = new Date();

    return date.getFullYear() + shift;
}

/**
 * Returns copy of template sheet
 *
 * @param template
 * @returns {Sheet}
 */
function getNewSheetFromTemplate(template){
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    ss.getSheetByName(template).copyTo(ss);

    return ss.getSheetByName("Copy of "+template);
}


function getRangesInUse(sheet, excludeFrozen=true) {
    var frozenRows = excludeFrozen ? sheet.getFrozenRows() : 0;

    var rowsAmt = sheet.getLastRow() - frozenRows;
    var maxCols = sheet.getLastColumn();

    var ranges = [];

    for(var i=0; i<rowsAmt; i++){
        var row = i + 1 + frozenRows; // Add 1 as Rows start at 1 and offset by frozen rows
        var range = sheet.getRange(row, 1, 1, maxCols);

        if(range.getValues()[0].some(val => val !== ""))
            ranges.push(range); // If the range contains at least one value add to array
    }

    return ranges;
}

/**
 * Checks if range is empty. Can ignore Checkbox cells (which are always filled).
 * @param {Range} range Range to check
 * @param {boolean} ignoreCheckbox Whether to
 * @returns {boolean} Whether the range is empty or not
 */
function isRangeEmpty(range, ignoreCheckbox=true){
    if(!ignoreCheckbox)
        return range.isBlank();

    range.getValues()[0].forEach((val, i) => {
        if(val !== "" && range.getDataValidations()[0][i] !== SpreadsheetApp.DataValidationCriteria.CHECKBOX)
            return false;
    });
    return true;
}

/**
 * Calculate the digest of given range. Default settings skip the first cell (timestamp for form submissions) and use
 * SHA1 as the digest algorithm
 * @param {Range} range Range to calculate the digest of
 * @param {number} skip How many columns to skip
 * @param {Utilities.DigestAlgorithm} encoding Digest algorithm encoding to use
 * @returns {string} Calculated digest
 */
function getDigest(range, skip=1, encoding=Utilities.DigestAlgorithm.SHA_1){
    var values = range.getValues()[0].slice(skip);

    Logger.log(values);

    return Utilities.base64Encode(Utilities.computeDigest(encoding, values.join()));
}

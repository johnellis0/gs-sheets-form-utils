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
 * @param {boolean} ignoreCheckboxes Whether to ignore checkbox cells when determining if row is empty
 * @returns {Range} New range
 * @example
 function onFormSubmit(e){ // Move all form submissions to sheet "Responses"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.moveToFirstEmptyRow(e.range, sheet);
}
 * @example
 function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.moveToFirstEmptyRow(e.range, sheet, true, (range) => { // Move range to sheet "Responses" with duplicate callback
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
 */
function moveToFirstEmptyRow(range,
                             sheet,
                             digest=false,
                             duplicateCallback=null,
                             useLock=true,
                             ignoreCheckboxes=true){
    if(useLock){
        var lock = LockService.getScriptLock();
        lock.waitLock(300000); //Wait to get lock
    }

    let destination = copyToFirstEmptyRow(range, sheet, digest, duplicateCallback, false, ignoreCheckboxes);
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
 * @param {boolean} ignoreCheckboxes Whether to ignore checkbox cells when determining if row is empty
 * @returns {Range} New range
 * @example
 function onFormSubmit(e){ // Copy all form submissions to sheet "Responses"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.copyToFirstEmptyRow(e.range, sheet);
}
 * @example
 function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.copyToFirstEmptyRow(e.range, sheet, true, (range) => { // Copy range to sheet "Responses" with duplicate callback
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
 */
function copyToFirstEmptyRow(range,
                             sheet,
                             digest=false,
                             duplicateCallback=null,
                             useLock=true,
                             ignoreCheckboxes=true){
    if(useLock){
        var lock = LockService.getScriptLock();
        lock.waitLock(300000); //Wait to get lock
    }

    let duplicate = duplicateCallback !== null && isDuplicate(range, sheet, digest);
    let firstEmpty = getFirstEmptyRow(sheet, 1, ignoreCheckboxes);
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
 * @param {String} templateName Name of template sheet for sheet creation
 * @returns {Sheet}
 * @example
 // For example if the date were 01/01/2020 it would give the following sheet names:
 FormUtils.getPeriodicSheet("month"); // JAN20
 FormUtils.getPeriodicSheet("year"); // 2020
 FormUtils.getPeriodicSheet("month", false); // January 2020
 FormUtils.getPeriodicSheet("month", true, 1); // MAR20
 FormUtils.getPeriodicSheet("month", true, -1); // DEC19

 var templateName = "Template";
 FormUtils.getPeriodicSheet("month", true, 0, templateName); // Will make a copy of "Template" called 'JAN20'
 */
function getPeriodicSheet(period="month",
                          abbreviated=true,
                          shift=0,
                          templateName=null){
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
        if(templateName){
            sheet = getNewSheetFromTemplate(templateName);
            sheet.setName(sheetName);
        }else{
            sheet = ss.insertSheet(sheetName);
        }
    }

    return sheet;
}

/**
 * Checks sheet for to see if range is a duplicate of an existing row.
 *
 * Setting `useDigest` to true will treat the column after the last column in the range as a digest column, which can
 * be included by {@link addDigest}, allowing it to find duplicates more efficiently. It is recommended to use this mode
 * and just 'hide' the digest column on the spreadsheet.
 *
 * @param {Range} range Range to check for a duplicate of
 * @param {Sheet} sheet Sheet to check for duplicates
 * @param {boolean} useDigest Whether to use a digest column or if to check row values individually
 * @param {number} last Last row to check
 * @param {number} skip Columns to skip when calculating digest
 * @returns {boolean} Whether the range occurs on the sheet or not
 */
function isDuplicate(range,
                     sheet,
                     useDigest=true,
                     last=null,
                     skip=1){
    last = last !== null ? last : sheet.getLastRow();

    if(last === 0)
        return false;

    if(useDigest){
        var digest = getDigest(range, skip);
        var col = range.getNumColumns() + 1;

        var data = sheet.getRange(1, col, last, 1).getValues();
        return data.flat().includes(digest);
    }else{
        throw "Digest required for duplication check";
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

function addDateFormat(range, dateCol=1, time=false){
    range.getCell(1, dateCol).dateRange.setNumberFormat("dd/MM/yyyy"); //Set new format
    return range; // Return range for chaining
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
 * Returns range with specified amount of columns removed from the range
 * @param {Range} range Range to trim
 * @param {number} amount How many columns to remove from range. Positive values remove from the end, negative values
 * from the front
 * @returns {Range} Shortened range
 * @ignore
 */
function trimRowRange(range, amount){
    if(amount < 0){ // Remove from start
        return range.getSheet().getRange(range.getRow(), range.getColumn() + amount, 1, range.getNumColumns() + amount);
    }else{ // Remove from end
        return range.getSheet().getRange(range.getRow(), range.getColumn(), 1, range.getNumColumns() - amount);
    }
}

/**
 *
 * @param {Sheet} sheet
 * @param index
 * @param ignoreCheckboxes
 * @returns {Range}
 * @ignore
 */
function getFirstEmptyRow(sheet, index=1, ignoreCheckboxes=true){
    var row, max;

    if(ignoreCheckboxes){
        row = sheet.getLastRow() + 1; // Last row with data + 1
        max = sheet.getMaxRows(); // Max row
    }else{
        var cols = sheet.getMaxColumns();
        var getRow = (num) => sheet.getRange(num, 1, 1, cols);

        var min = sheet.getFrozenRows() + 1; // Row must come after the title rows
        max = sheet.getLastRow(); // Last row with data

        var range = sheet.getRange(1, index, max, 1);
        var data = range.getValues().flat();

        var i = max;
        do{ i--; }while(i>=min && data[i] === "");
        i--; // Subtract 1 to compensate for initial i++ in below loop
        do{ i++; }while(i<max && !isRangeEmpty(getRow(i+1), ignoreCheckboxes))

        row = i + 1; // Add 1 for index offset
    }

    if(row > max){
        // If all rows are filled then insert 1 at end
        sheet.insertRowAfter(max);
        return sheet.getRange(max+1, 1);
    }else{
        // Insert row before empty row to keep same amt of empty rows
        sheet.insertRowBefore(row);
        return sheet.getRange(row, 1);
    }
}

/**
 * Returns full/abbreviated sheet name for current month (or shifted by +/- x months)
 *
 * See {@link getPeriodicSheet} for examples
 *
 * @param abbreviated - Whether to use abbreviated names
 * @param shift - Time periods to shift by (+/-)
 * @ignore
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
 * @ignore
 *
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
 * @ignore
 */
function getNewSheetFromTemplate(template){
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    ss.getSheetByName(template).copyTo(ss);

    return ss.getSheetByName("Copy of "+template);
}

/**
 * Gets sheet named `sheetName`, or creates from template if it doesn't exist
 * @param sheetName
 * @param template
 * @returns {Sheet}
 */
function getOrCreateSheet(sheetName, template){
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = ss.getSheetByName(sheetName);

    return sheet ? sheet : getNewSheetFromTemplate(template);
}

/**
 *
 * @param sheet
 * @param excludeFrozen
 * @returns {Range[]}
 * @ignore
 */
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
 * @ignore
 */
function isRangeEmpty(range, ignoreCheckbox=true){
    if(!ignoreCheckbox)
        return range.isBlank();

    var values = range.getValues().flat();
    var validations = range.getDataValidations().flat();

    for(var i=0; i<values.length; i++){
        if(values[i] !== "" &&
            !(validations[i] !== null &&
                validations[i].getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX))
            return false;
    }

    return true;
}

/**
 * Calculate the digest of given range. Default settings skip the first cell (timestamp for form submissions) and use
 * SHA1 as the digest algorithm.
 * @param {Range} range Range to calculate the digest of
 * @param {number} skip How many columns to skip
 * @param {Utilities.DigestAlgorithm} encoding Digest algorithm encoding to use
 * @returns {string} Calculated digest
 */
function getDigest(range, skip=1, encoding=Utilities.DigestAlgorithm.SHA_1){
    var values = range.getValues()[0].slice(skip);

    return Utilities.base64Encode(Utilities.computeDigest(encoding, values.join()));
}

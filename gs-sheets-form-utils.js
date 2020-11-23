/* Google Scripts Sheets Form Utilities
 * By John Ellis
 * https://github.com/johnellis0/gs-sheets-form-utils
 *
 * Provides utilities to merge form responses into another sheet, check & remove duplicates, automatically rotate
 * form submission sheet monthly/yearly
 */

/**
 * Moves range to sheet (destructive)
 *
 * @param range
 * @param sheet
 */
function moveSubmissionToSheet(range, sheet){
    copySubmissionToSheet(range, sheet);
    range.getSheet().deleteRow(range.getRow());
}

/**
 * Copies range to sheet
 *
 * @param range
 * @param sheet
 */
function copySubmissionToSheet(range, sheet){
    let destination = getFirstEmptyRange(sheet);
    range.copyTo(destination, {contentsOnly:true});
}

/**
 * Returns a 1x1 range at the beginning of the first empty row in the given Sheet
 *
 * @param sheet
 * @returns {*}
 */
function getFirstEmptyRange(sheet){
    var first_empty_row = sheet.getLastRow() + 1;
    sheet.insertRowBefore(first_empty_row);

    return sheet.getRange(first_empty_row, 1);
}

/**
 * Returns current periodic sheet.
 * Sheet will be created if it does not exist - a named template shift can be supplied
 *
 * @param period - "month" / "year"
 * @param abbreviated - Whether to use abbreviated names
 * @param shift - Time periods to shift by (+/-)
 * @param template - Name of template sheet for sheet creation
 * @returns Sheet
 */
function getPeriodicSheet(period="month", abbreviated=true, shift=0, template=null){
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheetName = period == "month" ? getMonthlySheetName(abbreviated, shift) : getYearlySheetName(shift);
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
 * Returns full/abbreviated sheet name for current month (or shifted by +/- x months)
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
 * @returns {*}
 */
function getNewSheetFromTemplate(template){
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    ss.getSheetByName(template).copyTo(ss);

    return ss.getSheetByName("Copy of "+template);
}
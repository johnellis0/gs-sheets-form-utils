# Google Scripts Spreadsheet Form Utilities
Google Scripts utility library for using Google Forms with Google Sheets.

Features:
- Detect & remove/flag duplicate submissions
- Split form submissions into per-month / per-year sheets
- Move submissions to bottom of specific sheet, allowing both form submissions and manual entries to be ordered sequentially
- Automatically archive submissions
- Check if row is blank despite checkbox data validations

# Example usage

## Move form submissions onto monthly sheets

This will move the form submissions onto a different sheet each month.

```javascript
function onFormSubmit(e){
    var range = e.range;
    var sheet = getPeriodicSheet("month");

    range = moveToFirstEmptyRow(range, sheet);
}
```

## Highlight duplicated submissions

This will highlight any duplicate entries.

This uses a digest column which is appended to the end of the moved range, this column can be hidden on the sheet if you don't want it visible.

Using a digest column improves performance as only one column needs to be checked vs how many are in the range.

### Calculate duplicate yourself

This allows for more customization of actions to take if it is a duplicate; eg. not moving the range at all.

```javascript
function onFormSubmit(e){
    var range = e.range;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    var duplicate = isDuplicate(range, sheet);

    range = moveToFirstEmptyRow(range, sheet);
    range = addDigest(range);

    if(duplicate)
        range.setBackground("red");
}
```

### Using duplicate callback

`moveToFirstEmptyRow` and `copyToFirstEmptyRow` can take a callback as the 4th argument which will be ran if the range is detected as a duplicate. The 3rd argument is whether to use a digest column or not. More info and examples available in the API Reference below.

```javascript
function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    moveToFirstEmptyRow(e.range, sheet, true, (range) => { // Move range to sheet "Responses" with callback if duplicate
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
```

## Process form submissions in bulk

```javascript
function scheduledBulkProcess(){
    var submissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    var sheet = getPeriodicSheet("month");


}
```

# API Reference
## Functions

<dl>
<dt><a href="#moveToFirstEmptyRow">moveToFirstEmptyRow(range, sheet, digest, duplicateCallback)</a> ⇒ <code>Range</code></dt>
<dd><p>Moves range to first empty row in sheet</p>
</dd>
<dt><a href="#copyToFirstEmptyRow">copyToFirstEmptyRow(range, sheet, digest, duplicateCallback)</a> ⇒ <code>Range</code></dt>
<dd><p>Copies range to first empty row in sheet</p>
</dd>
<dt><a href="#trimRowRange">trimRowRange(range, amount)</a> ⇒ <code>Range</code></dt>
<dd><p>Returns range with specified amount of columns removed from the end</p>
</dd>
<dt><a href="#getFirstEmptyRow">getFirstEmptyRow(sheet)</a> ⇒ <code>Range</code></dt>
<dd></dd>
<dt><a href="#getPeriodicSheet">getPeriodicSheet(period, abbreviated, shift, template)</a> ⇒</dt>
<dd><p>Returns current periodic sheet.</p>
<p>Sheet will be created if it does not exist - a named template sheet can be supplied for this.</p>
</dd>
<dt><a href="#getMonthlySheetName">getMonthlySheetName(abbreviated, shift)</a></dt>
<dd><p>Returns full/abbreviated sheet name for current month (or shifted by +/- x months)</p>
<p>See <a href="#getPeriodicSheet">getPeriodicSheet</a> for examples</p>
</dd>
<dt><a href="#getYearlySheetName">getYearlySheetName(shift)</a> ⇒ <code>number</code></dt>
<dd><p>Return sheet name for current year (or shifted by +/- x years)</p>
</dd>
<dt><a href="#getNewSheetFromTemplate">getNewSheetFromTemplate(template)</a> ⇒ <code>Sheet</code></dt>
<dd><p>Returns copy of template sheet</p>
</dd>
<dt><a href="#isRangeEmpty">isRangeEmpty(range, ignoreCheckbox)</a> ⇒ <code>boolean</code></dt>
<dd><p>Checks if range is empty. Can ignore Checkbox cells (which are always filled).</p>
</dd>
<dt><a href="#sweep">sweep(sheetFrom, sheetTo, deleteFromSource)</a></dt>
<dd></dd>
<dt><a href="#getDigest">getDigest(range, skip, encoding)</a> ⇒ <code>string</code></dt>
<dd><p>Calculate the digest of given range. Default settings skip the first cell (timestamp for form submissions) and use
SHA1 as the digest algorithm</p>
</dd>
<dt><a href="#isDuplicate">isDuplicate(range, sheet, useDigest, last, skip)</a> ⇒ <code>boolean</code></dt>
<dd><p>Checks sheet for a duplicate of range</p>
</dd>
<dt><a href="#addDigest">addDigest(range, skip, digest)</a> ⇒ <code>Range</code></dt>
<dd><p>Appends cell to end of range containing digest of range values.</p>
<p>Can be used to create a digest column and used with <a href="#isDuplicate">isDuplicate</a> to check for duplicates more efficiently.</p>
</dd>
</dl>

<a name="moveToFirstEmptyRow"></a>

## moveToFirstEmptyRow(range, sheet, digest, duplicateCallback) ⇒ <code>Range</code>
Moves range to first empty row in sheet

**Kind**: global function  
**Returns**: <code>Range</code> - New range  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to move |
| sheet | <code>Sheet</code> |  | Sheet to insert range into |
| digest | <code>boolean</code> | <code>false</code> | Whether to add digest to destination range |
| duplicateCallback | <code>function</code> | <code></code> | Callback for if the range is detected as a duplicate. Will use digest duplicate detection if `digest` is set to true. Callback will be called with the destination range as the first parameter. |

**Example**  
```js
function onFormSubmit(e){ // Move all form submissions to sheet "Responses"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    moveToFirstEmptyRow(e.range, sheet);
}
```
**Example**  
```js
function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    moveToFirstEmptyRow(e.range, sheet, true, (range) => { // Move range to sheet "Responses" with duplicate callback
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
```
<a name="copyToFirstEmptyRow"></a>

## copyToFirstEmptyRow(range, sheet, digest, duplicateCallback) ⇒ <code>Range</code>
Copies range to first empty row in sheet

**Kind**: global function  
**Returns**: <code>Range</code> - New range  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to copy |
| sheet | <code>Sheet</code> |  | Sheet to insert range into |
| digest | <code>boolean</code> | <code>false</code> | Whether to add digest to destination range |
| duplicateCallback | <code>function</code> | <code></code> | Callback for if the range is detected as a duplicate. Will use digest duplicate detection if `digest` is set to true. Callback will be called with the destination range as the first parameter. |

**Example**  
```js
function onFormSubmit(e){ // Copy all form submissions to sheet "Responses"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    copyToFirstEmptyRow(e.range, sheet);
}
```
**Example**  
```js
function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    copyToFirstEmptyRow(e.range, sheet, true, (range) => { // Copy range to sheet "Responses" with duplicate callback
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
```
<a name="trimRowRange"></a>

## trimRowRange(range, amount) ⇒ <code>Range</code>
Returns range with specified amount of columns removed from the end

**Kind**: global function  
**Returns**: <code>Range</code> - Shortened range  

| Param | Type | Description |
| --- | --- | --- |
| range | <code>Range</code> | Range to trim |
| amount | <code>number</code> | How many columns to remove from range |

<a name="getFirstEmptyRow"></a>

## getFirstEmptyRow(sheet) ⇒ <code>Range</code>
**Kind**: global function  

| Param | Type |
| --- | --- |
| sheet | <code>Sheet</code> | 

<a name="getPeriodicSheet"></a>

## getPeriodicSheet(period, abbreviated, shift, template) ⇒
Returns current periodic sheet.

Sheet will be created if it does not exist - a named template sheet can be supplied for this.

**Kind**: global function  
**Returns**: Sheet  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| period | <code>String</code> | <code>month</code> | Sheet period, values from: "month", "year" |
| abbreviated | <code>boolean</code> | <code>true</code> | Whether to use abbreviated names (eg. AUG / August) |
| shift | <code>number</code> | <code>0</code> | Time periods to shift by (+ or -) |
| template | <code>String</code> | <code></code> | Name of template sheet for sheet creation |

**Example**  
```js
// For example if the date were 01/01/2020 it would give the following sheet names:
getPeriodicSheet("month"); // JAN20
getPeriodicSheet("year"); // 2020
getPeriodicSheet("month", false); // January 2020
getPeriodicSheet("month", true, 1); // MAR20
getPeriodicSheet("month", true, -1); // DEC19

var templateName = "Template";
getPeriodicSheet("month", true, 0, templateName); // Will make a copy of "Template" called 'JAN20'
```
<a name="getMonthlySheetName"></a>

## getMonthlySheetName(abbreviated, shift)
Returns full/abbreviated sheet name for current month (or shifted by +/- x months)

See [getPeriodicSheet](#getPeriodicSheet) for examples

**Kind**: global function  

| Param | Default | Description |
| --- | --- | --- |
| abbreviated | <code>true</code> | Whether to use abbreviated names |
| shift | <code>0</code> | Time periods to shift by (+/-) |

<a name="getYearlySheetName"></a>

## getYearlySheetName(shift) ⇒ <code>number</code>
Return sheet name for current year (or shifted by +/- x years)

**Kind**: global function  

| Param | Default |
| --- | --- |
| shift | <code>0</code> | 

<a name="getNewSheetFromTemplate"></a>

## getNewSheetFromTemplate(template) ⇒ <code>Sheet</code>
Returns copy of template sheet

**Kind**: global function  

| Param |
| --- |
| template | 

<a name="isRangeEmpty"></a>

## isRangeEmpty(range, ignoreCheckbox) ⇒ <code>boolean</code>
Checks if range is empty. Can ignore Checkbox cells (which are always filled).

**Kind**: global function  
**Returns**: <code>boolean</code> - Whether the range is empty or not  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to check |
| ignoreCheckbox | <code>boolean</code> | <code>true</code> | Whether to |

<a name="sweep"></a>

## sweep(sheetFrom, sheetTo, deleteFromSource)
**Kind**: global function  

| Param | Default |
| --- | --- |
| sheetFrom |  | 
| sheetTo |  | 
| deleteFromSource | <code>true</code> | 

<a name="getDigest"></a>

## getDigest(range, skip, encoding) ⇒ <code>string</code>
Calculate the digest of given range. Default settings skip the first cell (timestamp for form submissions) and use
SHA1 as the digest algorithm

**Kind**: global function  
**Returns**: <code>string</code> - Calculated digest  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to calculate the digest of |
| skip | <code>number</code> | <code>1</code> | How many columns to skip |
| encoding | <code>Utilities.DigestAlgorithm</code> |  | Digest algorithm encoding to use |

<a name="isDuplicate"></a>

## isDuplicate(range, sheet, useDigest, last, skip) ⇒ <code>boolean</code>
Checks sheet for a duplicate of range

**Kind**: global function  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to check for a duplicate of |
| sheet | <code>Sheet</code> |  | Sheet to check for duplicates |
| useDigest | <code>boolean</code> | <code>true</code> | Whether to use a digest column or if to check row values individually |
| last | <code>number</code> | <code></code> | Last row to check |
| skip | <code>number</code> | <code>1</code> | Columns to skip when calculating digest |

<a name="addDigest"></a>

## addDigest(range, skip, digest) ⇒ <code>Range</code>
Appends cell to end of range containing digest of range values.

Can be used to create a digest column and used with [isDuplicate](#isDuplicate) to check for duplicates more efficiently.

**Kind**: global function  
**Returns**: <code>Range</code> - Range with digest cell appended  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to add digest to |
| skip | <code>number</code> | <code>1</code> | Columns to skip from start of range (eg. to avoid timestamp) |
| digest | <code>String</code> | <code></code> | Calculated digest (will be calculated if not provided) |


#About
### Authors
- John Ellis - [GitHub](https://github.com/johnellis0) / [Portfolio](https://johnellis.dev)

### License
Released under [MIT](/LICENSE)

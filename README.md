# Google Scripts Spreadsheet Form Utilities
Google Scripts utility library for using Google Forms with Google Sheets.

Features:
- Detect & remove/flag duplicate submissions
- Split form submissions into per-month / per-year sheets
- Move submissions to bottom of specific sheet, allowing both form submissions and manual entries to be ordered sequentially
- Automatically archive submissions
- Check if row is blank despite checkbox data validations

# Install

## As library

To install as a library directly in Google Scripts follow the below instructions:

*Google Scripts Editor > Resources > Libraries > Enter **Script ID** in Add a library > Select latest version > Save*

**Script ID:** `1Rez6KOQDFg6RpI1sNXoaxSneLcwjXUT4eHTROuYcE5L9BuTs1D06pcbn`

![image](https://user-images.githubusercontent.com/34400721/101267807-fb6df500-3754-11eb-80da-c423aaf38c27.png)

If you install like this you will need to prefix all the methods by whatever is in the **Identifier** box, with the default being FormUtils.

```javascript
function onFormSubmit(e){
    var range = e.range;
    var sheet = FormUtils.getPeriodicSheet("month");

    range = FormUtils.moveToFirstEmptyRow(range, sheet);
}
```

## Manually

You can also copy the source file, `src/FormUtils.js` directly into your Scripts project.

If you install via this method you do not have to prefix the methods with a module name, example:

```javascript
function onFormSubmit(e){
    var range = e.range;
    var sheet = getPeriodicSheet("month");

    range = moveToFirstEmptyRow(range, sheet);
}
```

# Example usage

## Move form submissions onto monthly sheets

This will move the form submissions onto a different sheet each month.

```javascript
function onFormSubmit(e){
    var range = e.range;
    var sheet = FormUtils.getPeriodicSheet("month");

    range = FormUtils.moveToFirstEmptyRow(range, sheet);
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

    var duplicate = FormUtils.isDuplicate(range, sheet);

    range = FormUtils.moveToFirstEmptyRow(range, sheet);
    range = FormUtils.addDigest(range);

    if(duplicate)
        range.setBackground("red");
}
```

### Using duplicate callback

`moveToFirstEmptyRow` and `copyToFirstEmptyRow` can take a callback as the 4th argument which will be ran if the range is detected as a duplicate. The 3rd argument is whether to use a digest column or not. More info and examples available in the API Reference below.

```javascript
function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.moveToFirstEmptyRow(e.range, sheet, true, (range) => { // Move range to sheet "Responses" with callback if duplicate
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
```

## Process form submissions in bulk

```javascript
function scheduledBulkProcess(){
    var submissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    var sheet = FormUtils.getPeriodicSheet("month");


}
```

# API Reference
## Functions

<dl>
<dt><a href="#moveToFirstEmptyRow">moveToFirstEmptyRow(range, sheet, digest, duplicateCallback, useLock, ignoreCheckboxes)</a> ⇒ <code>Range</code></dt>
<dd><p>Moves range to first empty row in sheet</p>
</dd>
<dt><a href="#copyToFirstEmptyRow">copyToFirstEmptyRow(range, sheet, digest, duplicateCallback, useLock, ignoreCheckboxes)</a> ⇒ <code>Range</code></dt>
<dd><p>Copies range to first empty row in sheet</p>
</dd>
<dt><a href="#getPeriodicSheet">getPeriodicSheet(period, abbreviated, shift, templateName)</a> ⇒ <code>Sheet</code></dt>
<dd><p>Returns current periodic sheet.</p>
<p>Sheet will be created if it does not exist - a named template sheet can be supplied for this.</p>
</dd>
<dt><a href="#isDuplicate">isDuplicate(range, sheet, useDigest, last, skip)</a> ⇒ <code>boolean</code></dt>
<dd><p>Checks sheet for to see if range is a duplicate of an existing row.</p>
<p>Setting <code>useDigest</code> to true will treat the column after the last column in the range as a digest column, which can
be included by <a href="#addDigest">addDigest</a>, allowing it to find duplicates more efficiently. It is recommended to use this mode
and just &#39;hide&#39; the digest column on the spreadsheet.</p>
</dd>
<dt><a href="#addDigest">addDigest(range, skip, digest)</a> ⇒ <code>Range</code></dt>
<dd><p>Appends cell to end of range containing digest of range values.</p>
<p>Can be used to create a digest column and used with <a href="#isDuplicate">isDuplicate</a> to check for duplicates more efficiently.</p>
</dd>
<dt><a href="#sweep">sweep(sheetFrom, sheetTo, deleteFromSource)</a></dt>
<dd></dd>
<dt><a href="#getOrCreateSheet">getOrCreateSheet(sheetName, template)</a> ⇒ <code>Sheet</code></dt>
<dd><p>Gets sheet named <code>sheetName</code>, or creates from template if it doesn&#39;t exist</p>
</dd>
<dt><a href="#getDigest">getDigest(range, skip, encoding)</a> ⇒ <code>string</code></dt>
<dd><p>Calculate the digest of given range. Default settings skip the first cell (timestamp for form submissions) and use
SHA1 as the digest algorithm.</p>
</dd>
</dl>

<a name="moveToFirstEmptyRow"></a>

## moveToFirstEmptyRow(range, sheet, digest, duplicateCallback, useLock, ignoreCheckboxes) ⇒ <code>Range</code>
Moves range to first empty row in sheet

**Kind**: global function  
**Returns**: <code>Range</code> - New range  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to move |
| sheet | <code>Sheet</code> |  | Sheet to insert range into |
| digest | <code>boolean</code> | <code>false</code> | Whether to add digest to destination range |
| duplicateCallback | <code>function</code> | <code></code> | Callback for if the range is detected as a duplicate. Will use digest duplicate detection if `digest` is set to true. Callback will be called with the destination range as the first parameter. |
| useLock | <code>boolean</code> | <code>true</code> | Whether to use a Lock to prevent concurrent excecutions |
| ignoreCheckboxes | <code>boolean</code> | <code>true</code> | Whether to ignore checkbox cells when determining if row is empty |

**Example**  
```js
function onFormSubmit(e){ // Move all form submissions to sheet "Responses"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.moveToFirstEmptyRow(e.range, sheet);
}
```
**Example**  
```js
function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.moveToFirstEmptyRow(e.range, sheet, true, (range) => { // Move range to sheet "Responses" with duplicate callback
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
```
<a name="copyToFirstEmptyRow"></a>

## copyToFirstEmptyRow(range, sheet, digest, duplicateCallback, useLock, ignoreCheckboxes) ⇒ <code>Range</code>
Copies range to first empty row in sheet

**Kind**: global function  
**Returns**: <code>Range</code> - New range  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to copy |
| sheet | <code>Sheet</code> |  | Sheet to insert range into |
| digest | <code>boolean</code> | <code>false</code> | Whether to add digest to destination range |
| duplicateCallback | <code>function</code> | <code></code> | Callback for if the range is detected as a duplicate. Will use digest duplicate detection if `digest` is set to true. Callback will be called with the destination range as the first parameter. |
| useLock | <code>boolean</code> | <code>true</code> | Whether to use a Lock to prevent concurrent excecutions |
| ignoreCheckboxes | <code>boolean</code> | <code>true</code> | Whether to ignore checkbox cells when determining if row is empty |

**Example**  
```js
function onFormSubmit(e){ // Copy all form submissions to sheet "Responses"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.copyToFirstEmptyRow(e.range, sheet);
}
```
**Example**  
```js
function onFormSubmit(e){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

    FormUtils.copyToFirstEmptyRow(e.range, sheet, true, (range) => { // Copy range to sheet "Responses" with duplicate callback
        range.setBackground("red"); // Highlight moved range in red if it is a duplicate
    })
}
```
<a name="getPeriodicSheet"></a>

## getPeriodicSheet(period, abbreviated, shift, templateName) ⇒ <code>Sheet</code>
Returns current periodic sheet.

Sheet will be created if it does not exist - a named template sheet can be supplied for this.

**Kind**: global function  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| period | <code>String</code> | <code>month</code> | Sheet period, values from: "month", "year" |
| abbreviated | <code>boolean</code> | <code>true</code> | Whether to use abbreviated names (eg. AUG / August) |
| shift | <code>number</code> | <code>0</code> | Time periods to shift by (+ or -) |
| templateName | <code>String</code> | <code></code> | Name of template sheet for sheet creation |

**Example**  
```js
// For example if the date were 01/01/2020 it would give the following sheet names:
 FormUtils.getPeriodicSheet("month"); // JAN20
 FormUtils.getPeriodicSheet("year"); // 2020
 FormUtils.getPeriodicSheet("month", false); // January 2020
 FormUtils.getPeriodicSheet("month", true, 1); // MAR20
 FormUtils.getPeriodicSheet("month", true, -1); // DEC19

 var templateName = "Template";
 FormUtils.getPeriodicSheet("month", true, 0, templateName); // Will make a copy of "Template" called 'JAN20'
```
<a name="isDuplicate"></a>

## isDuplicate(range, sheet, useDigest, last, skip) ⇒ <code>boolean</code>
Checks sheet for to see if range is a duplicate of an existing row.

Setting `useDigest` to true will treat the column after the last column in the range as a digest column, which can
be included by [addDigest](#addDigest), allowing it to find duplicates more efficiently. It is recommended to use this mode
and just 'hide' the digest column on the spreadsheet.

**Kind**: global function  
**Returns**: <code>boolean</code> - Whether the range occurs on the sheet or not  

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

<a name="sweep"></a>

## sweep(sheetFrom, sheetTo, deleteFromSource)
**Kind**: global function  

| Param | Default |
| --- | --- |
| sheetFrom |  | 
| sheetTo |  | 
| deleteFromSource | <code>true</code> | 

<a name="getOrCreateSheet"></a>

## getOrCreateSheet(sheetName, template) ⇒ <code>Sheet</code>
Gets sheet named `sheetName`, or creates from template if it doesn't exist

**Kind**: global function  

| Param |
| --- |
| sheetName | 
| template | 

<a name="getDigest"></a>

## getDigest(range, skip, encoding) ⇒ <code>string</code>
Calculate the digest of given range. Default settings skip the first cell (timestamp for form submissions) and use
SHA1 as the digest algorithm.

**Kind**: global function  
**Returns**: <code>string</code> - Calculated digest  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| range | <code>Range</code> |  | Range to calculate the digest of |
| skip | <code>number</code> | <code>1</code> | How many columns to skip |
| encoding | <code>Utilities.DigestAlgorithm</code> |  | Digest algorithm encoding to use |


#About
### Authors
- John Ellis - [GitHub](https://github.com/johnellis0) / [Portfolio](https://johnellis.dev)

### License
Released under [MIT](/LICENSE)

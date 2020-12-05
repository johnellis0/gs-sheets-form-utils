# Google Scripts Spreadsheet Form Utilities
Google Apps Script utilities for using Google Forms with Google Sheets

# API Reference
## Functions

<dl>
<dt><a href="#moveSubmissionToSheet">moveSubmissionToSheet(range, sheet)</a></dt>
<dd><p>Moves range to sheet (destructive)</p>
</dd>
<dt><a href="#copySubmissionToSheet">copySubmissionToSheet(range, sheet)</a></dt>
<dd><p>Copies range to sheet</p>
</dd>
<dt><a href="#trimRowRange">trimRowRange(range, amount)</a></dt>
<dd><p>Returns range with end values omitted.
Useful to remove unnecessary form fields</p>
</dd>
<dt><a href="#getFirstEmptyRange">getFirstEmptyRange(sheet)</a> ⇒ <code>*</code></dt>
<dd><p>Returns a 1x1 range at the beginning of the first empty row in the given Sheet</p>
</dd>
<dt><a href="#getPeriodicSheet">getPeriodicSheet(period, abbreviated, shift, template)</a> ⇒</dt>
<dd><p>Returns current periodic sheet.
Sheet will be created if it does not exist - a named template shift can be supplied</p>
</dd>
<dt><a href="#getMonthlySheetName">getMonthlySheetName(abbreviated, shift)</a></dt>
<dd><p>Returns full/abbreviated sheet name for current month (or shifted by +/- x months)</p>
</dd>
<dt><a href="#getYearlySheetName">getYearlySheetName(shift)</a> ⇒ <code>number</code></dt>
<dd><p>Return sheet name for current year (or shifted by +/- x years)</p>
</dd>
<dt><a href="#getNewSheetFromTemplate">getNewSheetFromTemplate(template)</a> ⇒ <code>Sheet</code></dt>
<dd><p>Returns copy of template sheet</p>
</dd>
<dt><a href="#sweep">sweep(sheetFrom, sheetTo, deleteFromSource)</a></dt>
<dd></dd>
</dl>

<a name="moveSubmissionToSheet"></a>

## moveSubmissionToSheet(range, sheet)
Moves range to sheet (destructive)

**Kind**: global function  

| Param |
| --- |
| range | 
| sheet | 

<a name="copySubmissionToSheet"></a>

## copySubmissionToSheet(range, sheet)
Copies range to sheet

**Kind**: global function  

| Param |
| --- |
| range | 
| sheet | 

<a name="trimRowRange"></a>

## trimRowRange(range, amount)
Returns range with end values omitted.
Useful to remove unnecessary form fields

**Kind**: global function  

| Param | Description |
| --- | --- |
| range |  |
| amount | Amount of values to omit from end of range. |

<a name="getFirstEmptyRange"></a>

## getFirstEmptyRange(sheet) ⇒ <code>\*</code>
Returns a 1x1 range at the beginning of the first empty row in the given Sheet

**Kind**: global function  

| Param |
| --- |
| sheet | 

<a name="getPeriodicSheet"></a>

## getPeriodicSheet(period, abbreviated, shift, template) ⇒
Returns current periodic sheet.
Sheet will be created if it does not exist - a named template shift can be supplied

**Kind**: global function  
**Returns**: Sheet  

| Param | Default | Description |
| --- | --- | --- |
| period | <code>month</code> | "month" / "year" |
| abbreviated | <code>true</code> | Whether to use abbreviated names |
| shift | <code>0</code> | Time periods to shift by (+/-) |
| template | <code></code> | Name of template sheet for sheet creation |

<a name="getMonthlySheetName"></a>

## getMonthlySheetName(abbreviated, shift)
Returns full/abbreviated sheet name for current month (or shifted by +/- x months)

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

<a name="sweep"></a>

## sweep(sheetFrom, sheetTo, deleteFromSource)
**Kind**: global function  

| Param | Default |
| --- | --- |
| sheetFrom |  | 
| sheetTo |  | 
| deleteFromSource | <code>true</code> | 


#About
### Authors
- John Ellis - [GitHub](https://github.com/johnellis0) / [Portfolio](https://johnellis.dev)

### License
Released under [MIT](/LICENSE)

# XLSXImporter
Alternative utility class used to import data from XLSX files without using Data Sources.

# Manual
This class enables the import of an Excel file into a selected table, giving you complete control over each stage of the process.

## Operation Result Structure

The result of the imports are returned in an araay of object with the following keys:
  -  **row**: Number of row in the Excel file
  -  **code**: Internal code representing the result of parsing for this row
  -  **message**: A human readable message representing the result of parsing for this row
  -  **target**: SysID of the created or updated record (if created/updated)
  -  **error**: Unexpected error occourred while parsing the row (if an error occours)

## Fields Mapping
The file's columns are initially mapped to the fields in the table that share the same label.

However, you may use the mapping method to assign a specific header to a field.

```javascript
importer.map("% Gen", "u_january"); // Header "% Gen" goes into field "u_january"
```

## Transform Methods
The transform method lets you specify a custom function to transform a given value for a particular field.

Multiple functions can be created and added to process the value of a given field.

```javascript
var stringToBool = function(value) { return value === "YES"; }
var negateBool = function (value) { return !value; }
importer.transform("u_frozen", stringToBool); // Transform cell values with "YES" into boolean true
importer.transform("u_frozen", negateBool); // Convert true to false and viceversa
```

## Coalescing Fields
The coalesce method can define one or more fields as coalescing fields to determine whether to create or update a preexisting record.

```javascript
importer.coalesce("u_name");
importer.coalesce("u_last_name");
```

## Require or ignore a header
It's possible to set as mandatory or ignore a specific header.

```javascript
importer.ignore("User Name"); // Will skip the header if found
importer.require("User ID")); // Will skip the row if missing
```

## Validation Methods
For every field one or more validation methods could be defined:\
Said methods could accept only two parameters, the row data and the name of the current field evaluated, and must return a boolean value (true when the row is valid, false otherwise).

They will be executed in a FIFO stack and if even one of the methods fail, the entire row will be skipped.

```javascript
var isEmpty(row, field) { return !gs.nil(row[field]); }
importer.validate("u_description", isEmpty) // Will skip the row if the "u_description" field in the row is empty
```

## Event Callbacks
Multiple callbacks are available while parsing a row, every callback accept a signle parameter "data" that could contains:
 -  **row**: Contains the header names as keys and the cell values as values
 -  **index**: Number of the current row
 -  **sys_id**: SysID of the created record

As follows the available callbacks with the possible parameters:
- **onRowRead** (Parameters: row, index): 
  -  This is executed after the single row is read from the file but transformation functions are not yet performed.
  -  The "row" parameter contains the header names as keys and the cell values as values.
  -  By returning a boolean value of "false", the current row will be skipped and not inserted.
- **onRowValidating** (Parameters: row, index): 
  -  This is executed after the mapping but before the single row is validated.
  -  The "row" parameter contains the field names as keys and the cell values as values.
  -  By returning a boolean value of "false", the current row will be skipped and not inserted.
- **onRowValidated** (Parameters: row, index): 
  -  This is executed after the single row is successfully validated.
  -  The "row" parameter contains the field names as keys and the cell values as values.
  -  By returning a boolean value of "false", the current row will be skipped and not inserted.
- **onRowTransformed** (Parameters: row, index):
  -  This is executed after the transformation and mapping functions are executed.
  -  The "row" parameter contains the field names as keys and the data to be inserted as values.
  -  By returning a boolean value of "false", the current row will be skipped and not inserted.
- **onRowImported** (Parameters: sys_id, row, index) 
  -  This is executed after creating or updating a record.
  -  The "sys_id" parameter represents the SysID of the created/updated record.
  -  The "row" parameter contains the field names as keys and the data to be inserted as values.
  -  The data parameter also uses final field names as keys, with values representing the data to be inserted.
  -  No return value would be kept in consideration 

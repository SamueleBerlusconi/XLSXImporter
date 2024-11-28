# XLSXImporter

A utility class designed to import data from XLSX files into ServiceNow tables, bypassing the use of Data Sources. Offers granular control over the import process, enabling advanced mapping, transformation, validation, and event handling.

# Overview

The XLSXImporter class provides a highly customizable way to import Excel files into a target table. It allows full control over data mapping, transformation, and validation, ensuring flexibility and precision throughout the process.

## Fields Mapping

Columns in the file are automatically mapped to fields in the table with matching labels.

However, you may use the mapping method to assign a specific header to a field.

```javascript
var importer = new XLSXImporter();
importer.map("% Gen", "u_january"); // Header "% Gen" goes into field "u_january"
```

## Transform Methods

The transform method lets you specify a custom function to transform a given value for a particular field.

Multiple functions can be created and added to process the value of a given field and will be processed sequentially.

```javascript
var stringToBool = function(value) { return value === "YES"; }
var negateBool = function (value) { return !value; }

var importer = new XLSXImporter();
importer.transform("u_frozen", stringToBool); // Transform cell values with "YES" into boolean true
importer.transform("u_frozen", negateBool); // Convert true to false and viceversa
```

## Coalescing Fields

The `coalesce` method allows you to define fields that determine whether to create a new record or update an existing one.

Specifically, if all the coalescing fields from the file row are the same of an existing record in the table, the data will be updated, otherwise a new record will be created.

```javascript
var importer = new XLSXImporter();
importer.coalesce("u_name");
importer.coalesce("u_last_name");
```

## Require or ignore a header

You can designate a header as mandatory or ignore it entirely during the import process:

A **mandatory** header triggers an error if it is missing from the file. The import process will fail, returning the code `XLSXImporter.STATES.MISSING_REQUIRED_HEADER`.

An **ignored** header ensures that the corresponding column is excluded from the import for all rows, while the remaining column values are processed as usual.

A mandatory header cannot be ignored and vice-versa.

```javascript
var importer = new XLSXImporter();
importer.ignore("User Name"); // The specific column will not be imported
importer.require("User ID"); // The entire import will fails if the column is missing
```

## Validation Methods

The `validate` method lets you define validation rules for specific fields.\
Validation methods accept two parameters: `row` (the current row's data) and `field` (the field being validated).

These methods must return a boolean value, with `false` skipping the row.

Validation methods are executed in a FIFO order, and failure of any method will prevent the row from being processed further.

```javascript
var isEmpty(row, field) { return !gs.nil(row[field]); }

var importer = new XLSXImporter();
importer.validate("u_description", isEmpty) // Skip the row if the "u_description" field is empty
```

## Event Callbacks

Multiple events are available while parsing a row, every callback accept a single parameter `data` that could contains:

| Key      | Type     | Value                                                                       |
|:---------|:---------|:----------------------------------------------------------------------------|
| `row`    | `Object` | The current row's data, with header names as keys and cell values as values |
| `index`  | `Number` | The row index in the file                                                   |
| `sys_id` | `String` | The SysID of the created or updated record (if applicable)                  |

The following events are available:

| Event              | Execution                                                             | Parameters                    | Return value                                             |
|:-------------------|:----------------------------------------------------------------------|:------------------------------|:---------------------------------------------------------|
| `onRowRead`        | Executed after reading a row from file but before any transformations | `row` (header names), `index` | Return `false` to skip the row                           |
| `onRowValidating`  | Executed after mapping but before row validation                      | `row` (field names), `index`  | Return `false` to skip the row                           |
| `onRowValidated`   | Executed after successful row validation                              | `row` (field names), `index`  | Return `false` to skip the row                           |
| `onRowTransformed` | Executed after applying transformations and mappings functions        | `row` (field names), `index`  | Return `false` to skip the row                           |
| `onRowImported`    | Executed after the record is created or updated                       | `row` (field names), `index`, `sys_id` |This callback is informational; its return value is ignored|

## Result Structures

The result of the import operation is returned as an object with the following structure:
| Key       | Type      | Value                                                                                   |
|:----------|:----------|:----------------------------------------------------------------------------------------|
| `success` | `Boolean` | Success result of the import                                                            |
| `code`    | `String`  | Specific `XLSXImporter.STATES` result code of the operation                             |
| `message` | `String`  | Message related to the current operation                                                |
| `rows`    | `Number`  | Number of processed rows                                                                |
| `elapsed` | `Number`  | Time elapsed for the import process in milliseconds                                     |
| `data`    | `Object`  | Optional data object, will contains an array of row results if the import is successful |

Every row parsed will also create a result object, structured as follows:
| Key       | Type     | Value                                                                            |
|:----------|:---------|:---------------------------------------------------------------------------------|
| `row`     | `Number` | Number of row in the Excel file                                                  |
| `code`    | `Number` | `XLSXImporter.RCODES` code representing the result of parsing for this row       |
| `message` | `String` | A human readable message representing the result of parsing for this row         |
| `target`  | `String` | SysID of the created or updated record (if created/updated) or target field name |
| `error`   | `Error`  | Unexpected error occourred while parsing the row (if an error occours)           |

### Response Codes

The `XLSXImporter.STATES` object contains codes that are used as response codes for the entire import operation:

| Code                      | Value                     | Description                                                                 |
|:--------------------------|:--------------------------|:----------------------------------------------------------------------------|
| `SUCCESS`                 | `success`                 | Returned when the import is correctly terminated                            |
| `PARSING_ERROR`           | `parsing_error`           | Returned when the `sn_impex.GlideExcelParser` is not correctly instantiated |
| `MISSING_REQUIRED_HEADER` | `missing_required_header` | Returned when a required header is missing in the XLSX file                 |

The `XLSXImporter.RCODES` object contains codes that are used as response codes for the single row parsing:

| Code                      | Value | Description                                                                                                                                      |
|:--------------------------|:------|:-------------------------------------------------------------------------------------------------------------------------------------------------|
| `SUCCESS`                 | `0`   | Returned when a row is correclty parsed                                                                                                          |
| `SKIPPED_EMPTY`           | `1`   | Returned when the entire row is empty                                                                                                            |
| `SKIPPED_EVENT`           | `2`   | Returned when the following events return a `false` value: `onRowRead`, `onRowValidating`, `onRowValidated`, `onRowTransformed`, `onRowImported` |
| `SKIPPED_VALIDATION`      | `3`   | Returned when a row validation fails                                                                                                             |
| `ERROR`                   | `4`   | Returned when an unhandled error occours while parsing a row                                                                                     |

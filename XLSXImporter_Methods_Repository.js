/**
 * Repository containing some common transformation and validation methods to use with the XLSX Importer class.
 *
 * @author Samuele Berlusconi (GitHub: @SamueleBerlusconi)
 * @license Apache-2.0
 */
var XLSXImporter_Methods_Repository = Class.create();

XLSXImporter_Methods_Repository.TRANSFORMS = {};
XLSXImporter_Methods_Repository.VALIDATIONS = {};

/**
 * Transform a string value into a boolean one.
 *
 * Case insensitive, supports the following strings: Yes, Y, True, 1
 *
 * All non-positive strings will be treated as negative boolean values.
 */
XLSXImporter_Methods_Repository.TRANSFORMS.STRING_TO_BOOLEAN = function (value) {
	var _value = value.trim().toUpperCase();
	var positive = ["YES", "Y", "TRUE", "1"];
	return positive.includes(_value);
};

/**
 * Transform a string value into a float one.
 */
XLSXImporter_Methods_Repository.TRANSFORMS.STRING_TO_FLOAT = function (value) {
	// Remove empty spaces
	var _value = gs.nil(value) ? "" : value.trim();
	
	// Remove dots (thousand-separator) and replace comma (decimal separator) with dot
	_value = _value.replaceAll(".", ""); // 100.000.000,00 -> 100000000,00
	_value = _value.replaceAll(",", "."); // 100000000,00 -> 100000000.00
	
	// Convert and return the float value
	return parseFloat(_value);
};

/**
 * Transform a string value into a number with Excel formatting.
 */
XLSXImporter_Methods_Repository.TRANSFORMS.STRING_TO_EXCEL_NUMBER = function (value) {
	/**
	 * Regex that obtain only the last dot in a string.
	 */
	var DECIMAL_SEPARATOR_REGEX = /\.(?=[^.]*$)/;

	// Remove empty spaces
	var _value = gs.nil(value) ? "" : value.trim();

	// Replace the decimal separator with a comma (if exists) and remove dots (thousand-separator)
	_value = _value.replace(DECIMAL_SEPARATOR_REGEX, ","); // 100.000.000.00 -> 100.000.000,00
	_value = _value.replaceAll(".", ""); // 100.000.000,00 -> 100000000,00
	
	// Return string
	return _value;
};

/**
 * Transform a "Date" type cell with a numeric format into the string equivalent in the format "dd/MM/YYYY".
 * 
 * Excel define a date using the number of days elapsed since January 1, 1904.
 */
XLSXImporter_Methods_Repository.TRANSFORMS.EXCEL_DATE_TO_STRING = function (value) {
	// Parse the value into a number object
	var days = parseInt(value);

	/*
	 * Excel define the cells of type "Date" using the number of days since 01/01/1900 (for PC, 01/01/1904 for MAC)
	 */
	var target = new Date(Date.UTC(0, 0, days));

	// Convert the date into a string in the DD/MM/YYYY format
	return target.getDate() + "/" + "/" + target.getFullYear();
};

/**
 * Return true when the field is not empty.
 */
XLSXImporter_Methods_Repository.VALIDATIONS.IS_NOT_EMPTY = function (row, field) {
	// Extract the value
	var value = row[field];

	// Verify if the value is undefined, null or a empty string
	var isEmpty = gs.nil(value);
	
	return !isEmpty;
};

/**
 * Verify if the value is a boolean one.
 *
 * Case insensitive, supports the following values: Yes, Y, True, No, N, False, 1, 0
 */
XLSXImporter_Methods_Repository.VALIDATIONS.IS_BOOLEAN = function (row, field) {
	// Extract the value
	var value = row[field] || "";

	var allowed = ["YES", "Y", "TRUE", "1", "NO", "N", "FALSE", "0"];
	return allowed.includes(value.toUpperCase());
};

/**
 * Return true when the value has a reference in the binded table.
 * 
 * To use this methods in necessary to bind an object with the following values:
 * - table: The table where to search for the value
 * - field: The field in which looking for the value
 * - allowEmpty: True to pass the check even if the value is empty, false otherwise
 * 
 * Example:
 * var context = {};
 * context.table = "task";
 * context.field = "number";
 * context.allowEmpty = false;
 * XLSXImporter_Methods_Repository.VALIDATIONS.RECORD_EXISTS.bind(context);
 */
XLSXImporter_Methods_Repository.VALIDATIONS.RECORD_EXISTS = function (row, field) {
	// Extract the context binded to this method
	var context = this;
	
	// Extract the value
	var value = row[field];

	// Return a value based on the empty permission
	if (gs.nil(value)) return context.allowEmpty;

	// Get the value from the table
	var grRecord = new GlideRecord(context.table);
	var exists = grRecord.get(context.field, value);

	return exists;
};

/**
 * Return true when the value exists in the binded array.
 * 
 * The allowed values are case-insensitive and trimmed at the beginning and at the end.
 * 
 * To use this methods in necessary to bind an object with the following values:
 * - allowed: Array of values allowed for this field
 * - allowEmpty: True to pass the check even if the value is empty, false otherwise
 * 
 * Example:
 * var context = {};
 * context.allowed = ["open", "close", "edit"];
 * context.allowEmpty = false;
 * XLSXImporter_Methods_Repository.VALIDATIONS.VALUE_IN_LIST.bind(context);
 */
XLSXImporter_Methods_Repository.VALIDATIONS.VALUE_IN_LIST = function (row, field) {
	// Extract the context binded to this method
	var context = this;
	
	// Extract the value
	var value = row[field];

	// Return a value based on the empty permission
	if (gs.nil(value)) return context.allowEmpty;

	// Define a helper method
	var _normalize = function(value) { return value.trim().toUpperCase(); };

	// Parse the value
	value = _normalize(value);

	// Parse the allowed values
	var allowed = context.allowed.map(_normalize);

	return allowed.includes(value);
};

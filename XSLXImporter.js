var XLSXImporter = Class.create();

/**
 * Possible import operation results.
 */
XLSXImporter.STATES = {
    SUCCESS: "success",
    PARSING_ERROR: "parsing_error",
    MISSING_REQUIRED_HEADER: "missing_required_header",
};

/**
 * Possible codes to return for every row used to indicate the result of the parsing.
 */
XLSXImporter.RCODES = {
    SUCCESS: 0,
    SKIPPED_EMPTY: 1,
    SKIPPED_EVENT: 2,
    SKIPPED_VALIDATION: 3,
    ERROR: 4
};

/**
 * Import XLSX file without using Data Sources.
 *
 * @author Samuele Berlusconi (GitHub @SamueleBerlusconi)
 * @license Apache-2.0
 */
XLSXImporter.prototype = {
    /**
     * Initialize the class specifying the target table.
     *
     * @param {String} table Existing table where writing the new records.
     */
    initialize: function(table) {
        // Validate parameters
        if (gs.nil(table) || typeof table != "string") throw new Error("Invalid parameter: the 'table' parameter is empty or not a string");

        // Verify if the table effectively exists
        var _table = table.toLowerCase();
        if (!gs.tableExists(_table)) throw new Error("Invalid parameter: no table with name '" + _table + "' exists in the database");

        /**
         * List of allowed events for this class.
         */
        this.ALLOWED_EVENTS = ["onCoalesce", "onRowRead", "onRowValidating", "onRowValidated", "onRowTransformed", "onRowImported"];

        /**
         * Possible messages to return for every row used to indicate the result of the parsing.
         */
        this.MESSAGES = {};
        this.MESSAGES[XLSXImporter.RCODES.SUCCESS] = "Record imported correctly";
        this.MESSAGES[XLSXImporter.RCODES.SKIPPED_EMPTY] = "Record skipped because empty";
        this.MESSAGES[XLSXImporter.RCODES.SKIPPED_EVENT] = "Record skipped after result of event: ";
        this.MESSAGES[XLSXImporter.RCODES.SKIPPED_VALIDATION] = "Record skipped after validation failed: ";
        this.MESSAGES[XLSXImporter.RCODES.ERROR] = "Record skipped due to unexpected error";

        /**
         * Enabling the debug mode will allow to log in details what happens under the hood of the import process.
         */
        this._debug = false;
        /**
         * A virtual import will run the execute the import without writing information in the database, useful for validation-only imports.
         */
        this._virtual = false;
        /**
         * A solppy import will run the import skipping all the validation steps.
         * 
         * Note: Imported data could be damaged and/or invalid.
         */
        this._sloppy = false;
        /**
         * Target table for the imported records.
         */
        this.table = _table;
        /**
         * Object containing all the mapping from header to field name.
         *
         * Has as key the Excel headers and as values the fields' names.
         */
        this._mappings = {};
        /**
         * Object containing all the value trasformation functions.
         *
         * Has as key the fields' names and as value an array of transform functions.
         *
         * Every transform method must accept only one parameter and return only one transformed value.
         */
        this._transforms = {};
        /**
         * Object containing all the value validation functions.
         *
         * Has as key the fields' names and as value an array of object with the validation functions and related error messages.
         *
         * Every validation method must accept only two value (row value and field name) and return only one value (validation value).
         */
        this._validations = {};
        /**
         * Array containing all the fields used to verify if a record already exists in the table.
         */
        this._coalescing = [];
        /**
         * Array containing all the headers to ignore.
         */
        this._ignored = [];
        /**
         * Array containing all the required headers of the file.
         */
        this._required = [];
        /**
         * Object containing all the methods to execute on a specific event.
         *
         * Has as key the event names and as values callback to execute.
         */
        this._events = {};

        // Create the default mapping for this tale
        this._mapLabelWithName();

        this._trace("Class correctly initialized for table: " + this.table);
    },

    /* ################################ Start Public Methods ################################ */

    /**
     * Execute the import process from the given Excel file.
     *
     * @param {SysID} attachment_sys_id SysID of the source Excel file in the attachment table
     * @return {object} Result of the operation
     */
    import: function(attachment_sys_id) {
        // Save the start time of the process
        this._start = new Date();
        this._trace("Starting import process at " + this._start.toString());
        this._trace("Attachment SysID used for import: " + attachment_sys_id);

        // Notify the mode the import is running on
        if (this._virtual) this._trace("Import running in VIRTUAL mode: no data will be saved in the database");
        if (this._sloppy) this._trace("Import running in SLOPPY mode: no validation controls will be performed");

        // Get a stream from the attachment file
        var attachment = new GlideSysAttachment();
        var stream = attachment.getContentStream(attachment_sys_id);
        this._trace("File stream correctly initializated for attachment with SysID: " + attachment_sys_id);

        // Instantiate the parser class
        var parser = new sn_impex.GlideExcelParser();
        var success = parser.parse(stream);
        if (!success) return this._createReturnValue(XLSXImporter.STATES.PARSING_ERROR, parser.getErrorMessage());
        this._trace("GlideExcelParsed correctly initializated");

        // Verify if all the required headers are in the file
        if (!this._sloppy) {
            var result = this._validateHeaders(parser.getColumnHeaders());
            if (!result.success) return this._createReturnValue(XLSXImporter.STATES.MISSING_REQUIRED_HEADER, "One or more required headers are missing", result.missing);
            this._trace("Headers correctly validated");
        } else this._trace("Headers validation skipped because running in SLOPPY mode");

        /**
         * Index of the current row of the file.
         *
         * Starts from 2 because the first row (0) is the header.
         * 
         * Also, another 1 is added because is a zero-based index.
         */
        var index = 2;

        /**
         * List of results of import for every row in the input file.
         */
        var results = [];

        // Parse every row in the file
        while (parser.next()) {
            // Parse the current row and get the result
            var obj = this._parseRow(parser, index);
            this._trace("Correctly parsed row " + index + " with result: " + JSON.stringify(obj));

            // Memorize the row result
            results.push(obj);

            // Increment the row index
            index++;
        }

        // Close the connection to the input stream and release the document
        parser.close();
        this._trace("GlideExcelParsed correctly closed");

        this._trace("Ending import process at " + (new Date()).toString());
        return this._createReturnValue(XLSXImporter.STATES.SUCCESS, "Import completed successfully", results, results.length);
    },

    /**
     * Map an Excel header to a record field.
     *
     * @param {String} header Name of the header in the Excel file
     * @param {String} field Name of the field on the target import table to map the header to
     */
    map: function(header, field) {
        // Validate parameters
        if (gs.nil(header) || typeof header != "string") throw new Error("Invalid parameter: the 'header' parameter is empty or not a string");
        if (gs.nil(field) || typeof field != "string") throw new Error("Invalid parameter: the 'field' parameter is empty or not a string");

        // Verify if the "field" parameter exists on the table
        var _field = this._normalize(field);
        if (this._fields().indexOf(_field) === -1) throw new Error("Invalid parameter: no field with name '" + _field + "' exists on the table '" + this.table + "'");

        // Finally associate the header with the field
        var _header = this._normalize(header);
        this._mappings[_header] = _field;
        this._trace("Defined mapping function: " + _header + " -> " + _field);
    },

    /**
     * Add a transformation map for the specified field.
     *
     * @param {String} field Name of the field on the target import table on which the transformation must apply
     * @param {Function} f Transform method, accept only one value (cell value) and return only one value (transformed value)
     */
    transform: function(field, f) {
        // Validate parameters
        if (gs.nil(field) || typeof field != "string") throw new Error("Invalid parameter: the 'field' parameter is empty or not a string");
        if (gs.nil(f) || typeof f != "function") throw new Error("Invalid parameter: the 'f' parameter is empty or not a function");

        // Verify if the "field" parameter exists on the table
        var _field = this._normalize(field);
        if (this._fields().indexOf(_field) === -1) throw new Error("Invalid parameter: no field with name '" + _field + "' exists on the table '" + this.table + "'");

        // Create the transform array for the field if not already created
        if (gs.nil(this._transforms[_field])) this._transforms[_field] = [];

        // Finally push the transform function into the array
        this._transforms[_field].push(f);
        this._trace("Defined transform function for: " + _field + " (Methods Pipeline: " + this._transforms[_field].length + ")");
    },

    /**
     * Validate the record after it has been parsed.
     *
     * @param {String} field Name of the field on the target import table on which the validation must apply
     * @param {Function} f Validation method, accept only two value (row value and field name) and return only one value (validation value)
     * @param {String} message Error message to return in the row data if the validaton method fails.
     */
    validate: function(field, f, message) {
        // Validate parameters
        if (gs.nil(field) || typeof field != "string") throw new Error("Invalid parameter: the 'field' parameter is empty or not a string");
        if (gs.nil(f) || typeof f != "function") throw new Error("Invalid parameter: the 'f' parameter is empty or not a function");
        if (gs.nil(message) || typeof message != "string") throw new Error("Invalid parameter: the 'message' parameter is empty or not a string");

        // Verify if the "field" parameter exists on the table
        var _field = this._normalize(field);
        if (this._fields().indexOf(_field) === -1) throw new Error("Invalid parameter: no field with name '" + _field + "' exists on the table '" + this.table + "'");

        // Create the validation array for the field if not already created
        if (gs.nil(this._validations[_field])) this._validations[_field] = [];

        // Create an object containing the function and the related message
        var obj = {};
        obj.method = f;
        obj.message = message;

        // Finally push the validation function into the array
        this._validations[_field].push(obj);

        this._trace("Defined validation function for " + _field + " (Methods Pipeline: " + this._validations[_field].length + ")");
    },

    /**
     * Execute the specific function when the event is triggered.
     *
     * @param {String} event Event on which the function must be executed. Parameter differs from event to event.
     * @param {Function} f Callback method to execute
     */
    callback: function(event, f) {
        // Validate parameters
        if (gs.nil(event) || typeof event != "string") throw new Error("Invalid parameter: the 'event' parameter is empty or not a string");
        if (gs.nil(f) || typeof f != "function") throw new Error("Invalid parameter: the 'f' parameter is empty or not a function");

        // Verify if the "event" parameter is a valid one
        function lower(s) { return s.toLowerCase(); }
        var allowed = this.ALLOWED_EVENTS.map(lower);
        var _event = this._normalize(event);
        if (allowed.indexOf(_event) === -1) throw new Error("Invalid parameter: no event allowed with name '" + event + "'");

        // Set the callback for the event
        this._events[_event] = f;

        this._trace("Defined callback function for: " + _event + " (Methods Pipeline: " + this._events[_event].length + ")");
    },

    /**
     * Add a coalescing field to use as a key to determine if a record already exists in the table.
     *
     * @param {String} field Name of the field on the target import table to use as coalescing field
     */
    coalesce: function(field) {
        // Validate parameters
        if (gs.nil(field) || typeof field != "string") throw new Error("Invalid parameter: the 'field' parameter is empty or not a string");

        // Verify if the "field" parameter exists on the table
        var _field = this._normalize(field);
        if (this._fields().indexOf(_field) === -1) throw new Error("Invalid parameter: no field with name '" + _field + "' exists on the table '" + this.table + "'");

        // Finally push the coalescing field in the array
        this._coalescing.push(_field);
        this._trace("Set field '" + _field + "' as coalescent");
    },

    /**
     * Ignore the specified header while processing the Excel file.
     *
     * @param {String} header Name of the header in the Excel file
     */
    ignore: function(header) {
        // Validate parameters
        if (gs.nil(header) || typeof header != "string") throw new Error("Invalid parameter: the 'header' parameter is empty or not a string");

        // Verify that the header is not in the required list
        var _header = this._normalize(header);
        if (this._required.indexOf(_header) != -1) throw new Error("Invalid parameter: the field '" + _field + "' is also a required one");

        // Save the header to ignore
        if (this._ignored.indexOf(_header) === -1) this._ignored.push(_header);
        this._trace("Set header '" + _header + "' as ignored");
    },

    /**
     * Set the specified field as mandatory.
     *
     * @param {String} header Name of the header in the Excel file
     */
    require: function(header) {
        // Validate parameters
        if (gs.nil(header) || typeof header != "string") throw new Error("Invalid parameter: the 'header' parameter is empty or not a string");

        // Verify that the header is not in the ignored list
        var _header = this._normalize(header);
        if (this._ignored.indexOf(_header) != -1) throw new Error("Invalid parameter: the field '" + _field + "' is set to be ignored");

        // Save the header to require
        if (this._required.indexOf(_header) == -1) this._required.push(_header);
        this._trace("Set header '" + _header + "' as required");
    },

    /**
     * Enable or disable the debug mode, logging every step in details.
     * 
     * Note: It's sugsested to NOT enable this mode in production instance.
     *
     * @param {Boolean} active Value used to enable/disable the functionality
     */
    debug: function(active) {
        // Validate parameters
        if (gs.nil(active) || typeof active != "boolean") throw new Error("Invalid parameter: the 'active' parameter is empty or not a boolean");

        // Update the mode
        this._debug = active;
        this._trace("Debug mode has now state: " + (this._debug ? "ENABLED" : "DISABLED"));
    },

    /**
     * Enable or disable the virtual import that execute the import without actually updating records in the database.
     * 
     * Useful to validate data before the actual import.
     *
     * @param {Boolean} active Value used to enable/disable the functionality
     */
    virtual: function(active) {
        // Validate parameters
        if (gs.nil(active) || typeof active != "boolean") throw new Error("Invalid parameter: the 'active' parameter is empty or not a boolean");

        // Update the mode
        this._virtual = active;
        this._trace("Virtual import mode has now state: " + (this._virtual ? "ENABLED" : "DISABLED"));
    },

    /**
     * Run the import without executing the validation of the data.
     * 
     * CAUTION: Be careful when using this method as you can import non-parsed or invalid data.
     *
     * @param {Boolean} active Value used to enable/disable the functionality
     */
    sloppy: function(active) {
        // Validate parameters
        if (gs.nil(active) || typeof active != "boolean") throw new Error("Invalid parameter: the 'active' parameter is empty or not a boolean");

        // Update the mode
        this._sloppy = active;
        this._trace("Sloppy import mode has now state: " + (this._sloppy ? "ENABLED" : "DISABLED"));
    },

    /* ################################# End Public Methods ################################# */

    /* ################################ Start Private Methods ################################ */

    /**
     * @typedef {object} ImportResult Result of the import operation
     * @property {Boolean} success Result of the operation
     * @property {String} code Specific result code of the operation
     * @property {String} message Message related to the current operation
     * @property {number} rows Number of processed rows
     * @property {number} elapsed Time elapsed for the import process in milliseconds
     * @property {object} [data] Optional data object
     */

    /**
     * @typedef {object} RowResult Result of the row import
     * @property {number} Index of the current row
     * @property {String} code Response code for the result of the parsing of the current row
     * @property {String} message Human readable message representing the result of the parsing of the current row
     * @property {SysID|String} target SysID of the imported record (when created or updated) or field name when the result is created after a validation function
     * @property {Error} Error object when an error occour
     */

    /**
     * Create a standard return message.
     *
     * @return {ImportResult} Result of the import operation
     */
    _createReturnValue: function(code, message, data, rows) {
        var obj = {};
        /**
         * Success of the import process.
         */
        obj.success = code === XLSXImporter.STATES.SUCCESS;
        /**
         * Specific code result of the operation.
         */
        obj.code = code;
        /**
         * Generic message to return.
         */
        obj.message = message || "";
        /**
         * Number of processed rows.
         */
        obj.rows = rows || 0;
        /**
         * Time elapsed for the import process in milliseconds.
         */
        obj.elapsed = new Date() - this._start;
        /**
         * Generic data object to return.
         */
        obj.data = data || null;

        this._trace("Return value created: " + JSON.stringify(obj));

        return obj;
    },

    /**
     * Get all the non system fields for the current import target table.
     *
     * @return {String[]} List of non-system field names for the instance table
     */
    _fields: function() {
        // Create an empty record (without insertion) for the current target table
        var record = new GlideRecord(this.table);
        record.initialize();

        // Extract all the fields for the record
        var fields = Object.keys(record);

        // Ignore system fields
        function isNotSystemField(field) { return !field.startsWith("sys_"); }
        fields = fields.filter(isNotSystemField);

        return fields;
    },

    /**
     * Normalize and clean the string.
     */
    _normalize: function(value) {
        // Remove beginning and trailing spaces
        var _value = value.trim();

        // Make the string lowercase
        _value = _value.toLowerCase();

        // Remove special chars
        _value = _value.replace(/[\t\n\r]/gm, "");

        // Remove multiple whitespaces
        _value = _value.replace(/\s+/gm, " ");

        return _value;
    },

    /**
     * Normalize the keys of the row object, making them lowercase and without starting and trailing spaces.
     */
    _normalizeRowHeaders: function(row) {
        var obj = {};

        // Extract the keys from the row
        var keys = Object.keys(row);

        // For every row, normalize the key
        for (var i = 0; i < keys.length; i++) {
            // Extract the key from the object
            var key = keys[i];

            // Normalize the key
            var nkey = this._normalize(key);

            obj[nkey] = row[key];
        }

        return obj;
    },

    /**
     * Create a result object for a signle row parsed.
     *
     * @return {RowResult} Import result of the current row
     */
    _createRowResult: function(index, code, info, target, error) {
        /**
         * This object represents the result of the single row and contains
         * all the information of the import for the current row.
         */
        var obj = {};
        /**
         * Index of the current row.
         */
        obj.row = index;
        /**
         * Response code for the result of the parsing of the current row.
         */
        obj.code = code;

        var isValidationFailed = XLSXImporter.RCODES.SKIPPED_VALIDATION === code;
        /**
         * Human readable message representing the result of the parsing of the current row.
         * 
         * If the validation of the row failed, this value will contains the message linked to the validation function.
         */
        obj.message = isValidationFailed ? info : this.MESSAGES[code];

        // Add additional information for the allowed response code
        var codeNeedAdditionalInfo = [XLSXImporter.RCODES.SKIPPED_EVENT].indexOf(code) != -1;
        if (codeNeedAdditionalInfo && !gs.nil(info)) obj.message += info;

        /**
         * SysID of the imported record (when created or updated) or field
         * name when the result is created after a validation function.
         */
        obj.target = target || null;

        /**
         * Error object when an error occour.
         */
        obj.error = error || null;
        return obj;
    },

    /**
     * Create a default mapping where we suppose that the Excel's cells have the same value as the fields' labels.
     */
    _mapLabelWithName: function() {
        // Create an empty record (without insertion) for the current target table
        var record = new GlideRecord(this.table);
        record.initialize();

        // Get all non system fields
        var fields = this._fields();

        for (var i = 0; i < fields.length; i++) {
            // Extract the field name
            var field = fields[i];

            // Get the representation of the field
            var element = record.getElement(field);

            // Add the default mapping for this field
            this.map(element.getLabel(), element.getName());
        }
        this._trace("Default mapping executed for " + fields.length + " fields for record in table: " + this.table);
    },

    /**
     * Parse a single row of the Excel file.
     */
    _parseRow: function(parser, index) {
        try {
            /**
             * This flag will be used to skip or not the current row.
             * 
             * It will be set form exterior of this class through callbacks or when the validation of this row fails.
             */
            var valid = true;

            // Extract the current row object
            var row = this._normalizeRowHeaders(parser.getRow());

            // Skip the current row if all the cells are empty
            if (this._isEmptyRow(row)) return this._createRowResult(index, XLSXImporter.RCODES.SKIPPED_EMPTY);

            // Execute this callback when the row is read from file
            valid = this._triggerEvent("onRowRead", row, index);
            if (!gs.nil(valid) && !valid) return this._createRowResult(index, XLSXImporter.RCODES.SKIPPED_EVENT, "onRowRead");

            // Ignore fields by removing them from the row object
            row = this._ignoreRowHeaders(row);

            // Map the headers to the table fields
            row = this._mapHeadersToFields(row);

            // Ignore the validation steps if the import is run in SLOPPY mode
            if (!this._sloppy) {
                // Execute this callback before validating the row
                valid = this._triggerEvent("onRowValidating", row, index);
                if (!gs.nil(valid) && !valid) return this._createRowResult(index, XLSXImporter.RCODES.SKIPPED_EVENT, "onRowValidating");

                // Validate the row using the functions provided, if an error is found the related message is returned
                var result = this._validateRowValues(row);
                if (!gs.nil(result)) return this._createRowResult(index, XLSXImporter.RCODES.SKIPPED_VALIDATION, result.message, result.field);

                // Execute this callback after validating the row
                valid = this._triggerEvent("onRowValidated", row, index);
                if (!gs.nil(valid) && !valid) return this._createRowResult(index, XLSXImporter.RCODES.SKIPPED_EVENT, "onRowValidated");
            }

            // Apply transformations on the values of the row
            row = this._transformRowValues(row);

            // Execute this callback after transforming the row
            valid = this._triggerEvent("onRowTransformed", row, index);
            if (!gs.nil(valid) && !valid) return this._createRowResult(index, XLSXImporter.RCODES.SKIPPED_EVENT, "onRowTransformed");

            // Do not save anything on the database if the import is run in VIRTUAL mode
            if (!this._virtual) {
                // Create the record with the elaborated values
                var sys_id = this._createUpdateRecord(row);

                // Elaborate the record after the insertion/update
                valid = this._triggerEvent("onRowImported", row, index, sys_id);
                if (!gs.nil(valid) && !valid) return this._createRowResult(index, XLSXImporter.RCODES.SKIPPED_EVENT, "onRowImported");
            }

            // Return the successful log (the record SysID will be null if run in VIRTUAL mode)
            return this._createRowResult(index, XLSXImporter.RCODES.SUCCESS, null, this._virtual ? null : sys_id);
        } catch (ex) {
            // An unexpected error occourred, return also the error object
            return this._createRowResult(index, XLSXImporter.RCODES.ERROR, null, null, ex);
        }
    },

    /**
     * Verify if the row object has all values set to null.
     */
    _isEmptyRow: function(row) {
        function hasValue(key) {
            return !gs.nil(this.obj[key]);
        }
        var context = {};
        context.obj = row;

        var keys = Object.keys(row);
        return keys.filter(hasValue, context).length === 0;
    },

    /**
     * Convert every row header in the appropriate field name using
     * the label for the table or the custom mapping defined.
     */
    _mapHeadersToFields: function(row) {
        // Extract all the headers from the row
        var headers = Object.keys(row);

        /**
         * Object containing all the mapped values from the row object.
         *
         * The key is the mapped field name and the value the same cell value.
         */
        var data = {};

        // Iterate over every header of the row
        for (var i = 0; i < headers.length; i++) {
            // Extract the header
            var header = headers[i];

            // Map the field
            var field = this._mappings[this._normalize(header)];

            // No target field found, exclude this element from import
            if (!field) {
                gs.warn(this.type + " | Skipping header: '" + header + "', no field with this name found in the table '" + this.table + "'");
                continue;
            }

            // Save in the transformed 
            data[field] = row[header];
        }

        return data;
    },

    /**
     * Transform every row value using the transformation methods defined.
     * 
     * Executed after mapping and validation.
     */
    _transformRowValues: function(row) {
        // Extract all the fields from the row
        var fields = Object.keys(row);

        /**
         * Object containing all the trasformed values to insert in the record.
         *
         * The key is the field name and the value the trasformed Excel cell value.
         */
        var data = {};

        // Iterate over every field of the row
        for (var i = 0; i < fields.length; i++) {
            // Extract the mapped field name
            var field = fields[i];

            // Extract the Excel cell value
            var value = row[field];

            // Execute transformations for this field (if any)
            var transformations = this._transforms[field];
            if (!gs.nil(transformations)) {
                var execute = function(f) {
                    this.value = f(this.value);
                };
                var context = {};
                context.value = value;
                transformations.forEach(execute, context);
                value = context.value;
            }

            // Assign the trasformed value from the Excel to the object
            // that will be used to create the new record
            data[field] = value;
        }

        return data;
    },

    /**
     * Validate every row value using the validation methods defined.
     * 
     * Executed after field mapping.
     * 
     * Return an error message with the field name or null if the row is valid.
     */
    _validateRowValues: function(row) {
        // Extract all the fields from the row
        var fields = Object.keys(row);

        // Iterate over every field of the row
        for (var i = 0; i < fields.length; i++) {
            // Extract the field
            var field = fields[i];

            // Extract validation objects (if any)
            var farray = this._validations[field];

            // No validation method for this field, skip iteration
            if (gs.nil(farray)) continue;

            // Execute the validation method
            for (var j = 0; j < farray.length; j++) {
                // Extract the validation method
                var f = farray[j].method;

                // Execute the method
                var valid = f(row, field);

                // The validation method failed, skip this row
                if (!valid) {
                    var obj = {};
                    obj.message = farray[j].message;
                    obj.field = field;
                    return obj;
                }
            }
        }

        return null;
    },

    /**
     * Trigger the specified event with the value provided.
     */
    _triggerEvent: function(event, row, index, sys_id) {
        // Normalize the vent name
        var _event = this._normalize(event);

        // Get the callback from the map object
        var f = this._events[_event];

        // Fail fast if no callback is found
        if (gs.nil(f)) return null;

        // Create the data object to pass to callback
        var data = {};
        if (!gs.nil(row)) data.row = row;
        if (!gs.nil(index)) data.index = index;
        if (!gs.nil(sys_id)) data.sys_id = sys_id;

        // Execute the callback and return the result
        return f(data);
    },

    /**
     * Create or update a record in the target table with the specified values based on the coalesce fields.
     *
     * @param data: Obejct having as keys the fields' names and as values the values to insert.
     * @returns SysID of the newly created record or null if an error occours
     */
    _createUpdateRecord: function(data) {
        // Create and initialize the record
        var grRecord = new GlideRecord(this.table);

        // Get the sys_id of a matching record based on the coalescing fields
        var sys_id = this._getCoalescenceRecord(data);

        // If a record exists, update it, otherwise create it
        if (gs.nil(sys_id)) grRecord.newRecord(); // Automagically set default values for the fields
        else grRecord.get(sys_id);

        // Write the data object into the newly created record
        var fields = Object.keys(data);
        for (var i = 0; i < fields.length; i++) {
            // Get the current field
            var field = fields[i];

            // Set the value in the record
            grRecord.setValue(field, data[field]);
        }

        // Insert/update the record in the table and return the sys_id
        return grRecord.update();
    },

    /**
     * Based on the coalescing fields, get a matching record.
     *
     * @param {object} data Object having as keys the fields' names and as values the values to insert
     * @returns {SysID|null} SysID of the matching record or null if nothing match
     */
    _getCoalescenceRecord: function(data) {
        // If no coalescence field is defined, return false because nothing matches
        if (gs.nil(this._coalescing) || this._coalescing.length === 0) return null;

        // Create and initialize the record
        var grRecord = new GlideRecord(this.table);

        // Build the query with the coalescence fields
        for (var i = 0; i < this._coalescing.length; i++) {
            // Extract the field name
            var field = this._coalescing[i];

            // If the coalescence field is not defined, skip this iteration
            if (gs.nil(data[field])) {
                gs.warn(this.type + " | Skipping coalescing field: '" + field + "', no field with this name found in the table '" + this.table + "'");
                continue;
            }

            // Add the "AND" condition
            grRecord.addQuery(field, data[field]);
        }

        // Execute the query
        grRecord.query();

        // Cycle the records and execute the callback (if available)
        // If no callback is provided, get the first record available
        var valid = false;
        while (grRecord.next()) {
            // Execute the callback passing the current GlideRecord object
            valid = this._triggerEvent("onCoalesce", null, null, grRecord.getValue("sys_id"));

            // If one record with the selected keys exists, we need to update and not insert the record
            // When gs.nil(valid) is true it means that no callback is registered and we take the first record found
            if (gs.nil(valid) || valid) return grRecord.getValue("sys_id");
        }

        // No coalescing record found, we need to create a new record
        return null;
    },

    /**
     * Remove from the row all the fields that the user decided to ignore.
     */
    _ignoreRowHeaders: function(row) {
        // Extract all the headers from the row
        var headers = Object.keys(row);

        /**
         * Object containing only the not ignored headers of the row.
         */
        var data = {};

        // Iterate over every header of the row
        for (var i = 0; i < headers.length; i++) {
            // Extract the header
            var header = headers[i];

            // Ignore the current row if the header is in the exclusion list
            if (this._ignored.indexOf(header) != -1) continue;

            // Replicate the current value in the return object
            data[header] = row[header];
        }

        return data;
    },

    /**
     * Verify that all the required headers are present in the list of headers from the file.
     */
    _validateHeaders: function(headers) {
        // Normalize the headers from the file
        var _norm = function(header) { return this._normalize(header); };
        var _headers = headers.map(_norm.bind(this));

        /**
         * Container of the headers validation operation.
         */
        var result = {};
        /**
         * Success of the operation, if this value is false at least one header is missing.
         */
        result.success = true;
        /**
         * List of missing headers (if any)
         */
        result.missing = [];

        for (var i = 0; i < this._required.length; i++) {
            // Extract the required header
            var required = this._required[i];

            // Verify if the header is containined in the list of headers
            var missing = _headers.indexOf(required) == -1;

            // The required header is in the list, go on
            if (!missing) continue;

            // The required header is not in the list, save the missing header and validate the others
            result.success = false;
            result.missing.push(required);
        }

        return result;
    },

    /**
     * Write a debug message when the debug mode is active.
     */
    _trace: function(message) {
        if (this._debug) gs.info(this.type + " | " + this.table + " | " + message);
    },

    /* ################################# End Private Methods ################################# */

    type: "XLSXImporter"
};

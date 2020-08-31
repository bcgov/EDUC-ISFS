"use strict";

var ISFS = ISFS || {};
ISFS.DataImport = ISFS.DataImport || {
};

const FORM_STATE = {
    UNDEFINED: 0,
    CREATE: 1,
    UPDATE: 2,
    READ_ONLY: 3,
    DISABLED: 4,
    BULK_EDIT: 6
};


/**
 * Handle form load, set record name
 * @param {any} executionContext the form execution context
 */
ISFS.DataImport.onLoad = function (executionContext) {
    ISFS.DataImport.UpdateNameField(executionContext);
}

/**
 * Handle Import Type field changing, set record name
 * @param {any} executionContext the form execution context
 */
ISFS.DataImport.isfs_importtype_onchange = function (executionContext) {
    ISFS.DataImport.UpdateNameField(executionContext);
}

/**
 * Handle when import file changes, disable delete button after file is uploaded
 * @param {any} executionContext the form execution context
 */
ISFS.DataImport.isfs_importfile_onchange = function (executionContext) {
    ISFS.DataImport.HandleFileControl(executionContext);
}

/**
 * Handles hiding the import details section, sets import record name, and calls function to update the file field.
 * @param {any} executionContext execution context
 */
ISFS.DataImport.UpdateNameField = function (executionContext) {
    //debugger; 

    // Set variables
    var formContext = executionContext.getFormContext();
    var formState = formContext.ui.getFormType();
    var name = 'Data Import';

    if (formState === FORM_STATE.CREATE) formContext.ui.tabs.get('tab_General').sections.get('section_ImportDetails').setVisible(false);
    else formContext.ui.tabs.get('tab_General').sections.get('section_ImportDetails').setVisible(true);

    var fileAttr = formContext.getAttribute('isfs_importfile');
    var fileVal = fileAttr.getValue();
    if (fileVal == null) {
        var importType = formContext.getAttribute('isfs_importtype').getText();
        var createdOn = formContext.getAttribute('createdon').getValue();
        var createdDate = new Date();

        if (importType !== null && importType.length > 0) name = importType + ' ' + name;
        if (createdOn !== null && createdOn.length > 0) createdDate = Date.parse(createdOn);

        name = name + ', ' + createdDate.toLocaleString();

        formContext.getAttribute('isfs_name').setValue(name);
    }

    // Update the File field
    ISFS.DataImport.HandleFileControl(executionContext);
}

/**
 * Handles setting the file field to required, and disabling the "Delete" button after a file is uploaded (to prevent multiple flows from triggering on same record).
 * @param {any} executionContext execution context
 */
ISFS.DataImport.HandleFileControl = function (executionContext) {
    //debugger;
    var count = 0;
    var formContext = executionContext.getFormContext();
    var fileAttr = formContext.getAttribute('isfs_importfile');
    fileAttr.setRequiredLevel('required');

    var fileVal = fileAttr.getValue();
    if (fileVal !== null) {
        var fileCtrl = formContext.getControl('isfs_importfile');
        fileCtrl.setDisabled();
        window.setInterval(function () {
            try {
                count++;
                // UNSUPPORTED CODE
                var btn = window.parent.document.querySelector('[data-id="isfs_importfile.fieldControl-file-control-remove-button"]');
                if (btn !== null || count > 250) {
                    btn.disabled = true;
                    window.clearTimeout();
                }
            } catch (e) {
                if (count > 250) {
                    window.clearTimeout();
                }
            }
        }, 500);
    }
}
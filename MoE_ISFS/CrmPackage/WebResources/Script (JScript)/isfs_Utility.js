/// <reference path="xrm_v9.js" />

if (typeof (ISFS) === "undefined") { ISFS = { __namespace: true }; }

const FormType_Undefined    = 0;
const FormType_Create       = 1;
const FormType_Update       = 2;
const FormType_ReadOnly     = 3;
const FormType_Disabled     = 4;

ISFS.Utility = {
    /**
     * Convert Guid string without "{}" and in lower case
     * @param {string} guid guid
     * @returns {string} guid
     */
    ConvertGuid: function (guid) {
        if (guid) {
            guid = guid.replace(/[{}]/g, '').toLowerCase();
        }
        return guid;
    },

    GetLookupValue: function (result, valueName) {
        var lookupValue = new Array();
        lookupValue[0] = new Object();

        var id = result[valueName];
        var name = result[valueName + "@OData.Community.Display.V1.FormattedValue"];
        var entityType = result[valueName + "@Microsoft.Dynamics.CRM.lookuplogicalname"];

        lookupValue[0].id = id;
        lookupValue[0].name = name;
        lookupValue[0].entityType = entityType;

        return lookupValue;
    },

    ConstructLookupValue: function (entityType, id, name) {
        var lookupValue = new Array();
        lookupValue[0] = new Object();

        lookupValue[0].id = id;
        lookupValue[0].name = name;
        lookupValue[0].entityType = entityType;

        return lookupValue;
    },


    __namespace: true
};
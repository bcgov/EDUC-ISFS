/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.FundingDetailForm = {
    /**
     */
    OnFormLoad: function (executionContext) {

        var formContext = executionContext.getFormContext();

    },

    /**
    */
    OnChange_GrantProgram: function (executionContext) {

        var formContext = executionContext.getFormContext();

        // Set SY from selected Grant Program
        var grantProgram = formContext.getAttribute("isfs_grantprogram").getValue();
        if (grantProgram !== null && grantProgram[0] !== null) {
            Xrm.WebApi.online.retrieveRecord(grantProgram[0].entityType, grantProgram[0].id, "?$select=_isfs_schoolyear_value").then(
                function success(result) {
                    formContext.getAttribute("isfs_schoolyear").setValue(ISFS.Utility.GetLookupValue(result, "_isfs_schoolyear_value"));
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
            ISFS.FundingDetailForm.SetFundingDetailName(executionContext);
        }
        else
            formContext.getAttribute("isfs_schoolyear").setValue(null);
    },

    /**
    */
    OnChange_ESGroup: function (executionContext) {
        ISFS.FundingDetailForm.SetFundingDetailName(executionContext);
    },

    /**
    */
    OnChange_ESGroupSubCategory: function (executionContext) {
        ISFS.FundingDetailForm.SetFundingDetailName(executionContext);
    },

    SetFundingDetailName(executionContext) {
        var formContext = executionContext.getFormContext();

        var esGroupSubCategory = formContext.getAttribute("isfs_esgroupsubcategory").getValue();

        if (esGroupSubCategory !== null && esGroupSubCategory[0] != null) {
            formContext.getAttribute("isfs_name").setValue(esGroupSubCategory[0].name + " Grant");
            return;
        }

        var esGroup = formContext.getAttribute("isfs_esgroup").getValue();

        if (esGroup !== null && esGroup[0] != null) {

            // Set Name to "Grant Program" Abbreviation + "ES Group Abbreviation" + "Grant"
            var programAbbreviation = null;
            var grantProgram = formContext.getAttribute("isfs_grantprogram").getValue();
            if (grantProgram !== null && grantProgram[0] !== null) {
                Xrm.WebApi.online.retrieveRecord(grantProgram[0].entityType, grantProgram[0].id, "?$select=isfs_abbreviation").then(
                    function success(result) {
                        programAbbreviation = result.isfs_abbreviation;

                        var esGroupAbbreviation = null;
                        Xrm.WebApi.online.retrieveRecord(esGroup[0].entityType, esGroup[0].id, "?$select=isfs_abbreviation").then(
                            function success(result) {

                                esGroupAbbreviation = result.isfs_abbreviation;

                                if (programAbbreviation != null && esGroupAbbreviation != null) {
                                    formContext.getAttribute("isfs_name").setValue(programAbbreviation + " " + esGroupAbbreviation + " Grant");
                                }
                            },
                            function (error) {
                                Xrm.Utility.alertDialog(error.message);
                            });
                    },
                    function (error) {
                        Xrm.Utility.alertDialog(error.message);
                    });
            }
        }

    },

    __namespace: true
};

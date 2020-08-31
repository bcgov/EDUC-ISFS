/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.FundingRestriction = {
    /**
     */
    OnFormLoad: function (executionContext) {

        var formContext = executionContext.getFormContext();

    },

    /**
    */
    OnChange_GrantProgram: function (executionContext) {
        ISFS.FundingRestriction.SetFundingRestrictionName(executionContext);
    },

    /**
    */
    OnChange_School: function (executionContext) {
        ISFS.FundingRestriction.SetFundingRestrictionName(executionContext);
    },

    SetFundingRestrictionName(executionContext) {

        var formContext = executionContext.getFormContext();

        var school = formContext.getAttribute("isfs_school").getValue();

        if (school !== null && school[0] != null) {

            // Set Name to "School MinCode" + "Grant Program Abbreviation" + " Funding Restriction"
            var programAbbreviation = null;
            var grantProgram = formContext.getAttribute("isfs_grantprogram").getValue();
            if (grantProgram !== null && grantProgram[0] !== null) {
                Xrm.WebApi.online.retrieveRecord(grantProgram[0].entityType, grantProgram[0].id, "?$select=isfs_abbreviation").then(
                    function success(result) {
                        programAbbreviation = result.isfs_abbreviation;
                        Xrm.WebApi.online.retrieveRecord(school[0].entityType, school[0].id, "?$select=edu_mincode").then(
                            function success(result) {
                                schoolMinCode = result.edu_mincode;
                                if (programAbbreviation != null && schoolMinCode != null) {
                                    formContext.getAttribute("isfs_name").setValue(schoolMinCode + " " + programAbbreviation + " Funding Restriction");
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

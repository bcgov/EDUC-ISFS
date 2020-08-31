/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.FundingRate = {
    /**
     */
    OnFormLoad: function (executionContext) {

        var formContext = executionContext.getFormContext();
    },
    /**
    */
    OnChange_SchoolYear: function (executionContext) {

        var formContext = executionContext.getFormContext();

        // Set SY from selected Grant Program
        var schoolYear = formContext.getAttribute("isfs_schoolyear").getValue();
        if (schoolYear !== null && schoolYear[0] !== null) {
            Xrm.WebApi.online.retrieveRecord(schoolYear[0].entityType, schoolYear[0].id, "?$select=isfs_startdate,isfs_enddate").then(
                function success(result) {
                    formContext.getAttribute("isfs_startdate").setValue(new Date(result.isfs_startdate));
                    formContext.getAttribute("isfs_enddate").setValue(new Date(result.isfs_enddate));
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
        }
    },

    __namespace: true
};

/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.ScheduleForm = {
    /**
     * Open Disbursement Creation Form from Ribbon button
     * @param {any} entityId - Funding Schedule Id
     */
    onCreateDisbursement: function (entityId) {

        var pageInput = {
            pageType: "entityrecord",
            entityName: "isfs_disbursementcreation",
            createFromEntity: { entityType: "isfs_fundingschedule", id:entityId}
        };
        var navigationOptions = {
            target: 2,
            height: { value: 80, unit: "%" },
            width: { value: 80, unit: "%" },
            position: 1
        };
        Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
            function success(result) {
            },
            function (error) {
                Xrm.Utility.alertDialog(error.message);
            }
        );
    },
    /**
     * @param {XrmBase.ExecutionContext<Form.isfs_fundingschedule.Main.Information>} executionContext execution context
     */
    OnFormLoad: function (executionContext) {

        var formContext = executionContext.getFormContext();

        // Cannot change Grant Program and Funding Schedule after Payment Details records created
        var paymentDetailedCreatedOn = formContext.getAttribute("isfs_disbursementcreatedon").getValue();

        if (paymentDetailedCreatedOn !== null) {
            formContext.getControl("isfs_grantprogram").setDisabled(true);
            formContext.getControl("isfs_disbursementdate").setDisabled(true);
        }
        else {
            formContext.getControl("isfs_grantprogram").setDisabled(false);
            formContext.getControl("isfs_disbursementdate").setDisabled(false);
        }
    },

    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_fundingschedule.Main.Information>} executionContext execution context
    */
    OnChange_GrantProgram: function (executionContext) {

        var formContext = executionContext.getFormContext();

        // Set SY from selected Grant Program
        var grantProgram = formContext.getAttribute("isfs_grantprogram").getValue();
        if (grantProgram !== null && grantProgram[0] !== null) {
            Xrm.WebApi.online.retrieveRecord(grantProgram[0].entityType, grantProgram[0].id, "?$select=_isfs_schoolyear_value").then(
                function success(result) {
                    var schoolYear = ISFS.Utility.GetLookupValue(result, "_isfs_schoolyear_value");
                    formContext.getAttribute("isfs_schoolyear").setValue(schoolYear);
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
            ISFS.ScheduleForm.OnChange_DisbursementDate(executionContext);
        }
        else
            formContext.getAttribute("isfs_schoolyear").setValue(null);
    },

    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_fundingschedule.Main.Information>} executionContext execution context
    */
    OnChange_DisbursementDate: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var disbursementDate = formContext.getAttribute("isfs_disbursementdate").getValue();

        if (disbursementDate !== null) {

            // Set Name to "Grant Program" Abbreviation + Disbursement Date 
            var grantProgram = formContext.getAttribute("isfs_grantprogram").getValue();
            if (grantProgram !== null && grantProgram[0] !== null) {
                Xrm.WebApi.retrieveRecord(grantProgram[0].entityType, grantProgram[0].id, "?$select=isfs_abbreviation").then(
                    function success(result) {
                        formContext.getAttribute("isfs_name").setValue(result.isfs_abbreviation + " " + disbursementDate.toISOString().substring(0,10));
                    },
                    function (error) {
                        Xrm.Utility.alertDialog(error.message);
                    });
            }
        }
    },
    __namespace: true
};
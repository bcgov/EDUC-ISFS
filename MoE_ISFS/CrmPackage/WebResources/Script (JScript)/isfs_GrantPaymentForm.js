/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") { ISFS = { __namespace: true }; }

ISFS.GrantPayment = {
    /**
     * @param {XrmBase.ExecutionContext<Form.isfs_GrantPayment.Main.Information>} executionContext execution context
     */
    OnFormLoad: function (executionContext) {

        var formContext = executionContext.getFormContext();

        // Cannot change Grant Program after Payment Details records created
        var paymentDetailedCreatedOn = formContext.getAttribute("isfs_paymentdetailcreatedon").getValue();

        if (paymentDetailedCreatedOn !== null) {
            formContext.getControl("isfs_grantprogram").setDisabled(true);
        }
        else {
            formContext.getControl("isfs_grantprogram").setDisabled(false);
        }
    }, 

    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_GrantPayment.Main.Information>} executionContext execution context
    */
     OnChange_GrantProgram: function(executionContext) {

        var formContext = executionContext.getFormContext();

        // Set School Year
        var grantProgram = formContext.getAttribute("isfs_grantprogram").getValue();
         if (grantProgram !== null && grantProgram[0] !== null) {
            ISFS.GrantPayment.SetGrantPaymentName(executionContext);        

            Xrm.WebApi.retrieveRecord(grantProgram[0].entityType, grantProgram[0].id, "?$select=_isfs_schoolyear_value").then(
                function success(result) {
                    formContext.getAttribute("isfs_schoolyear").setValue(ISFS.Utility.GetLookupValue(result, "_isfs_schoolyear_value"));
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
        }
    }, 
    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_fundingschedule.Main.Information>} executionContext execution context
    */
    OnChange_DisbursementDate: function (executionContext) {
        ISFS.GrantPayment.SetGrantPaymentName(executionContext);        
    },

    SetGrantPaymentName: function (executionContext) {
        var formContext = executionContext.getFormContext();

        var disbursementDate = formContext.getAttribute("isfs_disbursementdate").getValue();

        if (disbursementDate !== null) {

            var grantProgram = formContext.getAttribute("isfs_grantprogram").getValue();
            if (grantProgram !== null && grantProgram[0] !== null) {
                // Set Name to "Grant Program" Abbreviation + Disbursement Date 
                Xrm.WebApi.retrieveRecord(grantProgram[0].entityType, grantProgram[0].id, "?$select=isfs_abbreviation").then(
                    function success(result) {
                        formContext.getAttribute("isfs_name").setValue(result.isfs_abbreviation + " " + disbursementDate.toISOString().substring(0, 10) + " Payment");
                    },
                    function (error) {
                        Xrm.Utility.alertDialog(error.message);
                    });
            }
        }
    }, 

    __namespace: true
};
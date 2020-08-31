/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") { ISFS = { __namespace: true }; }

ISFS.AuthorityPayment = {
    /**
     * @param {XrmBase.ExecutionContext<Form.isfs_authoritypayment.Main.Information>} executionContext execution context
     */
    OnFormLoad: function (executionContext) {

        var formContext = executionContext.getFormContext();

        // Cannot change Grant Program and Funding Schedule after Payment Details records created
        var paymentDetailedCreatedOn = formContext.getAttribute("isfs_paymentdetailcreatedon").getValue();

        if (paymentDetailedCreatedOn !== null) {
            formContext.getControl("isfs_grantprogram").setDisabled(true);
            formContext.getControl("isfs_fundingschedule").setDisabled(true);
        }
        else {
            formContext.getControl("isfs_grantprogram").setDisabled(false);
            formContext.getControl("isfs_fundingschedule").setDisabled(false);
        }
    }, 

    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_authoritypayment.Main.Information>} executionContext execution context
    */
     OnChange_GrantProgram: function(executionContext) {

        var formContext = executionContext.getFormContext();

        formContext.getAttribute("isfs_fundingschedule").setValue(null);

        // Set GL Code from selected Grant Program
        var grantProgram = formContext.getAttribute("isfs_grantprogram").getValue();
        if (grantProgram !== null && grantProgram[0] !== null) {
            Xrm.WebApi.retrieveRecord(grantProgram[0].entityType, grantProgram[0].id, "?$select=isfs_glcode").then(
                function success(result) {
                    formContext.getAttribute("isfs_glcode").setValue(result.isfs_glcode);
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
        }
    }, 

    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_authoritypayment.Main.Information>} executionContext execution context
    */
    OnChange_FundingSchedule: function(executionContext) {

        var formContext = executionContext.getFormContext();

        // Set Name to "Funding Schedule" Name + " Payment"
        var fundingSchedule = formContext.getAttribute("isfs_fundingschedule").getValue();

        if (fundingSchedule !== null && fundingSchedule[0] !== null) {

            Xrm.WebApi.retrieveMultipleRecords("isfs_authoritypayment", "?$filter=_isfs_fundingschedule_value eq " + fundingSchedule[0].id).then(
                function success(result) {
                    var duplicate = false;
                    for (var i = 0; i < result.entities.length; i++) {
                        if (result.entities[i].isfs_authoritypaymentid !== ISFS.Utility.ConvertGuid(formContext.data.entity.getId())) {
                            duplicate = true;
                            break;
                        }
                    }
                    if (duplicate) {
                        formContext.getControl("isfs_fundingschedule").setNotification("Another Authority Payment already exists for the selected Funding Schedule.");
                    }
                    else {
                        formContext.getAttribute("isfs_name").setValue(fundingSchedule[0].name + " Payment");
                        formContext.getControl("isfs_fundingschedule").clearNotification();
                    }
                },
                function (error) {
                    console.log(error.message);
                }
            );
        }
    }, 
    __namespace: true
};
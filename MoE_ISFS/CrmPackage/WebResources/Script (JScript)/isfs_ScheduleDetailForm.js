/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.ScheduleDetailForm = {
    /**
     * @param {XrmBase.ExecutionContext<Form.isfs_fundingscheduledetail.Main.Information>} executionContext execution context
     */
    OnFormLoad: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var fundingSchedule = formContext.getAttribute("isfs_fundingschedule").getValue();
        if (fundingSchedule !== null && fundingSchedule[0] !== null) {
            formContext.getControl("isfs_fundingdetail").setDisabled(false);
        }
        else {
            formContext.getControl("isfs_fundingdetail").setDisabled(true);
        }

        ISFS.ScheduleDetailForm.OnChange_FundingDetail(executionContext);
        ISFS.ScheduleDetailForm.OnChange_NewDisbursement(executionContext);

        //var fundingDetail = formContext.getAttribute("isfs_fundingdetail").getValue();
        //if (fundingDetail !== null && fundingDetail[0] !== null) {
        //    formContext.getControl("isfs_esgroupsubcategory").setDisabled(false);
        //}
        //else {
        //    formContext.getControl("isfs_esgroupsubcategory").setDisabled(true);
        //}
    },

    OnChange_NewDisbursement: function (executionContext) {
        var formContext = executionContext.getFormContext();

        var newDisbursement = formContext.getAttribute("isfs_newdisbursement").getValue();
        if (newDisbursement == true) {
            formContext.getAttribute("isfs_previousfundingscheduledetail").setRequiredLevel("none");
            formContext.getAttribute("isfs_previousfundingscheduledetail").setValue(null);
            formContext.getControl("isfs_previousfundingscheduledetail").setDisabled(true);
        }
        else {
            formContext.getAttribute("isfs_previousfundingscheduledetail").setRequiredLevel("required");
            formContext.getControl("isfs_previousfundingscheduledetail").setDisabled(false);
        }
    },

    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_fundingscheduledetail.Main.Information>} executionContext execution context
    */
    OnChange_FundingSchedule: function (executionContext) {
        var formContext = executionContext.getFormContext();

        // Set SY and Grant Program from selected Funding Schedule
        var fundingSchedule = formContext.getAttribute("isfs_fundingschedule").getValue();
        if (fundingSchedule !== null && fundingSchedule[0] !== null) {
            Xrm.WebApi.online.retrieveRecord(fundingSchedule[0].entityType, fundingSchedule[0].id, "?$select=_isfs_schoolyear_value, _isfs_grantprogram_value").then(
                function success(result) {
                    var schoolYear = ISFS.Utility.GetLookupValue(result, "_isfs_schoolyear_value");
                    formContext.getAttribute("isfs_schoolyear").setValue(schoolYear);
                    var grantProgram = ISFS.Utility.GetLookupValue(result, "_isfs_grantprogram_value");
                    formContext.getAttribute("isfs_grantprogram").setValue(grantProgram);

                    formContext.getControl("isfs_fundingdetail").setDisabled(false);
                    formContext.getAttribute("isfs_schoolsubcategory").setValue(null);
                    formContext.getAttribute("isfs_esgroup").setValue(null);
                    formContext.getAttribute("isfs_esgroupsubcategory").setValue(null);
                    formContext.getControl("isfs_esgroupsubcategory").setDisabled(true);

                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
        }
        else {
            formContext.getAttribute("isfs_schoolyear").setValue(null);
            formContext.getAttribute("isfs_grantprogram").setValue(null);
            formContext.getAttribute("isfs_fundingdetail").setValue(null);
            formContext.getControl("isfs_fundingdetail").setDisabled(true);

            formContext.getAttribute("isfs_schoolsubcategory").setValue(null);
            formContext.getAttribute("isfs_esgroup").setValue(null);
            formContext.getAttribute("isfs_esgroupsubcategory").setValue(null);
            formContext.getControl("isfs_esgroupsubcategory").setDisabled(true);
        }
    },

    OnChange_FundingDetail: function (executionContext) {
        var formContext = executionContext.getFormContext();

        // Set IS School Sub-Category and ES Group from selected Funding Detail
        var fundingDetail = formContext.getAttribute("isfs_fundingdetail").getValue();
        if (fundingDetail !== null && fundingDetail[0] !== null) {
            Xrm.WebApi.online.retrieveRecord(fundingDetail[0].entityType, fundingDetail[0].id, "?$select=_isfs_esgroup_value,_isfs_esgroupsubcategory_value,_isfs_isschoolsubcategory_value").then(
                function success(result) {
                    var schoolCategory = ISFS.Utility.GetLookupValue(result, "_isfs_isschoolsubcategory_value");
                    formContext.getAttribute("isfs_schoolsubcategory").setValue(schoolCategory);

                    var esGroup = ISFS.Utility.GetLookupValue(result, "_isfs_esgroup_value");
                    formContext.getAttribute("isfs_esgroup").setValue(esGroup);

                    var esGroupSubCategory = ISFS.Utility.GetLookupValue(result, "_isfs_esgroupsubcategory_value");
                    if (esGroupSubCategory[0].id != null) {
                        formContext.getAttribute("isfs_esgroupsubcategory").setValue(esGroupSubCategory);
                        formContext.getControl("isfs_esgroupsubcategory").setDisabled(true);
                    }
                    else
                        formContext.getControl("isfs_esgroupsubcategory").setDisabled(false);


                    ISFS.ScheduleDetailForm.SetScheduleDetailName(executionContext);

                    formContext.getControl("isfs_esgroupsubcategory").removePreSearch(ISFS.ScheduleDetailForm.SetESGroupSubCategoryFilter);
                    formContext.getControl("isfs_esgroupsubcategory").addPreSearch(ISFS.ScheduleDetailForm.SetESGroupSubCategoryFilter);

                    //formContext.getControl("isfs_previousfundingscheduledetail").removePreSearch(ISFS.ScheduleDetailForm.SetFundingScheduleDetailFilter);
                    //formContext.getControl("isfs_previousfundingscheduledetail").addPreSearch(ISFS.ScheduleDetailForm.SetFundingScheduleDetailFilter);

                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
        }
        else {
            formContext.getAttribute("isfs_schoolsubcategory").setValue(null);
            formContext.getAttribute("isfs_esgroup").setValue(null);
            formContext.getAttribute("isfs_esgroupsubcategory").setValue(null);
            formContext.getControl("isfs_esgroupsubcategory").setDisabled(true);
            formContext.getControl("isfs_esgroupsubcategory").removePreSearch(ISFS.ScheduleDetailForm.SetESGroupSubCategoryFilter);
            //formContext.getAttribute("isfs_previousfundingscheduledetail").setValue(null);
            //formContext.getControl("isfs_previousfundingscheduledetail").setDisabled(true);
            //formContext.getControl("isfs_previousfundingscheduledetail").removePreSearch(ISFS.ScheduleDetailForm.SetFundingScheduleDetailFilter);
        }
    },

    SetESGroupSubCategoryFilter: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var schoolCategory = formContext.getAttribute("isfs_schoolsubcategory").getValue();
        var esGroup = formContext.getAttribute("isfs_esgroup").getValue();

        var fetchData = {
            isfs_isschoolsubcategory: schoolCategory[0].id,
            isfs_esgroup: esGroup[0].id
        };
        var fetchXml = [
            "<filter type='and'>",
            "      <condition attribute='isfs_isschoolsubcategory' operator='eq' value='", fetchData.isfs_isschoolsubcategory/*isfs_isschoolsubcategory*/, "'/>",
            "      <condition attribute='isfs_esgroup' operator='eq' value='", fetchData.isfs_esgroup/*isfs_esgroup*/, "'/>",
            "    </filter>",
        ].join("");

        formContext.getControl("isfs_esgroupsubcategory").addCustomFilter(fetchXml);
    },

    SetFundingScheduleDetailFilter: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var fundingDetail = formContext.getAttribute("isfs_fundingdetail").getValue();

        var fetchData = {
            isfs_fundingdetail: fundingDetail[0].id
        };
        var fetchXml = [
            "<filter type='and'>",
            "      <condition attribute='isfs_fundingdetail' operator='eq' value='", fetchData.isfs_fundingdetail/*isfs_isschoolsubcategory*/, "'/>",
            "    </filter>",
        ].join("");

        formContext.getControl("isfs_previousfundingscheduledetail").addCustomFilter(fetchXml);
    },

    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_fundingscheduledetail.Main.Information>} executionContext execution context
    */
    OnChange_ESGroupSubCategory: function (executionContext) {
        ISFS.ScheduleDetailForm.SetScheduleDetailName(executionContext);
    },

    OnChange_ESGroup: function (executionContext) {
        ISFS.ScheduleDetailForm.SetScheduleDetailName(executionContext);
    },
    /**
    * @param {XrmBase.ExecutionContext<Form.isfs_fundingscheduledetail.Main.Information>} executionContext execution context
    */
    SetScheduleDetailName: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var fundingSchedule = formContext.getAttribute("isfs_fundingschedule").getValue();
        if (fundingSchedule !== null && fundingSchedule[0] !== null) {

            var esGroupSub = formContext.getAttribute("isfs_esgroupsubcategory").getValue();
            if (esGroupSub !== null && esGroupSub[0] !== null) {
                formContext.getAttribute("isfs_name").setValue(fundingSchedule[0].name + " " + esGroupSub[0].name);
                return;
            }

            var esGroup = formContext.getAttribute("isfs_esgroup").getValue();
            if (esGroup !== null && esGroup[0] !== null) {
                //formContext.getAttribute("isfs_name").setValue(fundingSchedule[0].name + " " + esGroup[0].name);

                Xrm.WebApi.retrieveRecord(esGroup[0].entityType, esGroup[0].id, "?$select=isfs_abbreviation").then(
                    function success(result) {
                        formContext.getAttribute("isfs_name").setValue(fundingSchedule[0].name + " " + result.isfs_abbreviation);
                    },
                    function (error) {
                        Xrm.Utility.alertDialog(error.message);
                    });
            }
            else
                formContext.getAttribute("isfs_name").setValue(null);
        }
    },

    __namespace: true
};
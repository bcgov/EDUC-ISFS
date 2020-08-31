/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.FundingRateDetail = {
    /**
     */
    OnFormLoad: function (executionContext) {

        var formContext = executionContext.getFormContext();

        ISFS.FundingRateDetail.OnChange_FundingGroupType(executionContext);
        ISFS.FundingRateDetail.OnChange_FundingRateType(executionContext);
        ISFS.FundingRateDetail.OnChange_BaseRateType(executionContext);

        //// Cannot change Grant Program and Funding Schedule after Payment Details records created
        //var paymentDetailedCreatedOn = formContext.getAttribute("isfs_disbursementcreatedon").getValue();

        //if (paymentDetailedCreatedOn !== null) {
        //    formContext.getControl("isfs_grantprogram").setDisabled(true);
        //    formContext.getControl("isfs_disbursementdate").setDisabled(true);
        //}
        //else {
        //    formContext.getControl("isfs_grantprogram").setDisabled(false);
        //    formContext.getControl("isfs_disbursementdate").setDisabled(false);
        //}
    },

    OnChange_FundingRateType: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var fundingRateType = formContext.getAttribute("isfs_fundingratetype").getValue();

        if (fundingRateType == null || fundingRateType == 746910000) { // flat
            formContext.getControl("isfs_schooldistrict").setVisible(false);
            formContext.getAttribute("isfs_schooldistrict").setRequiredLevel("none");
            formContext.getAttribute("isfs_schooldistrict").setValue(null);

            formContext.getControl("isfs_school").setVisible(false);
            formContext.getAttribute("isfs_school").setRequiredLevel("none");
            formContext.getAttribute("isfs_school").setValue(null);
        }
        else if (fundingRateType == 746910001) { // District
            formContext.getAttribute("isfs_schooldistrict").setRequiredLevel("required");
            formContext.getControl("isfs_schooldistrict").setVisible(true);

            formContext.getControl("isfs_school").setVisible(false);
            formContext.getAttribute("isfs_school").setRequiredLevel("none");
            formContext.getAttribute("isfs_school").setValue(null);
        }
        else if (fundingRateType == 746910002) { // School
            formContext.getAttribute("isfs_school").setRequiredLevel("required");
            formContext.getControl("isfs_school").setVisible(true);

            formContext.getControl("isfs_schooldistrict").setVisible(false);
            formContext.getAttribute("isfs_schooldistrict").setRequiredLevel("none");
            formContext.getAttribute("isfs_schooldistrict").setValue(null);
        }
    },

    OnChange_FundingGroupType: function (executionContext) {
        var formContext = executionContext.getFormContext();

        var fundingGroupType = formContext.getAttribute("isfs_fundinggrouptype").getValue();

        if (fundingGroupType == null || fundingGroupType == 746910000) { // N/A
            formContext.getAttribute("isfs_fundinggroup").setRequiredLevel("none");
            formContext.getAttribute("isfs_fundinggroup").setValue(null);
            formContext.getControl("isfs_fundinggroup").setVisible(false);

            formContext.getAttribute("isfs_grouppercentage").setRequiredLevel("none");
            formContext.getAttribute("isfs_grouppercentage").setValue(null);
            formContext.getControl("isfs_grouppercentage").setVisible(false);

            formContext.getAttribute("isfs_esgroupsubcategory").setRequiredLevel("none");
            //formContext.getAttribute("isfs_esgroupsubcategory").setValue(null);
            //formContext.getControl("isfs_esgroupsubcategory").setVisible(false);

            formContext.getAttribute("isfs_isschoolsubcategory").setRequiredLevel("none");
        }
        else { // School, ES Group, ES Group Sub-Category
            formContext.getAttribute("isfs_fundinggroup").setRequiredLevel("required");
            formContext.getControl("isfs_fundinggroup").setVisible(true);

            formContext.getAttribute("isfs_grouppercentage").setRequiredLevel("required");
            formContext.getControl("isfs_grouppercentage").setVisible(true);

            if (fundingGroupType == 746910003) { // ES Group Sub-Category
                formContext.getAttribute("isfs_esgroupsubcategory").setRequiredLevel("required");
                //formContext.getControl("isfs_esgroupsubcategory").setVisible(true);

                formContext.getAttribute("isfs_isschoolsubcategory").setRequiredLevel("required");
            }
            else {
                formContext.getAttribute("isfs_esgroupsubcategory").setRequiredLevel("none");
                //formContext.getAttribute("isfs_esgroupsubcategory").setValue(null);
                //formContext.getControl("isfs_esgroupsubcategory").setVisible(false);

                formContext.getAttribute("isfs_isschoolsubcategory").setRequiredLevel("none");
            }
        }
    },

    OnChange_BaseRateType: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var baseRateType = formContext.getAttribute("isfs_baseratetype").getValue();

        if (baseRateType == 746910000) { // IS Grant Rate
            formContext.getControl("isfs_publicschoolrate").setVisible(false);
            formContext.getAttribute("isfs_publicschoolrate").setRequiredLevel("none");
            formContext.getAttribute("isfs_publicschoolrate").setValue(null);

            formContext.getControl("isfs_fngrantrate").setVisible(false);
            formContext.getAttribute("isfs_fngrantrate").setRequiredLevel("none");
            formContext.getAttribute("isfs_fngrantrate").setValue(null);

            formContext.getAttribute("isfs_isgrantrate").setRequiredLevel("required");
            formContext.getControl("isfs_isgrantrate").setVisible(true);
        }
        else if (baseRateType == 746910001) { // PS Grant Rate
            formContext.getControl("isfs_isgrantrate").setVisible(false);
            formContext.getAttribute("isfs_isgrantrate").setRequiredLevel("none");
            formContext.getAttribute("isfs_isgrantrate").setValue(null);

            formContext.getControl("isfs_fngrantrate").setVisible(false);
            formContext.getAttribute("isfs_fngrantrate").setRequiredLevel("none");
            formContext.getAttribute("isfs_fngrantrate").setValue(null);

            formContext.getAttribute("isfs_publicschoolrate").setRequiredLevel("required");
            formContext.getControl("isfs_publicschoolrate").setVisible(true);
        }
        else if (baseRateType == 746910002) { // FN Grant Rate
            formContext.getControl("isfs_publicschoolrate").setVisible(false);
            formContext.getAttribute("isfs_publicschoolrate").setRequiredLevel("none");
            formContext.getAttribute("isfs_publicschoolrate").setValue(null);

            formContext.getControl("isfs_isgrantrate").setVisible(false);
            formContext.getAttribute("isfs_isgrantrate").setRequiredLevel("none");
            formContext.getAttribute("isfs_isgrantrate").setValue(null);


            formContext.getAttribute("isfs_fngrantrate").setRequiredLevel("required");
            formContext.getControl("isfs_fngrantrate").setVisible(true);
        }
    },

    OnCalculateISRate: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var baseRateType = formContext.getAttribute("isfs_baseratetype").getValue();

        var baseRate = null;
        if (baseRateType == 746910000) { // IS Grant Rate
            baseRate = formContext.getAttribute("isfs_isgrantrate").getValue();
        }
        else if (baseRateType == 746910001) { // PS Grant Rate
            baseRate = formContext.getAttribute("isfs_publicschoolrate").getValue();
        }
        else if (baseRateType == 746910002) { // FN Grant Rate
            baseRate = formContext.getAttribute("isfs_fngrantrate").getValue();
        }

        if (baseRate != null) {
            var isRate = baseRate;

            var groupPercentage = formContext.getAttribute("isfs_grouppercentage").getValue();
            if (groupPercentage != null) {
                isRate = isRate * groupPercentage;
            }

            var grantPercentage = formContext.getAttribute("isfs_grantpercentage").getValue();
            if (grantPercentage != null) {
                isRate = isRate * grantPercentage;
            }

            formContext.getAttribute("isfs_isschoolrate").setValue(isRate);
        }
    },

    OnFundingRateDetailName: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var name = "";

        var school = formContext.getAttribute("isfs_school").getValue();
        if (school !== null && school[0] !== null) {
            name = name + " " + school[0].name;
        }
        else {
            var district = formContext.getAttribute("isfs_schooldistrict").getValue();
            if (district !== null && district[0] !== null) name = name + " " + district[0].name;
        }

        var esGroupSubcategory = formContext.getAttribute("isfs_esgroupsubcategory").getValue();
        if (esGroupSubcategory !== null && esGroupSubcategory[0] !== null) {
            name = name + " " + esGroupSubcategory[0].name;
        }
        else {
            var schoolCategory = formContext.getAttribute("isfs_isschoolsubcategory").getValue();
            if (schoolCategory !== null && schoolCategory[0] !== null) name = name + " " + schoolCategory[0].name;

            var esGroup = formContext.getAttribute("isfs_esgroup").getValue();
            if (esGroup !== null && esGroup[0] !== null) name = name + " " + esGroup[0].name;
        }

        var fundingGroup = formContext.getAttribute("isfs_fundinggroup").getValue();
        if (fundingGroup !== null && fundingGroup[0] !== null) name = name + " " + fundingGroup[0].name;

        formContext.getAttribute("isfs_name").setValue(name.trim());
    },
    /**
    */
    OnChange_ESGroup: function (executionContext) {

        var formContext = executionContext.getFormContext();

        formContext.getAttribute("isfs_esgroupsubcategory").setValue(null);
    },

    OnChange_ESGroupSubCategory: function (executionContext) {

        var formContext = executionContext.getFormContext();

        var esGroupSubCategory = formContext.getAttribute("isfs_esgroupsubcategory").getValue();

        if (esGroupSubCategory !== null && esGroupSubCategory[0] != null) {
            Xrm.WebApi.online.retrieveRecord(esGroupSubCategory[0].entityType, esGroupSubCategory[0].id, "?$select=_isfs_isschoolsubcategory_value").then(
                function success(result) {
                    var schoolCategory = ISFS.Utility.GetLookupValue(result, "_isfs_isschoolsubcategory_value");
                    formContext.getAttribute("isfs_isschoolsubcategory").setValue(schoolCategory);
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                });
        }
    },

    __namespace: true
};

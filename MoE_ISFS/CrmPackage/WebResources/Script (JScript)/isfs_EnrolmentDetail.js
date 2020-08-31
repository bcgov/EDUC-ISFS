/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.EnrolmentDetail = {
    /**
     */
    OnFormLoad: function (executionContext) {
        var formContext = executionContext.getFormContext();

        if (formContext.ui.getFormType() == FormType_Create)
            formContext.getControl("isfs_adjustedenrolmentnumber").setVisible(false);

        formContext.getControl("isfs_adjustmentnotes").setVisible(false);
        formContext.getAttribute("isfs_adjustmentnotes").setRequiredLevel("none");
    },

    /**
     */
    OnChange_Enrolment: function (executionContext) {
        var formContext = executionContext.getFormContext();

        if (formContext.ui.getFormType() != FormType_Update) return;

        var enrolmentNumber = formContext.getAttribute("isfs_enrolmentnumber").getIsDirty();
        var adjusted = formContext.getAttribute("isfs_adjustedenrolmentnumber").getIsDirty();
        if (enrolmentNumber == true || adjusted == true) {
            formContext.getControl("isfs_adjustmentnotes").setVisible(true);
            formContext.getAttribute("isfs_adjustmentnotes").setRequiredLevel("required");
            formContext.getAttribute("isfs_adjustmentnotes").setValue("");
            if(adjusted) formContext.getControl("isfs_adjustmentnotes").setFocus();
        }
        else {
            formContext.getControl("isfs_adjustmentnotes").setVisible(false);
            formContext.getAttribute("isfs_adjustmentnotes").setRequiredLevel("none");
        }
    },

    __namespace: true
};
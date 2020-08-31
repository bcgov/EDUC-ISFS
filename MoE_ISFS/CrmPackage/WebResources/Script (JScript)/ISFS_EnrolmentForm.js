if (typeof (ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.EnrolmentForm = {
    OnFormLoad: function (executionContext) {
        var formContext = executionContext.getFormContext();
        var collectiontype = formContext.getAttribute("isfs_collectiontype").getValue();
        if (collectiontype == 746910002) {
            Xrm.Page.ui.tabs.get("General").sections.get("ESAudit").setVisible(true);
            Xrm.Page.ui.tabs.get("General").sections.get("MayEnrolmentDetail").setVisible(true);
            Xrm.Page.ui.tabs.get("General").sections.get("{ce84a170-4354-4ce4-8b97-3ea8966429c8}_section_3").setVisible(false);
        } else {
            Xrm.Page.ui.tabs.get("General").sections.get("ESAudit").setVisible(false);
            Xrm.Page.ui.tabs.get("General").sections.get("MayEnrolmentDetail").setVisible(false);
            Xrm.Page.ui.tabs.get("General").sections.get("{ce84a170-4354-4ce4-8b97-3ea8966429c8}_section_3").setVisible(true);
        }

    }
    ,
    OnChange_auditor: function (executionContext) {
        var formContext = executionContext.getFormContext();
        // Set all auditor fields from the selected firm record
        var auditor = formContext.getAttribute("isfs_auditor").getValue();
        if (auditor !== null && auditor[0] !== null) {
            Xrm.WebApi.online.retrieveRecord(auditor[0].entityType, auditor[0].id,
                "?$select=isfs_practitioner, isfs_phonenumber, isfs_address1, isfs_address1_city, isfs_address1_province, isfs_address2, isfs_address2_city, isfs_address2_province, emailaddress").then(
                    function success(result) {
                        var practitioner = result.isfs_practitioner;
                        formContext.getAttribute("isfs_practitioner").setValue(practitioner);
                        var phoneNumber = result.isfs_phonenumber;
                        formContext.getAttribute("isfs_phonenumber").setValue(phoneNumber);
                        var address1 = result.isfs_address1;
                        formContext.getAttribute("isfs_address1").setValue(address1);
                        var address1City = result.isfs_address1_city;
                        formContext.getAttribute("isfs_address1_city").setValue(address1City);
                        var address1Province = result.isfs_address1_province;
                        formContext.getAttribute("isfs_address1_province").setValue(address1Province);
                        var emailAddress = result.emailaddress;
                        formContext.getAttribute("isfs_emailaddress").setValue(emailAddress);
                        var address2 = result.isfs_address2;
                        formContext.getAttribute("isfs_address2").setValue(address2);
                        var address2City = result.isfs_address2_city;
                        formContext.getAttribute("isfs_address2_city").setValue(address2City);
                        var address2Province = result.isfs_address2_province;
                        formContext.getAttribute("isfs_address2_province").setValue(address2Province);
                    },
                    function (error) {
                        Xrm.Utility.alertDialog(error.message);
                    }
                );
        }
        else {
            formContext.getAttribute("isfs_practitioner").setValue(null);
            formContext.getAttribute("isfs_phonenumber").setValue(null);
            formContext.getAttribute("isfs_address1").setValue(null);
            formContext.getAttribute("isfs_address1_city").setValue(null);
            formContext.getAttribute("isfs_address1_province").setValue(null);
            formContext.getAttribute("isfs_emailaddress").setValue(null);
            formContext.getAttribute("isfs_address2").setValue(null);
            formContext.getAttribute("isfs_address2_city").setValue(null);
            formContext.getAttribute("isfs_address2_province").setValue(null);
        }

    }
}

/// <reference path="xrm_v9.js" />
/// <reference path="isfs_utility.js" />

if (typeof(ISFS) === "undefined") {
    ISFS = {
        __namespace: true
    };
}

ISFS.DisbursementCreation = {
    OnCreateOnce: function (formContext) {
        var recordId = formContext.data.entity.getId();
        return recordId === "";
    },

    OnFormSave: function (executionContext) {
        // Disable Auto Save
        var eventArgs = executionContext.getEventArgs();
        if (eventArgs.getSaveMode() == 70 || eventArgs.getSaveMode() == 2) {
            eventArgs.preventDefault();
        }
    },

    OnCreateDisbursementClose: function (formContext) {
        ISFS.DisbursementCreation.OnCreateDisbursement(formContext, true);
    },

    OnCreateDisbursement: function (formContext, close) {

        var selectedEntityReferences = [];
        var selectedIDs = "";

        var selectedRows = formContext.getControl("gridEnrolment").getGrid().getSelectedRows();

        selectedRows.forEach(function (selectedRow, i) {
            selectedEntityReferences.push(selectedRow.getData().getEntity().getEntityReference());
        });

        //get all required data from each selected row    
        for (let i = 0; i < selectedEntityReferences.length; i++) {
            if (selectedIDs == "") {
                selectedIDs = ISFS.Utility.ConvertGuid(selectedEntityReferences[i].id.toString());
            }
            else {
                selectedIDs = selectedIDs + ";" + ISFS.Utility.ConvertGuid(selectedEntityReferences[i].id.toString());
            }
        }  

        if (selectedIDs === "") {
            formContext.ui.setFormNotification("Please select at least One School Enrolment.", "INFO", "checkSelection");
            return;
        }
        else {
            formContext.ui.clearFormNotification("checkSelection");
            formContext.getAttribute("isfs_selectedenrolmentids").setValue(selectedIDs);
        }

        formContext.data.save().then(
            function success(results) {
                if (close === true) {
                    formContext.ui.close();
                    return;
                }

                Xrm.Utility.showProgressIndicator("Creating School Disbursements ...");

                var t = setInterval(function () {
                    Xrm.WebApi.online.retrieveRecord(formContext.data.entity.getEntityName(), formContext.data.entity.getId(), "?$select=isfs_creationcompletedon").then(
                        function success(result) {
                            var isfs_creationcompletedon = result.isfs_creationcompletedon;
                            if (isfs_creationcompletedon != null) {
                                clearInterval(t);
                                Xrm.Utility.closeProgressIndicator();
                                ISFS.DisbursementCreation.OnFormReload(formContext);
                            }
                        },
                        function (error) {
                            clearInterval(t);
                            Xrm.Utility.closeProgressIndicator();
                            Xrm.Utility.alertDialog(error.message);
                        });
                }, 3000)

                setTimeout(function () {
                    clearInterval(t);
                    Xrm.Utility.closeProgressIndicator();
                }, 120000); // Cancel waiting after 2 minutes
                
            },
            function (error) {
                Xrm.Utility.alertDialog(error.message);
            });
    }, 

    /**
     */
    OnFormLoad: function (executionContext) {
        var formContext = executionContext.getFormContext();
        ISFS.DisbursementCreation.OnFormReload(formContext);
    },
    /**
     */
    OnFormReload: function (formContext) {

        var recordId = formContext.data.entity.getId();

        if (recordId === "") {
            var now = new Date();
            var nowString = now.toISOString();
            formContext.getAttribute("isfs_name").setValue("Disbursement Creation " + nowString.replace(/:/g, ""));


            //Xrm.WebApi.online.retrieveMultipleRecords("isfs_fundingrate", "?$select=isfs_fundingrateid,isfs_name&$orderby=createdon asc&$top=1").then(
            //    function success(results) {
            //        for (var i = 0; i < results.entities.length; i++) {
            //            var isfs_fundingrateid = results.entities[i]["isfs_fundingrateid"];
            //            var isfs_name = results.entities[i]["isfs_name"];
            //            formContext.getAttribute("isfs_fundingrate").setValue(ISFS.Utility.ConstructLookupValue("isfs_fundingrate", isfs_fundingrateid, isfs_name))
            //        }
            //    },
            //    function (error) {
            //        Xrm.Utility.alertDialog(error.message);
            //    }
            //);
        }

        ISFS.DisbursementCreation.OnChangeCollectionFormContext(formContext);

        // Filter Disbursements
        var fundingSchedule = formContext.getAttribute("isfs_fundingschedule").getValue();
        if (fundingSchedule !== null && fundingSchedule[0] !== null) {
            var fetchXml = [
                "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>",
                "  <entity name='isfs_schooldisbursement'>",
                "    <filter type='and'>",
                "      <condition attribute='isfs_fundingschedule' operator='eq' value='", ISFS.Utility.ConvertGuid(fundingSchedule[0].id), "'/>",
                "    </filter>",
                "  </entity>",
                "</fetch>",
            ].join("");

            formContext.getControl("gridSchoolDisbursement").setFilterXml(fetchXml);
            formContext.getControl("gridSchoolDisbursement").setVisible(true);
            formContext.getControl("gridSchoolDisbursement").refresh();
        }

        if (recordId !== "") {

            formContext.data.refresh();

            formContext.getControl("isfs_enrolmentcollection").setDisabled(true);
            formContext.getControl("isfs_fundingrate").setDisabled(true);
            formContext.getControl("isfs_replaceexistingdisbursements").setDisabled(true);

            formContext.getControl("gridEnrolment").setVisible(false);

            formContext.ui.tabs.get("tabCreationLog").setVisible(true);

            var fetchXml = [
                "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>",
                "  <entity name='isfs_schooldisbursement'>",
                "    <filter type='and'>",
                "      <condition attribute='isfs_disbursementcreation' operator='eq' value='", ISFS.Utility.ConvertGuid(recordId), "'/>",
                "    </filter>",
                "  </entity>",
                "</fetch>",
            ].join("");

            formContext.getControl("gridCreatedDisbursement").setFilterXml(fetchXml);
            formContext.getControl("gridCreatedDisbursement").setVisible(true);
            formContext.getControl("gridCreatedDisbursement").refresh();
        }
        else {
            formContext.getControl("isfs_enrolmentcollection").setDisabled(false);
            formContext.getControl("isfs_fundingrate").setDisabled(false);
            formContext.getControl("isfs_replaceexistingdisbursements").setDisabled(false);
            formContext.getControl("gridCreatedDisbursement").setVisible(false);
        }
    },

    FilterEnrolmentGrid(formContext, collectionId) {
        if (collectionId == null) {
            collectionId = "3320C9DB-0000-EA11-A810-000D3A0C8A65"; // dummy guid
        }

        formContext.getControl("gridEnrolment").setVisible(false);

        var fetchData = {
            isfs_enrolmentcollection: ISFS.Utility.ConvertGuid(collectionId)
        };
        var fetchXml = [
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>",
            "  <entity name='isfs_isenrolment'>",
            "    <attribute name='isfs_name' />",
            "    <attribute name='createdon' />",
            "    <attribute name='isfs_school' />",
            "    <attribute name='modifiedon' />",
            "    <attribute name='modifiedby' />",
            "    <attribute name='isfs_enrolmentcollection' />",
            "    <attribute name='createdby' />",
            "    <attribute name='isfs_isenrolmentid' />",
            "    <order attribute='isfs_name' descending='false' />",
            "    <filter type='and'>",
            "      <condition attribute='isfs_enrolmentcollection' operator='eq' value='", fetchData.isfs_enrolmentcollection, "'/>",
            "    </filter>",
            "  </entity>",
            "</fetch>",
        ].join("");

        formContext.getControl("gridEnrolment").setFilterXml(fetchXml);
        formContext.getControl("gridEnrolment").refresh();
        formContext.getControl("gridEnrolment").setVisible(true);
    }, 
    /**
    */
    OnChange_Collection: function (executionContext) {

        var formContext = executionContext.getFormContext();

        ISFS.DisbursementCreation.OnChangeCollectionFormContext(formContext);
    },

    /**
    */
    OnChangeCollectionFormContext: function (formContext) {

        // Filter IS Enrolment by Collection
        var collection = formContext.getAttribute("isfs_enrolmentcollection").getValue();
        if (collection !== null && collection[0] !== null) {
            ISFS.DisbursementCreation.FilterEnrolmentGrid(formContext, collection[0].id);
        }
        else
            ISFS.DisbursementCreation.FilterEnrolmentGrid(formContext, null);

    },

    __namespace: true
};
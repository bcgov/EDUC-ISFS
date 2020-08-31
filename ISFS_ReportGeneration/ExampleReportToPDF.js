/*
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
    <entity name="new_quoteproductmapping">
        <attribute name="new_quoteproductmappingid" />
        <attribute name="new_name" />
        <attribute name="createdon" />
        <link-entity name="quote" from="quoteid" to="new_quoteproductmapping" alias="aa" enableprefiltering="1">
            <attribute name="name" />
            <attribute name="customerid" />
            <attribute name="statecode" />
            <attribute name="totalamount" />
            <attribute name="quoteid" />
            <attribute name="createdon" />
            <attribute name="quotenumber" />
        </link-entity>
    </entity>
</fetch>
*/

var reportName = "QuoteWithHeader"; //Please specify your report name.
var reportId = null;
var fileName = "PDF01"; //Please specify a file which you want to save.
var fromFieldId = null;
var fromFieldEntityName = null;
var toFieldId = null;
var toFieldEntityName = null;
var emailId = null;

//This function is need to be called from action side e.g. Ribbon button click.
function EmailReport() {
    debugger;
    //This function does not work in add mode of form.
    var type = Xrm.Page.ui.getFormType();
    if (type != 1) {
        GetIds(); // Set ids in global variables.
        CreateEmail(); //Create Email
    }
}

//Gets the fromFieldId and toFieldEntityName used to set Email To & From fields
function GetIds() {
    debugger;
    // dev user id
    fromFieldId = "794F8045-05F4-4199-A852-9757172D3B7F";//The Guid which needs to be set in form field.
    fromFieldEntityName = "systemuser";//Please specify entity name for which you have specified above Guid. Most probably it's systemuser.
    toFieldId = "B55FFA62-8197-E711-812D-5065F38C8431"; //The Guid which needs to be set in to field.
    toFieldEntityName = "contact";//Please specify entity name for which you have specified above Guid.
}

//Create Email and link it with Order as Reagrding field
function CreateEmail() {
    debugger;
    var id = Xrm.Page.data.entity.getId();
    id = id.replace('{', "");
    id = id.replace('}', "");
    var entityLogicalName = Xrm.Page.data.entity.getEntityName();
    var regardingObjectId = new Sdk.EntityReference(entityLogicalName, id);

    var email = new Sdk.Entity("email");
    email.addAttribute(new Sdk.String("subject", "Your Booking"));
    email.addAttribute(new Sdk.Lookup("regardingobjectid", regardingObjectId));
    var fromParties = PrepareActivityParty(fromFieldId, fromFieldEntityName);
    email.addAttribute(new Sdk.PartyList("from", fromParties));
    var toParties = PrepareActivityParty(toFieldId, toFieldEntityName);
    email.addAttribute(new Sdk.PartyList("to", toParties));
    Sdk.Async.create(email, EmailCallBack, function (error) { alert(error.message); });

    GetReportId();

}

//This method get entity's id and logical name and return entitycollection of it.
function PrepareActivityParty(partyId, partyEntityName) {
    debugger;
    var activityParty = new Sdk.Entity("activityparty");
    activityParty.addAttribute(new Sdk.Lookup("partyid", new Sdk.EntityReference(partyEntityName, partyId)));
    var activityParties = new Sdk.EntityCollection();
    activityParties.addEntity(activityParty);
    return activityParties;
}

// Email Call Back function
function EmailCallBack(result) {
    debugger;
    emailId = result;
    //  GetReportId();
}

//This method will get the reportId based on a report name

function GetReportId() {
    debugger;
    var context = Xrm.Page.context;
    var serverUrl = context.getClientUrl();
    var ODataPath = serverUrl + "/XRMServices/2011/OrganizationData.svc";
    var retrieveResult = new XMLHttpRequest();

    retrieveResult.open("GET", ODataPath + "/ReportSet?$select=Name,ReportId&$filter=Name eq'" + reportName + "'", false);

    retrieveResult.setRequestHeader("Accept", "application/json");
    retrieveResult.setRequestHeader("Content-Type", "application/json; charset=utf-8?");
    retrieveResult.send();

    if (retrieveResult.readyState == 4 /* complete */) {
        if (retrieveResult.status == 200) {
            var retrieved = this.parent.JSON.parse(retrieveResult.responseText).d;
            var Result = retrieved.results;
            if (typeof Result !== "undefined") {

                reportId = Result[0].ReportId;

                var params = getReportingSession(reportName, reportId);
                EncodePdf(params);
            }
        }
    }
}

//Gets the report contents
function getReportingSession(reportName, reportGuid) {
    debugger;
    reportName = "QuoteWithHeader.rdl";
    var selectedIds = Xrm.Page.data.entity.getId();

    selectedIds = selectedIds.replace('{', '');

    selectedIds = selectedIds.replace('}', '');

    //use this pth for older version
    var pth = Xrm.Page.context.getClientUrl() + "/CRMReports/rsviewer/QuirksReportViewer.aspx";
    //use this pth for version 9.0
    var pth = Xrm.Page.context.getClientUrl() + "/CRMReports/rsviewer/reportviewer.aspx";


    var retrieveEntityReq = new XMLHttpRequest();

    var strParameterXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='quote'><all-attributes /><filter type='and'><condition attribute='quoteid' operator='eq' value='" + selectedIds + "' /> </filter></entity></fetch>";

    retrieveEntityReq.open("POST", pth, false);

    retrieveEntityReq.setRequestHeader("Accept", "*/*");

    retrieveEntityReq.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

    retrieveEntityReq.send("id=%7B" + reportGuid + "%7D&uniquename=" + Xrm.Page.context.getOrgUniqueName() + "&iscustomreport=true&reportnameonsrs=&reportName=" + reportName + "&isScheduledReport=false&p:CRM_quote=" + strParameterXML);

    var x = retrieveEntityReq.responseText.lastIndexOf("ReportSession=");

    var y = retrieveEntityReq.responseText.lastIndexOf("ControlID=");

    var ret = new Array();

    ret[0] = retrieveEntityReq.responseText.substr(x + 14, 24);

    ret[1] = retrieveEntityReq.responseText.substr(x + 10, 32);

    return ret;

}


function EncodePdf(params) {
    debugger;
    var retrieveEntityReq = new XMLHttpRequest();

    var pth = Xrm.Page.context.getClientUrl() + "/Reserved.ReportViewerWebControl.axd?ReportSession=" + params[0] + "&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=" + params[1] + "&OpType=Export&FileName=Public&ContentDisposition=OnlyHtmlInline&Format=PDF";
    retrieveEntityReq.open('GET', pth, true);
    retrieveEntityReq.setRequestHeader("Accept", "*/*");
    retrieveEntityReq.responseType = "arraybuffer";

    retrieveEntityReq.onload = function (e) {
        if (this.status == 200) {
            var uInt8Array = new Uint8Array(this.response);
            var base64 = Encode64(uInt8Array);
            createNote(base64);
        }
    };
    retrieveEntityReq.send();
}

var keyStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
function Encode64(input) {
    debugger;
    var output = new StringMaker();
    var chr1, chr2, chr3;
    var enc1, enc2, enc3, enc4;
    var i = 0;

    while (i < input.length) {
        chr1 = input[i++];
        chr2 = input[i++];
        chr3 = input[i++];

        enc1 = chr1 >> 2;
        enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
        enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
        enc4 = chr3 & 63;

        if (isNaN(chr2)) {
            enc3 = enc4 = 64;
        } else if (isNaN(chr3)) {
            enc4 = 64;
        }
        output.append(keyStr.charAt(enc1) + keyStr.charAt(enc2) + keyStr.charAt(enc3) + keyStr.charAt(enc4));
    }
    return output.toString();
}

var StringMaker = function () {
    this.parts = [];
    this.length = 0;
    this.append = function (s) {
        this.parts.push(s);
        this.length += s.length;
    }
    this.prepend = function (s) {
        this.parts.unshift(s);
        this.length += s.length;
    }
    this.toString = function () {
        return this.parts.join('');
    }
}

//Create attachment for the created email
function CreateEmailAttachment(encodedPdf) {
    debugger;

    //Get order number to name a newly created PDF report
    var QuoteNumber = Xrm.Page.getAttribute("quotenumber");
    var emailEntityReference = new Sdk.EntityReference("email", emailId);
    var newFileName = fileName + ".pdf";
    if (QuoteNumber != null)
        newFileName = fileName + QuoteNumber.getValue() + ".pdf";

    var activitymimeattachment = new Sdk.Entity("activitymimeattachment");
    activitymimeattachment.addAttribute(new Sdk.String("body", encodedPdf));
    activitymimeattachment.addAttribute(new Sdk.String("subject", "File Attachment"));
    activitymimeattachment.addAttribute(new Sdk.String("objecttypecode", "email"));
    activitymimeattachment.addAttribute(new Sdk.String("filename", newFileName));
    activitymimeattachment.addAttribute(new Sdk.Lookup("objectid", emailEntityReference));
    activitymimeattachment.addAttribute(new Sdk.String("mimetype", "application/pdf"));
    Sdk.Async.create(activitymimeattachment, ActivityMimeAttachmentCallBack, function (error) { alert(error.message); });

}

//ActivityMimeAttachment CallBack function
function ActivityMimeAttachmentCallBack(result) {
    debugger;
    Xrm.Utility.openEntityForm("email", emailId);
}
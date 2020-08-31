"use strict";  
var WorkingSY = window.WorkingSY || {};  


WorkingSY.getContext = function() {
    var context;  
    // GetGlobalContext defined by including reference to   
    // ClientGlobalContext.js.aspx in the HTML page.  
    if (typeof GetGlobalContext != "undefined") {  
     context = GetGlobalContext();  
    } else {  
     if (typeof Xrm != "undefined") {  
      // Xrm.Page.context defined within the Xrm.Page object model for form scripts.  
      context = Xrm.Page.context;  
     } else {  
      throw new Error("Context is not available.");  
     }  
    }
    return context;
}


/**  
 * @function getClientUrl   
 * @description Get the client URL.  
 * @return {string} The client URL.  
 */  
WorkingSY.getClientUrl = function () {  
 var context = WorkingSY.getContext();
 
 return context.getClientUrl();  
};  


var clientUrl = WorkingSY.getClientUrl();     // ie.: https://org.crm.dynamics.com  
var webAPIPath = "/api/data/v8.1";      // Path to the web API.  


/**  
 * @function getWebAPIPath   
 * @description Get the full path to the Web API.  
 * @return {string} The full URL of the Web API.  
 */  
WorkingSY.getWebAPIPath = function () {  
    return WorkingSY.getClientUrl() + webAPIPath;  
   }


/**  
* @function request  
* @description Generic helper function to handle basic XMLHttpRequest calls.  
* @param {string} action - The request action. String is case-sensitive.  
* @param {string} uri - An absolute or relative URI. Relative URI starts with a "/".  
* @param {object} data - An object representing an entity. Required for create and update actions.  
* @param {object} addHeader - An object with header and value properties to add to the request  
* @returns {Promise} - A Promise that returns either the request object or an error object.  
*/  
WorkingSY.request = function (action, uri, data, addHeader) {  
if (!RegExp(action, "g").test("POST PATCH PUT GET DELETE")) { // Expected action verbs.  
    throw new Error("WorkingSY.request: action parameter must be one of the following: " +  
        "POST, PATCH, PUT, GET, or DELETE.");  
}  
if (!typeof uri === "string") {  
    throw new Error("WorkingSY.request: uri parameter must be a string.");  
}  
if ((RegExp(action, "g").test("POST PATCH PUT")) && (!data)) {  
    throw new Error("WorkingSY.request: data parameter must not be null for operations that create or modify data.");  
}  
if (addHeader) {  
    if (typeof addHeader.header != "string" || typeof addHeader.value != "string") {  
    throw new Error("WorkingSY.request: addHeader parameter must have header and value properties that are strings.");  
    }  
}  
    
// Construct a fully qualified URI if a relative URI is passed in.  
if (uri.charAt(0) === "/") {  
    uri = clientUrl + webAPIPath + uri;  
}  
    
return new Promise(function (resolve, reject) {  
    var request = new XMLHttpRequest();  
    request.open(action, encodeURI(uri), true);  
    request.setRequestHeader("OData-MaxVersion", "4.0");  
    request.setRequestHeader("OData-Version", "4.0");  
    request.setRequestHeader("Accept", "application/json");  
    request.setRequestHeader("Content-Type", "application/json; charset=utf-8");  
    if (addHeader) {  
    request.setRequestHeader(addHeader.header, addHeader.value);  
    }  
    request.onreadystatechange = function () {  
    if (this.readyState === 4) {  
    request.onreadystatechange = null;  
    switch (this.status) {  
    case 200: // Success with content returned in response body.  
    case 204: // Success with no content returned in response body.  
    case 304: // Success with Not Modified  
        resolve(this);  
        break;  
    default: // All other statuses are error cases.  
        var error;  
        try {  
        error = JSON.parse(request.response).error;  
        } catch (e) {  
        error = new Error("Unexpected Error");  
        }  
        reject(error);  
        break;  
    }  
    }  
    };  
    request.send(JSON.stringify(data));  
});  
};  

WorkingSY.initialize = function () {  
    WorkingSY.loadSYs();
}

WorkingSY.loadSYs = function () {  
    // Get School Years
    return new Promise(function (resolve, reject) {  
        WorkingSY.getSYs().then(function (request) {  
            // Process result
            var results = JSON.parse(request.response);
            var resultsDiv = $('#divSYs');
            resultsDiv.empty();

            resultsDiv.append('<div class="SY"><input type="checkbox" id="chkAllNone" data-id="ALL" disabled onclick="WorkingSY.SaveAllNone(this);" /> ALL/NONE</div>');

            for (var i = 0; i < results.value.length; i++) {
                var schoolYearID = results.value[i]['isfs_schoolyearid'];
                var schoolYearName = results.value[i]['isfs_schoolyearname'];

                resultsDiv.append('<div class="SY"><input type="checkbox" id="chk' + schoolYearID + '" data-id="' + schoolYearID + '" disabled onclick="WorkingSY.SaveYear(this, true);" /> ' + schoolYearName + '</div>');
            }
        
            // Get User's Working School Years
            return WorkingSY.getUserSYs().then(function (request) {
                var results = JSON.parse(request.response);

                var workingYearsCount = $('div.SY > input[type=checkbox]').length - 1;
                var selectedYearsCount = 0;

                for (var i = 0; i < results.value.length; i++) {
                    var schoolYearID = results.value[i]['_isfs_schoolyearid_value'];
                    var workingSchoolYearID = results.value[i]['isfs_workingschoolyearid'];

                    var checkbox = $('div.SY > input[data-id="' + schoolYearID + '"]');
                    checkbox.prop("checked", true);
                    checkbox.data("workingschoolyearid", workingSchoolYearID);
                    selectedYearsCount++;
                }

                if (selectedYearsCount == 0 || selectedYearsCount == workingYearsCount) {
                    $("#chkAllNone").prop("checked", true);
                    if (selectedYearsCount == 0) {
                        $('div.SY > input[type=checkbox]').each(function(){
                            if (($this).attr(id) != 'chkAllNone') {
                                ($this).prop("checked", true);
                                WorkingSY.SaveYear(this, false);
                            }
                        });
                    }
                }

                WorkingSY.EnableDisableForm(true);
                $("#divNotification").empty();
            })
        })
        .then(function (request) {
            resolve();
        })
        .catch(function (err) {  
            reject("Error in WorkingSY.loadUserSYs function: " + err.message);  
        });  
    });  
}

WorkingSY.SaveAllNone = function(checkbox) {
    WorkingSY.EnableDisableForm(false);

    if (checkbox.checked==true) {
    }
    else {
        
    }

    WorkingSY.EnableDisableForm(true);
}

WorkingSY.SaveYear = function(checkbox, showMessage) {
    WorkingSY.EnableDisableForm(false);
    if (checkbox.checked==true) {
        // Create a working SY record for the user
        //alert('Create a working SY record for the user');

        var schoolYearId = $(checkbox).data("id");

        //return WorkingSY.request("POST", uri, workingSchoolYear)
        WorkingSY.saveWorkingSY(schoolYearId)
        .then(function(request) {
            var workingSYId = request.getResponseHeader("OData-EntityId");
            workingSYId = workingSYId.substring(workingSYId.indexOf('(') + 1);
            workingSYId = workingSYId.substring(0, workingSYId.length - 1);
            $(checkbox).data("workingschoolyearid", workingSYId);
            if (showMessage==true) WorkingSY.showMessage('Working year(s) updated');
            WorkingSY.EnableDisableForm(true);
        })
        .catch(function (err) {
            reject("Error in WorkingSY.SaveYear function: " + err.message);  
            WorkingSY.EnableDisableForm(true);
        });  
    }
    else {
        // Delete the user's working SY record
        //alert('Delete the user\'s working SY record');

        var workingSYId = $(checkbox).data("workingschoolyearid");
        var uri = WorkingSY.getWebAPIPath() + '/isfs_workingschoolyears(' + workingSYId + ')';
        WorkingSY.request("DELETE", uri)
        .then(function() {
            if (showMessage==true) WorkingSY.showMessage('Working year(s) updated');
            WorkingSY.EnableDisableForm(true);
        })
        .catch(function (err) {  
            reject("Error in WorkingSY.SaveYear function: " + err.message);  
            WorkingSY.EnableDisableForm(true);
        });  
    }
}

WorkingSY.saveWorkingSY = function(schoolYearId) {
    var context = WorkingSY.getContext();
    var userId = context.userSettings.userId.replace("{","").replace("}","");

    var workingSchoolYear = {
        "isfs_UserId@odata.bind": "/systemusers(" + userId + ")",
        "isfs_SchoolYearId@odata.bind": "/isfs_schoolyears(" + schoolYearId + ")"
    }

    var uri = WorkingSY.getWebAPIPath() + '/isfs_workingschoolyears';
    return WorkingSY.request("POST", uri, workingSchoolYear)
}

WorkingSY.getSYs = function() {
    return WorkingSY.request("GET", WorkingSY.getWebAPIPath() + "/isfs_schoolyears?$select=isfs_schoolyearid,isfs_schoolyearname&$orderby=isfs_startdate desc");
}

WorkingSY.getUserSYs = function() {
    var context = WorkingSY.getContext();
    var userId = context.userSettings.userId.replace("{","").replace("}","");

    return WorkingSY.request("GET", WorkingSY.getWebAPIPath() + "/isfs_workingschoolyears?$select=isfs_workingschoolyearid,_isfs_schoolyearid_value,_isfs_userid_value&$filter=_isfs_userid_value eq " + userId);
}

WorkingSY.showMessage = function(message) {
    $("#divNotification").empty();
    $("#divNotification").append("<div>" + message + "</div>").fadeIn();
    setTimeout(function(){
        $("#divNotification > div").fadeOut();
    }, 1500);
}

WorkingSY.EnableDisableForm = function(enable) {
    $("div.SY > input[type=checkbox]").prop("disabled", !enable);
}
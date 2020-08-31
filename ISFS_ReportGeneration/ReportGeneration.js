// 2020 - Mike McGuinness - ITK Consulting - BC Ministry of Education
'use strict';

// Function used by the grid toolbar to load the custom dialog window
function loadReportGenerator(primaryControl) {
	var context = primaryControl;
	var globalContext = Xrm.Utility.getGlobalContext();

	var windowOptions = { height: 260, width: 520 };

	var parameters =
		'orgName=' + globalContext.organizationSettings.uniqueName +
		',userId=' + globalContext.userSettings.userId +
		',dateSeparator=' + globalContext.userSettings.dateFormattingInfo.DateSeparator +
		',dateShortDatePattern=' + globalContext.userSettings.dateFormattingInfo.ShortDatePattern;

	Xrm.Navigation.openWebResource("isfs_GenerateReports_html", windowOptions, encodeURIComponent(parameters));
}

// Function used to format the date in the user's specified format
Date.prototype.formatForCRM = function FormatDate() {
	var formatedDate = '';

	var dateparts = _DateShortDatePattern.toLowerCase().split('/');
	if (dateparts.length === 0) dateparts = _DateShortDatePattern.toLowerCase().split('-');

	for (var i = 0; i < dateparts.length; i++) {
		if (i > 0) formatedDate += _DateSeparator;
		if (dateparts[i].indexOf('d') > -1) {
			var day = this.getDate();
			if (dateparts[i].length === 2 && day.toString().length === 1) formatedDate += '0';

			formatedDate += day.toString();
		}
		else if (dateparts[i].indexOf('m') > -1) {
			var month = (this.getMonth() + 1);
			if (dateparts[i].length === 2 && month.toString().length === 1) formatedDate += '0';

			formatedDate += month.toString();
		}
		else if (dateparts[i].indexOf('y') > -1) {
			formatedDate += this.getFullYear();
		}
    }

	return formatedDate;
}

// Global Variables
var _TotalCount = 0;
var _CurrentCount = 0;
var _Zip = new JSZip();
var _Complete = false;
var _OrgName = '';
var _UserId = '';
var _DateSeparator = '/';
var _DateShortDatePattern = 'yyyy/MM/dd';
var _ReportGenerating = false;

function getXrmObj() {
	var xrm;

	if (window.opener && window.opener.Xrm) xrm = window.opener.Xrm;
	else if (window.parent && window.parent.Xrm) xrm = window.parent.Xrm;
	else xrm = null;

	return xrm;
}

// Function used to load form parameters from URL
function initializeForm() {
	var params = decodeURIComponent(getUrlVars()['data']);
	params = decodeURIComponent(params);
	if (params !== null && params.length > 0) {
		var paramArray = params.split(',');

		var param = paramArray[0].split('=');
		_OrgName = param[1];

		param = paramArray[1].split('=');
		_UserId = param[1];

		param = paramArray[2].split('=');
		_DateSeparator = param[1];

		param = paramArray[3].split('=');
		_DateShortDatePattern = param[1];
	}
}

function loadSchoolYears() {
	initializeForm();
	$('#btnGenerateReports').hide();

	var fetchXml = '<fetch distinct="true" useraworderby="false" no-lock="false" mapping="logical">' +
		'<entity name="isfs_schooldisbursement">' +
			'<link-entity name="isfs_fundingschedule" to="isfs_fundingschedule" from="isfs_fundingscheduleid" link-type="inner" alias="FundingSchedule">' +
				'<link-entity name="isfs_schoolyear" to="isfs_schoolyear" from="isfs_schoolyearid" link-type="inner" alias="SY">' +
					'<attribute name="isfs_schoolyearid" />' +
					'<attribute name="isfs_schoolyearname" />' +
				'</link-entity>' +
			'</link-entity>' +
		'</entity>' +
	'</fetch>';

	retrieveWebAPIDataFetchQuery('isfs_schooldisbursements', fetchXml, renderSchools);
}

function renderSchools(xhr) {
	var schoolYears = JSON.parse(xhr.responseText, dateReviver);

	if (schoolYears !== null && schoolYears.value !== null && schoolYears.value.length > 0) {
		var ddlSY = $('#ddlSY');
		for (var i = 0; i < schoolYears.value.length; i++) {
			ddlSY.append('<option value="' + schoolYears.value[i]['SY.isfs_schoolyearid'] + '">' + schoolYears.value[i]['SY.isfs_schoolyearname'] + '</option>');
		}
		ddlSY.prop('disabled', false);
	}
}

function loadGrantPrograms() {
	var syID = $('#ddlSY').val();

	$('#btnGenerateReports').prop('disabled', true);
	$('#btnGenerateReports').hide();
	resetDDL('ddlGrantProgram', 'Grant Program');
	resetDDL('ddlDisbursementDate', 'Disbursement Date');
	resetDDL('ddlAuthority', 'Authority');
	setStatus('Select report parameters.');

	if (syID !== null && syID.length > 0) {
		var fetchXml =
			'<fetch distinct="true" useraworderby="false" no-lock="false" mapping="logical">' +
				'<entity name="isfs_schooldisbursementdetail">' +
					'<link-entity name="isfs_fundingscheduledetail" to="isfs_fundingscheduledetail" from="isfs_fundingscheduledetailid" link-type="inner"></link-entity>' +
					'<link-entity name="isfs_schooldisbursement" to="isfs_schooldisbursement" from="isfs_schooldisbursementid" link-type="inner">' +
						'<filter type="and">' +
							'<condition attribute="isfs_ignoresyfilter" value="1" operator="ne" />' +
							'<condition attribute="statuscode" operator="eq" value="746910000" />' +
						'</filter>' +
						'<link-entity name="isfs_fundingschedule" to="isfs_fundingschedule" from="isfs_fundingscheduleid" link-type="inner" alias="FS">' +
							'<link-entity name="isfs_grantprogram" link-type="inner" to="isfs_grantprogram" from="isfs_grantprogramid" alias="GP">' +
								'<attribute name="isfs_grantprogramid" />' +
								'<attribute name="isfs_name" />' +
								'<link-entity name="isfs_schoolyear" link-type="inner" to="isfs_schoolyear" from="isfs_schoolyearid">' +
									'<filter type="and">' +
										'<condition attribute="isfs_schoolyearid" operator="eq" value="' + syID + '" />' +
									'</filter>' +
								'</link-entity>' +
							'</link-entity>' +
						'</link-entity>' +
						'<link-entity name="isfs_fundingrate" to="isfs_fundingrate" from="isfs_fundingrateid" link-type="inner" alias="FundingRate"></link-entity>' +
					'</link-entity>' +
				'</entity>' +
			'</fetch>';

		retrieveWebAPIDataFetchQuery('isfs_schooldisbursementdetails', fetchXml, renderGrantPrograms);
	}
}

function renderGrantPrograms(xhr) {
	var grantPrograms = JSON.parse(xhr.responseText, dateReviver);

	if (grantPrograms !== null && grantPrograms.value !== null && grantPrograms.value.length > 0) {
		var ddlGrantProgram = $('#ddlGrantProgram');
		for (var i = 0; i < grantPrograms.value.length; i++) {
			ddlGrantProgram.append('<option value="' + grantPrograms.value[i]['GP.isfs_grantprogramid'] + '">' + grantPrograms.value[i]['GP.isfs_name'] + '</option>');
		}
		ddlGrantProgram.prop('disabled', false);
	}
}

function loadDisbursementDates() {
	var syID = $('#ddlSY').val();
	var grantProgramID = $('#ddlGrantProgram').val();

	$('#btnGenerateReports').prop('disabled', true);
	$('#btnGenerateReports').hide();
	resetDDL('ddlDisbursementDate', 'Disbursement Date');
	resetDDL('ddlAuthority', 'Authority');
	setStatus('Select report parameters.');

	if (syID !== null && grantProgramID !== null && syID.length > 0 && grantProgramID.length > 0) {
		var fetchXml =
			'<fetch distinct="true" useraworderby="false" no-lock="false" mapping="logical">' +
				'<entity name="isfs_schooldisbursementdetail">' +
					'<link-entity name="isfs_fundingscheduledetail" to="isfs_fundingscheduledetail" from="isfs_fundingscheduledetailid" link-type="inner"></link-entity>' +
					'<link-entity name="isfs_schooldisbursement" to="isfs_schooldisbursement" from="isfs_schooldisbursementid" link-type="inner">' +
						'<filter type="and">' +
							'<condition attribute="isfs_ignoresyfilter" value="1" operator="ne" />' +
							'<condition attribute="statuscode" operator="eq" value="746910000" />' +
						'</filter>' +
						'<link-entity name="isfs_fundingschedule" to="isfs_fundingschedule" from="isfs_fundingscheduleid" link-type="inner" alias="FS">' +
							'<attribute name="isfs_disbursementdate" />' +
							'<filter type="and">' +
								'<condition attribute="isfs_grantprogram" operator="eq" value="' + grantProgramID + '" />' +
							'</filter>' +
						'</link-entity>' +
						'<link-entity name="isfs_fundingrate" to="isfs_fundingrate" from="isfs_fundingrateid" link-type="inner" alias="FundingRate"></link-entity>' +
					'</link-entity>' +
				'</entity>' +
			'</fetch>';

		retrieveWebAPIDataFetchQuery('isfs_schooldisbursementdetails', fetchXml, renderDisbursementDates);
	}
}

function renderDisbursementDates(xhr) {
	var disbursementDates = JSON.parse(xhr.responseText, dateReviver);

	if (disbursementDates !== null && disbursementDates.value !== null && disbursementDates.value.length > 0) {
		var ddlDisbursementDate = $('#ddlDisbursementDate');
		var lastDate = new Date();
		for (var i = 0; i < disbursementDates.value.length; i++) {
			var distDate = new Date(disbursementDates.value[i]['FS.isfs_disbursementdate']);
			if (distDate.getTime() !== lastDate.getTime()) {
				ddlDisbursementDate.append('<option value="' + disbursementDates.value[i]['FS.isfs_disbursementdate'] + '">' + distDate.formatForCRM() + '</option>');
				lastDate = distDate;
			}
		}
		ddlDisbursementDate.prop('disabled', false);
	}
}

function loadAuthorities() {
	var syID = $('#ddlSY').val();
	var grantProgramID = $('#ddlGrantProgram').val();
	var disbursementDate = new Date($('#ddlDisbursementDate').val());

	$('#btnGenerateReports').prop('disabled', true);
	$('#btnGenerateReports').hide();
	resetDDL('ddlAuthority', 'Authority');
	setStatus('Select report parameters.');

	var fetchXml =
		'<fetch distinct="true" useraworderby="false" no-lock="false" mapping="logical">' +
			'<entity name="isfs_schooldisbursementdetail">' +
				'<link-entity name="isfs_fundingscheduledetail" to="isfs_fundingscheduledetail" from="isfs_fundingscheduledetailid" link-type="inner"></link-entity>' +
				'<link-entity name="isfs_schooldisbursement" to="isfs_schooldisbursement" from="isfs_schooldisbursementid" link-type="inner">' +
					'<filter type="and">' +
						'<condition attribute="isfs_ignoresyfilter" value="1" operator="ne" />' +
						'<condition attribute="statuscode" operator="eq" value="746910000" />' +
					'</filter>' +
					'<link-entity name="isfs_fundingschedule" to="isfs_fundingschedule" from="isfs_fundingscheduleid" link-type="inner">' +
						'<filter type="and">' +
							'<condition attribute="isfs_disbursementdate" operator="on" value="' + disbursementDate.formatForCRM() + '"></condition>' +
							'<condition attribute="isfs_grantprogram" operator="eq" value="' + grantProgramID + '" />' +
						'</filter>' +
					'</link-entity>' +
					'<link-entity name="isfs_fundingrate" to="isfs_fundingrate" from="isfs_fundingrateid" link-type="inner" alias="FundingRate"></link-entity>' +
					'<link-entity name="edu_school" to="isfs_school" from="edu_schoolid" link-type="inner" alias="School"></link-entity>' +
					'<link-entity name="isfs_authority" to="isfs_authority" from="isfs_authorityid" link-type="inner" alias="Authority">' +
						'<attribute name="isfs_authorityid" />' +
						'<attribute name="isfs_authorityno" />' +
					'</link-entity>' +
				'</link-entity>' +
			'</entity>' +
		'</fetch>';

	if (syID !== null && grantProgramID !== null && disbursementDate !== null && syID.length > 0 && grantProgramID.length > 0) {
		retrieveWebAPIDataFetchQuery('isfs_schooldisbursementdetails', fetchXml, renderAuthorityResults);
	}
}

function renderAuthorityResults(xhr) {
	var authorities = JSON.parse(xhr.responseText, dateReviver);
	
	if (authorities !== null && authorities.value !== null && authorities.value.length > 0) {
		var ddlAuthority = $('#ddlAuthority');
		setStatus(authorities.value.length + ' authorities found.');
		for (var i = 0; i < authorities.value.length; i++) {
			ddlAuthority.append('<option value="' + authorities.value[i]['Authority.isfs_authorityid'] + '">' + authorities.value[i]['Authority.isfs_authorityno'] + '</option>');
		}
		ddlAuthority.prop('disabled', false);
		enableReportGeneration();
	}
	else setStatus('No matching authorities found for this date.');
}

function setStatus(status) {
	$('#divStatus').text(status);
}

function enableReportGeneration() {
	var syID = $('#ddlSY').val();
	var grantProgramID = $('#ddlGrantProgram').val();
	var disbursementDate = new Date($('#ddlDisbursementDate').val());
	var ddlAuthority = $('#ddlAuthority').val();
	
	var btnGenerateReports = $('#btnGenerateReports');

	if (syID !== null && grantProgramID !== null && disbursementDate !== null && ddlAuthority !== null && syID.length > 0 && grantProgramID.length > 0) {
		btnGenerateReports.prop('disabled', false);
		btnGenerateReports.show();
	}
	else {
		btnGenerateReports.prop('disabled', true);
		btnGenerateReports.hide();
	}
}

function resetDDL(id, displayName) {
	var ddlGrantProgram = $('#' + id);
	ddlGrantProgram.prop('disabled', true);
	ddlGrantProgram.val('');
	ddlGrantProgram.empty();
	ddlGrantProgram.append('<option value="">-- Select ' + displayName + ' --</option>');
}

function generateReports() {
	setStatus('Preparring reports...');

	var syID = $('#ddlSY').val();
	var syName = $("#ddlSY option:selected").text();
	var grantProgramID = $('#ddlGrantProgram').val();
	var grantProgramName = $("#ddlGrantProgram option:selected").text();
	var disbursementDate = new Date($('#ddlDisbursementDate').val());
	var dispDate = (new Date(disbursementDate)).formatForCRM();
	var authorityOptions = $('#ddlAuthority option');

	_TotalCount = authorityOptions.length - 1;
	_CurrentCount = 1;

	if (syID !== null && grantProgramID !== null && disbursementDate !== null && ddlAuthority !== null && syID.length > 0 && grantProgramID.length > 0) {
		setStatus('Generating ' + _TotalCount + ' reports, please wait...');
		var btnGenerateReports = $('#btnGenerateReports');
		btnGenerateReports.prop('disabled', true);
		btnGenerateReports.hide();

		var alertStrings = { confirmButtonLabel: 'Ok', text: _TotalCount + ' reports will now be generated for download. With a large number of reports, this may take a few moments. Click OK to begin.', title: 'Run Reports:' };
		var alertOptions = { height: 120, width: 260 };
		var xrm = getXrmObj();
		xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
			function success(result) {
				_Complete = false;
				_Zip = new JSZip();

				window.setTimeout(runReports, 50, syName, grantProgramID, grantProgramName, dispDate, disbursementDate);
			},
			function (error) {
				//console.log(error.message);
			}
		);
	}
	else {
		var alertStrings = { confirmButtonLabel: 'Ok', text: 'Please select a School Year, Grant Program and Disbursement Date.', title: 'Validation:' };
		var alertOptions = { height: 120, width: 260 };
		var xrm = getXrmObj();
		xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
	}
}

function runReports(syName, grantProgramID, grantProgramName, dispDate, disbursementDate) {
	var authorityOption = $('#ddlAuthority :nth-child(2)');
	var authorityId = authorityOption.val(); //$(this).val();
	var authorityNum = authorityOption.text(); //$(this).text();
	if (authorityNum !== null && authorityId.length > 0) {
		setStatus('Generating report ' + _CurrentCount + ' of ' + _TotalCount + ' reports.');
		executeReport(syName, grantProgramID, grantProgramName, dispDate, disbursementDate, authorityNum, authorityId, convertReport); //dispDate
	}
}

function convertReport(req, syName, grantProgramID, grantProgramName, dispDate, disbursementDate, authorityNum, authorityId) {
	//These variables captures the response and returns the response in an array.
	var x = req.responseText.lastIndexOf('ReportSession=');
	var y = req.responseText.lastIndexOf('ControlID=');

	var reportResponse = new Array();
	reportResponse[0] = req.responseText.substr(x + 14, 24);
	reportResponse[1] = req.responseText.substr(y + 10, 32);

	convertResponseToPDF(reportResponse, syName, grantProgramID, grantProgramName, dispDate, (new Date(disbursementDate)).formatForCRM(), authorityNum, authorityId);
}

function getClientURL() {
	return window.location.protocol + '//' + window.location.host.split(":")[0]; //Xrm.Page.context.getClientUrl();
}

function retrieveWebAPIDataFetchQuery(sEntityName, sFetchXML, callback) {
	sFetchXML = encodeURI(sFetchXML);

	var serverUrl = getClientURL();

	var ODATA_ENDPOINT = '/api/data/v9.1';

	var odataUri = serverUrl + ODATA_ENDPOINT + '/' + sEntityName + '?fetchXml=' + sFetchXML;

	var data;

	try {
		var req = new XMLHttpRequest();

		req.onload = function () {
			if (req.readyState === 4) {
				if (req.status === 200) {
					callback(req);
				} else {
					errorAlert('Error: retrieveWebAPIData - ' + req.statusText, 'Error');
				}
			}
		};
		req.open('GET', encodeURI(odataUri), true);
		req.setRequestHeader('Accept', 'application/json');
		req.setRequestHeader('Content-Type', 'application/json; charset=utf-8');
		req.send(null);
	}
	catch (ex) {
		errorAlert('Error: retrieveWebAPIData - Data Set =' + odataSetName + '; filter = ' + sFilter + '; select = ' + sSelect + '; Error = ' + ex.message, 'Error');
	}
}

function errorAlert(message) {
	var errorOptions = { message: message };

	var xrm = getXrmObj();
	xrm.Navigation.openErrorDialog(errorOptions);
}

function dateReviver(key, value) {
	if (typeof value === 'string') {
		var a = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*)?)Z$/.exec(value);
		if (a) {
			return new Date(Date.UTC(+a[1], +a[2] - 1, +a[3], +a[4], +a[5], +a[6]));
		}
	}
	return value;
};

function getFormContext(executionContext) {
     var formContext = null;
     if (executionContext !== null) {
         if (typeof executionContext.getAttribute === 'function') {
             formContext = executionContext; //most likely called from the ribbon.
         } else if (typeof executionContext.getFormContext === 'function' 
                 && typeof(executionContext.getFormContext()).getAttribute === 'function') {
            formContext = executionContext.getFormContext(); // most likely called from the form via a handler
         } else {
            throw 'formContext was not found'; //you could do formContext = Xrm.Page; if you like.
        }
    }
    return formContext;
}

function executeReport(syID, grantProgramID, grantProgramName, dispDate, disbursementDate, authorityNum, authorityId, callback) {
	// GUID of SSRS report in CRM.
	var reportGuid = '';

	var program = grantProgramName.toLowerCase();
	if (program.indexOf('distributed learning') > -1) reportGuid = '3d48c0d7-b0b9-ea11-a812-000d3a0c86a9';
	else reportGuid = '1668f008-05a5-ea11-a812-000d3a0c8a65'; //if (program.indexOf('first nations') > -1 || program.indexOf('independent school') > -1 || program.indexOf('special education') > -1)
	//else {
	//	alert('Grant Program "' + grantProgramName + '" is not currently supported for automated funding statement generation.');
	//	return;
    //}

	//Name of the report. Note: .RDL needs to be specified.
	var reportName = 'DL Funding Statement.rdl';

	// URL of the report server which will execute report and generate response.
	var pth = getClientURL() + '/CRMReports/rsviewer/reportviewer.aspx'; //'/CRMReports/rsviewer/QuirksReportViewer.aspx';

	//This is the filter that is passed to pre-filtered report. It passes GUID of the record using Xrm.Page.data.entity.getId() method.
	//This filter shows example for quote report. If you want to pass ID of any other entity, you will need to specify respective entity name.
	//var reportPrefilter = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false"><entity name="quote"><all-attributes /><filter type="and"><condition attribute="quoteid" operator="eq" value="' + Xrm.Page.data.entity.getId() + '" /></filter></entity></fetch>';

	//Prepare query to execute report.
	var query = 'id=%7B' + reportGuid + '%7D&uniquename=' + _OrgName +
		'&iscustomreport=true&reportnameonsrs=&reportName=' + encodeURI(reportName) + '&isScheduledReport=false' +
		'&p:SY=' + encodeURI(syID) + '&p:GrantProgram=' + encodeURI(grantProgramID) + '&p:DisbursementDate=' + encodeURI(dispDate);

	if (grantProgramName.toLowerCase().indexOf('distributed learning') >- 1)
		query += '&p:Authority=' + authorityNum;
	else
		query += '&p:AuthorityNumber=' + authorityNum;

	//Prepare request object to execute the report.
	var req = new XMLHttpRequest();

	req.onload = function () {
		if (req.readyState === 4) {
			if (req.status === 200) {
				callback(req, syID, grantProgramID, grantProgramName, dispDate, disbursementDate, authorityNum, authorityId);
			} else {
				errorAlert('Error: retrieveWebAPIData - ' + req.statusText, 'Error');
			}
		}
	};

	req.open('POST', pth, true);
	req.setRequestHeader('Accept', '*/*');
	req.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	req.send(query); //This statement runs the query and executes the report synchronously.
}

function convertResponseToPDF(arrResponseSession, syID, grantProgramID, grantProgramName, dispDate, disbursementDate, authorityNum, authorityId) {
    //Create query string that will be passed to Report Server to generate PDF version of report response.
	var pth = getClientURL() + '/Reserved.ReportViewerWebControl.axd?' +
		'ReportSession=' + arrResponseSession[0] +
		'&Culture=1033' +
		'&CultureOverrides=True' +
		'&UICulture=1033' +
		'&UICultureOverrides=True' +
		'&ReportStack=1' +
		'&ControlID=' + arrResponseSession[1] +
		'&OpType=Export' +
		'&FileName=Public' +
		'&ContentDisposition=OnlyHtmlInline' +
		'&Format=PDF';
		//'&p:SY=' + encodeURI(syID) + '&p:GrantProgram=' + encodeURI(grantProgramID) + '&p:DisbursementDate=' + encodeURI(disbursementDate) + '&p:Authority=' + authorityNum;

    //Create request object that will be called to convert the response in PDF base 64 string.
	var retrieveEntityReq = new XMLHttpRequest();

	var authId = authorityId;
	var authNo = authorityNum;
	var zipName = grantProgramName + ' Statements_' + syID + '_Disbursement-' + disbursementDate;
	var reportName = grantProgramName + ' Statement_' + syID + '_Disbursement-' + disbursementDate + '_Authority-' + authorityNum;
	reportName = reportName.replace('/', '-');

	retrieveEntityReq.open('GET', pth, true);
	retrieveEntityReq.timeout = '';
    retrieveEntityReq.setRequestHeader('Accept', '*/*');
	retrieveEntityReq.responseType = 'arraybuffer';
	retrieveEntityReq.onreadystatechange = function () { // This is the callback function.
		if (retrieveEntityReq.readyState === 4 && retrieveEntityReq.status === 200) {
			setStatus('Saving report ' + _CurrentCount + ' of ' + _TotalCount + ' for Authority ' + authNo);

			var binary = '';
			var bytes = new Uint8Array(this.response);

			for (var i = 0; i < bytes.byteLength; i++) {
				binary += String.fromCharCode(bytes[i]);
			}

			//This is the base 64 PDF formatted string and is ready to pass to the action as an input parameter.
			var base64PDFString = btoa(binary);

			if (retrieveEntityReq.responseURL.indexOf('errorhandler.aspx') === -1)
				reportName += '.pdf';
			else 
				reportName += '.html';

			_Zip.file(reportName, base64PDFString, { base64: true });

			//4. Call Action and pass base 64 string as an input parameter. That's it.
			createNote(base64PDFString, authId, reportName);

			if (_CurrentCount === _TotalCount) {
				_Complete = true;
				setStatus(_TotalCount + ' reports created.');

				window.setTimeout(downloadZip, 100, zipName);

				var btnGenerateReports = $('#btnGenerateReports');
				btnGenerateReports.prop('disabled', false);
				btnGenerateReports.show();
			}
			else {
				_CurrentCount++;

				var authorityOption = $('#ddlAuthority :nth-child(' + (_CurrentCount + 1) + ')');

				var authorityId = authorityOption.val(); //$(this).val();
				var authorityNum = authorityOption.text(); //$(this).text();
				if (authorityNum !== null && authorityId.length > 0) {
					setStatus('Generating report ' + _CurrentCount + ' of ' + _TotalCount + ' reports.');
					executeReport(syID, grantProgramID, grantProgramName, dispDate, disbursementDate, authorityNum, authorityId, convertReport);
				}
            }

			
		}
	};

	//This statement sends the request for execution asynchronously. Callback function will be called on completion of the request.
	retrieveEntityReq.send();

	_ReportGenerating = false;
}

function downloadZip(zipName) {
	var fileName = zipName;
	_Zip.generateAsync({ type: 'blob' })
		.then(function (content) {
			// Force down of the Zip file
			//window.location = "data:application/zip;base64," + content;
			saveAs(content, fileName + '.zip');
		});
}

function createNote(pdfFile, authorityId, reportName) {
	var note = {};

	note.isdocument = true;
	note.objecttypecode = 'isfs_authority';
	//note['ownerid@odata.bind'] = '/systemusers(E94126AC-64FB-E211-9BED-005056920E6D)';
	//note['owneridtype'] = 8;
	note.documentbody = pdfFile;
	note.mimetype = 'application/pdf';
	note.filename = reportName;
	note.subject = reportName;
	note['objectid_isfs_authority@odata.bind'] = '/isfs_authorities(' + authorityId + ')';

	try {
		var odataUri = getClientURL() + '/api/data/v9.1/annotations';

		$.ajax({
			type: "POST",
			contentType: "application/json; charset=utf-8",
			datatype: "json",
			url: odataUri,
			data: JSON.stringify(note),
			beforeSend: function (XMLHttpRequest) {
				XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
				XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
				XMLHttpRequest.setRequestHeader("Accept", "application/json");
			},
			async: true,
			success: function (data, textStatus, xhr) {
				var uri = xhr.getResponseHeader("OData-EntityId");
				var regExp = /\(([^)]+)\)/;
				var matches = regExp.exec(uri);
				var newEntityId = matches[1];
			},
			error: function (xhr, textStatus, errorThrown) {
				errorAlert(textStatus + " " + errorThrown, 'Error');
			}
		});
	}
	catch (ex) {
		errorAlert('Error: retrieveWebAPIData - ' + ex.message, 'Error');
	}
}

function getUrlVars() {
	var vars = [], hash;
	var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
	for (var i = 0; i < hashes.length; i++) {
		hash = hashes[i].split('=');
		vars.push(hash[0]);
		vars[hash[0]] = hash[1];
	}
	return vars;
}
var attachementSizeInKB=0;
function checkFileSize(attachement) {
	try {
		var attachementSize=0;
		//var attachement = document.getElementById('attachementLotFileID');
		var file = attachement.files[0];
		var bytes = file.size;
		var sizes = [ 'Bytes', 'KB', 'MB', 'GB', 'TB' ];
		
		if (bytes == 0)
			return 'n/a';
		var i = parseInt(Math.floor(Math.log(bytes) / Math.log(1024)));
		if (sizes[i] == "MB") {
			attachementSize = Math.round(bytes / Math.pow(1024, i), 2);
		} 
		else if(sizes[i] == "KB") {
			attachementSizeInKB = Math.round(bytes / Math.pow(1024, i), 2);
		}
		else if(sizes[i] == "Bytes") {
			attachementSizeInKB = bytes;
		}
		else {
			attachementSize = 0;
		}
	} catch (e) {
		attachementSize = 0;
	}
	return attachementSize;
}


var validateForFileSize_message = "File size exceed.";
var validateForFileSize_isExceed = false;
function validateForFileSize(inputFileObject, mbSize) {
	if (getFileSize(inputFileObject) > mbSize) {
		validateForFileSize_isExceed = true;
		if (typeof showErrorMessage == 'function') {
			showErrorMessage(validateForFileSize_message);
		} else {
			alert(validateForFileSize_message);
		}
	} else {
		validateForFileSize_isExceed = false;
	}
}
function removeOtherThenNumber(string) 
{	
	var tempString = '';
	var achar;
	for ( var n = 0; n < string.length; n++) {
		achar=string.charAt(n)
		switch (achar) {
			case '0' :
			case '1' :
			case '2' :
			case '3' :
			case '4' :
			case '5' :
			case '6' :
			case '7' :
			case '8' :
			case '9' :
				tempString = tempString+achar ;
				break;
			default :
				//ignore others;
				break;
		}
	}

	return tempString;
}

function isLeftClick() {
	if (!e)
		var e = window.event;
	if (navigator.appName == 'Netscape' && (e.which == 3 || e.which == 2))
		return true;
	else if (navigator.appName == 'Microsoft Internet Explorer'
			&& (event.button == 2 || event.button == 3))
		return true;
	return false;
}

function doProperURL(url) {

	while (url.indexOf('+') != -1) {
		url = url.replace('+', '%2b');
	}
	return url;
}

function openWindow(url, windowName, feature) {
	window.open(url, windowName, feature);
}
// openWindow("<%=SysSetting.CONTEXT%>/manageRFI.do?action=reqShowAnalytics&projectID="+projectID,'Show','status=yes,scrollbars=yes,resizable=yes');
// openWindow('<%=SysSetting.CONTEXT%>/manageRFIQuestionFile.do?action=<%=ActionConst.FILE_UPLOAD%>&projectID=<%=request.getParameter("projectID")%>&isCreate=1&qID=' + rfiQID + '&modify='+document.form1.modifyFlag.value+'&addendum=1','six16','width=900,height=450')
function ignoreSingleQuote(str) {
	str = replaceAllSingleChar(str, "'", "\\'");
	return str;
}
function ignoreDoubleQuote(str) {
	str = replaceAllSingleChar(str, "\"", "\\\"");
	return str;
}

function replaceAllSingleChar(str, charWhat, strWith) {
	var tempString = '';
	var tmpCH;
	var len = str.length
	for ( var n = 0; n < len; n++) {
		tmpCH = str.charAt(n);
		if (charWhat == tmpCH) {
			tempString = tempString + strWith;
		} else {
			tempString = tempString + tmpCH;
		}
	}
	return tempString;
}
function trimString(str) {
	while (str.charAt(0) == ' ')
		str = str.substring(1);
	while (str.charAt(str.length - 1) == ' ')
		str = str.substring(0, str.length - 1);
	return str;
}
function removePreceedingZero(num) {
	var str = num + "";
	if ((str.charAt(0) == '0' && str.charAt(1) == '.')
			|| (str.charAt(0) == '0' && str.charAt(1) == ',')) {
		return str;
	} else {
		while (str.charAt(0) == '0')
			str = str.substring(1);
		return str;
	}
}

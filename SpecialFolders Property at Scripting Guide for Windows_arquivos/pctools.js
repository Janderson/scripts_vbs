//if (top.location != location) {
//	top.location.href = document.location.href ;
//}

function getURL(uri) {
	uri.dir = location.href.substring(0, location.href.lastIndexOf('\/'));
	uri.protocol = window.location.protocol;
	uri.dom = uri.dir.substr(uri.protocol.length + 2);
	uri.host = window.location.host;
	uri.path = '';
	var pos = uri.dom.indexOf('\/'); if (pos > -1) {uri.path = uri.dom.substr(pos+1); uri.dom = uri.dom.substr(0,pos);}
	uri.page = location.href.substring(uri.dir.length+1, location.href.length+1);
	pos = uri.page.indexOf('?');if (pos > -1) {uri.page = uri.page.substring(0, pos);}
	pos = uri.page.indexOf('#');if (pos > -1) {uri.page = uri.page.substring(0, pos);}
	uri.ext = ''; pos = uri.page.indexOf('.');if (pos > -1) {uri.ext =uri.page.substring(pos+1); uri.page = uri.page.substr(0,pos);}
	uri.file = uri.page;
	if (uri.ext != '') uri.file += '.' + uri.ext;
	if (uri.file == '') uri.page = 'index';
	uri.args = location.search.substr(1).split("?");
	return uri;
}

function getParam(name,type) {
	name = name.replace(/[\[]/,"\\\[").replace(/[\]]/,"\\\]");
	if(type == "url") {
		var regexS = "[\\/]"+name+"\/([^\/\?&#]*)";
	} else {
		var regexS = "[\\?&]"+name+"=([^&#]*)";
	}
	var regex = new RegExp( regexS );
	var results = regex.exec( window.location.href );
	if( results == null ) {
		return "";
	} else {
		return results[1];
	}
}

function getRef(p) {
	if(!p || p == "") {
		p = 'ref';
	}
	var ref = getParam(p);
	if(ref == "") {
		ref = getParam(p,'url')
	}
	return ref;
}

var doc_loc = new String(document.location);
if (window.top.location !== document.location) {
	if (doc_loc.indexOf('google_pack')==-1 && doc_loc.indexOf('privacypolicy')==-1){
		window.top.location.replace(document.location.href);
	}
}


function show_sidebar(id) {
	if (document.getElementById(id+"_content").style.display == "none") {
		document.getElementById(id+"_content").style.display = ""; // Should this be 'block'?
		document.getElementById(id+"_cornerleft").src = "/res/images/sidebar/px.gif";
		document.getElementById(id+"_cornerright").src = "/res/images/sidebar/px.gif";
		document.getElementById(id+"_collapsebtn").src = "/res/images/sidebar/collapse.gif";
	} else {
		document.getElementById(id+"_content").style.display = "none";
		document.getElementById(id+"_cornerleft").src = "/res/images/sidebar/top_left_bottom_corner.gif";
		document.getElementById(id+"_cornerright").src = "/res/images/sidebar/top_right_bottom_corner.gif";
		document.getElementById(id+"_collapsebtn").src = "/res/images/sidebar/expand.gif";
	}
}


function validate(array, input_type) {
	switch (input_type) {
		case 'text':
		for(i=0;i<array.length;i++) {
			document.getElementById("validate_"+array[i]).value="/.+$/";
		}
		break;
		case 'select':
		for(i=0;i<array.length;i++) {
			document.getElementById("validate_"+array[i]).value="count=1";
		}
		break;
		default:
		for(i=0;i<array.length;i++) {
			document.getElementById("validate_"+array[i]).value="/.+$/";
		}
		break;
	}
}

function do_not_validate(array, input_type) {
	switch (input_type) {
		case 'text' :
		for(i=0;i<array.length;i++) {
			document.getElementById("validate_"+array[i]).value="";
		}
		break;
		case 'select':
		for(i=0;i<array.length;i++) {
			document.getElementById("validate_"+array[i]).value="count=0";
		}
		break;
		default:
		for(i=0;i<array.length;i++) {
			document.getElementById("validate_"+array[i]).value="";
		}
		break;
	}
}

function check_sc_select() {
	var no_selected = false;
	var temp = document.getElementById("sc_select");
	for(i=0;i<temp.length;i++) {
		if (temp[i].value == 'no' && temp[i].selected==true) {
			no_selected = true;
		}
	}
	if (no_selected == true) {
		for(i=0;i<temp.length;i++) {
			if (temp[i].value != 'no' && temp[i].selected==true) {
				temp[i].selected = false;
			}
		}
	}
}


function check_other(name) {

	var temp = 	document.getElementById(name+"_select");
	for(i=0;i<temp.length;i++) {
		if ((temp[i].value == 'other') && (temp[i].selected != false)) {
			document.getElementById("validate_"+name+"other").value="/.+$/";
		}
		else
			document.getElementById("validate_"+name+"other").value="";
	}
}


/*
displays hidden file upload fields
used in customer malware report form
*/
var file_counter=1;
function upload_more () {
	document.getElementById('file_upload'+file_counter).style.display="";
	file_counter++;
	if (file_counter >= 10) {
		document.getElementById('upload_more_link').style.display="none";
	}
}

/* function to check form fields
used in customer malware report form
should be modified to make it generic
*/
function check_form(myform) {
	var return_value = true;
	if (Trim(myform.name.value)=="") {
		document.getElementById('name_error').style.display="inline";
		myform.name.value="";
		return_value = false;
	}
	else {
		document.getElementById('name_error').style.display="none";
	}

	if (Trim(myform.email.value)=="") {
		document.getElementById('email_error').style.display="inline";
		myform.email.value="";
		return_value = false;
	}
	else {
		document.getElementById('email_error').style.display="none";
	}

	if (Trim(myform.description.value)=="") {
		document.getElementById('description_error').style.display="inline";
		myform.description.value="";
		return_value = false;
	}
	else {
		document.getElementById('description_error').style.display="none";
	}
	return return_value;
}


function Trim(trim_value){
	if(trim_value.length < 1){
		return"";
	}
	trim_value = RTrim(trim_value);
	trim_value = LTrim(trim_value);
	if(trim_value==""){
		return "";
	}
	else{
		return trim_value;
	}
}

function RTrim(r_value){
	var w_space = String.fromCharCode(32);
	var v_length = r_value.length;
	var strTemp = "";
	if(v_length < 0){
		return"";
	}
	var iTemp = v_length -1;

	while(iTemp > -1){
		if(r_value.charAt(iTemp) == w_space){
		}
		else{
			strTemp = r_value.substring(0,iTemp +1);
			break;
		}
		iTemp = iTemp-1;

	}
	return strTemp;

}

function LTrim(l_value){
	var w_space = String.fromCharCode(32);
	if(v_length < 1){
		return"";
	}
	var v_length = l_value.length;
	var strTemp = "";

	var iTemp = 0;

	while(iTemp < v_length){
		if(l_value.charAt(iTemp) == w_space){
		}
		else{
			strTemp = l_value.substring(iTemp,v_length);
			break;
		}
		iTemp = iTemp + 1;
	}
	return strTemp;
}

function check_password_fields() {
	if (document.getElementById("current_password").value != "") {
		document.getElementById("validate_password").value = "/.+$/";
		document.getElementById("validate_confirm").value = "/.+$/";
	}
}

/***
* Desc: Creates XMLHttpRequest Object
***/
function loadResults(product,url) {
	if(url == null) {
		var time = new Date();
		var random_time = time.getTime().toString()+time.getHours().toString()+time.getMinutes().toString()+time.getSeconds().toString();
		var url = "/dcounter.php?product="+product+"&time="+random_time;
	}
	// branch for native XMLHttpRequest object
	if (window.XMLHttpRequest) {
		req = new XMLHttpRequest();
		req.onreadystatechange = processReqChange;
		req.open("GET", url, true);
		req.send(null);
		// branch for IE/Windows ActiveX version
	} else if (window.ActiveXObject) {
		isIE = true;
		req = new ActiveXObject("Microsoft.XMLHTTP");
		if (req) {
			req.onreadystatechange = processReqChange;
			req.open("GET", url, true);
			req.send();
		}
	}
}

function processReqChange() {
	// only if req shows "loaded"
	if (req.readyState == 4) {
		// only if "OK"
		if (req.status == 200) {
			//alert(req.responseText);
			document.getElementById('counter').innerHTML = req.responseText;
		}
	}
}

function loadScript (url, callback) {
	var script = document.createElement('script');
	script.type = 'text/javascript';

	if (callback) {
		script.onload = function() {
			script.onload = script.onreadystatechange = null;
			callback();
		};
		script.onreadystatechange = function() {
			if (script.readyState == 'loaded' || script.readyState == 'complete') {
				script.onreadystatechange = script.onload = null;
				callback();
			}
		};
	}
	script.src = url;
	document.getElementsByTagName('head')[0].appendChild (script);
}

function addEvent(obj, eType, fnct){
	if (obj.addEventListener){
    	obj.addEventListener(eType, fnct, true);
    	return true;
 	} else if (obj.attachEvent){
    	var r = obj.attachEvent("on"+eType, fnct);
    	return r;
 	} else {
    	return false;
 	}
}

function getElementsByClassName(oElm, strTagName, strClassName){
/*
    Written by Jonathan Snook, http://www.snook.ca/jonathan
    Add-ons by Robert Nyman, http://www.robertnyman.com
*/
	var arrElements = (strTagName == "*" && oElm.all)? oElm.all : oElm.getElementsByTagName(strTagName);
	var arrReturnElements = new Array();
	strClassName = strClassName.replace(/\-/g, "\\-");
	var oRegExp = new RegExp("(^|\\s)" + strClassName + "(\\s|$)");
	var oElement;
	for(var i=0; i<arrElements.length; i++){
		oElement = arrElements[i];
		if(oRegExp.test(oElement.className)){
			arrReturnElements.push(oElement);
		}
	}
	return (arrReturnElements)
}

function load_tf_image() {
	document.getElementById('tf_image').src = 'http://a.tribalfusion.com/ti.ad?client=177773&ev=1';
	document.getElementById('tf_image2').src = 'http://ctxtad.tribalfusion.com/ctxtad/Lead?accountId=10232951';
}

function createCookie(name,value,days) {
	if (days) {
		var date = new Date();
		date.setTime(date.getTime()+(days*24*60*60*1000));
		var expires = "; expires="+date.toGMTString();
	}
	else var expires = "";
	document.cookie = name+"="+value+expires+"; path=/; domain=.pctools.com";
}

function readCookie(name) {
	var nameEQ = name + "=";
	var ca = document.cookie.split(';');
	for(var i=0;i < ca.length;i++) {
		var c = ca[i];
		while (c.charAt(0)==' ') c = c.substring(1,c.length);
		if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
	}
	return null;
}

function eraseCookie(name) {
	createCookie(name,"",-1);
}

function toggle_display(elementId) {
	var theElement;
	if (theElement = document.getElementById(elementId)) {
		if (theElement.style.display == 'none') {
			theElement.style.display = '';
		} else {
			theElement.style.display = 'none';
		}
	}
}

function track_downloads() {
	if (document.getElementById('download_tracking')) {
		theDiv = document.getElementById('download_tracking');
		content = '\n'
		+ '	<!-- Google Code for Download Conversion Page -->\n'
		+ '	<script type="text/javascript">\n'
		+ '	var google_conversion_id = 1041083675;\n'
		+ '	var google_conversion_language = "en_AU";\n'
		+ '	var google_conversion_format = "3";var google_conversion_color = "ffffff";\n'
		+ '	var google_conversion_label = "v-1BCKeBdRCb2rbwAw";\n'
		+ '	if (1.0) {\n'
		+ '		var google_conversion_value = 1.0;\n'
		+ '	}\n'
		+ '	</script>\n'
		+ '	<script type="text/javascript" src="http://www.googleadservices.com/pagead/conversion.js"></script>\n'
		+ '	<noscript>\n'
		+ '	<img height="1" width="1" border="0" src="http://www.googleadservices.com/pagead/conversion/1041083675/?value=1.0&amp;label=v-1BCKeBdRCb2rbwAw&amp;script=0">\n'
		+ '	</noscript>\n'
		+ '	<!-- End Google Code for Download Conversion Page -->\n'
		theDiv.innerHTML = content;
	}
}

function displayModalBox(boxId) {
	if (document.getElementById(boxId)) {
		// Need to work out where to put it...
		var viewportWidth = getViewportWidth();
		var viewportHeight = getViewportHeight();
		var boxLeft = 0;
		if (viewportWidth > 644) { // box is 644px wide
			var space = viewportWidth-644;
			boxLeft = Math.round(space/2);
		}
		var boxRight = (window.pageYOffset ? window.pageYOffset : document.documentElement.scrollTop) + 10;

		// Now we display it and set the relevant styles
		theBox = document.getElementById(boxId);
		theBox.style.display = 'block';
		theBox.style.position = 'absolute';
		theBox.style.top = boxRight+'px';
		theBox.style.left = boxLeft+'px';
		theBox.style.zIndex = 9999;

		// Add a box behind it to stop them from doing anything with the rest of the page
		var hideIt = document.createElement('div');
		hideIt.style.backgroundColor = 'white';
		hideIt.style.opacity = 0.7;
		hideIt.style.filter = 'alpha(opacity=70)'; // For IE :-(
		hideIt.style.position = 'absolute';
		hideIt.style.top = 0;
		hideIt.style.left = 0;
		hideIt.style.width = '110%';
		hideIt.style.height = '110%';
		hideIt.style.zIndex = 999;
		hideIt.style.marginTop = '-20px';
		hideIt.style.marginLeft = '-20px';
		hideIt.id = 'temp_hide_div';
		document.body.appendChild(hideIt);
	}
}

function hideModalBox(boxId) {
	if (document.getElementById(boxId)) {
		document.getElementById(boxId).style.display = 'none';
		if (document.getElementById('temp_hide_div')) {
			document.body.removeChild(document.getElementById('temp_hide_div'));
		}
	}
}

function getViewportWidth() {
	// the more standards compliant browsers (mozilla/netscape/opera/IE7+) use window.innerWidth and window.innerHeight
	if (typeof window.innerWidth != 'undefined')
	{
		return window.innerWidth;
	}

	// IE6 in standards compliant mode (i.e. with a valid doctype as the first line in the document)
	else if (typeof document.documentElement != 'undefined' && typeof document.documentElement.clientWidth != 'undefined' && document.documentElement.clientWidth != 0)
	{
		return document.documentElement.clientWidth;
	}

	// older versions of IE
	else
	{
		return document.getElementsByTagName('body')[0].clientWidth;
	}
	return 0;
}

function getViewportHeight() {
	// the more standards compliant browsers (mozilla/netscape/opera/IE7+) use window.innerWidth and window.innerHeight
	if (typeof window.innerHeight != 'undefined')
	{
		return window.innerHeight;
	}

	// IE6 in standards compliant mode (i.e. with a valid doctype as the first line in the document)
	else if (typeof document.documentElement != 'undefined' && typeof document.documentElement.clientHeight != 'undefined' && document.documentElement.clientHeight != 0)
	{
		return document.documentElement.clientHeight;
	}

	// older versions of IE
	else
	{
		return document.getElementsByTagName('body')[0].clientHeight;
	}
	return 0;
}
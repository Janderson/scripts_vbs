function ajaxObject(url) {                                   
	var that=this;                                
	var updating = false;                             
	this.callback = function() {}                           

	this.update = function(url) {
		if (updating==true) { return false; }                       
		updating = true;                                               
		var AJAX = null;                                               
		if (window.XMLHttpRequest) {                                  
			AJAX=new XMLHttpRequest();                                
		} else {                                                   
			AJAX=new ActiveXObject("Microsoft.XMLHTTP");               
		}                                                              
		if (AJAX==null) {                                             
			alert("Your browser doesn't support AJAX.");               		
			return false;                                              
		} else {
			AJAX.onreadystatechange = function() {                    
				if (AJAX.readyState==4 || AJAX.readyState=="complete") { 

					var result = AJAX.responseText;
					var values = result.split('=');
					if (values[0] == 'authenticated') {
						authenticated = values[1].substring(0,1);
					} 
					if (values[0] == 'pr_auth') {
						pr_auth = values[1].substring(0,1);
					}
					if (values[0] == 'iav_offered') {
						iav_offered = values[1].substring(0,1);
					}
					
					delete AJAX;                                          
					updating=false;                                       
					that.callback();                                      
				}                                                       
			}  
			                                                     
			AJAX.open("GET", url, true);                                
			AJAX.send(null);                                            
			return true;                                               
		}                                                              
	}
}                               

function setMyaccountStatus() {
	if (document.getElementById('login_box')) {
		document.getElementById('login_box').style.display = (authenticated == 1) ? 'none' : '';
	}
	if (document.getElementById('myaccount_login')) {

		document.getElementById('myaccount_login').style.display = (authenticated == 1) ? 'none' : '';
	}
	if (document.getElementById('myaccount')) {
		document.getElementById('myaccount').style.display = (authenticated == 1) ? '' : 'none';
	}
	if (document.getElementById('myaccount_logout')) {
		document.getElementById('myaccount_logout').style.display = (authenticated == 1) ? '' : 'none';
	}	
}

function setPressRoomLoginStatus() {
	if (document.getElementById('pr_login_box')) {
		document.getElementById('pr_login_box').style.display = (pr_auth == 1) ? 'none' : '';
	}
	if (document.getElementById('pr_help_box')) {
		document.getElementById('pr_help_box').style.display = (pr_auth == 1) ? '' : 'none';
	}	
}

function ShowIAV(){
	if (iav_offered==0) {
		ShowIAVOffer();
	}
}
/* Cookies Directive Disclosure Script
 * Version: 1.5
 * Author: Ollie Phillips
 * 20 June 2012 
 */

var bottomortop = 'top';
var sQSVD='';

function cookiesDirective(disclosurePos, displayTimes, privacyPolicyUri, cookieScripts) {
	
	// From v1.1 the position can be set to 'top' or 'bottom' of viewport
	var disclosurePosition = disclosurePos;

	// Better check it!
	if((disclosurePosition.toLowerCase() != 'top') && (disclosurePosition.toLowerCase() != 'bottom')) {
		
		// Set a default of top
		disclosurePosition = 'top';
	
	}
	
	// Start Test/Loader (improved in v1.1)
	var jQueryVersion = '1.5';
	
	// Test for JQuery and load if not available
	if (window.jQuery === undefined) {
		
		var s = document.createElement("script");
		s.src = "http://ajax.googleapis.com/ajax/libs/jquery/" + jQueryVersion + "/jquery.min.js";
		s.type = "text/javascript";
		s.onload = s.onreadystatechange = function() {
		
			if ((!s.readyState || s.readyState == "loaded" || s.readyState == "complete")) {
	
				// Safe to proceed	
				if(!cdReadCookie('cookiesDirective')) {
					
					if(displayTimes > 0) {
						
						// We want to limit the number of times this is displayed	
						// Record the view
						if(!cdReadCookie('cookiesDisclosureCount')) {

							cdCreateCookie('cookiesDisclosureCount',1,1);		

						} else {

							var disclosureCount = cdReadCookie('cookiesDisclosureCount');
							disclosureCount ++;
							cdCreateCookie('cookiesDisclosureCount',disclosureCount,1);

						}
						
						if(displayTimes >= cdReadCookie('cookiesDisclosureCount')) {
						
							// Cookies not accepted make disclosure
							if(cookieScripts) {
						
								cdHandler(disclosurePosition,privacyPolicyUri,cookieScripts);

							} else {
						
								cdHandler(disclosurePosition,privacyPolicyUri);

							}
						
						}
					

					} else {
						
						// No limit display on all pages	
						// Cookies not accepted make disclosure
						if(cookieScripts) {
						
							cdHandler(disclosurePosition,privacyPolicyUri,cookieScripts);

						} else {
						
							cdHandler(disclosurePosition,privacyPolicyUri);

						}
					
					}	
				
					
				} else {
					
					// Cookies accepted run script wrapper
					cookiesDirectiveScriptWrapper();
					
				}		
				
			}	
		
		}
	
		document.getElementsByTagName("head")[0].appendChild(s);		
	
	} else {
		
		// We have JQuery and right version
		if(!cdReadCookie('cookiesDirective')) {
			
			if(displayTimes > 0) {
				
				// We want to limit the number of times this is displayed	
				// Record the view
				if(!cdReadCookie('cookiesDisclosureCount')) {

					cdCreateCookie('cookiesDisclosureCount',1,1);		

				} else {

					var disclosureCount = cdReadCookie('cookiesDisclosureCount');
					disclosureCount ++;
					cdCreateCookie('cookiesDisclosureCount',disclosureCount,1);

				}
				
				if(displayTimes >= cdReadCookie('cookiesDisclosureCount')) {
				
					// Cookies not accepted make disclosure
					if(cookieScripts) {
				
						cdHandler(disclosurePosition,privacyPolicyUri,cookieScripts);

					} else {
				
						cdHandler(disclosurePosition,privacyPolicyUri);

					}
				
				}
			

			} else {
				
				// No limit display on all pages	
				// Cookies not accepted make disclosure
				if(cookieScripts) {
				
					cdHandler(disclosurePosition,privacyPolicyUri,cookieScripts);

				} else {
				
					cdHandler(disclosurePosition,privacyPolicyUri);

				}
			
			}	
		
			
		} else {
			
			// Cookies accepted run script wrapper
			cookiesDirectiveScriptWrapper();
			
		}
		
	}	
	// End Test/Loader
}

function detectIE789(){
 
	// Detect IE less than version 9.0
	var version;
	
	if (navigator.appName == 'Microsoft Internet Explorer') {

        var ua = navigator.userAgent;

        var re = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");

        if (re.exec(ua) != null) {

            version = parseFloat(RegExp.$1);

		}	
		
		if (version <= 8.0) {
			
			return true;
		
		} else {
			
			if(version == 9.0) {
				
				
				if(document.compatMode == "BackCompat") {
					
					// IE9 in quirks mode won't run the script properly, set to emulate IE8	
					var mA = document.createElement("meta");
					mA.content = "IE=EmulateIE8";				
					document.getElementsByTagName('head')[0].appendChild(mA);
				
					return true;
				
				} else {
				
					return false;
				
				}
			
			}	
			
			return false;
		
		}		

    } else {
	
		return false;
		
	}
	
}	

function cdHandler(disclosurePosition, privacyPolicyUri, cookieScripts) {
	
	// Our main disclosure script
		
	var displaySeconds = 100; // Alter this to remove the banner after number of seconds
	
	var epdApps;
	var epdAppsCount;
	var epdAppsDisclosure;
	var epdPrivacyPolicyUri;
	var epdDisclosurePosition;
	var epdCSSPosition = 'fixed';
	
	epdDisclosurePosition = disclosurePosition;
	epdPrivacyPolicyUri = privacyPolicyUri;
	
	if(detectIE789()) {
		
		// In IE 8 & presumably lower, position:fixed does not work
		// IE 9 in compatibility mode also means script won't work
		// Means we need to force to top of viewport and set position absolute
		epdDisclosurePosition = 'top';
		epdCSSPosition = 'absolute';
	
	}
	
	// What scripts must be declared, user passes these as comma delimited string
	if (cookieScripts) {
		
		epdApps = cookieScripts.split(',');
		epdAppsCount = epdApps.length;
		var epdAppsDisclosureText='';
		
		if(epdAppsCount>1) {
			
			for(var t=0; t < epdAppsCount - 1; t++) {
				
				epdAppsDisclosureText += epdApps[t] + ', ';	
							
			}	
			
			epdAppsDisclosure = ' We also use ' + epdAppsDisclosureText.substring(0, epdAppsDisclosureText.length - 2) + ' and ' + epdApps[epdAppsCount - 1] + ' scripts, which all use cookies. ';
		
		} else {
			
			epdAppsDisclosure = ' We also use a ' + epdApps[0] + ' script which uses cookies.';		
			
		}
			
	} else {
		
		epdAppsDisclosure = '';
		
	}
	
	// Create our overlay with message
	var divNode = document.createElement('div');
	divNode.setAttribute('id','epd');
	document.body.appendChild(divNode);
	
	// The disclosure narrative pretty much follows that on the Information Commissioners Office website		
	var disclosure = '<div id="cookiesdirective" style="position:'+ epdCSSPosition +';'+ epdDisclosurePosition + ':-300px;left:0px;width:100%;height:auto;background:#000000;opacity:.80; -ms-filter: “alpha(opacity=80)”; filter: alpha(opacity=80);-khtml-opacity: .80; -moz-opacity: .80; color:#FFFFFF;font-family:Arial;font-size:13px;text-align:center;z-index:1000;">';
	
	disclosure +='<div style="position:relative;height:auto;width:95%;padding:4px;margin-left:auto;margin-right:auto;">';	
	disclosure += '<div style="float:right" id="epdclose"><a onclick="$(\'#cookiesdirective\').animate({' + bottomortop + ':\'-300\'},1000,function(){$(\'#cookiesdirective\').remove()});return false;" href="#"><img src="' + sQSVD + '/js/gtk_close.png" style="border-style:none;padding:0px;padding:3px 12px 0px 0px" /></a></div>';
	disclosure += '<table style="border-style:none" align="center" cellspacing="0" cellpadding="0"><tr><td align="center" style="font-family:Arial;font-size:13px"><div style="text-align:center;font-family:Arial;font-size:13px;padding:2px" id="sCWText"></div></td></tr></table>';
	disclosure += '<div id="epdnotick" style="color:#ca0000;display:none;margin:2px;"><span id="sCWError" style="background:#cecece;padding:2px;"></span></div>'
	disclosure += '<table id="bCWUseAsNormalPP" style="border-style:none" align="center" cellspacing="0" cellpadding="0"><tr><td style="font-family:Arial;font-size:13px"><span id="sCWAccept" style="font-family:Arial;font-size:13px"></span></td><td><input type="checkbox" name="epdagree" id="epdagree" /></td><td><input type="submit" name="epdsubmit" id="epdsubmit" value="Continue"/></td></tr></table></div></div>';
	document.getElementById("epd").innerHTML= disclosure;
  	
	// Bring our overlay in
	if(epdDisclosurePosition.toLowerCase() == 'top') { 
	
		// Serve from top of page
		$('#cookiesdirective').animate({
		    top: '0'
		 }, 1000, function() {
			// Overlay is displayed, set a listener on the button
			$('#epdsubmit').click(function() {
		  
				if(document.getElementById('epdagree').checked){
					
					// Set a cookie to prevent this being displayed again
					cdCreateCookie('cookiesDirective',1,365);	
					// Close the overlay
					$('#cookiesdirective').animate({
						top:'-300'
					},1000,function(){
						
						// Remove the elements from the DOM and reload page, which should now
						// fire our the scripts enclosed by our wrapper function
						$('#cookiesdirective').remove();
						location.reload(true);
						
					});
					
				} else {
					
					// We need the box checked we want "explicit consent", display message
					document.getElementById('epdnotick').style.display = 'block';
						
				}
			});
			
			// Set a timer to remove the warning after 10 seconds
			setTimeout(function(){
					
				$('#cookiesdirective').animate({
					
					opacity:'0'
					
				},2000, function(){
					
					$('#cookiesdirective').css('top','-300px');
					
				});
				
			},displaySeconds * 1000);
			
		});
	
	} else {
	
		// Serve from bottom of page
		$('#cookiesdirective').animate({
		    bottom: '0'
		 }, 1000, function() {
			
			// Overlay is displayed, set a listener on the button
			$('#epdsubmit').click(function() {
		  
				if(document.getElementById('epdagree').checked) {
					 
					// Set a cookie to prevent this being displayed again
					cdCreateCookie('cookiesDirective',1,365);	
					// Close the overlay
					$('#cookiesdirective').animate({
						bottom:'-300'
					},1000,function() { 
						
						// Remove the elements from the DOM and reload page, which should now
						// fire our the scripts enclosed by our wrapper function
						$('#cookiesdirective').remove();
						location.reload(true);
						
					});
					
				} else {
					
					// We need the box checked we want "explicit consent", display message
					document.getElementById('epdnotick').style.display = 'block';	
					
				}
			});
			
			// Set a timer to remove the warning after 10 seconds
			setTimeout(function(){
					
				$('#cookiesdirective').animate({
					
					opacity:'0'
					
				},2000, function(){
					
					$('#cookiesdirective').css('bottom','-300px');
					
				});
				
			},displaySeconds * 1000);
			
		});
		
	}			
		
}	

function cdScriptAppend(scriptUri, myLocation) {
		
	// Reworked in Version 1.1 - needed a more robust loader

	var elementId = String(myLocation);
	var sA = document.createElement("script");
	sA.src = scriptUri;
	sA.type = "text/javascript";
	sA.onload = sA.onreadystatechange = function() {

		if ((!sA.readyState || sA.readyState == "loaded" || sA.readyState == "complete")) {
				
			return;
		
		} 	

	}

	switch(myLocation) {
		
		case 'head':
			document.getElementsByTagName('head')[0].appendChild(sA);
		  	break;
	
		case 'body':
			document.getElementsByTagName('body')[0].appendChild(sA);
		  	break;
	
		default: 
			document.getElementById(elementId).appendChild(sA);
			
	}	
		
}


// Simple Cookie functions from http://www.quirksmode.org/js/cookies.html - thanks!
function cdReadCookie(name) {
	
	var nameEQ = name + "=";
	var ca = document.cookie.split(';');
	for(var i=0;i < ca.length;i++) {
		var c = ca[i];
		while (c.charAt(0)==' ') c = c.substring(1,c.length);
		if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
	}
	
	return null;
	
}

function cdCreateCookie(name,value,days) {
	
	if (days) {
		var date = new Date();
		date.setTime(date.getTime()+(days*24*60*60*1000));
		var expires = "; expires="+date.toGMTString();
	}
	else var expires = "";
	document.cookie = name+"="+value+expires+"; path=/";
	
}
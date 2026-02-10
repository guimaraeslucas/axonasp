
<%function dumpJavaScript(includeJQuery)


dumpJavaScript="<script type=""text/javascript""><!--"& vbcrlf
dumpJavaScript=dumpJavaScript&"function openPopUpWindow(windowName,fileName,width,height) {var popUp = window.open(fileName,windowName,'top=10,left=10,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,top=50,left=50,target=_new,width='+width+',height='+height)} "
dumpJavaScript=dumpJavaScript&"function getIcon2(sCol1,sCol2,sCol3,iMode){try{var sColor;if (iMode==0){sColor='"&customer.publicIconColor&"';}else{sColor='"&customer.publicIconColorHover&"';}document.getElementById(sCol1).style.backgroundColor = sColor;document.getElementById(sCol2).style.backgroundColor = sColor;document.getElementById(sCol3).style.backgroundColor = sColor;}catch(err){}}"
dumpJavaScript=dumpJavaScript&"function getIcon(sLabel,sType,sUrl,sJs,sUniqueKey,sClassname){"
dumpJavaScript=dumpJavaScript&"try{document.write ('<TABLE cellSpacing=0 cellPadding=0 width=16 style=" & """" & "border-style:none;margin:0px;padding:0px;height:16px" & """" & " border=0><TR><TD width=1><\/TD>');"
dumpJavaScript=dumpJavaScript&"document.write ('<TD width=14 class=pI style=" & """" & "height:1px" & """" & " id=" & """" & "'+sUniqueKey+'1" & """" & "><\/TD><TD width=1><\/TD><\/TR>');"
dumpJavaScript=dumpJavaScript&"document.write ('<TR><TD colspan=3 width=14 class=pI style=" & """" & "height:14px" & """" & " id=" & """" & "'+sUniqueKey+'2" & """" & ">');"
dumpJavaScript=dumpJavaScript&"document.write ('<a '+sClassname+' href=" & """" & "'+sUrl+'" & """" & " title=" & """" & "'+sLabel+'" & """" & " onclick=" & """" & "'+sJs+'" & """" & "><img hspace=1 ');"
dumpJavaScript=dumpJavaScript&"document.write ('onmouseover=" & """" & "javascript:getIcon2(\''+sUniqueKey+'1\',\''+sUniqueKey+'2\',\''+sUniqueKey+'3\',1);" & """" & " ');"
dumpJavaScript=dumpJavaScript&"document.write ('onmouseout=" & """" & "javascript:getIcon2(\''+sUniqueKey+'1\',\''+sUniqueKey+'2\',\''+sUniqueKey+'3\',0);" & """" & " ');"
dumpJavaScript=dumpJavaScript&"document.write ('alt=" & """" & "'+sLabel+'" & """" & " src=" & """" & quotRepJS(C_DIRECTORY_QUICKERSITE)&"\/fixedImages\/public\/'+sType+'.gif" & """" & " border=0><\/a><\/TD><\/TR>');"
dumpJavaScript=dumpJavaScript&"document.write ('<TR><TD width=1><\/TD><TD width=14 class=pI id=" & """" & "'+sUniqueKey+'3" & """" & "><\/TD><TD width=1><\/TD><\/TR><\/TABLE>');}catch(err){}}"
dumpJavaScript=dumpJavaScript&"//--></script>"& vbcrlf

dumpJavaScript=dumpJavaScript & inputCssTweak & vbcrlf

dumpJavaScript=dumpJavaScript & "<style>.material-symbols-outlined {font-variation-settings:'FILL' 0,'wght' 400,'GRAD' 0,'opsz' 48}</style>"

if includeJQuery then
'JQUERY 183
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""https://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js""></script>" & vbcrlf
'local fallback
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"">if (typeof jQuery == 'undefined'){document.write(unescape(""%3Cscript src='" & C_DIRECTORY_QUICKERSITE & "/js/JQuery183.js' type='text/javascript'%3E%3C/script%3E""));}</script>"
'JQUERY UI
dumpJavaScript=dumpJavaScript&"<link media=""screen"" type=""text/css"" rel=""stylesheet"" href=""" & C_DIRECTORY_QUICKERSITE & "/js/JQueryUI.css"" />" & vbcrlf
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/JQueryUI.js""></script>" & vbcrlf
end if
dumpJavaScript=dumpJavaScript&"<link media=""all"" type=""text/css"" rel=""stylesheet"" href=""" & C_DIRECTORY_QUICKERSITE & "/css/media.css"" />" & vbcrlf
dumpJavaScript=dumpJavaScript&"<style type=""text/css"">" & vbcrlf
dumpJavaScript=dumpJavaScript&".ui-datepicker {font-family:Arial;font-size:12px}" & vbcrlf
dumpJavaScript=dumpJavaScript&".ui-dialog {font-family:Arial;font-size:12px;z-index:999}" & vbcrlf
dumpJavaScript=dumpJavaScript&".ui-dialog td {font-family:Arial;font-size:12px;}" & vbcrlf
dumpJavaScript=dumpJavaScript&".ui-accordion-header {background-image:none}" & vbcrlf
dumpJavaScript=dumpJavaScript&".ui-accordion-content {background-image:none}" & vbcrlf
dumpJavaScript=dumpJavaScript&".QSAccordion {" & customer.sQSAccordionMain & "}" & vbcrlf
dumpJavaScript=dumpJavaScript&".QSAccordion .QSAccordionHeader {" & customer.sQSAccordionHeader & "}" & vbcrlf
dumpJavaScript=dumpJavaScript&".QSAccordion .QSAccordionContent {" & customer.sQSAccordionContent & "}" & vbcrlf
dumpJavaScript=dumpJavaScript&"</style>" & vbcrlf

'nivo slider
if includeNS then
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/nivo-slider/jquery.nivo.slider.pack.js""></script>" & vbcrlf
dumpJavaScript=dumpJavaScript&"<link rel=""stylesheet"" type=""text/css"" href=""" & C_DIRECTORY_QUICKERSITE & "/js/nivo-slider/nivo-slider.css"" />" & vbcrlf
dumpJavaScript=dumpJavaScript&"<link rel=""stylesheet"" type=""text/css"" href=""" & C_DIRECTORY_QUICKERSITE & "/js/nivo-slider/themes/default/default.css"" media=""screen"" />" & vbcrlf
dumpJavaScript=dumpJavaScript&"<link rel=""stylesheet"" type=""text/css"" href=""" & C_DIRECTORY_QUICKERSITE & "/js/nivo-slider/themes/dark/dark.css"" media=""screen"" />" & vbcrlf
dumpJavaScript=dumpJavaScript&"<link rel=""stylesheet"" type=""text/css"" href=""" & C_DIRECTORY_QUICKERSITE & "/js/nivo-slider/themes/light/light.css"" media=""screen"" />" & vbcrlf
dumpJavaScript=dumpJavaScript&"<link rel=""stylesheet"" type=""text/css"" href=""" & C_DIRECTORY_QUICKERSITE & "/js/nivo-slider/themes/bar/bar.css"" media=""screen"" />" & vbcrlf
end if

'color picker
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/spectrum/spectrum.js""></script>" & vbcrlf
dumpJavaScript=dumpJavaScript&"<link rel=""stylesheet"" type=""text/css"" href=""" & C_DIRECTORY_QUICKERSITE & "/js/spectrum/spectrum.css"" />" & vbcrlf
'COLORBOX
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/colorbox/jquery.colorbox-min.js""></script>" & vbcrlf



dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"">"
dumpJavaScript=dumpJavaScript&"jQuery.colorbox.settings.maxWidth  = '95%';jQuery.colorbox.settings.maxHeight = '95%';var resizeTimer;"
dumpJavaScript=dumpJavaScript&"function resizeColorBoxQS() {if (resizeTimer) clearTimeout(resizeTimer);resizeTimer = setTimeout(function() {"
dumpJavaScript=dumpJavaScript&"if (jQuery('#cboxOverlay').is(':visible')) {jQuery.colorbox.load(true);}}, 300);}" & vbcrlf
dumpJavaScript=dumpJavaScript&"jQuery(window).resize(resizeColorBoxQS);window.addEventListener(""orientationchange"", resizeColorBoxQS, false);</script>"& vbcrlf




dumpJavaScript=dumpJavaScript&"<link media=""screen"" type=""text/css"" rel=""stylesheet"" href=""" & C_DIRECTORY_QUICKERSITE & "/js/colorbox/example" & customer.sGetPopupMode & "/colorbox.asp?qsd=" & server.urlEncode(C_DIRECTORY_QUICKERSITE) & """ />" & vbcrlf
dumpJavaScript=dumpJavaScript&textCounterJS & vbcrlf
dumpJavaScript=dumpJavaScript&resizeIframe & vbcrlf
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"">var slideShowTimerQS=4;</script>"& vbcrlf
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/slide.js""></script>" & vbcrlf
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/cycleJS.js""></script>" & vbcrlf
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/pollv2.js""></script>" & vbcrlf
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/qsAjax.js""></script>" & vbcrlf
if convertBool(selectedPage.bAccordeon) then
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"">$(function() {$( ""#accordion" & encrypt(selectedPage.iId) & """ ).accordion({heightStyle: ""content"",collapsible:""true"",active:" & iLPDEFAULTOpenerQS & "});});</script>"
end if
dim cH
for each ch in headerDictionary
dumpJavaScript=dumpJavaScript& headerDictionary(ch)
next
dumpJavaScript=dumpJavaScript & cPopup.dumpJS & vbcrlf
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"">" & vbcrlf
dumpJavaScript=dumpJavaScript&"$(document).ready(function(){" & vbcrlf
dumpJavaScript=dumpJavaScript&"$("".bPopupFullWidthNoReload"").colorbox({close: """ & quotrep(l("close")) & """, width:""90%"", height:""90%"", iframe:true}); " & vbcrlf
dumpJavaScript=dumpJavaScript&"$("".bPopupFullWidthNoReloadNoOverlayClose"").colorbox({close: """ & quotrep(l("close")) & """, width:""90%"", height:""90%"", iframe:true, overlayClose: false}); " & vbcrlf
dumpJavaScript=dumpJavaScript&"$("".bPopupFullWidthReload"").colorbox({close: """ & quotrep(l("close")) & """, width:""90%"", height:""90%"", iframe:true, onClosed:function(){location.reload(true);}}); " & vbcrlf
dumpJavaScript=dumpJavaScript&"$("".QSPP"").colorbox({close: """ & quotrep(l("close")) & """, width:""750"", height:""600"", iframe:true}); " & vbcrlf
dumpJavaScript=dumpJavaScript&"$("".QSPPR"").colorbox({close: """ & quotrep(l("close")) & """, width:""750"", height:""600"", iframe:true, onClosed:function(){location.reload(true);}}); " & vbcrlf
dumpJavaScript=dumpJavaScript&"$("".QSPPIMG"").colorbox({photo:true});" & vbcrlf
dumpJavaScript=dumpJavaScript&"$("".QSPPAVATAR"").colorbox({close: """ & quotrep(l("close")) & """, width:""720"", height:""280"", iframe:true});" & vbcrlf
dumpJavaScript=dumpJavaScript&"$("".QSFullScreen"").colorbox({close: """ & quotrep(l("close")) & """, width:""100%"", height:""100%"", iframe:true}); " & vbcrlf
dumpJavaScript=dumpJavaScript&"});" & vbcrlf & "</script>" & vbcrlf
dumpJavaScript=dumpJavaScript&sCBExtUrl
dumpJavaScript=dumpJavaScript&customer.getArrowCSS& vbcrlf

if customer.bPeelEnabled then
dumpJavaScript=dumpJavaScript&"<script type=""text/javascript"">"
dumpJavaScript=dumpJavaScript&"$(document).ready(function(){"
'dumpJavaScript=dumpJavaScript&"//Page Flip on hover" & vbcrlf
dumpJavaScript=dumpJavaScript&"$(""#pageflip"").hover(function(){"
dumpJavaScript=dumpJavaScript&"$(""#pageflip img,.msg_block"").stop()"
dumpJavaScript=dumpJavaScript&".animate({width:'" & customer.sPeelMOSize & "px',height:'" & round(convertgetal(customer.sPeelMOSize)*1.04,0) & "px'},500);} ,"
dumpJavaScript=dumpJavaScript&"function(){$(""#pageflip img"").stop()"
dumpJavaScript=dumpJavaScript&".animate({width:'" & customer.sPeelIdleSize & "px',height:'" & round(convertgetal(customer.sPeelIdleSize)*1.04,0) & "px'},220);"
dumpJavaScript=dumpJavaScript&"$("".msg_block"").stop()"
dumpJavaScript=dumpJavaScript&".animate({width:'" & customer.sPeelIdleSize & "px',height:'" & customer.sPeelIdleSize & "px'},200);"
dumpJavaScript=dumpJavaScript&"});" & vbcrlf
dumpJavaScript=dumpJavaScript&"});</script>" & vbcrlf
dumpJavaScript=dumpJavaScript&"<style type=""text/css"">"& vbcrlf
dumpJavaScript=dumpJavaScript&"#pageflip {"
dumpJavaScript=dumpJavaScript&"position:relative;"
dumpJavaScript=dumpJavaScript&"right:0;top:0;"
dumpJavaScript=dumpJavaScript&"float:right;"
dumpJavaScript=dumpJavaScript&"}" & vbcrlf
dumpJavaScript=dumpJavaScript&"#pageflip img {"
dumpJavaScript=dumpJavaScript&"width:" & customer.sPeelIdleSize & "px;height:" & round(convertgetal(customer.sPeelIdleSize)*1.04,0) & "px;"
dumpJavaScript=dumpJavaScript&"z-index:99;"
dumpJavaScript=dumpJavaScript&"position:absolute;"
dumpJavaScript=dumpJavaScript&"right:0;top:0;"
dumpJavaScript=dumpJavaScript&"-ms-interpolation-mode: bicubic;"
dumpJavaScript=dumpJavaScript&"}" & vbcrlf
dumpJavaScript=dumpJavaScript&"#pageflip .msg_block {"
dumpJavaScript=dumpJavaScript&"width:" & customer.sPeelIdleSize & "px;height:" & customer.sPeelIdleSize & "px;"
dumpJavaScript=dumpJavaScript&"overflow:hidden;"
dumpJavaScript=dumpJavaScript&"position:absolute;"
dumpJavaScript=dumpJavaScript&"right:0;top:0;"
'dumpJavaScript=dumpJavaScript&"background-color:#FFFFFF;" & vbcrlf
select case right(customer.sPeelImage,3) 
case "jpg"
dumpJavaScript=dumpJavaScript&"background: url(" & C_DIRECTORY_QUICKERSITE & "/showthumb.aspx?maxsize=" & customer.sPeelMOSize+1 & "&FSR=1&img=" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & customer.sPeelImage & ") no-repeat right top;"
case else
dumpJavaScript=dumpJavaScript&"background: url(" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & customer.sPeelImage & ") no-repeat right top;"
end select
dumpJavaScript=dumpJavaScript&"}" & vbcrlf
dumpJavaScript=dumpJavaScript & vbcrlf

dumpJavaScript=dumpJavaScript&"</style>" & vbcrlf
end if
end function



function inputCssTweak

inputCssTweak="<style>"
inputCssTweak=inputCssTweak & "input,textarea,select,radio,checkbox "
inputCssTweak=inputCssTweak & "{padding:5px;border-radius:4px;border:1px solid #AAA} " & vbcrlf
inputCssTweak=inputCssTweak & "</style>"

end function





function textCounterJS
textCounterJS="<script type=""text/javascript"">"& vbcrlf
textCounterJS=textCounterJS&"function textCounter(field, countfield, maxlimit){if (field.value.length > maxlimit){alert('"
textCounterJS=textCounterJS& quotrepJS(replace(l("err_maxchars"),"([NMBR])","",1,-1,1)) & "');"
textCounterJS=textCounterJS&"field.value = field.value.substring(0, maxlimit);"
textCounterJS=textCounterJS&"} else {"
textCounterJS=textCounterJS&"document.getElementById(countfield).innerHTML = maxlimit - field.value.length;"
textCounterJS=textCounterJS&"}"
textCounterJS=textCounterJS&"}"& vbcrlf
textCounterJS=textCounterJS&"</script>"& vbcrlf
end function
function resizeIframe
resizeIframe="<script type='text/javascript'><!--"& vbcrlf
resizeIframe=resizeIframe&"var iFrameWidth=465;var iFrameHeight=270;var iFrameAddW=80;var iFrameAddH=120;"
resizeIframe=resizeIframe&"function resizeiframe(iframe_id)"
resizeIframe=resizeIframe&"{if(iFrameWidth==465)"
resizeIframe=resizeIframe&"{var the_iframe = document.getElementById(iframe_id);"
resizeIframe=resizeIframe&"the_iframe.style.width=iFrameWidth+iFrameAddW+'px';"
resizeIframe=resizeIframe&"the_iframe.style.height=iFrameHeight+iFrameAddH+'px';"
resizeIframe=resizeIframe&"iFrameWidth=iFrameWidth+iFrameAddW;"
resizeIframe=resizeIframe&"iFrameHeight=iFrameHeight+iFrameAddH;"
resizeIframe=resizeIframe&"}"
resizeIframe=resizeIframe&"}"
resizeIframe=resizeIframe&"//--></script>"& vbcrlf
end function
function getMenuJS
getMenuJS="<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/qsAjax.js""></script>"& vbcrlf
end function%>

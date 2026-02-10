<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bFiles%><%dim fe
set fe=new cls_fileexplorer
application("mupath")=replace(Application("QS_CMS_C_VIRT_DIR") & Application("QS_CMS_userfiles") & request("sPath"),"//","/",1,-1,1)
application("UserIP")=convertStr(UserIP)
application("sessionID")=Session.SessionID
session("bHasSetUF")=false%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /><title>Uploadify Form</title><link href="uploadify214/css/default.css" rel="stylesheet" type="text/css" /><link href="uploadify214/css/uploadify.css" rel="stylesheet" type="text/css" /><script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.5/jquery.min.js"></script><script type="text/javascript" src="uploadify214/swfobject.js"></script><script type="text/javascript" src="uploadify214/jquery.uploadify.v2.1.4.min.js"></script><script>function Send_document()
{
$('#uploadify').uploadifyUpload();
}
</script><script type="text/javascript">$(document).ready(function() {
$("#uploadify").uploadify({
'uploader'       : 'uploadify214/uploadify.swf',
'script'         : 'uploader214.asp?sId=<%=Session.SessionID%>',
'cancelImg'      : 'uploadify214/cancel.png',
'sizeLimit'	 : 1324000,
'fileDesc'	 : '<%dim aft,callowedFileTypes,aftString
set callowedFileTypes=allowedFileTypes
for each aft in callowedFileTypes
aftString=aftString & ucase(aft) & " (*." & aft & "), "
next
response.write left(aftString,len(aftString)-2)%>',
'fileExt'	 : '<%aftString=""
for each aft in callowedFileTypes
aftString=aftString & "*." & lcase(aft) & ";"
next
response.write left(aftString,len(aftString)-1)%>',
'folder'         : '<%=quotRepJS(application("mupath"))%>',
'multi'          : true,
'auto'          : true,
onError: function (a, b, c, d) {
         if (d.status == 404)
            alert('Could not find upload script. Use a path relative to: '+'<?= getcwd() ?>');
         else if (d.type === "HTTP")
            alert('error aaa'+d.type+": "+d.status);
         else if (d.type ==="File Size")
            alert('File Size Limit: 1MB!');
         else
            alert('error '+d.type+": "+d.text);
},
onComplete	 : function(event, queueID, fileObj, response, data) {
     	var path = escape(fileObj.filePath);
$('#filesUploaded').append('<div class=\'uploadifyQueueItem\'><a href='+path+' target=\'_blank\'>'+fileObj.name+'</a></div>');
}
});
});
</script><link href="assetmanager/style.css" rel="stylesheet" type="text/css"></head><body><table width="100%" height="100%" align=center style="" cellpadding=4 cellspacing=0 border=0 ><tr><td valign=top style="background:url('assetmanager/bg.gif') no-repeat right bottom;padding-top:5px;padding-left:5px;padding-right:5px;padding-bottom:0px;"><form id="formIDdoc" name="formIDdoc" class="form" method="post" action="bs_multifileupload.asp"><p><strong><a href="assetmanager/assetmanagerIF.asp?QS_inpFolder=<%=server.urlencode(server.mappath(application("mupath")))%>">Back to file manager</a></strong></p><table align=left><tr><td>Select the folder to upload files to:</td></tr><tr><td><%=left(Application("QS_CMS_userfiles"),len(Application("QS_CMS_userfiles"))-1)%><select onchange="javascript:document.formIDdoc.submit();" name="sPath"><option value="/">/</option><%=fe.SelectBoxFolders(server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles")),request("sPath"))%></select></td></tr><%'if not isLeeg(request("sPath")) then%><tr><td style="height:10px">&nbsp;</td></tr><tr><td>Select one or more files, next click '<b>Upload files</b>'</p><p id="sending" name="sending"><a href="javascript:Send_document()"><b>Upload files</b></a></p><p><input class="text-input" name="uploadify" id="uploadify" type="file" size="20"></p><div id="filesUploaded"></div></td></tr><%'end if%></table></form></td></tr></table></body></html>

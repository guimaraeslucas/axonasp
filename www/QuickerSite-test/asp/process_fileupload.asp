
<%dim uploadPostID,uploadPost
set uploadPost=new cls_post
uploadPost.pick(decrypt(request.querystring("iPostID")))
if not uploadPost.theme.bFileUploads or not convertBool(logon.authenticatedIntranet) then response.end
'if convertGetal(uploadPost.iId)<>0 then
'	if convertGetal(logon.contact.iId)<>convertGetal(uploadPost.iContactID) then 
'	if convertGetal(logon.contact.iId)<>convertGetal(uploadPost.theme.iContactID) then response.end 
'	end if
'end if
Dim uploadsDirVar2,uploadfilename
uploadsDirVar2 = server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "userfiles") 
dim Upload2, uploadOK2, ks2, sShowFileName, iFileSize
Set Upload2 = New FreeASPUpload
uploadOK2=false
Upload2.Save uploadsDirVar2
ks2 = Upload2.UploadedFiles.keys
if Upload2.form("upload")<>"" then
if convertStr(Upload2.form("QSSEC")) <> convertStr(secCode) then response.end 
cpw=generatepassword
if (UBound(ks2) <> -1) then
for each fileKey in Upload2.UploadedFiles.keys
select case lcase(GetFileExtension(Upload2.UploadedFiles(fileKey).FileName))
case "jpg","png","doc","zip","rar","pdf","txt","mp3","xls","docx","xlsx","gif","css","rtf","wmv","mp4"
case else
strMessage="File-type not allowed!<br />"
end select 
if lcase(Upload2.UploadedFiles(fileKey).Length)>7000000 then
strMessage=strMessage&"Max 5MB!<br />"
end if
if not isLeeg(strMessage) then 
Upload2.UploadedFiles(fileKey).delete()
else
'create filename
uploadfilename=lcase(encrypt(uploadPost.theme.iId) & "_" & encrypt(uploadPost.iId) & "_" & cpw & "_upload." & lcase(GetFileExtension(Upload2.UploadedFiles(fileKey).FileName)))
Upload2.UploadedFiles(fileKey).rename uploadfilename,uploadsDirVar2
if not isLeeg(Upload2.form("sFileDesc")) then
sShowFileName=Upload2.form("sFileDesc")
else
sShowFileName=uploadfilename
end if
uploadOK2=true
iFileSize=round(convertGetal(Upload2.UploadedFiles(fileKey).Length)/1024,0)
end if
next
else
strMessage=strMessage& l("err_newFile") & "<br />"
end if
end if
on error resume next
if (UBound(ks2) <> -1) and not uploadOK2 then
for each fileKey in Upload2.UploadedFiles.keys
Upload2.UploadedFiles(fileKey).delete()
next
end if
on error goto 0%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"><html lang="en"><head><title>Upload file</title><style type="text/css">td {font-family:Verdana;font-size:10pt}
</style><%if uploadOK2 then%><script type="text/javascript"><%if convertGetal(request.querystring("RP"))<>0 then 'reply%>parent.document.getElementById('uSS<%=encrypt(uploadPost.iId)%>').innerHTML='The file <i><b><%=quotrep(quotrepJS(sShowFileName))%></b></i> is attached to this reply';
parent.document.replyForm<%=encrypt(uploadPost.iId)%>.sFileName.value='<%=uploadfilename%>';
parent.document.replyForm<%=encrypt(uploadPost.iId)%>.sFileDesc.value='<%=quotrepJS(Upload2.form("sFileDesc"))%>';
parent.document.replyForm<%=encrypt(uploadPost.iId)%>.iFileSize.value='<%=iFileSize%>';
<%else%>parent.document.getElementById('sFileD').innerHTML='The file <i><b><%=quotrep(quotrepJS(sShowFileName))%></b></i> is attached to this topic';
parent.document.postTopic.sFileName.value='<%=uploadfilename%>';
parent.document.postTopic.sFileDesc.value='<%=quotrepJS(Upload2.form("sFileDesc"))%>';
parent.document.postTopic.iFileSize.value='<%=iFileSize%>';
<%end if%>parent.$.fn.colorbox.close();
</script><%end if%></head><body bgcolor="#FFFFFF" style="background-color:#FFFFFF"><table style="height:100%;width:100%" cellpadding="0" cellspacing="0"><tr><td valign="middle"><tr><td valign="middle"><table cellpadding="10" align="center" style="border:1px solid #AAAAAA;background-color:#EEEEEE" border="0"><tr><td><form method="post" enctype="multipart/form-data" action="<%=C_DIRECTORY_QUICKERSITE%>/default.asp?RP=<%=convertGetal(request.querystring("RP"))%>&iThemeID=<%=request.querystring("iThemeID")%>&iPostID=<%=request.querystring("iPostID")%>&pageAction=fileupload" name="upload"> 
<%=QS_secCodeHidden%><%if convertGetal(uploadPost.iId)<>0 then%><p>File upload for <b><%=quotrep(uploadPost.sSubject)%></b></p><%end if%><table cellpadding="4" cellspacing="0"><tr><td align="right" valign="top"><i>Select a file:</i></td><td align=left><input type="file" name="fup" /><br /><small><i>(max 5MB - jpg,png,doc,zip,rar,pdf,txt,mp3,xls,docx,xlsx,gif,css,rtf)</i></small></td></tr><tr><td align="right"><i>Description:</i></td><td><input type="text" maxlength="50" size="40" value="<%=quotrep(Upload2.form("sFileDesc"))%>" name="sFileDesc" /></td></tr><%if not isLeeg(strMessage) then%><tr><td>&nbsp;</td><td style="background-color:yellow"><span style="color:Red"><strong><%=strMessage%></strong></span></td></tr><%end if%><tr><td align="right">&nbsp;</td><td><INPUT class="art-button" TYPE=SUBMIT VALUE="<%=l("upload")%>" id=SUBMIT1 name=upload onclick="javascript:this.value='Uploading... Please wait...';" /></td></tr></table></form></td></tr></table></td></tr></table></body></html><%set uploadPost=nothing
set Upload2=nothing%>

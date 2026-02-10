
<%if not customer.bUseAvatars then response.end
dim workingContact
set workingContact=new cls_contact
if Session(cId & "isAUTHENTICATED")	= true	or Session(cId & "isAUTHENTICATEDSecondAdmin")= true then
if convertGetal(decrypt(request.querystring("iContactID")))<>0 then
workingContact.pick(decrypt(request.querystring("iContactID")))
else
set workingContact=logon.contact
end if
else
if convertGetal(logon.contact.iId)=0 then response.end
set workingContact=logon.contact
end if
Dim uploadsDirVar
uploadsDirVar = server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "userfiles") 
dim Upload, uploadOK
Set Upload = New FreeASPUpload
uploadOK=false
on error resume next
dim fsoC
set fsoC=server.createobject("scripting.filesystemobject")
if not fsoC.folderexists(uploadsDirVar) then
fsoC.createFolder(uploadsDirVar)
end if
on error goto 0
dim ks, fileKey, strMessage,bShowAvatar
bShowAvatar=false
Upload.Save uploadsDirVar
ks = Upload.UploadedFiles.keys
if Upload.form("upload")<>"" then
if convertStr(Upload.form("QSSEC")) <> convertStr(secCode) then response.end 
workingContact.pick(decrypt(Upload.form("iContactID")))
isSecure(workingContact.iId)
dim cpw
cpw=generatepassword
if (UBound(ks) <> -1) then
for each fileKey in Upload.UploadedFiles.keys
if lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))<>"jpg" and lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))<>"gif" then
strMessage="Please select a JPG/GIF file...<br />"
end if
if lcase(Upload.UploadedFiles(fileKey).Length)>650000 then
strMessage=strMessage&"Max 500kB!<br />"
end if
if not isLeeg(strMessage) then 
Upload.UploadedFiles(fileKey).delete()
else
'hernoemen naar userfiles.jpg
cpw=cpw & "." & lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))
Upload.UploadedFiles(fileKey).rename workingContact.iId & "_" & cpw,uploadsDirVar
uploadOK=true
end if
next
else
strMessage=strMessage& l("err_newFile") & "<br />"
end if
if isLeeg(strMessage) then
'save avatar
workingContact.saveAvatar(cpw)
end if
end if
on error resume next
if (UBound(ks) <> -1) and not uploadOK then
for each fileKey in Upload.UploadedFiles.keys
Upload.UploadedFiles(fileKey).delete()
next
end if
if request.querystring("delete")<>"" then
if convertStr(Request.querystring("QSSEC"))=convertStr(secCode) then
isSecure(workingContact.iId)
workingContact.removeAvatar()
end if
end if
on error goto 0%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"><html lang="en"><head><title>Avatar</title><style type="text/css">td {font-family:Verdana;font-size:10pt}
</style></head><body bgcolor="#FFFFFF" style="background-color:#FFFFFF"><table style="height:100%;width:100%" cellpadding="0" cellspacing="0"><tr><td valign="middle"><tr><td valign="middle"><table cellpadding="10" align="center" style="border:1px solid #AAAAAA;background-color:#EEEEEE" border="0"><tr><%if not isLeeg(workingContact.sAvatar) then%><td><table width="120" cellpadding="15" cellspacing="0" border="0" style="width:120px;height:120px"><tr><td style="text-align:center"><%=workingContact.sImgTagAvatar(100)%><br /><a href="<%=C_DIRECTORY_QUICKERSITE%>/default.asp?pageAction=avataredit&amp;iContactID=<%=encrypt(workingContact.iId)%>&amp;delete=true&amp;<%=QS_secCodeURL%>" onclick="javascript: return confirm('<%=l("areyousure")%>');"><%=l("delete")%></a></td></tr></table></td><%end if%><td><form method="post" enctype="multipart/form-data" action="<%=C_DIRECTORY_QUICKERSITE%>/default.asp?pageAction=avataredit" name="avatar"> 
<%=QS_secCodeHidden%><input type=hidden name=iContactID value="<%=encrypt(workingContact.iId)%>" /><table cellpadding="4" cellspacing="0"><tr><td align="right" valign="top"><i>Select new Avatar:</i></td><td align=left><input type="file" name="avatar" /><br /><small><i>(max 500kB)</i></small></td></tr><%if not isLeeg(strMessage) then%><tr><td>&nbsp;</td><td style="background-color:yellow"><span style="color:Red"><strong><%=strMessage%></strong></span></td></tr><%end if%><tr><td align="right">&nbsp;</td><td><INPUT class="art-button" TYPE=SUBMIT VALUE="<%=l("upload")%>" id=SUBMIT1 name=upload onclick="javascript:this.value='Uploading... Please wait...';" /></td></tr></table></form></td></tr></table></td></tr></table></body></html><%function isSecure(wid)
if Session(cId & "isAUTHENTICATED")	= true	or Session(cId & "isAUTHENTICATEDSecondAdmin")= true then exit function
if convertGetal(wid)<>convertGetal(logon.contact.iId) then response.end
end function
application("doresize")=false%>

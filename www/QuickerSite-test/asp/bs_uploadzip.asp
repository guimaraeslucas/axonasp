<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bTemplates%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_template)%><!-- #include file="bs_templateBack.asp"--><%Dim zipper
set zipper=new cls_zipper
if not customer.supportZipper then response.redirect("bs_templateList.asp")
Dim uploadsDirVar
uploadsDirVar = server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles")) 
dim Upload
Set Upload = New FreeASPUpload
Upload.Save uploadsDirVar
dim uploaderror
uploaderror=""
if err.number<>0 then
uploaderror="<p align=""center""><font color=Red><strong>Something went wrong when uploading. In most cases this is related to incorrect folder permissions, or to a limit of 200kB for file uploads.<br />Please contact your server administrator to reset folder permissions and/or increase the file-upload size limit for your website.</strong></font></p>"
end if
checkCSRF_Upload(Upload.form("QSSEC"))
dim ks, fileKey, strMessage
ks = Upload.UploadedFiles.keys
if (UBound(ks) <> -1) then
for each fileKey in Upload.UploadedFiles.keys
select case lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))
case "zip"
case else
message.AddError "err_fileType"
end select
if message.hasErrors then 
Upload.UploadedFiles(fileKey).delete()
else
Response.Redirect ("bs_unzip.asp?zip=" & Upload.UploadedFiles(fileKey).FileName)
end if
next
else
strMessage=strMessage&"err_newFile,"
end if
response.write uploaderror%><FORM method="post" ENCTYPE="multipart/form-data" ACTION="bs_uploadzip.asp" id=form1 name=uploadForm> 
<%=QS_secCodeHidden%><p align=center>Below you can upload a template made with <a target="_blank" href="https://jstemplates.com"><b>JStemplates.com</b></a></p><p align="center">Template upload processing requires a freeware ZIP-component named <a href="http://www.mitdata.com/" target="_blank">aspEasyZip</a> by John Lohmeyer.</p><table align=center cellpadding=5 border=0><tr><td colspan=2 class=header><%=l("upload")%></td></tr><tr><td class=QSlabel>Select ZIP:*</td><td><input type=file size=40 maxlength=255 name="sName" value="" /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td>&nbsp;</td><td><INPUT class="art-button" TYPE=button onclick="javascript:this.value='Please wait...';document.uploadForm.submit();this.disabled=true;" VALUE="<%=l("upload")%>" id=SUBMIT1 name=SUBMIT1 /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->

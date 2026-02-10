<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%Dim uploadsDirVar
uploadsDirVar = server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles")) 
dim Upload
Set Upload = New FreeASPUpload
Upload.Save uploadsDirVar
checkCSRF_Upload(Upload.form("QSSEC"))
dim ks, fileKey, strMessage
ks = Upload.UploadedFiles.keys
if (UBound(ks) <> -1) then
for each fileKey in Upload.UploadedFiles.keys
if lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))<>"ico" then
strMessage=strMessage&"err_ICO_file,"
end if
if lcase(Upload.UploadedFiles(fileKey).Length)>20000 then
strMessage=strMessage&"err_ICO_fileSize,"
end if
if not isLeeg(strMessage) then 
Upload.UploadedFiles(fileKey).delete()
else
'hernoemen naar favicon.ico
Upload.UploadedFiles(fileKey).rename "favicon.ico",uploadsDirVar
end if
next
else
strMessage=strMessage&"err_newFile,"
end if
if isLeeg(strMessage) then
customer.hasFavicon=true
if customer.save then Response.Redirect ("bs_favicon.asp")
else
Response.Redirect ("bs_favicon.asp?strMessage="& server.urlEncode(strMessage))
end if%>

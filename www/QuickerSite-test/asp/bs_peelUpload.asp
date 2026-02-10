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
select case GetFileExtension(Upload.UploadedFiles(fileKey).FileName)
case "png","jpg","gif"
case else
strMessage=strMessage&"err_fileType,"
end select
'if lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))<>"png" and lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))<>"jpg" and lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))<>"gif" then
'strMessage=strMessage&"err_fileType,"
'end if
if lcase(Upload.UploadedFiles(fileKey).Length)>2000000 then
strMessage=strMessage&"err_fileSize,"
end if
if not isLeeg(strMessage) then 
Upload.UploadedFiles(fileKey).delete()
else
'hernoemen naar favicon.ico
dim peelName
peelName="peel_" & generatePassword & "." & GetFileExtension(Upload.UploadedFiles(fileKey).FileName)
Upload.UploadedFiles(fileKey).rename peelName,uploadsDirVar
customer.sPeelImage=peelName
end if
next
else
strMessage=strMessage&"err_newFile,"
end if
if isLeeg(strMessage) then
customer.hasFavicon=true
if customer.save then Response.Redirect ("bs_peel.asp")
else
Response.Redirect ("bs_selectPeel.asp?strMessage="& server.urlEncode(strMessage))
end if%>

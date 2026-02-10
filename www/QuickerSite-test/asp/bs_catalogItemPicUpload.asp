<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCatalog%><%Dim uploadsDirVar
uploadsDirVar = server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles")) 
dim Upload
Set Upload = New FreeASPUpload
Upload.Save uploadsDirVar
checkCSRF_Upload(Upload.form("QSSEC"))
dim catalogItem
set catalogItem=new cls_catalogItem
catalogItem.pick(decrypt(Upload.form("iItemID")))
dim ks, fileKey, strMessage
ks = Upload.UploadedFiles.keys
if (UBound(ks) <> -1) then
for each fileKey in Upload.UploadedFiles.keys
'type controleren
if not allowedFileTypesforThumbing.exists(lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))) then
strMessage=strMessage&"err_fileType,"
end if
'grootte controleren
if lcase(Upload.UploadedFiles(fileKey).Length)>5550000 then
strMessage=strMessage&"err_fileSize,"
end if
if not isLeeg(strMessage) then 
Upload.UploadedFiles(fileKey).delete()
else

if isLeeg(catalogItem.catalog.sFilePath) then
catalogItem.sPicExt	= encrypt(catalogItem.iId) & "_" & GeneratePassWord() & "." & GetFileExtension(Upload.UploadedFiles(fileKey).FileName)
catalogItem.save()
Upload.UploadedFiles(fileKey).rename catalogItem.sPicExt,uploadsDirVar
else
catalogItem.sPicExt	= Upload.UploadedFiles(fileKey).FileName
catalogItem.save()
Upload.UploadedFiles(fileKey).move catalogItem.sPicExt,uploadsDirVar,server.mappath(Application("QS_CMS_userfiles") & catalogItem.catalog.correFP)
end if


end if
next
else
strMessage=strMessage&"err_newFile,"
end if
if isLeeg(strMessage) then
Response.Redirect ("bs_catalogItemEdit.asp?iItemID="& encrypt(catalogItem.iId))
else
Response.Redirect ("bs_catalogItemPic.asp?iItemID="& encrypt(catalogItem.iId) &"&strMessage="& server.urlEncode(strMessage))
end if%>

<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%dim shopProduct
set shopProduct=new cls_shopProduct
shopProduct.pick(decrypt(request.querystring("iShopProductID")))
dim fso
set fso=server.createobject("scripting.filesystemobject")
if not fso.folderexists(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/")) then
fso.createFolder server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/")
end if
if not fso.folderexists(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & shopProduct.iId)) then
fso.createFolder server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & shopProduct.iId)
end if
set fso=nothing
Dim uploadsDirVar
uploadsDirVar = server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & shopProduct.iId) 
dim Upload
Set Upload = New FreeASPUpload
Upload.Save uploadsDirVar
checkCSRF_Upload(Upload.form("QSSEC"))
dim ks, fileKey, strMessage
ks = Upload.UploadedFiles.keys
if (UBound(ks) <> -1) then
for each fileKey in Upload.UploadedFiles.keys
if lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))<>"jpg" then
strMessage=strMessage&"err_fileType,"
end if
if not isLeeg(strMessage) then 
Upload.UploadedFiles(fileKey).delete()
end if
next
else
strMessage=strMessage&"err_newFile,"
end if
if isLeeg(strMessage) then
if customer.save then Response.Redirect ("bs_shopProductImg.asp?iShopProductID=" & encrypt(shopProduct.iId))
else
Response.Redirect ("bs_shopProductImg.asp?iShopProductID=" & encrypt(shopProduct.iId) & "&strMessage="& server.urlEncode(strMessage))
end if%>

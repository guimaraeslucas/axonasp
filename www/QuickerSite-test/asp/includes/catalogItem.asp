
<%class cls_catalogItem
Public iId, sTitle, dUpdatedTS, dCreatedTS, fields, iCatalogID, dOnlineUntill, dOnlineFrom, dDate, sPicExt, bForm
Private p_catalog
Private Sub Class_Initialize
iId=null
set p_catalog	= nothing
set fields	= server.CreateObject ("scripting.dictionary")
bForm=true
On Error Resume Next
iCatalogID=convertLng(decrypt(request("iCatalogID")))
pick(decrypt(request("iItemID")))
On Error Goto 0
end sub
public function picURL
picURL=C_VIRT_DIR & Application("QS_CMS_userfiles") & catalog.correFP & sPicExt
end function
Public function showPic(iThumbnailSize)
if not isLeeg(sPicExt) then
dim iTSize
if isNumeriek(iThumbnailSize) then
iTSize=iThumbnailSize
else
iTSize=catalog.iMaxThumbSize
end if
if not catalog.bAutoThumb then
showPic="<img border=""0"" src='" & C_VIRT_DIR & Application("QS_CMS_userfiles") & catalog.correFP & sPicExt & "' alt=" & """" & sanitize(sTitle) & """" & " />"
else
if not isLeeg(catalog.sFullImage) then
showPic=getImageLinkLB(catalog.sResizePicTo,catalog.correFP & sPicExt,iTSize,iId,sTitle,catalog.iMaxThumbSize)
showPic=showPic&"<br /><span style='margin:3px;font-size:smaller'>" & catalog.sFullImage & "</span>"
else
showPic="<img border=""0"" src='"& C_DIRECTORY_QUICKERSITE & "/showThumb.aspx?maxSize=" & catalog.iMaxThumbSize & "&FSR=0&img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & catalog.correFP & sPicExt & "' alt=" & """" & sanitize(sTitle) & """" & " />"
end if
end if
'if catalog.bUseShadow then
'showPic=wrapInShadowCSS(showPic)
'end if
end if
end function
Public function removePic
dim fso
set fso=server.CreateObject ("scripting.filesystemobject")
if fso.FileExists (server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalog.correFP& sPicExt)) then
fso.DeleteFile server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalog.correFP & sPicExt)
end if
set fso=nothing
sPicExt=""
end function
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblCatalogItem where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
sTitle	= rs("sTitle")
iCatalogID	= rs("iCatalogID")
dUpdatedTS	= rs("dUpdatedTS")
dCreatedTS	= rs("dCreatedTS")
dOnlineFrom	= rs("dOnlineFrom")
dOnlineUntill	= rs("dOnlineUntill")
dDate	= rs("dDate")
sPicExt	= rs("sPicExt")
bForm	= rs("bForm")
dim rsFields, sValue,iFieldID
set rsFields=db.execute("select iFieldId,sValue from tbCatalogItemFields where iItemID="&iId)
fields.RemoveAll()
while not rsFields.eof
iFieldID	= rsFields(0)
sValue	= rsFields(1)
fields.Add iFieldID,sValue
rsFields.movenext
wend
set rsFields=nothing
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sTitle) then
check=false
message.AddError("err_mandatory")
exit function
end if
dim catalogFields, catalogField
set catalogFields=catalog.fields("")
for each catalogField in catalogFields
if catalogFields(catalogField).bMandatory then
if isLeeg(fields(catalogField)) then
message.AddError("err_mandatory")
check=false
exit function
end if
end if
next
set catalogFields=nothing
End Function
Public Function Save()
dim rs
if check() then
save=true
else
save=false
exit function
end if
set rs = db.GetDynamicRS
if convertGetal(iId)=0 then
rs.Open "select * from tblCatalogItem where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblCatalogItem where iId="& iId
end if
rs("sTitle")	= left(sTitle,255)
rs("dUpdatedTS")	= now()
rs("iCatalogID")	= iCatalogID
rs("dOnlineUntill")	= dOnlineUntill
rs("dOnlineFrom")	= dOnlineFrom
rs("dDate")	= dDate
rs("sPicExt")	= sPicExt
rs("bForm")	= bForm
rs.Update()
iId = clng(rs("iId"))
dUpdatedTS	= rs("dUpdatedTS")
dCreatedTS	= rs("dCreatedTS")
rs.close
Set rs = nothing
'fields
dim rsFields
set rsFields=db.execute("delete from tbCatalogItemFields where iItemID="&iId)
set rsFields=nothing
set rsFields=db.GetDynamicRS
rsFields.Open "select * from tbCatalogItemFields where 1=2"
dim field
for each field in fields
rsFields.AddNew
rsFields("iItemID")=iID
rsFields("iFieldID")=field
rsFields("sValue")=fields(field)
rsFields.Update
next
'Response.End 
rsFields.close
set rsFields=nothing
catalog.clearRSSCache
end function
public function getRequestValues()
sTitle	= convertStr(Request.Form ("sTitle"))
dOnlineFrom	= convertDateFromPicker(Request.Form ("dOnlineFrom"))
dOnlineUntill	= convertDateFromPicker(Request.Form ("dOnlineUntill"))
dDate	= convertDateFromPicker(Request.Form ("dDate"))
iCatalogID	= convertLng(decrypt(Request.Form ("iCatalogID")))
bForm	= convertBool(Request.Form ("bForm"))
dim catalogFields, catalogField
set catalogFields=catalog.fields("")
for each catalogField in catalogFields
select case catalogFields(catalogField).sType
case sb_date
fields(catalogField)=convertCalcDate(Request.Form(encrypt(catalogField)))
case else
fields(catalogField)=removeEmptyP(Request.Form(encrypt(catalogField)))
end select
next
set catalogFields=nothing
end function
public function catalog
if p_catalog is nothing then
set p_catalog=new cls_catalog
p_catalog.pick(iCatalogID)
end if
set catalog=p_catalog
end function
public function remove
removePic()
dim cFiles,cFile
set cFiles=files
for	each cFile in cFiles
cFiles(cFile).remove()
next
dim cSubmissions,cSub
set cSubmissions=submissions
for	each cSub in cSubmissions
cSubmissions(cSub).remove(catalog.form.fields)
next
set cSubmissions=nothing
catalog.clearRSSCache
dim rs
set rs=db.execute("delete from tbCatalogItemFields where iItemID="&iId)
set rs=nothing
set rs=db.execute("delete from tblCatalogItem where iId="&iId)
set rs=nothing
end function
public sub removeFile(iFileID)
dim file
set file=new cls_catalogItemFile
file.pick(iFileID)
file.remove()
set file=nothing
end sub
public function files
set files=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, file
set rs=db.execute("select iId from tblCatalogItemFiles where iItemId="& iId)
while not rs.eof
set file=new cls_catalogItemFile
file.pick(rs(0))
files.Add file.iId,file
rs.movenext
set file=nothing
wend
set rs=nothing
end if
end function
public function submissions
set submissions=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, submission
set rs=db.execute("select iId from tblFormSubmission where iItemId="& iId)
while not rs.eof
set submission=new cls_submission
submission.pick(rs(0))
submissions.Add submission.iId,submission
rs.movenext
set submission=nothing
wend
set rs=nothing
end if
end function
Public Property get sDateAndTitle
if isLeeg(dDate) then
sDateAndTitle=sTitle
else
sDateAndTitle=convertEuroDate(dDate)& ": " & sTitle
end if
end property
public function sFicheRSS(sItemTemplate)
sFicheRSS=treatConstants(sItemTemplate,true)
sFicheRSS=replace(sFicheRSS,"{ITEMTITLE}",sanitize(sTitle),1,-1,1)
sFicheRSS=replace(sFicheRSS,"{ITEMDATE}",convertEuroDate(dDate),1,-1,1)
sFicheRSS=replace(sFicheRSS,"{ITEMPICTURE}",picURL,1,-1,1)
sFicheRSS=replace(sFicheRSS,"{ITEMENCID}",encrypt(iId),1,-1,1)
'sFicheRSS=replace(sFicheRSS,"{ITEMFORMLINK}","default.asp?pageAction=itemform&amp;iPageID="& encrypt(selectedPage.iId) &"&amp;iItemID="& encrypt(iId),1,-1,1)
dim catalogFields, catalogField, sItem
set catalogFields=catalog.fields("public")
for each catalogField in catalogFields
sItem=""
select case catalogFields(catalogField).sType 
case sb_date
sItem=convertDateToPicker(fields(catalogField))
case sb_checkbox
sItem=convertCheckedYesNo(fields(catalogField))
case else
sItem=fields(catalogField)
end select
if isNull(sItem) then sItem=""
sFicheRSS=replace(sFicheRSS,"{" & ucase(catalogFields(catalogField).sName) & "}",sItem,1,-1,1) 
next
'files
dim fileTypes, ft, f, copyfiles
set fileTypes=catalog.fileTypes
set copyfiles=files
for each ft in fileTypes
for each f in copyfiles
if convertGetal(copyfiles(f).iFileTypeID)=convertGetal(ft) then
sFicheRSS=replace(sFicheRSS,"{" & ucase(fileTypes(ft).sName) & "}",copyfiles(f).simpleurl,1,-1,1) 
end if
next
sFicheRSS=replace(sFicheRSS,"{" & ucase(fileTypes(ft).sName) & "}","",1,-1,1) 
next
set copyfiles=nothing
end function
public function sFicheITEMVIEW
sFicheITEMVIEW=catalog.sItemView
sFicheITEMVIEW=replace(sFicheITEMVIEW,"{ITEMTITLE}",sanitize(sTitle),1,-1,1)
sFicheITEMVIEW=replace(sFicheITEMVIEW,"{ITEMDATE}",convertEuroDate(dDate),1,-1,1)
sFicheITEMVIEW=replace(sFicheITEMVIEW,"{ITEMPICTURE}",picURL,1,-1,1)
sFicheITEMVIEW=replace(sFicheITEMVIEW,"{ITEMENCID}",encrypt(iId),1,-1,1)
sFicheITEMVIEW=treatConstants(sFicheITEMVIEW,true)
dim catalogFields, catalogField, sItem
set catalogFields=catalog.fields("public")
for each catalogField in catalogFields
sItem=""
select case catalogFields(catalogField).sType 
case sb_date
sItem=convertDateToPicker(fields(catalogField))
case sb_checkbox
sItem=convertCheckedYesNo(fields(catalogField))
case else
sItem=fields(catalogField)
end select
if isNull(sItem) then sItem=""
sFicheITEMVIEW=replace(sFicheITEMVIEW,"{" & ucase(catalogFields(catalogField).sName) & "}",sItem,1,-1,1) 
next
'files
dim fileTypes, ft, f, copyfiles
set fileTypes=catalog.fileTypes
set copyfiles=files
for each ft in fileTypes
for each f in copyfiles
if convertGetal(copyfiles(f).iFileTypeID)=convertGetal(ft) then
sFicheITEMVIEW=replace(sFicheITEMVIEW,"{" & ucase(fileTypes(ft).sName) & "}",copyfiles(f).simpleurl,1,-1,1) 
end if
next
sFicheITEMVIEW=replace(sFicheITEMVIEW,"{" & ucase(fileTypes(ft).sName) & "}","",1,-1,1) 
next
set copyfiles=nothing
end function
Public Property Get sFiche
'template driven?
select case Request.QueryString ("viewtype")
case "1"
if not isLeeg(catalog.sRSSView1) then
sFiche=sFicheRSS(catalog.sRSSView1)
exit property
end if
case "2"
if not isLeeg(catalog.sRSSView2) then
sFiche=sFicheRSS(catalog.sRSSView2)
exit property
end if
case "3"
if not isLeeg(catalog.sRSSView3) then
sFiche=sFicheRSS(catalog.sRSSView3)
exit property
end if
end select
dim catalogFields, catalogField
set catalogFields=catalog.fields("public")
'picture
if not isLeeg(sPicExt) then
sFiche=sFiche & "<b>" & l("picture") &":</b><br />"
sFiche=sFiche & showPic(400) 
end if
'values
for each catalogField in catalogFields
if not isLeeg(fields(catalogField)) or catalogFields(catalogField).sType=sb_checkbox then
sFiche=sFiche & "<br /><i><b>" & catalogFields(catalogField).sName & "</b></i><br />" 
select case catalogFields(catalogField).sType 
case sb_text,sb_textarea,sb_select,sb_url,sb_email
sFiche=sFiche & LinkURLs(fields(catalogField))
case sb_richtext
sFiche=sFiche & fields(catalogField)
case sb_date
sFiche=sFiche & convertDateToPicker(fields(catalogField))
case sb_checkbox
sFiche=sFiche & convertCheckedYesNo(fields(catalogField))
end select
sFiche=sFiche&"<br />"
end if
next
'files
dim fileTypes, ft, f, copyfiles
set fileTypes=catalog.fileTypes
set copyfiles=files
if copyfiles.count>0 then
sFiche=sFiche & "<br /><i><b>" & l("files") & "</b></i><br />"
for each ft in fileTypes
for each f in copyfiles
if convertGetal(copyfiles(f).iFileTypeID)=convertGetal(ft) then
sFiche=sFiche & copyfiles(f).url & " - "
end if
next
next
sFiche=left(sFiche,len(sFiche)-3)
end if
set copyfiles=nothing
end property
Public sub checkOnline
if isLeeg(iId) then
Response.Redirect ("default.asp")
end if
if convertGetal(catalog.iCustomerID)<>convertGetal(cId) then
Response.Redirect ("default.asp")
end if
if not bForm then
Response.Redirect ("default.asp")
end if
if not catalog.bOnline then
Response.Redirect ("default.asp")
end if
end sub
Public function copy
if isNumeriek(iId) then
dim cFiles,cFile
set cFiles=files
iId=null
sTitle=l("copyof")&" "&sTitle
save()
for each cFile in cFiles
cFiles(cFile).copy(iId)
next
set cFiles=nothing
dim fso2, pw
set fso2=server.CreateObject ("scripting.filesystemobject")
pw=GeneratePassWord()
if fso2.FileExists (server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalog.correFP & sPicExt)) then
fso2.CopyFile server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalog.correFP & sPicExt),server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalog.correFP & encrypt(iId) & "_" & pw & "." & GetFileExtension(sPicExt))
sPicExt=encrypt(iId) & "_" & pw & "." & GetFileExtension(sPicExt)
save()
end if
set fso2=nothing
end if
end function
public property get sSeoTitle
sSeoTitle=quotrep(treatConstants(sTitle & " | " & catalog.sName & " | " & customer.siteTitle,true))
end property
end class%>

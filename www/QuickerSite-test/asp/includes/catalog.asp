
<%class cls_catalog
Public iId, sName, sItemName, sOrderItemsBy, iCustomerID, bOnline, dUpdatedTS, dCreatedTS, sFormTitle, iFormID
Public iPageSize, bAutoThumb, iMaxThumbSize, iAutoClose, sResizePicTo, bUseShadow, bSearchable, bPushRSS, sFullImage, overruleCID
Public sRSSView1,sRSSView2,sRSSView3,sItemView

Public sFilePath

public function correFP

	if not isLeeg(sFilePath) then

		if left(sFilePath,1)="/" then
			correFP=right(sFilePath,len(sFilePath)-1)
		end if
		
		if right(correFP,1)<>"/" then
			correFP=correFP& "/"
		end if

	end if

end function

Private Sub Class_Initialize
On Error Resume Next
iId=null
bOnline=true
bAutoThumb=false
bUseShadow=false
bSearchable=true
iFormID=null
iPageSize=10
iMaxThumbSize=200
sResizePicTo=640
iAutoClose=10
bPushRSS=false
sFullImage="Click image for full version"
overruleCID=null
pick(decrypt(request("iCatalogID")))
On Error Goto 0
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblCatalog where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
sName	= rs("sName")
sItemName	= rs("sItemName")
sOrderItemsBy	= rs("sOrderItemsBy")
iCustomerID	= rs("iCustomerID")
bOnline	= rs("bOnline")
dUpdatedTS	= rs("dUpdatedTS")
dCreatedTS	= rs("dCreatedTS")
sFormTitle	= rs("sFormTitle")
iFormID	= rs("iFormID")
iPageSize	= rs("iPageSize")
bAutoThumb	= rs("bAutoThumb")
iMaxThumbSize	= rs("iMaxThumbSize")
iAutoClose	= rs("iAutoClose")
sResizePicTo	= rs("sResizePicTo")
bUseShadow	= rs("bUseShadow")
bSearchable	= rs("bSearchable")
bPushRSS	= rs("bPushRSS")
sFullImage	= rs("sFullImage")
sRSSView1	= rs("sRSSView1")
sRSSView2	= rs("sRSSView2")
sRSSView3	= rs("sRSSView3")
sItemView	= rs("sItemView")
sFilePath	= rs("sFilePath")
end if
set RS = nothing
end if
end function
Public property get status
if bOnline then
status="Online"
else
status="Offline"
end if
end property
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sItemName) then
check=false
message.AddError("err_mandatory")
end if
if not isLeeg(iFormID) then
if isLeeg(sFormTitle) then
check=false
message.AddError("err_mandatory")
end if
end if
End Function
Public Function Save
if check() then
save=true
else
save=false
exit function
end if

dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblCatalog where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblCatalog where iId="& iId
end if
if isNull(overruleCID) then
rs("iCustomerID")	= cId
else
rs("iCustomerID")	= overruleCID
end if
rs("sName")	= sName
rs("sItemName")	= sItemName
rs("sOrderItemsBy")	= sOrderItemsBy
rs("bOnline")	= bOnline
rs("dUpdatedTS")	= now()
rs("sFormTitle")	= sFormTitle
rs("iFormID")	= iFormID
rs("iPageSize")	= iPageSize
rs("bAutoThumb")	= bAutoThumb
rs("iMaxThumbSize")	= iMaxThumbSize
rs("iAutoClose")	= iAutoClose
rs("sResizePicTo")	= sResizePicTo
rs("bUseShadow")	= bUseShadow
rs("bSearchable")	= bSearchable
rs("bPushRSS")	= bPushRSS
rs("sFullImage")	= sFullImage
rs("sRSSView1")	= sRSSView1
rs("sRSSView2")	= sRSSView2
rs("sRSSView3")	= sRSSView3
rs("sItemView")	= sItemView
rs("sFilePath")	= sFilePath
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
clearRSSCache
end function
public function clearRSSCache
dim rs, pageID
set rs=db.execute("select iId from tblPage where iCatalogId="& iId)
while not rs.eof
pageID=rs(0)
application("RSS"&cId&pageID)=""
application("RSS"&cId&pageID&"1")=""
application("RSS"&cId&pageID&"2")=""
application("RSS"&cId&pageID&"3")=""
rs.movenext
wend
set rs=nothing
end function
public function getRequestValues()
sName	= convertStr(Request.Form ("sName"))
sItemName	= convertStr(Request.Form ("sItemName"))
sOrderItemsBy	= convertStr(Request.Form ("sOrderItemsBy"))
bOnline	= convertBool(Request.Form ("bOnline"))
sFormTitle	= convertStr(Request.Form ("sFormTitle"))
iFormID	= convertLng(decrypt(Request.Form ("iFormID")))
iPageSize	= convertGetal(Request.Form ("iPageSize"))
bAutoThumb	= convertBool(Request.Form ("bAutoThumb"))
iMaxThumbSize	= convertGetal(Request.Form ("iMaxThumbSize"))
iAutoClose	= convertGetal(Request.Form ("iAutoClose"))
sResizePicTo	= convertGetal(Request.Form ("sResizePicTo"))
bUseShadow	= convertBool(Request.Form ("bUseShadow"))
bSearchable	= convertBool(Request.Form ("bSearchable"))
bPushRSS	= convertBool(Request.Form ("bPushRSS"))
sFullImage	= convertStr(Request.Form ("sFullImage"))
sRSSView1	= convertStr(Request.Form ("sRSSView1"))
sRSSView2	= convertStr(Request.Form ("sRSSView2"))
sRSSView3	= convertStr(Request.Form ("sRSSView3"))
sItemView	= convertStr(Request.Form ("sItemView"))
sFilePath	= convertStr(request.form("sFilePath"))
if iMaxThumbSize=0 then
iMaxThumbSize=200
iAutoClose=10
sResizePicTo=1024
end if
end function
Public function form
set form=new cls_form
form.pick(iFormID)
end function 
public property get isOnline
if bOnline then 
isOnline=l("yes")
else
isOnline=l("no")
end if
end property
public function remove
if not isLeeg(iId) then
dim copyItems
set copyItems=items
dim copyFields
set copyFields=fields("")
dim copyFiletypes
set copyFiletypes=filetypes
dim key
for each key in copyItems
copyItems(key).remove
next
for each key in copyFields
copyFields(key).remove
next
for each key in copyFiletypes
copyFiletypes(key).remove
next
dim rs
set rs=db.execute("update tblPage set iCatalogID=null where iCatalogID="&iId)
set rs=nothing
set rs=db.execute("delete from tblCatalogItem where iCatalogID="&iId)
set rs=nothing 
set rs=db.execute("delete from tblCatalog where iId="&iId)
set rs=nothing
end if
end function
public function filetypes
set filetypes=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, filetype, sql
sql="select iId from tblCatalogfiletype where iCatalogID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set filetype=new cls_catalogfiletype
filetype.pick(rs(0))
filetypes.Add filetype.iID,filetype
set filetype=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function fields(sType)
set fields=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, field, sql
sql="select iId from tblCatalogField where iCatalogID=" & iId
sType=split(sType,",")
dim tKey
for tKey=lbound(sType) to ubound(sType)
if sType(tKey)="search" then sql=sql&" and bSearchField="&getSQLBoolean(true)&" "
if sType(tKey)="public" then sql=sql&" and bPublic="&getSQLBoolean(true)&" "
next
sql=sql&" order by iRang asc"
set rs=db.execute(sql)
while not rs.eof
set field=new cls_catalogField
field.pick(rs(0))
fields.Add field.iID,field
set field=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function items
set items=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, cItem
set rs=db.execute("select iId from tblCatalogItem where iCatalogID="& iId & " order by sTitle")
while not rs.eof
set cItem=new cls_catalogItem
cItem.pick(rs(0))
items.Add cItem.iId, cItem
set cItem=nothing
rs.movenext
wend
set rs=nothing
end if
end function
Public Function showSelectedFileType(mode,selected)
dim copyCats
set copyCats=filetypes
select case mode
case "single"
showSelectedFileType=copyCats(convertGetal(decrypt(selected))).sName
case "option"
Dim key
For each key in copyCats
showSelectedFileType = showSelectedFileType & "<option value='" & encrypt(key) & "'"
If convertStr(selected) = convertStr(encrypt(key)) Then
showSelectedFileType = showSelectedFileType & " selected"
End If
showSelectedFileType = showSelectedFileType & ">" & copyCats(key).sName & "</option>"
Next
end select
End Function
public function copy()
if isNumeriek(iId) then
dim cFields, fKey, cFiletypes
set cFields=fields("")
set cFiletypes=filetypes
iId=null
save()
for each fKey in cFields
cFields(fKey).copy(iId)
next
for each fKey in cFiletypes
cFiletypes(fKey).copy(iId)
next
set cFiletypes=nothing
set cFields=nothing
end if
end function
public function copyToCustomer(oCID)
if isNumeriek(iId) then
overruleCID	= oCID
dim oldId
oldID=iId
dim cFields, fKey, cFiletypes
set cFields=fields("")
set cFiletypes=filetypes
iId=null
save()
for each fKey in cFields
cFields(fKey).copy(iId)
next
for each fKey in cFiletypes
cFiletypes(fKey).copy(iId)
next
set cFiletypes=nothing
set cFields=nothing
db.execute("update tblPage set iCatalogId=" & iId & " where iCatalogId=" & oldID & " and iCustomerID=" & oCID)
end if
end function
end class%>

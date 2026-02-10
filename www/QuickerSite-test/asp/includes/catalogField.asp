
<%class cls_catalogField
Public iId
Public iCatalogID
Public sName
Public sType
Public sValues
Public iRang
Public bMandatory
Public bSearchField
Public bPublic
Private Sub Class_Initialize
On Error Resume Next
iId=null
bMandatory=false
bSearchField=true
bPublic=true
iCatalogID=convertLng(decrypt(request("iCatalogID")))
pick(decrypt(request("iFieldID")))
On Error Resume Next
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblCatalogField where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
iCatalogID	= rs("iCatalogID")
sName	= rs("sName")
sType	= rs("sType")
sValues	= rs("sValues")
iRang	= rs("iRang")
bMandatory	= rs("bMandatory")
bSearchField	= rs("bSearchField")
bPublic	= rs("bPublic")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sType) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
if sType=sb_select and isLeeg(sValues) then
check=false
message.AddError("err_mandatory")
end if
End Function
Public Function Save
dim rs, isNew
isNew=false
if check() then
save=true
else
save=false
exit function
end if
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblCatalogField where 1=2"
rs.AddNew
isNew=true
dim rangRS
set rangRS=db.execute("select count(*) from tblCatalogField where iCatalogID=" & iCatalogID)
iRang=convertGetal(clng(rangRS(0)))+1
else
rs.Open "select * from tblCatalogField where iId="& iId
end if
rs("iCatalogID")	= iCatalogID
rs("sName")	= sName
rs("sType")	= sType
rs("sValues")	= sValues
rs("iRang")	= iRang
rs("bMandatory")	= bMandatory
rs("bSearchField")	= bSearchField
rs("bPublic")	= bPublic
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
if isNew then
dim list
set list=db.execute("select distinct(tblCatalogItem.iId) from tblCatalogItem where " & sqlNull("tblCatalogItem.iCatalogID",iCatalogID))
set rs = db.GetDynamicRS
rs.open "select * from tbCatalogItemFields where 1=2"
while not list.eof
rs.AddNew()
rs("iItemID")=list(0)
rs("iFieldID")=iId
rs("sValue")=""
rs.Update()
list.movenext
wend
rs.close()
end if
end function
public function getRequestValues()
iCatalogID	= convertLng(decrypt(Request.Form ("iCatalogID")))
sType	= convertStr(Request.Form ("sType"))
sName	= convertStr(Request.Form ("sName"))
sValues	= Request.Form ("sValues")
bMandatory	= convertBool(Request.Form ("bMandatory"))
bSearchField	= convertBool(Request.Form ("bSearchField"))
bPublic	= convertBool(Request.Form ("bPublic"))
end function
public property get isMandatory
if bMandatory then 
isMandatory=l("yes")
else
isMandatory=l("no")
end if
end property
public property get isSearchField
if bSearchField then 
isSearchField=l("yes")
else
isSearchField=l("no")
end if
end property
public property get isPublic
if bPublic then 
isPublic=l("yes")
else
isPublic=l("no")
end if
end property
public function catalog
set catalog=new cls_catalog
catalog.pick(iCatalogID)
end function
public function moveUp
if isNumeriek(iId) then
if convertGetal(iRang)=1 then
exit function
else
db.execute("update tblCatalogField set iRang=iRang-1 where " & sqlNull("iCatalogID",iCatalogID)  & " and iId="& iId)
db.execute("update tblCatalogField set iRang=iRang+1 where " & sqlNull("iCatalogID",iCatalogID)  & " and iId<>"& iId & " and iRang=" & iRang-1)
end if
end if
end function
public function moveDown
if isNumeriek(iId) then
if convertGetal(iRang)=convertGetal(catalog.fields("").count) then
exit function
else
db.execute("update tblCatalogField set iRang=iRang+1 where " & sqlNull("iCatalogID",iCatalogID)  & " and iId="& iId)
db.execute("update tblCatalogField set iRang=iRang-1 where " & sqlNull("iCatalogID",iCatalogID)  & " and iId<>"& iId & " and iRang=" & iRang+1)
end if
end if
end function
public function remove
if isNumeriek(iId) then
dim rs
set rs=db.execute("update tblCatalogField set iRang=iRang-1 where " & sqlNull("iCatalogID",iCatalogID)  & " and iRang>"& iRang)
set rs=nothing
set rs=db.execute("delete from tbCatalogItemFields where iFieldId="& iId)
set rs=nothing
set rs=db.execute("delete from tblCatalogField where iId="& iId)
set rs=nothing
end if
end function
Public Function showSelected(selected)
dim listArr, i, list
set list=server.CreateObject ("scripting.dictionary")
listArr=split(sValues,vbcrlf)
for i=lbound(listArr) to ubound(listArr)
On Error Resume Next
list.Add listArr(i),""
On Error Goto 0
next
Dim key
For each key in list
showSelected = showSelected & "<option value='" & key & "'"
If selected = key Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & key & "</option>"
Next
End Function
public function copy(catID)
if isNumeriek(iId) then
iCatalogID=catID
iId=null
save()
end if
end function
end class%>

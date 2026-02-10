
<%class cls_CatalogFileType
Public iId
Public sName
Public iCatalogID
Private Sub Class_Initialize
On Error Resume Next
iId=null
iCatalogID=convertLng(decrypt(request("iCatalogID")))
pick(decrypt(request("iFileTypeID")))
On Error Goto 0
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblCatalogFileType where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
sName	= rs("sName")
iCatalogID	= rs("iCatalogID")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
End Function
Public Function Save
dim rs
if check() then
save=true
else
save=false
exit function
end if
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblCatalogFileType where 1=2"
rs.AddNew
else
rs.Open "select * from tblCatalogFileType where iId="& iId
end if
rs("iCatalogID")	= iCatalogID
rs("sName")	= sName
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function getRequestValues()
sName	= convertStr(Request.Form ("sName"))
iCatalogID	= convertLng(decrypt(Request.Form ("iCatalogID")))
end function
public function remove
if isNumeriek(iId) then
dim rs
set rs=db.execute("delete from tblCatalogItemFiles where iFileTypeID="& iID)
set rs=nothing
set rs=db.execute("delete from tblCatalogFileType where iId="& iID)
set rs=nothing
end if
end function
public function catalog
set catalog=new cls_catalog
catalog.pick(iCatalogID)
end function
public function copy(catID)
if isNumeriek(iId) then
iCatalogID=catID
iId=null
save()
end if
end function
end class%>

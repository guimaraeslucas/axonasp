
<%class cls_shopCategory
Public iId,bOnline,iParentCatID,sName,dCreatedTS,dUpdatedTS
Private Sub Class_Initialize
iId=null
bOnline=true
iParentCatID=null
end sub
Private Sub Class_Terminate
end sub
public function subcategories
set subcategories=server.createobject("scripting.dictionary")
dim rs,category
set rs=db.execute("select iId from tblQShopCategory where iCustomerID=" & cId & " and iParentCatID=" & iId & " order by sName")
while not rs.eof
set category=new cls_shopCategory
category.pick(rs(0))
subcategories.Add category.iId,category
rs.movenext
set category=nothing
wend
set rs=nothing
end function
public function list
set list=server.createobject("scripting.dictionary")
dim rs,category
set rs=db.execute("select iId from tblQShopCategory where iParentCatID is null and iCustomerID=" & cId & " order by sName")
while not rs.eof
set category=new cls_shopCategory
category.pick(rs(0))
list.Add category.iId,category
rs.movenext
set category=nothing
wend
set rs=nothing
end function
Public Function Pick(id)
dim sql, RS, sValue
if isNumeriek(id) then
sql = "select * from tblQShopCategory where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
dCreatedTS	= rs("dCreatedTS")
dUpdatedTS	= rs("dUpdatedTS")
bOnline	= rs("bOnline")
iParentCatID	= rs("iParentCatID")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
exit function
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
if convertGetal(iParentCatID)=0 then iParentCatID=null
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblQShopCategory where 1=2"
rs.AddNew
rs("dCreatedTS")	= now()
else
rs.Open "select * from tblQShopCategory where iId="& iId
end if
rs("sName")	= trim(left(sName,255))
rs("bOnline")	= convertBool(bOnline)
rs("iParentCatID")	= convertLng(iParentCatID)
rs("dUpdatedTS")	= now()
rs("iCustomerID")	= cId
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
if not convertBool(bOnline) then
db.execute("update tblQShopCategory set bOnline=" & getSQLBoolean(false) & " where iParentCatID=" & convertGetal(iId))
end if
end function
public function delete
db.execute("update  tblQShopCategory set iParentCatID=null where iParentCatID=" & convertGetal(iId))
db.execute("delete from tblQShopProdCat where iCategoryID=" & convertGetal(iId))
db.execute("delete from tblQShopCategory where iId=" & convertGetal(iId))
end function
Public property get bShowParentCatDropDown
dim rs
set rs=db.execute("select iId,sName from tblQShopCategory where iParentCatID =" & convertGetal(iId))
if not rs.eof then
bShowParentCatDropDown=false
else
bShowParentCatDropDown=true
end if
set rs=nothing
end property
Public property get bShowOnlineCB
bShowOnlineCB=true
if convertGetal(iParentCatID)<>0 then
dim rs,helpBool
set rs=db.execute("select bOnline from tblQShopCategory where iId =" & convertGetal(iParentCatID))
helpBool=rs(0)
if convertBool(helpBool) then
bShowOnlineCB=true
else
bShowOnlineCB=false
end if
set rs=nothing
end if
end property
Public Function showParentCat(mode, selected)
dim rs
set rs=db.execute("select iId,sName from tblQShopCategory where iCustomerID=" & cId & " and iParentCatID is null and iId<>" & convertGetal(iId) & " order by sName")
dim catDict,catID,catName
set catDict=server.createobject("scripting.dictionary")
while not rs.eof
catID=rs(0)
catName=rs(1)
catDict.Add catID,catName
rs.movenext
wend
Select Case mode
Case "single"
showParentCat = catDict(convertGetal(selected))
Case "option"
Dim key
For each key in catDict
showParentCat = showParentCat & "<option value=""" & key & """"
If convertStr(selected) = convertStr(key) Then
showParentCat = showParentCat & " selected"
End If
showParentCat = showParentCat & ">" & catDict(key) & "</option>"
Next
End Select
End Function
end class%>


<%class cls_shopMake
Public iId,bOnline,sName,dCreatedTS,dUpdatedTS,sLogo
Private Sub Class_Initialize
iId=null
bOnline=true
end sub
Private Sub Class_Terminate
end sub
public function list
set list=server.createobject("scripting.dictionary")
dim rs,make
set rs=db.execute("select iId from tblQShopMake where iCustomerID=" & cId & " order by sName")
while not rs.eof
set make=new cls_shopmake
make.pick(rs(0))
list.Add make.iId,make
rs.movenext
set make=nothing
wend
set rs=nothing
end function
Public Function Pick(id)
dim sql, RS, sValue
if isNumeriek(id) then
sql = "select * from tblQShopmake where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
dCreatedTS	= rs("dCreatedTS")
dUpdatedTS	= rs("dUpdatedTS")
bOnline	= rs("bOnline")
sLogo	= rs("sLogo")
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
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblQShopmake where 1=2"
rs.AddNew
rs("dCreatedTS")	= now()
else
rs.Open "select * from tblQShopmake where iId="& iId
end if
rs("sName")	= trim(left(sName,255))
rs("bOnline")	= convertBool(bOnline)
rs("sLogo")	= trim(left(sLogo,255))
rs("dUpdatedTS")	= now()
rs("iCustomerID")	= cId
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function delete
'db.execute("update  tblQShopCategory set iParentCatID=null where iParentCatID=" & convertGetal(iId))
db.execute("delete from tblQShopMake where iId=" & convertGetal(iId))
end function
Public Function showMake(mode, selected)
dim rs
set rs=db.execute("select iId,sName from tblQShopMake where iCustomerID=" & cId & " order by sName")
dim makeDict,makename,makeid
set makeDict=server.createobject("scripting.dictionary")
while not rs.eof
makeid=rs(0)
makename=rs(1)
makeDict.Add makeid,makename
rs.movenext
wend
Select Case mode
Case "single"
showMake = makeDict(convertGetal(selected))
Case "option"
Dim key
For each key in makeDict
showMake = showMake & "<option value=""" & key & """"
If convertStr(selected) = convertStr(key) Then
showMake = showMake & " selected"
End If
showMake = showMake & ">" & makeDict(key) & "</option>"
Next
End Select
End Function
end class%>

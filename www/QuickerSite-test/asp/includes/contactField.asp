
<%class cls_contactField
Public iId
Public iCustomerID
Public sFieldName
Public sType
Public sValues
Public iRang
Public bMandatory
Public bSearchField
Public bProfile
Private Sub Class_Initialize
on error resume next
iId=null
bMandatory=true
bSearchField=true
bProfile=true
pick(decrypt(request("iFieldID")))
on error goto 0
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblContactField where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
iCustomerID	= rs("iCustomerID")
sFieldName	= rs("sFieldName")
sType	= rs("sType")
sValues	= rs("sValues")
iRang	= rs("iRang")
bMandatory	= rs("bMandatory")
bSearchField	= rs("bSearchField")
bProfile	= rs("bProfile")
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
if isLeeg(sFieldName) then
check=false
message.AddError("err_mandatory")
end if
if sType=sb_select and isLeeg(sValues) then
check=false
message.AddError("err_mandatory")
end if
if bMandatory and not bProfile then
check=false
message.AddError("err_mandatoryprofile")
end if
End Function
Public Function Save
dim rs, isNew
if check() then
save=true
else
save=false
exit function
end if
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblContactField where 1=2"
rs.AddNew
isNew=true
dim rangRS
set rangRS=db.execute("select count(*) from tblContactField where iCustomerID=" & cID)
iRang=convertGetal(rangRS(0))+1
else
rs.Open "select * from tblContactField where iId="& iId
isNew=false
end if
rs("iCustomerID")	= cId
rs("sFieldName")	= sFieldName
rs("sType")	= sType
rs("sValues")	= sValues
rs("iRang")	= iRang
rs("bMandatory")	= bMandatory
rs("bSearchField")	= bSearchField
rs("bProfile")	= bProfile
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
if isNew then
dim list
set list=db.execute("select distinct(tblContact.iId) from tblContact where tblContact.iCustomerID="& cId)
set rs = db.GetDynamicRS
rs.open "select * from tblContactValues where 1=2"
while not list.eof
rs.AddNew()
rs("iContactID")=list(0)
rs("iFieldID")=iId
rs("sValue")=""
rs.Update()
list.movenext
wend
rs.close()
end if
end function
public function getRequestValues()
iCustomerID	= cId
sType	= convertStr(Request.Form ("sType"))
sFieldName	= convertStr(Request.Form ("sFieldName"))
sValues	= Request.Form ("sValues")
bMandatory	= convertBool(Request.Form ("bMandatory"))
bSearchField	= convertBool(Request.Form ("bSearchField"))
bProfile	= convertBool(Request.Form ("bProfile"))
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
public property get isProfileField
if bProfile then 
isProfileField=l("yes")
else
isProfileField=l("no")
end if
end property
public function moveUp
if isNumeriek(iId) then
if convertGetal(iRang)=1 then
exit function
else
db.execute("update tblContactfield set iRang=iRang-1 where "& sqlCustIdCF &" and iId="& iId)
db.execute("update tblContactfield set iRang=iRang+1 where "& sqlCustIdCF &" and iId<>"& iId & " and iRang=" & iRang-1)
end if
end if
end function
public function moveDown
if isNumeriek(iId) then
if convertGetal(iRang)=convertGetal(customer.contactFields(false).count) then
exit function
else
db.execute("update tblContactfield set iRang=iRang+1 where "& sqlCustIdCF &" and iId="& iId)
db.execute("update tblContactfield set iRang=iRang-1 where "& sqlCustIdCF &" and iId<>"& iId & " and iRang=" & iRang+1)
end if
end if
end function
public function remove
if isNumeriek(iId) then
dim rs
set rs=db.execute("update tblContactfield set iRang=iRang-1 where "& sqlCustIdCF &" and iRang>"& iRang)
set rs=nothing
set rs=db.execute("delete from tblContactValues where iFieldId="& iId)
set rs=nothing
set rs=db.execute("delete from tblContactfield where iId="& iId)
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
showSelected = showSelected & "<option value=" & """" & quotrep(key) &  """"
If selected = key Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & key & "</option>"
Next
End Function
end class%>

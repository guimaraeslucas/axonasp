
<%class cls_contactSearch
Public recordCount
Public searchFields
Public inputFields
Public dateFields
Public pageSize
Public origPageSize
Public sEmail
Public rowCount
Public mode
Public iStatus
Public sNickName
private h
Private sub class_initialize
on error resume next
set h=new md5
set inputFields=server.CreateObject ("scripting.dictionary")
set dateFields=server.CreateObject ("scripting.dictionary")
rowCount=0
if not isLeeg(Request.Form ("pageSize")) then
pageSize=Request.Form ("pageSize")
else
if customer.hasManyContacts then
pageSize=100
else
pageSize=100000
end if
end if
origPageSize=pageSize
mode=l("search")
if not isLeeg(Request.Form("btnaction")) then
mode=Request.Form("btnaction")
end if
on error goto 0
end sub
Public function getRequestValues
dim sfId
for each sfId in searchFields
if not isLeeg(Request.Form(encrypt(sfId))) then
inputFields.Add sfId,Request.Form(encrypt(sfId))
end if
if not isLeeg(Request.Form("from"&encrypt(sfId))) then
dateFields.Add "from"&sfId,Request.Form("from"&encrypt(sfId))
end if
if not isLeeg(Request.Form("untill"&encrypt(sfId))) then
dateFields.Add "untill"&sfId,Request.Form("untill"&encrypt(sfId))
end if
next
sEmail=Request.Form ("sEmail")
iStatus=Request.Form ("iStatus")
sNickName=Request.Form ("sNickName")
end function
Public Function getRows()
On Error Resume Next
initSF()
dim searchFieldsIDs, sfId
for each sfId in searchFields
searchFieldsIDs=searchFieldsIDs&sfId&","
next
if not isLeeg(searchFieldsIDs) then
searchFieldsIDs=left(searchFieldsIDs,len(searchFieldsIDs)-1)
end if
dim sql, where, orderBy
sql="select [TOP] tblContactValues.iFieldID, tblContactValues.iContactID, tblContactValues.sValue, tblContact.sEmail, tblContact.iStatus, tblContact.sNickname, tblContact.sAvatar, tblContact.dLastLoginTS "
sql=sql&" from (tblContactValues INNER JOIN tblContact on tblContact.iID=tblContactValues.iContactID) "
sql=sql&" INNER JOIN tblContactField on tblContactField.iId=tblContactValues.iFieldID"
where=" where tblContact.iCustomerID="& cId
if not isLeeg(searchFieldsIDs) then
where=where& " and tblContactValues.iFieldID in ("&searchFieldsIDs&")"
end if
if not isLeeg(sEmail) then
where=where& " and tblContact.sEmail like ('%" & cleanUp(sEmail) & "%') "
end if
if not isLeeg(sNickName) then
where=where& " and tblContact.sNickName like ('%" & cleanUp(sNickName) & "%') "
end if
if not isLeeg(iStatus) then
where=where& " and tblContact.iStatus in ("&iStatus&")"
end if
dim inputField
for each inputField in inputFields
if not isLeeg(inputFields(inputField)) then
dim gCF
set gCF=new cls_contactField
gCF.pick(inputField)
select case convertStr(gCF.sType)
case convertStr(sb_select)
where=where& " and tblContact.iID in (select iContactID from tblContactValues where tblContactValues.iFieldID="& inputField &" and tblContactValues.sValue like ('"& cleanUp(inputFields(inputField)) &"')) "
case else
where=where& " and tblContact.iID in (select iContactID from tblContactValues where tblContactValues.iFieldID="& inputField &" and tblContactValues.sValue like ('%"& cleanUp(inputFields(inputField)) &"%')) "
end select
set gCF=nothing
end if
next
dim fieldName, fromDate, untillDate, hasDates
hasDates=false
for each inputField in dateFields
if not isLeeg(dateFields(inputField)) then
fieldName=replace(inputField,"from","")
fieldName=replace(fieldName,"untill","")
if (left(inputField,4))="from" then
hasDates=true
fromDate	= convertCalcDate(convertEuroDate(roomDate(convertDateFromPicker(dateFields(inputField)),"down")))
end if
if (left(inputField,6))="untill" then
hasDates=true
untillDate	= convertCalcDate(convertEuroDate(roomDate(convertDateFromPicker(dateFields(inputField)),"up")))
end if
end if
next
if hasDates and isLeeg(fromDate) then
fromDate	= convertCalcDate(convertEuroDate(roomDate("","down")))
end if
if hasDates and isLeeg(untillDate) then
untillDate	= convertCalcDate(convertEuroDate(roomDate("","up")))
end if
if hasDates then
where=where& " and tblContact.iID in (select iContactID from tblContactValues where tblContactValues.iFieldID="& fieldName &" and tblContactValues.sValue>='"& fromDate &"' and tblContactValues.sValue<='"& untillDate &"')"
end if
orderBy=" order by  tblContactValues.iContactID desc, tblContactField.iRang asc"
dim rs
set rs = db.GetDynamicRS
sql=sql	& where & orderBy & "[LIMIT]"
sql=getTOPSyntax(getUp(pageSize*searchFields.count),sql)
'Response.Write sql
'Response.Flush 
'Response.End 
rs.open sql
if not rs.EOF then
getrows=rs.getrows
recordCount=rs.recordCount
else
getrows=null
end if
set rs = nothing
dumperror sql & where & orderBy,err
On Error Goto 0
end function
Public function resultTable
initSF()
dim p_getrows, j, searchField, headerRow, copyContactID, row, rows, start, copyContactEmail,copyIstatus,copyNickname,copyAvatar,copyLastLoginTS
p_getrows=getRows
if not isNull(p_getrows) then
'headerRow
headerRow=headerRow&"<tr>"
for each searchField in searchFields
headerRow=headerRow&"<th><b>"& cleanUpStr(searchFields(searchField).sFieldname) &"</b></th>"
next
if customer.bUseAvatars then
headerRow=headerRow&"<th>&nbsp;</th>"
end if
headerRow=headerRow&"<th>iId</th>"
headerRow=headerRow&"<th>"&l("nickname")&"</th><th><b>E-mail</b></th>"
headerRow=headerRow&"<th>&nbsp;</th><th>last login</th><th>&nbsp;</th><th>&nbsp;</th>"
headerRow=headerRow&"</tr>"
'bepalen pageSize
if searchFields.count>0 then
if pageSize<>"all" then
pageSize=pageSize*searchFields.count
end if 
end if
if pageSize="all" then pageSize=recordCount
if convertGetal(pageSize)>convertGetal(recordCount) then pageSize=recordCount
for j=0 to pageSize-1
if copyContactID<>p_getrows(1,j) and not isLeeg(copyContactID) then
AddRow rows,row,copyContactID,copyContactEmail,copyIstatus,copyNickname,copyAvatar,copyLastLoginTS
end if
copyContactID=p_getrows(1,j)
copyContactEmail=p_getrows(3,j)
copyIstatus=p_getrows(4,j)
copyNickname=p_getrows(5,j)
copyAvatar=p_getrows(6,j)
copyLastLoginTS=p_getrows(7,j)
for each searchField in searchFields
if p_getrows(0,j)=searchField then
row=row&"<td>"
select case searchFields(searchField).sType
case sb_text,sb_textarea,sb_select
row=row&sanitize(p_getrows(2,j))
case sb_richtext
row=row&removeHTML(p_getrows(2,j))
case sb_checkbox
row=row&convertCheckedYesNo(sanitize(p_getrows(2,j)))
case sb_date
row=row&convertDateToPicker(sanitize(p_getrows(2,j)))
case sb_url
if not isLeeg(p_getrows(2,j)) then
row=row&"<a target='"& generatePassword &"' href="""& sanitize(p_getrows(2,j)) & """>" & sanitize(p_getrows(2,j)) & "</a>"
end if
case sb_email
if not isLeeg(p_getrows(2,j)) then
row=row&"<a href='mailto:"& p_getrows(2,j) & "'>" & sanitize(p_getrows(2,j)) & "</a>"
end if
end select
row=row&"&nbsp;</td>"
end if
next
if j=pageSize-1  then
AddRow rows,row,p_getrows(1,j),p_getrows(3,j),p_getrows(4,j),p_getrows(5,j),p_getrows(6,j),p_getrows(7,j)
end if
next
end if
if searchFields.count>0 then
recordCount=recordCount/searchFields.count
end if
select case mode
case l("search")
resultTable=""
if customer.hasManyContacts then
if not convertBool(Request.Form ("postBack")) then
resultTable="<p align=center>"& l("total") &": <b>" & customer.nmrbContacts & "</b> "& l("contacts") &"</p>"
end if
end if
resultTable=resultTable&"<p align=center>"& l("totalshown") &": <b>" & convertGetal(recordCount) & "</b> "& l("contacts") &"</p>"
if recordCount>0 then
resultTable=resultTable&"<table align=""center"" cellpadding=""4"" cellspacing=""0"" class=""sortable"" id=""contactSearch"" border=""1"">"& headerRow & rows & "</table>"
end if
case l("excel")
dim excelfile
set excelfile=new cls_excelfile
excelfile.export("<table>" & headerRow & rows & "</table>")
resultTable="<p align=center><b>"& excelfile.downloadLink &"</b></p>"
set excelfile=nothing
end select
end function
private function addRow(byref rows, byref row, contactID, contactEmail, iStatus, sNickname, sAvatar, lastLoginTS)
select case convertGetal(iStatus)
case 10
iStatus="cs_silent"
case 20
iStatus="cs_profile"
case 30
iStatus="cs_read"
case 40
iStatus="cs_write"
case 50
iStatus="cs_readwrite"
end select 
dim encContactID
encContactID=encrypt(contactID)
dim fixStyle
if iStatus="cs_readwrite" then
fixStyle=" style=""color:#FFFFFF"" "
else
fixStyle=""
end if
rows=rows&"<tr>" & replace(row,"<td>","<td class='"&iStatus&"'>",1,-1,1)
if customer.bUseAvatars then
if not isLeeg(sAvatar) then
if instr(sAvatar,".")=0 then
sAvatar=sAvatar & ".jpg"
end if
select case right(sAvatar,3)
case "jpg"
rows=rows & "<td class="""& iStatus &""" align=""center""><img src=""" & C_DIRECTORY_QUICKERSITE & "/showThumb.aspx?FSR=1&amp;maxSize=60&amp;img="
rows=rows & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & contactID & "_" & sAvatar & """ /></td>"
case else
rows=rows & "<td class="""& iStatus &""" align=""center""><img width=""60"" height=""60"" style=""width:60px;height:60px"" src=""" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & "userfiles/" & contactID & "_" & sAvatar & """ /></td>"
end select
else
rows=rows & "<td class="""& iStatus &""" align=""center"">"
rows=rows & "<img width=""60"" height=""60"" src=""http://www.gravatar.com/avatar.php?gravatar_id=" & h.hash(contactEmail) & "&amp;default=" & customer.sQSUrl & "/fixedImages/avatar.jpg"" />"
rows=rows&"</td>"
end if
end if
rows=rows & "<td class="""& iStatus &""" align='center'>"&contactID&"</td>"
rows=rows & "<td class="""& iStatus &""">"&quotrep(sNickname)&"</td>"
rows=rows & "<td class="""& iStatus &"""><a " & fixStyle & " href='mailto:"&contactEmail&"'>"&contactEmail&"</a></td>"
if customer.bUseAvatars then
rows=rows & "<td class="""& iStatus &"""><a " & fixStyle & " href=""" & C_DIRECTORY_QUICKERSITE &"/default.asp?pageAction=avataredit&iContactID=" & encContactID &""" class=""QSPPAVATAR"">Avatar</a></td>"
end if
if mode=l("search") then
rows=rows & "<td class="""& iStatus &""" align='center'>"&convertEuroDate(lastLoginTS)&"</td><td class="""& iStatus &""" align=center><input type=checkbox name=iContactIDM value='"& encContactID &"' checked /></td>"
rows=rows & "<td class="""& iStatus &""">"& getIconPP(l("managepermissions"),"unlock","bs_contactPage.asp?iContactID="& encContactID,"","unlock" & encContactID,"class=bPopupFullWidthNoReload") & "</td>"
rows=rows & "<td class="""& iStatus &""">"& getIconPP(l("modify"),"edit","bs_contactEdit.asp?iContactID="& encContactID,"","edit" & encContactID,"class=bPopupFullWidthReload") & "</td>"
else
rows=rows & "<td class="""& iStatus &""" align='center'>"&convertEuroDate(lastLoginTS)&"</td>"
end if
rows=rows & "</tr>"
row=""
rowCount=rowCount+1
end function
private function getUp(value)
if convertGetal(value)=0 then
getUp=pageSize
else
getUp=value
end if
end function
private function initSF
if searchFields.count=0 then
set searchFields=customer.contactFields(false)
end if
end function
public function getContactsByStatus(iStatus)
dim rs,sql
sql="select iId, sNickName, sEmail from tblContact where iCustomerID=" & cId
if not isLeeg(iStatus) then
sql=sql & " and iStatus in ("&iStatus&")"
end if
sql=sql&" order by sNickname"
set rs = db.GetDynamicRS
rs.open sql
if not rs.EOF then
getContactsByStatus=rs.getrows
else
getContactsByStatus=null
end if
set rs = nothing
end function
public function showSelected(iStatus,selected)
selected=convertGetal(selected)
dim arrContacts,arrKey
arrContacts=getContactsByStatus(iStatus)
if not isNull(arrContacts) then
for arrKey=lbound(arrContacts,2) to ubound(arrContacts,2)
showSelected=showSelected&"<option value='"&arrContacts(0,arrKey)&"' "
if selected=arrContacts(0,arrKey) then
showSelected=showSelected& "selected"
end if
showSelected=showSelected&">"&quotrep(arrContacts(1,arrKey))&" - " & arrContacts(2,arrKey) &"</option>"
next
end if
end function
end class%>

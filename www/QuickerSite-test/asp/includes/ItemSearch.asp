
<%class cls_itemSearch
Public recordCount
public pageCount
Public searchFields
Public inputFields
Public dateFields
Public pageSize
Public absolutepage
Public origPageSize
Public sTitle
Public rowCount
public mode
public iCatalogID
public bOnline
public iSf
private p_catalog
Private sub class_initialize
On Error Resume Next
set inputFields=server.CreateObject ("scripting.dictionary")
set dateFields=server.CreateObject ("scripting.dictionary")
rowCount=0
if isNumeriek(Request.Form ("pageSize")) then
pageSize=convertGetal(Request.Form ("pageSize"))
else
pageSize=9999
end if
if isNumeriek(Request.Form("absolutepage")) then
absolutepage=convertGetal(Request.Form("absolutepage"))
else
absolutepage=1
end if
origPageSize=pageSize
mode=l("search")
if not isLeeg(Request.Form("btnaction")) then
mode=Request.Form("btnaction")
end if
bOnline=false
set p_catalog=nothing
On Error Goto 0
end sub
public function fixedSQL
initSF()
dim searchFieldsIDs, sfId
for each sfId in searchFields
searchFieldsIDs=searchFieldsIDs&sfId&","
next
if not isLeeg(searchFieldsIDs) then
searchFieldsIDs=left(searchFieldsIDs,len(searchFieldsIDs)-1)
end if
dim sql, where
sql=sql&" from (tbCatalogItemFields INNER JOIN tblCatalogItem on tblCatalogItem.iID=tbCatalogItemFields.iItemID) "
sql=sql&" INNER JOIN tblCatalogField on tblCatalogField.iId=tbCatalogItemFields.iFieldID"
where=" where " & sqlNull("tblCatalogField.iCatalogID",iCatalogID)
if not isLeeg(searchFieldsIDs) then
where=where& " and tbCatalogItemFields.iFieldID in ("&searchFieldsIDs&")"
end if
dim splitValues, lV
if not isLeeg(sTitle) then
where=where & " and ("
splitValues=split(sTitle," ")
for lV=lbound(splitValues) to ubound(splitValues)
where=where&" tblCatalogItem.sTitle like '%" & cleanup(splitValues(lV)) & "%' OR"
next
where=left(where,len(where)-2)& ")"
end if
if bOnline then
where=where& " and (tblCatalogItem.dOnlineFrom is null or tblCatalogItem.dOnlineFrom<="&getSQLDateFunction&") and (tblCatalogItem.dOnlineUntill is null or tblCatalogItem.dOnlineUntill>="&getSQLDateFunction&")  " 
end if
dim inputField
for each inputField in inputFields
if not isLeeg(inputFields(inputField)) then
where=where& " and tblCatalogItem.iID in (select iItemID from tbCatalogItemFields where tbCatalogItemFields.iFieldID="& inputField &" and tbCatalogItemFields.sValue like ('%"& cleanUp(inputFields(inputField)) &"%')) "
end if
next
dim fieldName, fromDate, untillDate, hasDates
for each inputField in dateFields
if not isLeeg(dateFields(inputField)) then
hasDates=false
fromDate=""
untillDate=""
fieldName=""
fieldName=replace(inputField,"from","")
fieldName=replace(fieldName,"untill","")
fieldName=replace(fieldName,"solo","")
if (left(inputField,4))="from" then
hasDates=true
fromDate	= convertCalcDate(convertEuroDate(roomDate(convertDateFromPicker(dateFields(inputField)),"down")))
end if
if (left(inputField,6))="untill" then
hasDates=true
untillDate	= convertCalcDate(convertEuroDate(roomDate(convertDateFromPicker(dateFields(inputField)),"up")))
end if
if (left(inputField,4))="solo" then
hasDates=true
untillDate	= convertCalcDate(convertEuroDate(convertDateFromPicker(dateFields(inputField))))
fromDate	= untillDate
end if
if hasDates and isLeeg(fromDate) then
fromDate	= convertCalcDate(convertEuroDate(roomDate("","down")))
end if
if hasDates and isLeeg(untillDate) then
untillDate	= convertCalcDate(convertEuroDate(roomDate("","up")))
end if
if hasDates then
where=where& " and tblCatalogItem.iID in (select iItemID from tbCatalogItemFields where tbCatalogItemFields.iFieldID="& fieldName &" and tbCatalogItemFields.sValue>='"& fromDate &"' and tbCatalogItemFields.sValue<='"& untillDate &"')"
end if
end if
next
dim orderBy
orderBy=" order by tblCatalogItem."& catalog.sOrderItemsBy &", tbCatalogItemFields.iItemID desc, tblCatalogField.iRang asc"
fixedSQL=sql & where & orderBy
end function
Public function getRequestValues
On Error Resume Next
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
if not isLeeg(Request.Form("solo"&encrypt(sfId))) then
dateFields.Add "solo"&sfId,Request.Form("solo"&encrypt(sfId))
end if
next
sTitle	= Request.Form ("sTitle")
iCatalogID	= convertLng(decrypt(request("iCatalogID")))
On Error Goto 0
end function
Public Function getRows()
initSF()
dim rs
set rs = db.GetDynamicRS
dim sql
sql="select [TOP] tbCatalogItemFields.iFieldID, tbCatalogItemFields.iItemID, tbCatalogItemFields.sValue, tblCatalogItem.sTitle, tblCatalogItem.dOnlineFrom, tblCatalogItem.dOnlineUntill, tblCatalogItem.sPicExt "
sql=sql	& fixedSQL & "[LIMIT]"
sql=getTOPSyntax(pageSize * getUp(searchFields.count),sql)
'Response.Write sql
'Response.flush
rs.open sql
if not rs.EOF then
getrows=rs.getrows
recordCount=rs.recordCount
else
getrows=null
end if
set rs = nothing
end function
public function resultItems
initSF()
dim searchFieldsCount
searchFieldsCount=getUp(searchFields.count)
set resultItems=server.CreateObject ("scripting.dictionary")
dim sql
sql="select tblCatalogItem.iId "
dim rs
set rs = db.GetDynamicRS
dim pS
pS=catalog.iPageSize*searchFieldsCount
if convertGetal(pS)=0 then pS=100
rs.pagesize = pS
'Response.Write sql & fixedSQL
'Response.Flush 
rs.open sql	& fixedSQL
if not rs.EOF then
rs.absolutepage = absolutepage
end if
dim i, item, id
for i = 1 to rs.pagesize 
if not rs.EOF then
id=rs(0)
if not resultItems.Exists (id) then
set item=new cls_catalogItem
item.pick(rs(0))
resultItems.Add	id,item
set item=nothing
end if
rs.moveNext
end if
next
pageCount	= convertGetal(rs.PageCount)
recordCount	= convertGetal(rs.RecordCount)
end function
public function browseTable
initSF()
if Request.Form("pageAction")="catalog" or absolutepage < pageCount then
browseTable="<div id='QS_catalogbrowsetable'>"
if recordcount=0 then
browseTable=browseTable& "0 " & l("resultsforsearch")
else
if absolutepage > 1 then
browseTable=browseTable & "<a href='#' onclick="& """" &"javascript:document.catalog.absolutepage.value='"& absolutepage-1 &"';document.catalog.submit();return false"& """" &">&lt;&lt;</a>&nbsp;"
end if
browseTable=browseTable & First_Record & " - " & Last_Record & " " & l("of") & " " & recordcount/getUp(searchfields.count)
if absolutepage < pageCount then
browseTable=browseTable & "&nbsp;<a href='#' onclick="& """" &"javascript:document.catalog.absolutepage.value='"& absolutepage+1 &"';document.catalog.submit();return false"& """" &">&gt;&gt;</a>"
end if
end if
browseTable=browseTable&"</div>"
end if
end function
Public Property Get First_Record
First_Record = (absolutepage * catalog.iPageSize) - (catalog.iPageSize - 1)
end Property
Public Property Get Last_Record
Last_Record = absolutepage * catalog.iPageSize
if clng(Last_Record) > clng(recordcount/getUp(searchfields.count)) then
Last_Record = clng(recordcount/getUp(searchfields.count))
end if
end Property
Public function resultTable
initSF()
dim p_getrows, j, searchField, headerRow, copyItemID, row, rows, start, copyItemTitle, copyStartDate, copyEndDate, copyPicExt
p_getrows=getRows
if not isNull(p_getrows) then
headerRow=headerRow&"<tr>"
headerRow=headerRow&"<th>"& l("title") &"</th>"
for each searchField in searchFields
headerRow=headerRow&"<th>"& searchFields(searchField).sName &"</th>"
next
headerRow=headerRow&"<th class='sorttable_nosort'>&nbsp;</th>"
if catalog.bAutoThumb then
headerRow=headerRow&"<th class='sorttable_nosort'>&nbsp;</th>"
end if
headerRow=headerRow&"<th class='sorttable_nosort'>&nbsp;</th></tr>"
'bepalen pageSize
if searchFields.count>0 then
if pageSize<>"all" then
pageSize=pageSize*searchFields.count
end if 
end if
if pageSize="all" then pageSize=recordCount
if convertGetal(pageSize)>convertGetal(recordCount) then pageSize=recordCount
for j=0 to pageSize-1
if copyItemID<>p_getrows(1,j) and not isLeeg(copyItemID) then
AddRow rows,row,copyItemID,copyItemTitle,isBetween(copyStartDate,date,copyEndDate),copyPicExt
end if
copyItemID=p_getrows(1,j)
copyItemTitle=p_getrows(3,j)
copyStartDate=p_getrows(4,j)
copyEndDate=p_getrows(5,j)
copyPicExt=p_getrows(6,j)
for each searchField in searchFields
if p_getrows(0,j)=searchField then
row=row&"<td valign=top>"
select case searchFields(searchField).sType
case sb_text,sb_textarea,sb_select
row=row&p_getrows(2,j)
case sb_richtext
row=row&removeHTML(p_getrows(2,j))
case sb_checkbox
row=row&convertCheckedYesNo(p_getrows(2,j))
case sb_date
row=row&convertDateToPicker(p_getrows(2,j))
case sb_url
if not isLeeg(p_getrows(2,j)) then
row=row&"<a target='"& generatePassword &"' href='"& p_getrows(2,j) & "'>" & p_getrows(2,j) & "</a>"
end if
case sb_email
if not isLeeg(p_getrows(2,j)) then
row=row&"<a href='mailto:"& p_getrows(2,j) & "'>" & p_getrows(2,j) & "</a>"
end if
end select
row=row&"&nbsp;</td>"
end if
next
if j=pageSize-1  then
AddRow rows,row,p_getrows(1,j),p_getrows(3,j),isBetween(p_getrows(4,j),date,p_getrows(5,j)),p_getrows(6,j)
end if
next
end if
if searchFields.count>0 then
recordCount=recordCount/searchFields.count
end if
select case mode
case l("search")
resultTable="<p align=center>"& l("total") &": <b>" & convertGetal(recordCount) & "</b> " & l("items") & "</p>"
if recordCount>0 then 
resultTable=resultTable&"<table align=""center"" style=""width:90%"" class=""sortable"" id=""contactSearch"" border=""1"" cellpadding=""5"" cellspacing=""0"">"& headerRow & rows & "</table>"
end if
case l("excel")
dim excelFile
set excelFile=new cls_excelFile
excelFile.export("<table>" & headerRow & rows & "</table>")
resultTable="<p align=""center""><b>"&excelFile.downloadLink&"</b></p>"
set excelFile=nothing
end select
end function
private function addRow(byref rows, byref row, itemID, itemTitle, status, sPicExt)
dim encitemID
encitemID=encrypt(itemID)
rows=rows & "<tr>" 
if mode=l("search") then
rows=rows & "<td valign=top><a name='"&encitemID&"'></a><strong><a href='bs_catalogItemEdit.asp?iItemID="& encitemID & "'>"&itemTitle&"</a></strong></td>"
else
rows=rows & "<td>"&itemTitle&"</td>"
end if
rows=rows & row
if mode=l("search") then
if catalog.bAutoThumb then
rows=rows&"<td align=""center"" valign=""middle"">"
if not isLeeg(sPicExt) then
rows=rows&getImageLink(catalog.iAutoClose,catalog.sResizePicTo,catalog.correFP & sPicExt,100,itemID,itemTitle)
else
rows=rows&"&nbsp;"
end if
rows=rows&"</td>"
end if
rows=rows&"<td align=center valign=middle style=""" & convertTFS(status) & """ class='" & convertTF(status) & "' width=15>" & convertV(status) & "</td>"
rows=rows&"<td>" & getIcon(l("copyitem"),"copyItem","bs_catalogItemSearch.asp?iCatalogID="&encrypt(iCatalogID)&"&amp;iItemID="& encitemID,"","copy"&itemID) & "</td>"
else
rows=rows & "<td>&nbsp;</td><td>&nbsp;</td>"
end if
rows=rows & "</tr>"
row=""
rowCount=rowCount+1
end function
Public function catalog
if p_catalog is nothing then
set p_catalog=new cls_catalog
p_catalog.pick(iCatalogID)
end if
set catalog=p_catalog
end function
private function getUp(value)
if convertGetal(value)=0 then
getUp=1
else
getUp=value
end if
end function
private function initSF
if searchFields.count=0 then
set searchFields=catalog.fields("all")
end if
end function
end class%>

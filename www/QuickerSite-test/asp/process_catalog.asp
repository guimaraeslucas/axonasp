
<%'cataloog ophalen
set catalog=selectedPage.catalog
if catalog.bOnline then
set catalogFields=catalog.fields("public")
set searchFields=catalog.fields("public,search")
set fileTypes=catalog.fileTypes
set itemSearch=new cls_itemSearch
set itemSearch.searchFields=searchFields
itemSearch.getRequestValues()
itemSearch.bOnline=true
itemSearch.iCatalogID=catalog.iId
set inputFields=itemSearch.inputFields
set dateFields=itemSearch.dateFields
'-------------------------------------------------------------------------------------------------------------------
'BEGIN FORM
'-------------------------------------------------------------------------------------------------------------------
catForm=catForm&"<form name='catalog' method='post' action='default.asp'>"
catForm=catForm&"<input type='hidden' name='pageAction' value='catalog' />"
catForm=catForm&"<input type='hidden' name='iID' value='"& encrypt(selectedPage.iId) &"' />"
catForm=catForm&"<input type='hidden' name='pageSize' value="""& sanitize(catalog.iPageSize) &""" />"
catForm=catForm&"<input type='hidden' name='absolutepage' value="""& sanitize(itemSearch.absolutepage) &""" />"
if catalog.bSearchable then
pageCatalog=pageCatalog&catForm
pageCatalog=pageCatalog&"<div id='QS_catalog'>"
pageCatalog=pageCatalog& catalog.sItemName & ":&nbsp;<input style=""width:70px"" onclick='javascript:this.select();' type='text' name='sTitle' size='5' value=" & """" & sanitize(itemSearch.sTitle) & """" &" />&nbsp;"
'hier komen de zoekvelden
for each sf in searchFields
select case searchFields(sf).sType
case sb_text,sb_textarea,sb_url,sb_email,sb_richtext
pageCatalog=pageCatalog& searchFields(sf).sName&"&nbsp;"
pageCatalog=pageCatalog&"<input onclick='javascript:this.select();' style=""width:70px"" type='text' size='5' name='" & encrypt(sf) &"' value=" & """" & sanitize(inputFields(sf)) & """" &" /> "
case sb_date
pageCatalog=pageCatalog& searchFields(sf).sName&"&nbsp;"
pageCatalog=pageCatalog&"<input type=""text"" id=""solo" & encrypt(sf) & """ name=""solo" & encrypt(sf) & """ value=""" & sanitize(dateFields("solo" & sf)) & """ />" & JQDatePicker("solo" & encrypt(sf)) & " "
case sb_checkbox
pageCatalog=pageCatalog& searchFields(sf).sName&"&nbsp;"
pageCatalog=pageCatalog&"<input type='checkbox' name='" & encrypt(sf) &"' value='checked' " & inputFields(sf) &" /> "
case sb_select
pageCatalog=pageCatalog&"<select name='" & encrypt(sf) & "'><option value=''>"& searchFields(sf).sName &"</option>" & searchFields(sf).showSelected(inputFields(sf)) &"</select> "
end select
next
pageCatalog=pageCatalog&"<input class=""art-button"" type='button' onclick=" & """" & "javascript:document.catalog.absolutepage.value='1';document.catalog.submit();" & """" &" name='btnaction' value=""" & sanitize(l("search")) & """ />&nbsp;<input type='button' class=""art-button"" onclick=" & """" & "javascript:location.assign('" & selectedPage.getSimpleLink & "')" & """" & " name='reset' value=""" & sanitize(l("reset")) & """ />"
pageCatalog=pageCatalog&"</div>"
pageCatalog=pageCatalog&"</form>" 
end if
'-------------------------------------------------------------------------------------------------------------------
'EINDE FORM
'-------------------------------------------------------------------------------------------------------------------
 
set resultItems=itemSearch.resultItems
pageCatalog=pageCatalog & itemSearch.browseTable 
pageCatalog=pageCatalog&"<div id='QS_catalog'>"
pageCatalog=pageCatalog&"<script type='text/javascript'>var bAutoStartSS=false</script>"
for each itemKey in resultItems
set ifields=resultItems(itemKey).fields
pageCatalog=pageCatalog&"<div class='QS_catalogitem'>"
pageCatalog=pageCatalog&"<div class='QS_fieldline'>"
pageCatalog=pageCatalog&"<div class='QS_itemtitle'><a name='"& encrypt(itemKey) &"'></a>" & resultItems(itemKey).sDateAndTitle &"</div>"
pageCatalog=pageCatalog&"</div>"
'foto
if not isLeeg(resultItems(itemKey).sPicExt) then
pageCatalog=pageCatalog&"<div class='QS_fieldline'>"
pageCatalog=pageCatalog&"<div class='QS_fieldlabel'>" & l("picture") &":</div>"
pageCatalog=pageCatalog&"<div class='QS_itempicture'>" & resultItems(itemKey).showPic("") & "&nbsp;</div>"
pageCatalog=pageCatalog&"</div>"
end if
'fields
for each catalogField in catalogFields
if not isLeeg(ifields(catalogField)) or catalogFields(catalogField).sType=sb_checkbox then
pageCatalog=pageCatalog&"<div class='QS_fieldline'>"
pageCatalog=pageCatalog&"<div class='QS_fieldlabel'>" & catalogFields(catalogField).sName &"</div>"
pageCatalog=pageCatalog&"<div class='QS_fieldvalue'>"
select case catalogFields(catalogField).sType 
case sb_text,sb_textarea,sb_select,sb_url,sb_email
pageCatalog=pageCatalog & LinkURLs(ifields(catalogField))
case sb_richtext
pageCatalog=pageCatalog & ifields(catalogField)
case sb_date
pageCatalog=pageCatalog & convertDateToPicker(ifields(catalogField))
case sb_checkbox
pageCatalog=pageCatalog & convertCheckedYesNo(ifields(catalogField))
end select
pageCatalog=pageCatalog&"</div>"
pageCatalog=pageCatalog&"</div>"
end if
next
'files
if fileTypes.count>0 then
set files=resultItems(itemKey).files
if files.count>0 then
pageCatalog=pageCatalog&"<div class='QS_fieldline'>"
pageCatalog=pageCatalog&"<div class='QS_fieldlabel'>"& l("files") &":</div>"
pageCatalog=pageCatalog&"<div class='QS_fieldvalue'>"
for each ft in fileTypes
for each f in files
if convertGetal(files(f).iFileTypeID)=convertGetal(ft) then
pageCatalog=pageCatalog & files(f).url & " - "
end if
next
next
pageCatalog=left(pageCatalog,len(pageCatalog)-3)
pageCatalog=pageCatalog&"</div>"
pageCatalog=pageCatalog&"</div>"
end if
set files=nothing
end if
'inschrijvingsformulier
if isNumeriek(catalog.iFormID) then
if resultItems(itemKey).bForm then
pageCatalog=pageCatalog&"<div class='QS_fieldline'>"
pageCatalog=pageCatalog&"<div class='QS_fieldlabel'>&nbsp;</div>"
pageCatalog=pageCatalog&"<div class='QS_formlink'><a href='default.asp?pageAction=itemform&amp;iPageID="& encrypt(selectedPage.iId) &"&amp;iItemID="& encrypt(itemKey) &"'>-> " & catalog.sFormTitle &"</a></div>"
pageCatalog=pageCatalog&"</div>"
end if
end if
pageCatalog=pageCatalog&"</div>"
set ifields	= nothing
pageCatalog=treatConstants(pageCatalog,true)
next
pageCatalog=pageCatalog&"</div>"
if itemSearch.recordCOunt>0 then pageCatalog=pageCatalog & itemSearch.browseTable 
if not catalog.bSearchable then	pageCatalog=pageCatalog&catForm&"</form>" 
set fileTypes	= nothing
set catalog	= nothing
set catalogFields	= nothing
set itemSearch	= nothing
set resultItems	= nothing
end if 'einde cataloog%>


<%pageListPage="<div id='QS_list'>" 'start general DIV
set listitems=selectedPage.fastlistitems
LPCounter=0
dim pageHasForm,pageCondition
if isNumeriek(selectedPage.iFormID) then
pageHasForm=true
else
pageHasForm=false
end if
dim pageListPageItem
if not convertBool(selectedPage.bAccordeon) then
for each listkey in listitems
pageListPageItem=""
set selectedListItem=listitems(listkey)
selectedListItem.bHideDate=selectedPage.bHideDate
isLPEU=not isLeeg(selectedListItem.sLPExternalURL) and selectedListItem.sLPExternalURL<>"http://"
encryptedListkey=encrypt(listkey)
if pageHasForm then
pageCondition=false
else
pageCondition=Request.Form ("item")=encryptedListkey
end if
if Request.QueryString ("item")=encryptedListkey then
selectedPage.sSEOtitle=quotrep(selectedListItem.sTitle & " | "  & customer.siteTitle)
end if
pageListPageItem=pageListPageItem&"<div class='QS_listitem'>"
pageListPageItem=pageListPageItem&"<div class='QS_listminusplus'>"
if isLPEU then
pageListPageItem=pageListPageItem&"<a href=" & """" & selectedListItem.sLPExternalURL & """" 
if convertBool(selectedListItem.bLPExternalOINW) then
pageListPageItem=pageListPageItem & " target='_blank'"
end if
pageListPageItem=pageListPageItem & "><img alt='' align='top' border='0' style='margin:0px;border-style:none' src='"& C_DIRECTORY_QUICKERSITE &"/fixedImages/link"& tdir &".gif' /></a>"
else
if printReplies or pageCondition or (LPCounter<convertGetal(selectedPage.iLPOpenByDefault) and Request.QueryString ("item")<>"close" and convertGetal(decrypt(Request.Querystring ("item")))=0) or Request.QueryString ("item")=encryptedListkey or (convertBool(selectedPage.bOpenOnload) and isLeeg(Request.QueryString ("item")))	 then
pageListPageItem=pageListPageItem&"<a href='default.asp?iID="& encrypt(selectedPage.iId)&"&amp;item=close'><img alt='' style='margin:0px;border-style:none' align='top' border='0' src='"& C_DIRECTORY_QUICKERSITE &"/fixedImages/minus2.gif' /></a>"
else
pageListPageItem=pageListPageItem&"<a href='default.asp?iID="& encrypt(selectedPage.iId) &"&amp;item="& encryptedListkey &"#"& encryptedListkey &"'>{#QSLP1#}</a>"
end if
end if
pageListPageItem=pageListPageItem&"</div>"
if printReplies or pageCondition or (LPCounter<convertGetal(selectedPage.iLPOpenByDefault) and Request.QueryString ("item")<>"close" and convertGetal(decrypt(Request.Querystring ("item")))=0) or Request.QueryString ("item")=encryptedListkey or (convertBool(selectedPage.bOpenOnload) and isLeeg(Request.QueryString ("item"))) then
pageListPageItem=pageListPageItem&"<div class='QS_listitemtitle'><a name='"& encryptedListkey &"'></a>"
if isLPEU then
pageListPageItem=pageListPageItem&"<a href="& """" & selectedListItem.sLPExternalURL & """" 
if convertBool(selectedListItem.bLPExternalOINW) then
pageListPageItem=pageListPageItem&" target='_blank'"
end if
pageListPageItem=pageListPageItem & ">"
else
pageListPageItem=pageListPageItem&"<a href='default.asp?iID="& encrypt(selectedPage.iId)&"&amp;item=close'>"
end if
pageListPageItem=pageListPageItem& selectedListItem.sDateAndTitle&"</a></div>"
if not isLPEU then 
pageListPageItem=pageListPageItem&"<div class='QS_listitemvalue'>"

	pageListPageItem=pageListPageItem & selectedListItem.listitemPicIMGTag

	pageListPageItem=pageListPageItem& selectedListItem.page.insertMedia(selectedListItem.sValue) & "</div>"
end if
if not isLeeg(selectedListItem.iFeedId) then pageListPageItem=pageListPageItem&"<div class='QS_listitemvalue'>"&selectedListItem.Feed.build()&"</div>"
else
pageListPageItem=pageListPageItem&"<div class='QS_listitemtitle'><a name='"& encryptedListkey &"'></a>"
if isLPEU then
pageListPageItem=pageListPageItem&"<a href="& """" & selectedListItem.sLPExternalURL & """" 
if convertBool(selectedListItem.bLPExternalOINW) then
pageListPageItem=pageListPageItem&" target='_blank'"
end if
pageListPageItem=pageListPageItem & ">"
else
pageListPageItem=pageListPageItem&"<a href='default.asp?iID="& encrypt(selectedPage.iId) &"&amp;item="& encryptedListkey &"#"& encryptedListkey &"'>"
end if
pageListPageItem=pageListPageItem&selectedListItem.sDateAndTitle&"</a></div>"
end if
pageListPageItem=pageListPageItem&"</div>"
pageListPage=pageListPage&pageListPageItem
LPCounter=LPCounter+1
set selectedListItem=nothing
next
end if
if convertBool(selectedPage.bAccordeon) then
pageListPage=selectedPage.getAccordeonList
else
pageListPage=treatConstants(pageListPage,true)	& "</div>" 'end general DIV
pageListPage=replace(pageListPage,"{#QSLP1#}","<img alt='' style='margin:0px;border-style:none' align='top' border='0' src='"& C_DIRECTORY_QUICKERSITE &"/fixedImages/plus2.gif' />",1,-1,1)
end if
set listitems=nothing%>

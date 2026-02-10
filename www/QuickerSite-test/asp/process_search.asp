
<%search.value=left(Request.Form ("svalue"),100)
if search.value="" then search.value=left(Request.QueryString  ("svalue"),100)
search.bIntranet=Request.Form ("bIntranet")
search.includeListItems=true
search.bIncludeHideFromSearch=false
set results=search.results
if results.count>0 then
searchSnippet="<div id='QS_searchResults'>"
for each resultPage in results
if not isLeeg(results(resultPage).sApplication) then
searchSnippet=searchSnippet& "<div class='QS_searchResultsTitle'><a href='default.asp?iId=" & EnCrypt(resultPage) & "'>" & results(resultPage).sTitle & "</a></div>"
searchSnippet=searchSnippet& "<div class='QS_searchResultsText'>...</div>"
elseif convertGetal(results(resultPage).iListPageID)<>0  then
if isleeg(results(resultPage).sLPExternalURL) then
searchSnippet=searchSnippet& "<div class='QS_searchResultsTitle'><a href='default.asp?iId=" & EnCrypt(results(resultPage).listPage.iID) & "&amp;item="& EnCrypt(resultPage)&"#" & EnCrypt(resultPage) & "'>" & results(resultPage).sTitle & "</a></div>"
searchSnippet=searchSnippet& "<div class='QS_searchResultsText'>" & replace(left(convertStr(treatConstants(RemoveHTMLComments(results(resultPage).sValueTextOnly),false)),150),search.value,"<b>"& search.value &"</b>",1,-1,1) & "...</div>"
else
searchSnippet=searchSnippet& "<div class='QS_searchResultsTitle'><a "
if results(resultPage).bLPExternalOINW then
searchSnippet=searchSnippet& " target='" &	generatePassword & "' "
end if
searchSnippet=searchSnippet& " href=""" & results(resultPage).sLPExternalURL & """>" & results(resultPage).sTitle & "</a></div>"
searchSnippet=searchSnippet& "<div class='QS_searchResultsText'>" & replace(left(convertStr(treatConstants(RemoveHTMLComments(results(resultPage).sValueTextOnly),false)),150),search.value,"<b>"& search.value &"</b>",1,-1,1) & "...</div>"
end if
elseif isLeeg(results(resultPage).sExternalURL) then
if convertGetal(results(resultPage).iPostID)=0 then
searchSnippet=searchSnippet& "<div class='QS_searchResultsTitle'><a href='default.asp?iId=" & EnCrypt(resultPage) & "'>" & results(resultPage).sTitle & "</a></div>"
searchSnippet=searchSnippet& "<div class='QS_searchResultsText'>"& (replace(left(convertStr(treatConstants(RemoveHTMLComments(results(resultPage).sValueTextOnly),false)),150),search.value,"<b>"& search.value &"</b>",1,-1,1)) & "...</div>"
else
searchSnippet=searchSnippet& "<div class='QS_searchResultsTitle'><a href='default.asp?pageAction=showPost&amp;iPostID=" & resultPage & "'>" & quotrep(results(resultPage).sTitle) & "</a></div>"
searchSnippet=searchSnippet& "<div class='QS_searchResultsText'>"& replace(left(convertStr(treatConstants(RemoveHTMLComments(results(resultPage).sValueTextOnly),false)),150),search.value,"<b>"& search.value &"</b>",1,-1,1) & "...</div>"
end if
else
searchSnippet=searchSnippet& "<div class='QS_searchResultsTitle'><a target='"& EnCrypt(resultPage) &"' href='" & results(resultPage).sExternalURLPrefix & results(resultPage).sExternalURL &"'>"& results(resultPage).sTitle & "</a></div>"
searchSnippet=searchSnippet& "<div class='QS_searchResultsText'>&nbsp;</div>"
end if
next 
pageBody=treatConstants(searchSnippet,false)&"</div>"
pageBody=replace(pageBody,"[","",1,-1,1)
pageBody=replace(pageBody,"]","",1,-1,1)
end if
pageTitle=l("thereAre") & " "& results.count & " " & l("resultsForSearch") & ": '"& sanitize(treatConstants(search.value,false)) &"'"%>

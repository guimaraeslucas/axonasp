<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%if logon.currentPW=customer.secondAdmin.sPassword then Response.redirect ("bs_default.asp")%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader("")%><p align=center><%=l("sitesearchexpl")%></p><form method="post" name="bs" action="bs_search.asp"><input type="hidden" name="postback" value="<%=true%>" /><%=QS_secCodeHidden%><table align=center> 
<tr><td valign="middle"><input required type="text" value="<%=sanitize(Request.Form ("svalue"))%>" name="svalue" /></td><td valign="middle"><input class="art-button" type="submit" name="btnaction" value="<%=sanitize(l("search"))%>" /></td></tr></table></form><script type="text/javascript">document.bs.svalue.focus();</script><%dim postback
postback=convertBool(Request.Form ("postback"))
if postback and not isLeeg(Request.Form ("svalue")) then
checkCSRF()
dim fsearch
set fsearch=new cls_fullSearch
fsearch.pattern="("&treatParentheses(Request.Form ("svalue"))&")"
dim result
result=fsearch.search()
if not isLeeg(result) then
Response.Write result
else
Response.Write "<p align=center>" & l("thereAre") & " 0 " & l("resultsForSearch") & ": '<b>"& sanitize(treatConstants(Request.Form ("svalue"),false)) &"</b>'</p>"
end if
set fsearch=nothing
end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->

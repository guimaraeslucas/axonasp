
<%function getVisitorDetails
Dim variable
For Each variable In Request.ServerVariables
select case ucase(variable)
case "LOCAL_ADDR","REMOTE_ADDR","REMOTE_HOST","HTTP_ACCEPT_LANGUAGE","HTTP_REFERER","HTTP_USER_AGENT","HTTP_X_FORWARDED_FOR","HTTP_HOST"
if not isLeeg(Request.ServerVariables(variable)) then
getVisitorDetails=getVisitorDetails & "<li>"& variable & ": " & sanitize(Request.ServerVariables(variable)) & "</li>"
end if
end select
Next
if not isLeeg(session("extenalReferrer")) then
getVisitorDetails=getVisitorDetails & "<li>External referring site: " & quotrep(session("extenalReferrer")) & "</li>"
end if
if not isLeeg(getVisitorDetails) then
getVisitorDetails="<ul>" & getVisitorDetails & "</ul>"
end if
end function%>

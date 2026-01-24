<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write "Testing aspLite initialization<br>"
Response.Write "aspl TypeName: " & TypeName(aspl) & "<br>"
Response.Write "Request.QueryString: " & Request.QueryString & "<br>"

' Try direct Request access
if Request.QueryString("asplEvent") <> "" then
    Response.Write "Request.QueryString(""asplEvent""): " & Request.QueryString("asplEvent") & "<br>"
else
    Response.Write "Request.QueryString(""asplEvent"") is empty<br>"
end if

%>

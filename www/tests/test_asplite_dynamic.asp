<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
dim asplEvent : asplEvent=aspl.getRequest("asplEvent")

Response.Write "asplEvent = " & asplEvent & "<br>"

if not aspl.isEmpty(asplEvent) then
	Response.Write "Found asplEvent, executing: ebook/" & asplEvent & ".inc<br>"
	'dynamically execute the scriptname in asplEvent
	aspl("ebook/" & asplEvent & ".inc")
else
	Response.Write "asplEvent is empty<br>"
end if
%>

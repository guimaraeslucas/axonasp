<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Dim asplEvent
asplEvent = aspl.getRequest("asplEvent")

If aspl.isEmpty(asplEvent) Then
	asplEvent = "fields"
End If

aspl("aspLite-master/ebook/" & asplEvent & ".inc")
%>

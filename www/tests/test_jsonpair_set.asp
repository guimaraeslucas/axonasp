<%
Option Explicit
dim obj
set obj = new JSONpair
obj.name = "test"
Response.Write "Set name to: " & obj.name & "<br>"
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

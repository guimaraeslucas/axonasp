<%
Option Explicit
dim jsonObj
set jsonObj = new JSONobject
jsonObj.debug = false

dim result
result = jsonObj.pairs
Response.Write "Got pairs: " & TypeName(result) & "<br>"
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

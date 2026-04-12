<%
Option Explicit
dim jsonObj
set jsonObj = new JSONobject
jsonObj.debug = false

dim result
result = jsonObj.version
Response.Write "Version: " & result & "<br>"
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

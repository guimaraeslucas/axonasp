<%
Option Explicit
dim jsonObj, jsonString, outputObj
set jsonObj = new JSONobject
jsonObj.debug = false

jsonString = "{ ""test"" : ""value"" }"

dim start
start = timer()
set outputObj = jsonObj.parse(jsonString)
Response.Write "Parse time: " & (timer() - start) & " s<br>"
Response.Write "Parse successful<br>"
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

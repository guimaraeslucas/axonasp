<%
Option Explicit
Response.LCID = 1046
dim jsonObj, jsonString, outputObj
set jsonObj = new JSONobject
jsonObj.debug = false

jsonString = "[{ ""test"" : ""value"" }]"

dim start
start = timer()
set outputObj = jsonObj.parse(jsonString)

On Error Resume Next
jsonObj.add "nome", "test"

if Err.Number <> 0 then
    Response.Write "ERROR before: " & Err.Description & "<br>"
else
    Response.Write "Add succeeded<br>"
end if
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

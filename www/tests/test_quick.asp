<%
Option Explicit
on error resume next

dim jsonObj
set jsonObj = new JSONobject
dim result
result = jsonObj.parse("[{""test"": ""value""}]")

if err.number <> 0 then
    Response.Write "PARSE_FAILED: " & err.description
else
    err.Clear
    jsonObj.add "key", "value"
    if err.number <> 0 then
        Response.Write "ADD_FAILED: " & err.description
    else
        Response.Write "SUCCESS"
    end if
end if
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

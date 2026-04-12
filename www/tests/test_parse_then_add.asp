<%
Response.LCID = 1046
on error resume next

dim jsonObj
set jsonObj = new JSONobject

Response.Write "Step 1: Created JSONobject<br>"

' Try to parse
dim jsonString
jsonString = "[{""test"": ""value""}]"
set jsonObj = jsonObj.parse(jsonString)

Response.Write "Step 2: Parsed JSON<br>"

' Now try to add
jsonObj.add "newkey", "newvalue"

if err.number <> 0 then
    Response.Write "ERROR at line " & Erl & ": " & err.description & " (Error " & err.number & ")<br>"
else
    Response.Write "Step 3: Successfully added to JSON<br>"
end if
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

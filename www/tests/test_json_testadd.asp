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
Response.Write "Parse time: " & (timer() - start) & " s<br>"

Response.Write "Starting testAdd block<br>"

if true then
    dim arr, multArr, nestedObject
    arr = Array(1, "teste", 234.56, "mais teste", "234", now)
    
    redim multArr(2, 3)
    
    Response.Write "After redim<br>"
end if

Response.Write "testAdd block completed<br>"
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

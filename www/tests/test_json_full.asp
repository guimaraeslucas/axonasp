<%
Option Explicit
Response.LCID = 1046
dim jsonObj, jsonString, outputObj
set jsonObj = new JSONobject
jsonObj.debug = false

jsonString = "[{" & """test""" & ": " & """value""" & "}]"

dim start
start = timer()
set outputObj = jsonObj.parse(jsonString)
Response.Write "Load time: " & (timer() - start) & " s<br>"

if true then
    dim arr, multArr, nestedObject
    arr = Array(1, "test", 234.56)
    
    redim multArr(2, 3)
    multArr(0, 0) = "0,0"
    
    Response.Write "Before jsonObj.add<br>"
    jsonObj.add "nome", "Test"
    Response.Write "After jsonObj.add<br>"
end if
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

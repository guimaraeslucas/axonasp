<!--#include file="json-teste/jsonObject.class.asp" -->
<%
Option Explicit
Response.LCID = 1046
Dim jsonObj, jsonString, outputObj
Set jsonObj = New JSONobject
jsonObj.debug = False

jsonString = "[{ ""test"" : ""value"" }]"

Dim start
start = Timer()
Set outputObj = jsonObj.parse(jsonString)
Response.Write "Parse time: " & (Timer() - start) & " s<br>"

Response.Write "Starting testAdd block<br>"

If True Then
    Dim arr, multArr, nestedObject
    arr = Array(1, "teste", 234.56, "mais teste", "234", Now)

    ReDim multArr(2, 3)

    Response.Write "Adding to jsonObj<br>"
    jsonObj.Add "nome", "Jozé"
    jsonObj.Add "ficticio", True
    jsonObj.Add "idade", 25
    Response.Write "Added successfully<br>"
End If

Response.Write "testAdd block completed<br>"
%>

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

' Manually inline add logic with error handling
On Error Resume Next
Dim prop, obj
prop = "nome"
obj = "Jozé"

Dim p
jsonObj.getProperty prop, p

Response.Write "Type of p: " & TypeName(p) & "<br>"

If GetTypeName(p) <> "JSONpair" Then
    Response.Write "Creating new JSONpair<br>"
    Dim Item
    Set Item = New JSONpair

    Response.Write "Setting item.name<br>"
    Item.name = prop

    If Err.Number <> 0 Then
        Response.Write "ERROR: " & Err.Description & "<br>"
    Else
        Response.Write "item.name = " & Item.name & "<br>"
    End If
Else
    Response.Write "Property already exists<br>"
End If
%>
<!--#include file="json-teste/jsonObject.class.asp" -->

<%
Option Explicit
On Error Resume Next
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj, jsonString, s
Set jsonObj = New JSONobject
jsonString = "[{ ""a"": ""b"", ""arr"": [1,2] }]"
Set outputObj = jsonObj.parse(jsonString)

Err.Clear
s = outputObj.Serialize()
If Err.Number <> 0 Then Response.Write "E1:" & Err.Description & "|" Else Response.Write "OK1|"

Err.Clear
s = outputObj(0).Serialize()
If Err.Number <> 0 Then Response.Write "E2:" & Err.Description & "|" Else Response.Write "OK2|"

Err.Clear
s = outputObj(0).serializeArray(Array(1,2))
If Err.Number <> 0 Then Response.Write "E3:" & Err.Description & "|" Else Response.Write "OK3|"

Err.Clear
s = outputObj(0).EscapeCharacters("x")
If Err.Number <> 0 Then Response.Write "E4:" & Err.Description & "|" Else Response.Write "OK4|"
%>

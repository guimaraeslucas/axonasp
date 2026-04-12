<!--#include file="jsonObject.class.asp" -->
<%
Option Explicit
Dim jsonObj, outputObj
Set jsonObj = New JSONobject
Set outputObj = jsonObj.parse("[{""a"":1}]")
Response.Write "TypeName(jsonObj)=" & TypeName(jsonObj) & "<br>"
Response.Write "TypeName(outputObj)=" & TypeName(outputObj) & "<br>"
On Error Resume Next
jsonObj.Add "data", Now
Response.Write "Err=" & Err.Number & "|" & Err.Description & "<br>"
%>

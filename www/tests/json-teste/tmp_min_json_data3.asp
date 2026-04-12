<!--#include file="jsonObject.class.asp" -->
<%
Option Explicit
Dim jsonObj, outputObj
Set jsonObj = New JSONobject
Set outputObj = jsonObj.parse("[{""a"":1}]")
Response.Write "defaultPropertyName=" & jsonObj.defaultPropertyName & "<br>"
Response.Write "value(data) type=" & TypeName(jsonObj.value("data")) & "<br>"
On Error Resume Next
jsonObj.Add "data", 123
Response.Write "Err=" & Err.Number & "|" & Err.Description & "<br>"
%>

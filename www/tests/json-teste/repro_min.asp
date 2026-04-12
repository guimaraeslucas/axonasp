<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj, jsonString
Set jsonObj = New JSONobject
jsonString = "[{ ""a"": ""b"" }]"
Set outputObj = jsonObj.parse(jsonString)
Response.Write outputObj.Write
%>

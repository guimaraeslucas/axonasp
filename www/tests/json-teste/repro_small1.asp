<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj
Set jsonObj = New JSONobject
Set outputObj = jsonObj.parse("[{ ""a"": ""b"" }]")
Response.Write outputObj.Serialize()
%>

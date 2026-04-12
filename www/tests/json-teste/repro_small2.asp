<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj, s
Set jsonObj = New JSONobject
s = "[{ ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]] }]"
Set outputObj = jsonObj.parse(s)
Response.Write outputObj.Serialize()
%>

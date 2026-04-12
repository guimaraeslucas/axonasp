<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj, rootObj
Set jsonObj = New JSONobject
Set outputObj = jsonObj.parse("[{ ""emptyObject"": {}, ""objects"": { ""prop1"": ""x"" } }]")
Set rootObj = outputObj(0)
Response.Write "ROOT=" & rootObj.Serialize() & "|OBJ=" & rootObj("objects").Serialize() & "|EO=" & rootObj("emptyObject").Serialize()
%>

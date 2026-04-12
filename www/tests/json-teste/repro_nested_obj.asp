<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj, rootObj
Set jsonObj = New JSONobject
Set outputObj = jsonObj.parse("[{ ""emptyObject"": {}, ""objects"": { ""prop1"": ""x"" } }]")
Set rootObj = outputObj(0)
Response.Write "EO_ISOBJ=" & IsObject(rootObj("emptyObject")) & "|"
Response.Write "EO_TYPE=" & TypeName(rootObj("emptyObject")) & "|"
Response.Write "OBJ_ISOBJ=" & IsObject(rootObj("objects")) & "|"
Response.Write "OBJ_TYPE=" & TypeName(rootObj("objects")) & "|"
If IsObject(rootObj("objects")) Then Response.Write rootObj("objects").Serialize()
%>

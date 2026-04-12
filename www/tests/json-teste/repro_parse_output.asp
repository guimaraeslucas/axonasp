<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj, jsonString
Set jsonObj = New JSONobject
jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]], ""emptyArray"": [], ""emptyObject"": {}, ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] }, ""multiline"": ""Texto com\r\nMais de\r\numa linha e escapado com \\."" }"
jsonString = "[" & jsonString & "]"
Set outputObj = jsonObj.parse(jsonString)
Response.Write outputObj.Write
%>

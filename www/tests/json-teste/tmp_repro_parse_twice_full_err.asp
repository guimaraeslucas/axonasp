<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, s
Set jsonObj = New JSONobject
s = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]], ""emptyArray"": [], ""emptyObject"": {}, ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] }, ""multiline"": ""Texto com\r\nMais de\r\numa linha e escapado com \\."" }"
s = "[" & s & "]"
On Error Resume Next
jsonObj.parse s
Response.Write "E1=" & Err.Number & " src=" & Err.Source & " desc=" & Err.Description & "<br>"
Err.Clear
jsonObj.parse s
Response.Write "E2=" & Err.Number & " src=" & Err.Source & " desc=" & Err.Description & "<br>"
Err.Clear
On Error Goto 0
Response.Write "done"
%>

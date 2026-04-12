<%
Option Explicit
Response.LCID = 1046
Response.buffer = True
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, jsonString, jsonArr, outputObj
Set jsonObj = New JSONobject
Set jsonArr = New jsonArray
jsonObj.debug = False
jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]], ""emptyArray"": [], ""emptyObject"": {}, ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] }, ""multiline"": ""Texto com\r\nMais de\r\numa linha e escapado com \\."" }"
jsonString = "[" & jsonString & "]"
Set outputObj = jsonObj.parse(jsonString)
Response.Flush
Set outputObj = jsonObj.parse(jsonString)
Set jsonArr = outputObj
jsonObj.Add "nome", "Jozé"
jsonObj.Add "idade", 25
jsonObj.defaultPropertyName = "CustomName"
Response.Write "C:" & jsonObj("aNonExistantPropertyName") & "(" & TypeName(jsonObj("aNonExistantPropertyName")) & ")<br>"
If IsObject(jsonObj(jsonObj.defaultPropertyName)) Then
  Response.Write "D-obj<br>"
Else
  Response.Write "D-scalar:" & jsonObj(jsonObj.defaultPropertyName) & "<br>"
End If
Response.Write "END"
%>

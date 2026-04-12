<%
Option Explicit
Response.LCID = 1046
Response.buffer = True
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, jsonString, jsonArr, outputObj
Dim arr, multArr, nestedObject
Set jsonObj = New JSONobject
Set jsonArr = New jsonArray
jsonObj.debug = False
jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]], ""emptyArray"": [], ""emptyObject"": {}, ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] }, ""multiline"": ""Texto com\r\nMais de\r\numa linha e escapado com \\."" }"
jsonString = "[" & jsonString & "]"
Set outputObj = jsonObj.parse(jsonString)
Set jsonArr = outputObj
arr = Array(1, "teste", 234.56, "mais teste", "234", Now)
ReDim multArr(2, 3)
multArr(0, 0) = "0,0": multArr(0, 1) = "0,1": multArr(0, 2) = "0,2": multArr(0, 3) = "0,3"
multArr(1, 0) = "1,0": multArr(1, 1) = "1,1": multArr(1, 2) = "1,2": multArr(1, 3) = "1,3"
multArr(2, 0) = "2,0": multArr(2, 1) = "2,1": multArr(2, 2) = "2,2": multArr(2, 3) = "2,3"
jsonObj.Add "nome", "Jozé"
jsonObj.Add "ficticio", True
jsonObj.Add "idade", 25
jsonObj.Add "saldo", -52
jsonObj.Add "bio", "Nascido em São Paulo\Brasil" & vbcrlf & "Sem filhos" & vbcrlf & vbtab & "Jogador de WoW"
jsonObj.Add "data", Now
jsonObj.Add "lista", arr
jsonObj.Add "lista2", multArr
Set nestedObject = New JSONobject
nestedObject.Add "sub1", "value of sub1"
nestedObject.Add "sub2", "value of ""sub2"""
jsonObj.Add "nested", nestedObject
jsonObj.Remove "numbers"
jsonObj.Remove "aNonExistantPropertyName"
jsonObj.defaultPropertyName = "CustomName"
jsonArr.defaultPropertyName = "CustomArrName"
Response.Write "A:" & jsonObj.value("nome") & "<br>"
Response.Write "B:" & jsonObj("idade") & "<br>"
Response.Write "C:" & jsonObj("aNonExistantPropertyName") & "(" & TypeName(jsonObj("aNonExistantPropertyName")) & ")<br>"
Response.Write "D-before<br>"
If IsObject(jsonObj(jsonObj.defaultPropertyName)) Then
  Response.Write "D-obj<br>"
Else
  Response.Write "D-scalar:" & jsonObj(jsonObj.defaultPropertyName) & "<br>"
End If
Response.Write "END"
%>

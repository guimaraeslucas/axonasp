<%
Option Explicit
Response.LCID=1046
Response.Buffer=True
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, jsonString, jsonArr, outputObj
Dim arr, multArr, nestedObject, newJson
Dim start
Set jsonObj = New JSONobject
Set jsonArr = New jsonArray
jsonObj.debug = False
jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]], ""emptyArray"": [], ""emptyObject"": {}, ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] }, ""multiline"": ""Texto com\r\nMais de\r\numa linha e escapado com \\."" }"
jsonString = "[" & jsonString & "]"
Set outputObj = jsonObj.parse(jsonString)
Set outputObj = jsonObj.parse(jsonString)
Set jsonArr = outputObj
Response.Write "CP1<br>"
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
Response.Write "CP2<br>"
On Error Resume Next
Response.Write "nome:" & jsonObj.value("nome") & "<br>"
Response.Write "idade:" & jsonObj("idade") & "<br>"
Response.Write "non:" & jsonObj("aNonExistantPropertyName") & "(" & TypeName(jsonObj("aNonExistantPropertyName")) & ")<br>"
Response.Write "isobj? " & IsObject(jsonObj(jsonObj.defaultPropertyName)) & "<br>"
If Err.Number <> 0 Then Response.Write "ERR-A:" & Err.Number & " src=" & Err.Source & " desc=" & Err.Description & "<br>": Err.Clear
Response.Write "CP3<br>"
Response.Write "before change" & "<br>"
jsonObj.change "nome", "Mario"
If Err.Number <> 0 Then Response.Write "ERR-B:" & Err.Number & " src=" & Err.Source & " desc=" & Err.Description & "<br>": Err.Clear
Response.Write "after change" & "<br>"
Set newJson = New JSONobject
newJson.Add "newJson", "property"
newJson.Add "version", newJson.version
jsonArr.Push newJson
jsonArr.Push 1
jsonArr.Push "strings too"
If Err.Number <> 0 Then Response.Write "ERR-C:" & Err.Number & " src=" & Err.Source & " desc=" & Err.Description & "<br>": Err.Clear
Response.Write "CP4<br>"
Set newJson = Nothing
Set outputObj = Nothing
Set jsonObj = Nothing
Set jsonArr = Nothing
If Err.Number <> 0 Then Response.Write "ERR-D:" & Err.Number & " src=" & Err.Source & " desc=" & Err.Description & "<br>": Err.Clear
Response.Write "CP5<br>"
On Error Goto 0
%>

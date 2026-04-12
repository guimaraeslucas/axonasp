<!--#include file="jsonObject.class.asp" -->
<%
Option Explicit
Dim jsonObj, jsonArr, outputObj, jsonString
Dim arr, multArr, nestedObject
Set jsonObj = New JSONobject
Set jsonArr = New jsonArray

jsonString = "[{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4], ""emptyArray"": [], ""emptyObject"": {} }]"
Set outputObj = jsonObj.parse(jsonString)
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

Response.Write "A<br>"
Response.Write "nome=" & jsonObj.value("nome") & "<br>"
Response.Write "idade=" & jsonObj("idade") & "<br>"
Response.Write "missing=" & TypeName(jsonObj("aNonExistantPropertyName")) & "<br>"

On Error Resume Next
Response.Write "isDefaultObject=" & IsObject(jsonObj(jsonObj.defaultPropertyName)) & "<br>"
Response.Write "Err1=" & Err.Number & "|" & Err.Description & "<br>"
Err.Clear

If IsObject(jsonObj(jsonObj.defaultPropertyName)) Then
    Response.Write "SERIALIZE...<br>"
    Response.Write jsonObj(jsonObj.defaultPropertyName).Serialize() & "<br>"
    Response.Write "SERIALIZE_OK<br>"
Else
    Response.Write "DEFAULT_NOT_OBJECT<br>"
End If
Response.Write "DONE<br>"
%>

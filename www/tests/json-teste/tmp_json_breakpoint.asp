<!--#include file="jsonObject.class.asp" -->
<%
Option Explicit
Dim jsonObj, jsonArr, outputObj, jsonString
Set jsonObj = New JSONobject
Set jsonArr = New jsonArray

jsonString = "[{ ""strings"" : ""valorTexto"", ""numbers"": 123.456 }]"
Set outputObj = jsonObj.parse(jsonString)
Set outputObj = jsonObj.parse(jsonString)
If Left(jsonString, 1) = "[" Then Set jsonArr = outputObj

jsonObj.Add "nome", "Jozé"
jsonObj.Add "idade", 25
jsonObj.Add "data", Now

Response.Write "nome: " & jsonObj.value("nome") & "<br>"
Response.Write "idade: " & jsonObj("idade") & "<br>"
Response.Write "missing: " & jsonObj("aNonExistantPropertyName") & "(" & TypeName(jsonObj("aNonExistantPropertyName")) & ")<br>"

On Error Resume Next
Response.Write "dname=" & jsonObj.defaultPropertyName & "<br>"
Response.Write "isval=" & IsObject(jsonObj(jsonObj.defaultPropertyName)) & "<br>"
Response.Write "err=" & Err.Number & "|" & Err.Description & "<br>"
Err.Clear

If IsObject(jsonObj(jsonObj.defaultPropertyName)) Then
    Response.Write "serialize_start<br>"
    Response.Write jsonObj(jsonObj.defaultPropertyName).Serialize()
    Response.Write "<br>serialize_end<br>"
Else
    Response.Write "notobject<br>"
End If
%>

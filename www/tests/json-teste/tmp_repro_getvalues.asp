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
jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456 }"
jsonString = "[" & jsonString & "]"
Set outputObj = jsonObj.parse(jsonString)
Set jsonArr = outputObj
jsonObj.Add "nome", "Jozé"
jsonObj.Add "ficticio", True
jsonObj.Add "idade", 25
jsonObj.Add "saldo", -52
jsonObj.Add "data", Now
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

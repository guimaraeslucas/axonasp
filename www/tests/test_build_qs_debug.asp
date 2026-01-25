<%
Dim dict
Set dict = CreateObject("Scripting.Dictionary")

' Test que dict foi criado
Response.Write "Dict created: " & TypeName(dict) & vbCrLf

' Adiciona um valor
dict.Add "test", "value"
Response.Write "Dict.Count after Add: " & dict.Count & vbCrLf

' Testa if Keys() funciona
Dim keys
Set keys = dict.Keys
Response.Write "Keys type: " & TypeName(keys) & vbCrLf

' Testa AxBuildQueryString
Dim qs
qs = AxBuildQueryString(dict)
Response.Write "Query String result: '" & qs & "'" & vbCrLf
Response.Write "Query String length: " & Len(qs) & vbCrLf
%>
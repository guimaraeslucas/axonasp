<%
Response.Buffer = True
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj, jsonString
Set jsonObj = New JSONobject
jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true }"
Set outputObj = jsonObj.parse("[" & jsonString & "]")
%>
<pre><%= outputObj.Write %></pre>

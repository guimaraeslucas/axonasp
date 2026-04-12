<%
Option Explicit
Response.buffer = True
%>
<!--#include file="json-teste/jsonObject.class.asp" -->
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <title>JSON Simple Test</title>
    </head>
    <body>
        <h1>Simple JSON Test (No Recordset)</h1>
        <%
        Dim jsonObj, jsonString, outputObj

        Set jsonObj = New JSONobject

        ' Create a simple test JSON without parsing
        jsonObj.Add "text", "hello"
        jsonObj.Add "number", 123
        jsonObj.Add "bool", True

        Response.Write "<h3>Direct Serialization (no parsing)</h3>"
        jsonString = jsonObj.Serialize()
        Response.Write "<pre>" & jsonString & "</pre>"

        Response.Write "<h3>Parse Input</h3>"
        Dim inputJson
        inputJson = "{ ""name"" : ""John"", ""age"": 30, ""active"": true }"
        Response.Write "<pre>" & inputJson & "</pre>"

        Response.Write "<h3>Parse Output</h3>"
        Set outputObj = jsonObj.parse(inputJson)
        jsonString = outputObj.Serialize()
        Response.Write "<pre>" & jsonString & "</pre>"

        Set jsonObj = Nothing
        Set outputObj = Nothing
        %>
    </body>
</html>

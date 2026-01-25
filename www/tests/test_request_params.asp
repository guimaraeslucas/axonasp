<%@ LANGUAGE="VBScript" %>
<% Option Explicit %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Test Request Parameters</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .section { border: 1px solid #ccc; padding: 15px; margin: 10px 0; }
        .pass { color: green; }
        .fail { color: red; }
        pre { background: #f0f0f0; padding: 10px; }
    </style>
</head>
<body>
<h1>Test Request Parameters</h1>

<div class="section">
    <h2>AxGetGet (Query Parameters)</h2>
    <p><strong>URL: </strong><%= Request.ServerVariables("REQUEST_URI") %></p>
    <pre>
<%
    Dim getParams
    Set getParams = AxGetGet()
    If TypeName(getParams) = "DictionaryLibrary" Then
        Response.Write "Class: DictionaryLibrary" & vbCrLf
        Response.Write "Count: " & getParams.Count & vbCrLf
        Dim i
        For Each i in getParams.Keys()
            Response.Write "  " & i & " = " & getParams.Item(i) & vbCrLf
        Next
    Else
        Response.Write "Type: " & TypeName(getParams) & vbCrLf
        Response.Write "Value: " & CStr(getParams) & vbCrLf
    End If
%>
    </pre>
</div>

<div class="section">
    <h2>AxGetPost (Form Parameters)</h2>
    <pre>
<%
    Dim postParams
    Set postParams = AxGetPost()
    If TypeName(postParams) = "DictionaryLibrary" Then
        Response.Write "Class: DictionaryLibrary" & vbCrLf
        Response.Write "Count: " & postParams.Count & vbCrLf
        Dim j
        For Each j in postParams.Keys()
            Response.Write "  " & j & " = " & postParams.Item(j) & vbCrLf
        Next
    Else
        Response.Write "Type: " & TypeName(postParams) & vbCrLf
        Response.Write "Value: " & CStr(postParams) & vbCrLf
    End If
%>
    </pre>
</div>

<div class="section">
    <h2>AxGetRequest (GET + POST)</h2>
    <pre>
<%
    Dim allParams
    Set allParams = AxGetRequest()
    If TypeName(allParams) = "DictionaryLibrary" Then
        Response.Write "Class: DictionaryLibrary" & vbCrLf
        Response.Write "Count: " & allParams.Count & vbCrLf
        Dim k
        For Each k in allParams.Keys()
            Response.Write "  " & k & " = " & allParams.Item(k) & vbCrLf
        Next
    Else
        Response.Write "Type: " & TypeName(allParams) & vbCrLf
        Response.Write "Value: " & CStr(allParams) & vbCrLf
    End If
%>
    </pre>
</div>

<div class="section">
    <h2>Test Links</h2>
    <ul>
        <li><a href="test_request_params.asp">No Parameters</a></li>
        <li><a href="test_request_params.asp?name=John&age=25&city=NewYork">With GET Parameters</a></li>
        <li><a href="test_request_params.asp?x=1&y=2&z=3">Simple GET</a></li>
    </ul>
</div>

</body>
</html>

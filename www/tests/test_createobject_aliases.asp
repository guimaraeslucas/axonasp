<%@ Page Language="VBScript" %>
<html>
<head>
    <title>Server.CreateObject Aliases Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; }
        h1 { color: #0066cc; border-bottom: 3px solid #0066cc; padding-bottom: 10px; }
        h2 { color: #333; margin-top: 30px; border-left: 4px solid #0066cc; padding-left: 10px; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        table td { padding: 8px; border: 1px solid #ddd; }
        table td:first-child { font-weight: bold; background: #f0f0f0; width: 400px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Server.CreateObject - Alias Test</h1>
        <p>Testing all supported object names and aliases</p>
        <table>
            <tr><th>Test</th><th>Result</th></tr>
<%
On Error Resume Next
Dim obj

' G3 Libraries
Set obj = Server.CreateObject("G3JSON")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>G3JSON</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>G3JSON</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("JSON")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>JSON (alias)</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>JSON (alias)</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("G3FILES")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>G3FILES</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>G3FILES</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("G3HTTP")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>G3HTTP</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>G3HTTP</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("G3TEMPLATE")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>G3TEMPLATE</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>G3TEMPLATE</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("TEMPLATE")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>TEMPLATE (alias)</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>TEMPLATE (alias)</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("G3MAIL")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>G3MAIL</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>G3MAIL</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("MAIL")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>MAIL (alias)</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>MAIL (alias)</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

' ADODB
Set obj = Server.CreateObject("ADODB.Connection")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>ADODB.Connection</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>ADODB.Connection</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("ADODB")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>ADODB (alias)</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>ADODB (alias)</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("ADODB.Recordset")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>ADODB.Recordset</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>ADODB.Recordset</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("ADODB.Stream")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>ADODB.Stream</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>ADODB.Stream</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

' Scripting
Set obj = Server.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>Scripting.FileSystemObject</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>Scripting.FileSystemObject</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("FSO")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>FSO (alias)</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>FSO (alias)</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("Scripting.Dictionary")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>Scripting.Dictionary</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>Scripting.Dictionary</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

' MSXML2
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>MSXML2.ServerXMLHTTP</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>MSXML2.ServerXMLHTTP</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("XMLHTTP")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>XMLHTTP (alias)</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>XMLHTTP (alias)</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear

Set obj = Server.CreateObject("MSXML2.DOMDocument")
If Err.Number = 0 And IsObject(obj) Then
    Response.Write "<tr><td>MSXML2.DOMDocument</td><td class='success'>OK</td></tr>"
Else
    Response.Write "<tr><td>MSXML2.DOMDocument</td><td class='error'>Failed</td></tr>"
End If
Set obj = Nothing
Err.Clear
%>
        </table>
        <footer style="margin-top: 40px; padding-top: 20px; border-top: 2px solid #0066cc; text-align: center;">
            <p><strong>G3 AxonASP</strong> - Complete Alias Support</p>
        </footer>
    </div>
</body>
</html>

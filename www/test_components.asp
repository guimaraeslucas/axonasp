<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - COM Components Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3, h4 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - COM Components Test</h1>
        <div class="intro">
            <p>Tests Server.CreateObject with various COM components: Scripting.Dictionary, MSXML2.XMLHTTP and ADODB.Connection.</p>
        </div>
        <div class="box">

<%
Response.Write "<h3>Testing Server.CreateObject Factory</h3>"

' --- Dictionary Test ---
Response.Write "<h4>Scripting.Dictionary</h4>"
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")

If IsObject(dict) Then
    Response.Write "Dictionary created successfully.<br>"
    
    dict.Add "Key1", "Value1"
    dict.Add "Key2", "Value2"
    
    Response.Write "Count: " & dict.Count & "<br>"
    Response.Write "Item('Key1'): " & dict.Item("Key1") & "<br>"
    
    If dict.Exists("Key2") Then
        Response.Write "Key2 exists.<br>"
    End If
    
    dict.Remove "Key1"
    Response.Write "Count after remove: " & dict.Count & "<br>"
    
    dict.RemoveAll
    Response.Write "Count after RemoveAll: " & dict.Count & "<br>"
Else
    Response.Write "Failed to create Dictionary.<br>"
End If

' --- XMLHTTP Test ---
Response.Write "<h4>MSXML2.XMLHTTP</h4>"
Dim http
Set http = Server.CreateObject("MSXML2.XMLHTTP")

If IsObject(http) Then
    Response.Write "XMLHTTP created successfully.<br>"
    
    ' Simple GET (assuming internet access or local)
    ' We use a dummy URL or safe test.
    ' Let's try to fetch this very page (assuming localhost:4050)
    
    Dim url
    url = "http://localhost:4050/test_basics.asp"
    
    Response.Write "Fetching: " & url & "<br>"
    
    On Error Resume Next
    http.Open "GET", url, False
    http.Send
    
    If Err.Number <> 0 Then
        Response.Write "Error fetching: " & Err.Description & "<br>"
        Err.Clear
    Else
        Response.Write "Status: " & http.Status & "<br>"
        Response.Write "Response Length: " & Len(http.ResponseText) & "<br>"
        ' Response.Write "Body: " & Left(http.ResponseText, 100) & "...<br>"
    End If
Else
    Response.Write "Failed to create XMLHTTP.<br>"
End If

' --- ADODB Stub Test ---
Response.Write "<h4>ADODB.Connection (Stub)</h4>"
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
If IsObject(conn) Then
    conn.Open "Provider=SQLOLEDB;Data Source=Localhost;"
    Response.Write "Connection State: " & conn.State & "<br>"
    conn.Close
    Response.Write "Connection State after Close: " & conn.State & "<br>"
End If

%>
        </div>
    </div>
</body>
</html>

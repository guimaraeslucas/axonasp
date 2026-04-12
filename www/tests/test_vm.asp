<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>VM Test</title>
</head>
<body>
    <h1>G3pix AxonASP VM Test</h1>
    <%
        ' Test basic variables
        name = "AxonASP"
        version = 3.0
        
        ' Test Response.Write
        Response.Write("<p>Hello from " & name & " v" & version & "!</p>")
        
        ' Test built-in functions
        text = "   hello world   "
        Response.Write("<p>Original: '" & text & "'</p>")
        Response.Write("<p>Trimmed: '" & Trim(text) & "'</p>")
        Response.Write("<p>Upper: " & UCase(text) & "</p>")
        Response.Write("<p>Lower: " & LCase(text) & "</p>")
        Response.Write("<p>Length: " & Len(text) & "</p>")
        
        ' Test conditional
        If Len(Trim(text)) > 5 Then
            Response.Write("<p>Text is long enough!</p>")
        Else
            Response.Write("<p>Text is too short!</p>")
        End If
        
        ' Test loop
        Response.Write("<p>Counting: ")
        For i = 1 To 5
            Response.Write(i & " ")
        Next
        Response.Write("</p>")
        
        ' Test Session
        Session("counter") = 42
        Response.Write("<p>Session counter: " & Session("counter") & "</p>")
    %>
</body>
</html>

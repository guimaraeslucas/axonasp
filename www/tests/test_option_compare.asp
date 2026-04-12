<%
@ Language = "VBScript"
Option Compare Text
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <title>Option Compare Text Test</title>
    <style>
        body { font-family: Tahoma, Verdana, Segoe UI, sans-serif; background: #ECE9D8; margin: 0; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; background: #fff; border: 1px solid #808080; padding: 20px; }
        h1 { margin-top: 0; color: #003399; border-bottom: 2px solid #003399; padding-bottom: 8px; }
        .test { border: 1px solid #808080; background: #f8f8f8; padding: 12px; margin: 10px 0; }
        .pass { color: #2e7d32; font-weight: bold; }
        .fail { color: #c62828; font-weight: bold; }
    </style>
</head>
<body>
<div class="container">
    <h1>Option Compare Text Enforcement</h1>

    <%
    Dim passCount
    passCount = 0

    If "ABC" = "abc" Then
        passCount = passCount + 1
        Response.Write "<div class='test'><span class='pass'>PASS</span> '=' is case-insensitive.</div>"
    Else
        Response.Write "<div class='test'><span class='fail'>FAIL</span> '=' expected True.</div>"
    End If

    If Not ("ABC" <> "abc") Then
        passCount = passCount + 1
        Response.Write "<div class='test'><span class='pass'>PASS</span> '<>' is case-insensitive.</div>"
    Else
        Response.Write "<div class='test'><span class='fail'>FAIL</span> '<>' expected False.</div>"
    End If

    If (Not ("A" < "a")) And (Not ("A" > "a")) And ("A" <= "a") And ("A" >= "a") Then
        passCount = passCount + 1
        Response.Write "<div class='test'><span class='pass'>PASS</span> '<', '>', '<=', '>=' use text compare ordering.</div>"
    Else
        Response.Write "<div class='test'><span class='fail'>FAIL</span> relational operators mismatch in text mode.</div>"
    End If

    Dim selectResult
    selectResult = ""
    Select Case "Alpha"
        Case "alpha"
            selectResult = "matched"
        Case Else
            selectResult = "not matched"
    End Select

    If selectResult = "matched" Then
        passCount = passCount + 1
        Response.Write "<div class='test'><span class='pass'>PASS</span> Select Case is case-insensitive in text mode.</div>"
    Else
        Response.Write "<div class='test'><span class='fail'>FAIL</span> Select Case expected case-insensitive match.</div>"
    End If
    %>

    <div class="test"><strong>Summary:</strong> <%= passCount %>/4 passed</div>
</div>
</body>
</html>

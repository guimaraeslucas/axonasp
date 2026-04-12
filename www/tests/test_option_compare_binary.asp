<%
@ Language = "VBScript"
Option Compare Binary
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <title>Option Compare Binary Test</title>
        <style>
            body {
                font-family:
                    Tahoma,
                    Verdana,
                    Segoe UI,
                    sans-serif;
                background: #ece9d8;
                margin: 0;
                padding: 20px;
            }
            .container {
                max-width: 900px;
                margin: 0 auto;
                background: #fff;
                border: 1px solid #808080;
                padding: 20px;
            }
            h1 {
                margin-top: 0;
                color: #003399;
                border-bottom: 2px solid #003399;
                padding-bottom: 8px;
            }
            .test {
                border: 1px solid #808080;
                background: #f8f8f8;
                padding: 12px;
                margin: 10px 0;
            }
            .pass {
                color: #2e7d32;
                font-weight: bold;
            }
            .fail {
                color: #c62828;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Option Compare Binary Coexistence</h1>

            <%
            Dim passCount
            passCount = 0

            If ("ABC" = "abc") = False Then
                passCount = passCount + 1
                Response.Write "<div class='test'><span class='pass'>PASS</span> '=' remains case-sensitive in binary mode.</div>"
            Else
                Response.Write "<div class='test'><span class='fail'>FAIL</span> '=' expected False in binary mode.</div>"
            End If

            If ("A" < "a") = True Then
                passCount = passCount + 1
                Response.Write "<div class='test'><span class='pass'>PASS</span> '<' keeps binary ordering in binary mode.</div>"
            Else
                Response.Write "<div class='test'><span class='fail'>FAIL</span> '<' expected True for A < a in binary mode.</div>"
            End If

            Dim selectResult
            selectResult = ""
            Select Case "Alpha"
                Case "alpha"
                    selectResult = "matched"
                Case Else
                    selectResult = "not matched"
            End Select

            If selectResult = "not matched" Then
                passCount = passCount + 1
                Response.Write "<div class='test'><span class='pass'>PASS</span> Select Case remains case-sensitive in binary mode.</div>"
            Else
                Response.Write "<div class='test'><span class='fail'>FAIL</span> Select Case expected no match in binary mode.</div>"
            End If
            %>

            <div class="test">
                <strong>Summary:</strong>
                <%= passCount %>/3 passed
            </div>
        </div>
    </body>
</html>

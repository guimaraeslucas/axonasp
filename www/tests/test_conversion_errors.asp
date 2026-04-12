<%
@ Language = "VBScript"
%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <title>Conversion Error Handling Test</title>
        <style>
            body {
                font-family: "Segoe UI", Tahoma, Arial;
                margin: 20px;
                background: #ece9d8;
            }
            h1 {
                color: #003399;
                border-bottom: 3px solid #003399;
                padding-bottom: 5px;
            }
            .test {
                background: white;
                border: 1px solid #808080;
                padding: 15px;
                margin: 10px 0;
            }
            .success {
                color: green;
                font-weight: bold;
            }
            .error {
                color: red;
                font-weight: bold;
            }
            table {
                border-collapse: collapse;
                margin: 10px 0;
            }
            th,
            td {
                border: 1px solid #808080;
                padding: 8px;
                text-align: left;
            }
            th {
                background: #335ea8;
                color: white;
            }
        </style>
    </head>
    <body>
        <h1>Conversion Error Handling Test - AxonVM</h1>

        <%
        Const TYPE_MISMATCH_HRESULT = -2146828275

        Dim testValue, errNum, errDesc

        ' Test CInt with invalid string
        Response.Write "<div class='test'>"
        Response.Write "<h3>Test 1: CInt(""invalid"")</h3>"
        On Error Resume Next
        Err.Clear
        testValue = CInt("invalid")
        errNum = Err.Number
        errDesc = Err.Description
        On Error Goto 0
        Response.Write "<table>"
        Response.Write "<tr><th>Property</th><th>Value</th></tr>"
        Response.Write "<tr><td>Result</td><td>" & testValue & "</td></tr>"
        Response.Write "<tr><td>Err.Number</td><td>" & errNum & "</td></tr>"
        Response.Write "<tr><td>Err.Description</td><td>" & Server.HTMLEncode(errDesc) & "</td></tr>"
        Response.Write "</table>"
        If errNum = TYPE_MISMATCH_HRESULT Then
            Response.Write "<span class='success'>✓ PASS: Error correctly captured</span>"
        Else
            Response.Write "<span class='error'>✗ FAIL: Expected HRESULT " & TYPE_MISMATCH_HRESULT & ", got " & errNum & "</span>"
        End If
        Response.Write "</div>"

        ' Test CLng with invalid string
        Response.Write "<div class='test'>"
        Response.Write "<h3>Test 2: CLng(""not-a-number"")</h3>"
        On Error Resume Next
        Err.Clear
        testValue = CLng("not-a-number")
        errNum = Err.Number
        errDesc = Err.Description
        On Error Goto 0
        Response.Write "<table>"
        Response.Write "<tr><th>Property</th><th>Value</th></tr>"
        Response.Write "<tr><td>Result</td><td>" & testValue & "</td></tr>"
        Response.Write "<tr><td>Err.Number</td><td>" & errNum & "</td></tr>"
        Response.Write "<tr><td>Err.Description</td><td>" & Server.HTMLEncode(errDesc) & "</td></tr>"
        Response.Write "</table>"
        If errNum = TYPE_MISMATCH_HRESULT Then
            Response.Write "<span class='success'>✓ PASS: Error correctly captured</span>"
        Else
            Response.Write "<span class='error'>✗ FAIL: Expected HRESULT " & TYPE_MISMATCH_HRESULT & ", got " & errNum & "</span>"
        End If
        Response.Write "</div>"

        ' Test CDbl with invalid string
        Response.Write "<div class='test'>"
        Response.Write "<h3>Test 3: CDbl(""xyz123"")</h3>"
        On Error Resume Next
        Err.Clear
        testValue = CDbl("xyz123")
        errNum = Err.Number
        errDesc = Err.Description
        On Error Goto 0
        Response.Write "<table>"
        Response.Write "<tr><th>Property</th><th>Value</th></tr>"
        Response.Write "<tr><td>Result</td><td>" & testValue & "</td></tr>"
        Response.Write "<tr><td>Err.Number</td><td>" & errNum & "</td></tr>"
        Response.Write "<tr><td>Err.Description</td><td>" & Server.HTMLEncode(errDesc) & "</td></tr>"
        Response.Write "</table>"
        If errNum = TYPE_MISMATCH_HRESULT Then
            Response.Write "<span class='success'>✓ PASS: Error correctly captured</span>"
        Else
            Response.Write "<span class='error'>✗ FAIL: Expected HRESULT " & TYPE_MISMATCH_HRESULT & ", got " & errNum & "</span>"
        End If
        Response.Write "</div>"

        ' Test CSng with invalid string
        Response.Write "<div class='test'>"
        Response.Write "<h3>Test 4: CSng(""abc"")</h3>"
        On Error Resume Next
        Err.Clear
        testValue = CSng("abc")
        errNum = Err.Number
        errDesc = Err.Description
        On Error Goto 0
        Response.Write "<table>"
        Response.Write "<tr><th>Property</th><th>Value</th></tr>"
        Response.Write "<tr><td>Result</td><td>" & testValue & "</td></tr>"
        Response.Write "<tr><td>Err.Number</td><td>" & errNum & "</td></tr>"
        Response.Write "<tr><td>Err.Description</td><td>" & Server.HTMLEncode(errDesc) & "</td></tr>"
        Response.Write "</table>"
        If errNum = TYPE_MISMATCH_HRESULT Then
            Response.Write "<span class='success'>✓ PASS: Error correctly captured</span>"
        Else
            Response.Write "<span class='error'>✗ FAIL: Expected HRESULT " & TYPE_MISMATCH_HRESULT & ", got " & errNum & "</span>"
        End If
        Response.Write "</div>"

        ' Test CByte with invalid string
        Response.Write "<div class='test'>"
        Response.Write "<h3>Test 5: CByte(""invalid"")</h3>"
        On Error Resume Next
        Err.Clear
        testValue = CByte("invalid")
        errNum = Err.Number
        errDesc = Err.Description
        On Error Goto 0
        Response.Write "<table>"
        Response.Write "<tr><th>Property</th><th>Value</th></tr>"
        Response.Write "<tr><td>Result</td><td>" & testValue & "</td></tr>"
        Response.Write "<tr><td>Err.Number</td><td>" & errNum & "</td></tr>"
        Response.Write "<tr><td>Err.Description</td><td>" & Server.HTMLEncode(errDesc) & "</td></tr>"
        Response.Write "</table>"
        If errNum = TYPE_MISMATCH_HRESULT Then
            Response.Write "<span class='success'>✓ PASS: Error correctly captured</span>"
        Else
            Response.Write "<span class='error'>✗ FAIL: Expected HRESULT " & TYPE_MISMATCH_HRESULT & ", got " & errNum & "</span>"
        End If
        Response.Write "</div>"

        ' Test valid conversion (should not error)
        Response.Write "<div class='test'>"
        Response.Write "<h3>Test 6: CInt(""123"") - Valid Conversion</h3>"
        On Error Resume Next
        Err.Clear
        testValue = CInt("123")
        errNum = Err.Number
        errDesc = Err.Description
        On Error Goto 0
        Response.Write "<table>"
        Response.Write "<tr><th>Property</th><th>Value</th></tr>"
        Response.Write "<tr><td>Result</td><td>" & testValue & "</td></tr>"
        Response.Write "<tr><td>Err.Number</td><td>" & errNum & "</td></tr>"
        Response.Write "<tr><td>Err.Description</td><td>" & Server.HTMLEncode(errDesc) & "</td></tr>"
        Response.Write "</table>"
        If errNum = 0 And testValue = 123 Then
            Response.Write "<span class='success'>✓ PASS: Valid conversion works correctly</span>"
        Else
            Response.Write "<span class='error'>✗ FAIL: Valid conversion should not error</span>"
        End If
        Response.Write "</div>"

        ' Test hex string conversion (should not error)
        Response.Write "<div class='test'>"
        Response.Write "<h3>Test 7: CInt(""&HFF"") - Hex Conversion</h3>"
        On Error Resume Next
        Err.Clear
        testValue = CInt("&HFF")
        errNum = Err.Number
        errDesc = Err.Description
        On Error Goto 0
        Response.Write "<table>"
        Response.Write "<tr><th>Property</th><th>Value</th></tr>"
        Response.Write "<tr><td>Result</td><td>" & testValue & "</td></tr>"
        Response.Write "<tr><td>Err.Number</td><td>" & errNum & "</td></tr>"
        Response.Write "<tr><td>Err.Description</td><td>" & Server.HTMLEncode(errDesc) & "</td></tr>"
        Response.Write "</table>"
        If errNum = 0 And testValue = 255 Then
            Response.Write "<span class='success'>✓ PASS: Hex conversion works correctly</span>"
        Else
            Response.Write "<span class='error'>✗ FAIL: Hex conversion should return 255</span>"
        End If
        Response.Write "</div>"
        %>
    </body>
</html>

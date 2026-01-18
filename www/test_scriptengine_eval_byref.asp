<%
On Error Resume Next

' Declare variables for test tracking
Dim passCount, failCount
passCount = 0
failCount = 0

Sub PrintTestResult(testName, result, expected)
    Response.Write "<tr><td>" & testName & "</td><td>" & result & "</td><td>" & expected & "</td>"
    If CStr(result) = CStr(expected) Then
        Response.Write "<td class='success'>PASS</td></tr>"
        passCount = passCount + 1
    Else
        Response.Write "<td class='error'>FAIL</td></tr>"
        failCount = failCount + 1
    End If
End Sub

Sub ModifyByRef(ByRef value)
    value = value * 2
End Sub

Sub ModifyByVal(ByVal value)
    value = value * 2
End Sub
%>
<html>
<head>
    <title>ScriptEngine, Eval, ByRef & ByVal Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; }
        h1 { color: #0066cc; border-bottom: 3px solid #0066cc; padding-bottom: 10px; }
        h2 { color: #333; margin-top: 30px; border-left: 4px solid #0066cc; padding-left: 10px; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th { background: #0066cc; color: white; padding: 10px; text-align: left; }
        td { padding: 8px; border: 1px solid #ddd; }
        td:first-child { font-weight: bold; background: #f0f0f0; width: 350px; }
        .test-section { margin: 20px 0; padding: 15px; background: #f9f9f9; border-left: 4px solid #0066cc; }
        .summary { margin-top: 30px; padding: 15px; background: #f0f0f0; border: 2px solid #0066cc; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <h1>ScriptEngine, Eval, TypeName, VarType, ByRef & ByVal Tests</h1>
        <p>Comprehensive test suite for VBScript engine properties and functions</p>

        <div class="test-section">
            <h2>ScriptEngine Properties</h2>
            <table>
                <tr><th>Test</th><th>Result</th><th>Expected</th><th>Status</th></tr>
<%
PrintTestResult "ScriptEngine()", ScriptEngine(), "VBScript"
PrintTestResult "ScriptEngineBuildVersion()", ScriptEngineBuildVersion(), 18702
PrintTestResult "ScriptEngineMajorVersion()", ScriptEngineMajorVersion(), 5
PrintTestResult "ScriptEngineMinorVersion()", ScriptEngineMinorVersion(), 8
%>
            </table>
        </div>

        <div class="test-section">
            <h2>TypeName Function Tests</h2>
            <table>
                <tr><th>Test</th><th>Result</th><th>Expected</th><th>Status</th></tr>
<%
PrintTestResult "TypeName(42)", TypeName(42), "Integer"
PrintTestResult "TypeName(""hello"")", TypeName("hello"), "String"
PrintTestResult "TypeName(True)", TypeName(True), "Boolean"
PrintTestResult "TypeName(3.14)", TypeName(3.14), "Double"
PrintTestResult "TypeName(Array(1,2))", TypeName(Array(1, 2)), "Variant()"
%>
            </table>
        </div>

        <div class="test-section">
            <h2>VarType Function Tests</h2>
            <table>
                <tr><th>Test</th><th>Result</th><th>Expected</th><th>Status</th></tr>
<%
PrintTestResult "VarType(42)", VarType(42), 2
PrintTestResult "VarType(""hello"")", VarType("hello"), 8
PrintTestResult "VarType(True)", VarType(True), 11
PrintTestResult "VarType(3.14)", VarType(3.14), 5
PrintTestResult "VarType(Array(1,2))", VarType(Array(1, 2)), 8204
%>
            </table>
        </div>

        <div class="test-section">
            <h2>Eval Function Tests</h2>
            <table>
                <tr><th>Test</th><th>Result</th><th>Expected</th><th>Status</th></tr>
<%
PrintTestResult "Eval(""42"")", Eval("42"), 42
PrintTestResult "Eval(""true"")", Eval("true"), True
%>
            </table>
        </div>

        <div class="test-section">
            <h2>ByRef Parameter Tests</h2>
            <table>
                <tr><th>Test</th><th>Result</th><th>Expected</th><th>Status</th></tr>
<%
Dim byRefVar
byRefVar = 5
ModifyByRef(byRefVar)
PrintTestResult "ByRef: value * 2 (5 becomes 10)", byRefVar, 10

Dim byValVar
byValVar = 5
ModifyByVal(byValVar)
PrintTestResult "ByVal: value unchanged (5 stays 5)", byValVar, 5
%>
            </table>
        </div>

        <div class="summary">
            <h2>Test Summary</h2>
            <p><strong>Passed:</strong> <span class="success"><% Response.Write passCount %></span></p>
            <p><strong>Failed:</strong> <span class="error"><% Response.Write failCount %></span></p>
            <p><strong>Total:</strong> <% Response.Write passCount + failCount %></p>
        </div>

        <footer style="margin-top: 40px; padding-top: 20px; border-top: 2px solid #0066cc; text-align: center;">
            <p><strong>G3 AxonASP</strong> - ScriptEngine, Eval, TypeName, VarType, ByRef/ByVal Implementation</p>
        </footer>
    </div>
</body>
</html>

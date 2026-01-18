<%
Response.Write "=== ScriptEngine Tests ===<br/>"
Response.Write "ScriptEngine: " & ScriptEngine() & " (expected: VBScript)<br/>"
Response.Write "ScriptEngineBuildVersion: " & ScriptEngineBuildVersion() & " (expected: 18702)<br/>"
Response.Write "ScriptEngineMajorVersion: " & ScriptEngineMajorVersion() & " (expected: 5)<br/>"
Response.Write "ScriptEngineMinorVersion: " & ScriptEngineMinorVersion() & " (expected: 8)<br/>"

Response.Write "<br/>=== TypeName Tests ===<br/>"
Response.Write "TypeName(42): " & TypeName(42) & " (expected: Integer)<br/>"
Response.Write "TypeName(""hello""): " & TypeName("hello") & " (expected: String)<br/>"
Response.Write "TypeName(True): " & TypeName(True) & " (expected: Boolean)<br/>"
Response.Write "TypeName(3.14): " & TypeName(3.14) & " (expected: Double)<br/>"

Response.Write "<br/>=== VarType Tests ===<br/>"
Response.Write "VarType(42): " & VarType(42) & " (expected: 2)<br/>"
Response.Write "VarType(""hello""): " & VarType("hello") & " (expected: 8)<br/>"
Response.Write "VarType(True): " & VarType(True) & " (expected: 11)<br/>"

Response.Write "<br/>=== Eval Tests ===<br/>"
Response.Write "Eval(""42""): " & Eval("42") & " (expected: 42)<br/>"
Response.Write "Eval(""true""): " & Eval("true") & " (expected: -1 or True)<br/>"

Response.Write "<br/>=== ByRef Tests ===<br/>"
Sub TestByRef(ByRef val)
    val = val * 2
End Sub

Dim testVar
testVar = 5
TestByRef(testVar)
Response.Write "ByRef test: " & testVar & " (expected: 10)<br/>"

Response.Write "<br/>=== Summary ===<br/>"
If ScriptEngine() = "VBScript" Then
    Response.Write "ScriptEngine: PASS<br/>"
Else
    Response.Write "ScriptEngine: FAIL<br/>"
End If

If ScriptEngineBuildVersion() = 18702 Then
    Response.Write "ScriptEngineBuildVersion: PASS<br/>"
Else
    Response.Write "ScriptEngineBuildVersion: FAIL<br/>"
End If

If ScriptEngineMajorVersion() = 5 Then
    Response.Write "ScriptEngineMajorVersion: PASS<br/>"
Else
    Response.Write "ScriptEngineMajorVersion: FAIL<br/>"
End If

If ScriptEngineMinorVersion() = 8 Then
    Response.Write "ScriptEngineMinorVersion: PASS<br/>"
Else
    Response.Write "ScriptEngineMinorVersion: FAIL<br/>"
End If

If TypeName(42) = "Integer" Then
    Response.Write "TypeName: PASS<br/>"
Else
    Response.Write "TypeName: FAIL<br/>"
End If

If VarType(42) = 2 Then
    Response.Write "VarType: PASS<br/>"
Else
    Response.Write "VarType: FAIL<br/>"
End If

If Eval("42") = 42 Then
    Response.Write "Eval: PASS<br/>"
Else
    Response.Write "Eval: FAIL<br/>"
End If

If testVar = 10 Then
    Response.Write "ByRef: PASS<br/>"
Else
    Response.Write "ByRef: FAIL<br/>"
End If
%>

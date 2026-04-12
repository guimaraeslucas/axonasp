<%
' test_null_propagation.asp
' Verifies VBScript Null propagation semantics in AxonASP.
' Classic ASP/VBScript rule: any arithmetic or comparison operation
' involving Null produces Null, except for specific logical short-circuits.

Dim x
x = Null

Response.Write "<h2>AxonASP - Null Propagation Tests</h2>"
Response.Write "<table border='1' cellpadding='4'>"
Response.Write "<tr><th>Expression</th><th>Result</th><th>Expected</th><th>Pass?</th></tr>"

Sub TestRow(expr, result, expected)
    Dim pass
    pass = (CStr(result) = CStr(expected))
    Dim rowColor
    If pass Then rowColor = "#ccffcc" Else rowColor = "#ffcccc"
    Response.Write "<tr style='background:" & rowColor & "'>"
    Response.Write "<td>" & expr & "</td>"
    Response.Write "<td>" & CStr(result) & "</td>"
    Response.Write "<td>" & CStr(expected) & "</td>"
    Response.Write "<td>" & IIf(pass, "PASS", "FAIL") & "</td>"
    Response.Write "</tr>"
End Sub

' Arithmetic propagation
TestRow "Null + 1", x + 1, "Null"
TestRow "Null - 1", x - 1, "Null"
TestRow "Null * 5", x * 5, "Null"
TestRow "Null / 2", x / 2, "Null"
TestRow "Null & &quot;hi&quot;", x & "hi", "Null"
TestRow "Null ^ 2", x ^ 2, "Null"

' Comparison propagation
TestRow "Null = 1", x = 1, "Null"
TestRow "Null <> 1", x <> 1, "Null"
TestRow "Null < 5", x < 5, "Null"
TestRow "Null > 0", x > 0, "Null"

' Unary propagation
TestRow "Not Null", Not x, "Null"
TestRow "-Null", -x, "Null"

' Logical short-circuits
TestRow "Null And False", x And False, "False"
TestRow "Null And True", x And True, "Null"
TestRow "Null Or True", x Or True, "True"
TestRow "Null Or False", x Or False, "Null"
TestRow "Null Xor True", x Xor True, "Null"
TestRow "False Imp Null", False Imp x, "True"
TestRow "Null Imp True", x Imp True, "True"
TestRow "Null Imp False", x Imp False, "Null"

Response.Write "</table>"

' IsNull() function check
If IsNull(x) Then
    Response.Write "<p><b>IsNull(x)</b>: True (PASS)</p>"
Else
    Response.Write "<p><b>IsNull(x)</b>: False (FAIL)</p>"
End If

If IsNull(1) Then
    Response.Write "<p><b>IsNull(1)</b>: True (FAIL)</p>"
Else
    Response.Write "<p><b>IsNull(1)</b>: False (PASS)</p>"
End If
%>

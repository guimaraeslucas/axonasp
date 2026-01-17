<%
Function testDimColon(val)
    Dim x : x = val * 2
    Dim y : y = val + 10
    testDimColon = x + y
End Function

Dim funcResult : funcResult = testDimColon(5)
Response.Write "Result: " & funcResult
%>

<%
Response.Write "Test IsNumeric Implementation" & vbCrLf
Response.Write "==============================" & vbCrLf & vbCrLf

Dim testValues(10)
testValues(0) = "IHJEF"
testValues(1) = "JEJDI"
testValues(2) = "123"
testValues(3) = "123.45"
testValues(4) = "ABC"
testValues(5) = "12AB"
testValues(6) = ""
testValues(7) = " "
testValues(8) = "0"
testValues(9) = "-123"

Dim i
For i = 0 To 9
    If Not IsEmpty(testValues(i)) Then
        Response.Write "IsNumeric(""" & testValues(i) & """): " & IsNumeric(testValues(i)) & vbCrLf
    End If
Next
%>

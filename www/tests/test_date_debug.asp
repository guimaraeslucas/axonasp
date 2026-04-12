<%
Response.Write "<h1>Date Literal Debug</h1>"

' Test 1: Direct assignment
Dim d1
d1 = #1 / 1 / 2023#
Response.Write "d1 type: " & TypeName(d1) & "<br>"
Response.Write "d1 value: " & d1 & "<br>"
Response.Write "d1 as string: " & CStr(d1) & "<br>"

' Test 2: VarType check
Response.Write "VarType(d1): " & VarType(d1) & " (7=Date)<br>"

' Test 3: DateSerial comparison
Dim d2
d2 = DateSerial(2023, 1, 1)
Response.Write "d2 = DateSerial(2023,1,1): " & d2 & "<br>"

' Test 4: Check if dates are equal
If d1 = d2 Then
    Response.Write "d1 equals d2<br>"
Else
    Response.Write "d1 does NOT equal d2<br>"
End If

' Test 5: Get day, month, year
Response.Write "Day(d1): " & Day(d1) & "<br>"
Response.Write "Month(d1): " & Month(d1) & "<br>"
Response.Write "Year(d1): " & Year(d1) & "<br>"
%>

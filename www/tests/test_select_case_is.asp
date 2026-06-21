<%@ Language="VBScript" %>
<%
Dim x, result

' Test 1: Case Is > comparison
x = 7
Select Case x
    Case Is > 5
        Response.Write "gt5|"
    Case Else
        Response.Write "else1|"
End Select

' Test 2: Case Is < comparison
x = 3
Select Case x
    Case Is < 5
        Response.Write "lt5|"
    Case Else
        Response.Write "else2|"
End Select

' Test 3: Case Is >= comparison
x = 5
Select Case x
    Case Is >= 5
        Response.Write "gte5|"
    Case Else
        Response.Write "else3|"
End Select

' Test 4: Case Is <= comparison
x = 0
Select Case x
    Case Is <= 0
        Response.Write "lte0|"
    Case Else
        Response.Write "else4|"
End Select

' Test 5: Case Is <> comparison
x = 42
Select Case x
    Case Is <> 99
        Response.Write "neq99|"
    Case Else
        Response.Write "else5|"
End Select

' Test 6: Case Is = comparison
x = 10
Select Case x
    Case Is = 10
        Response.Write "eq10|"
    Case Else
        Response.Write "else6|"
End Select

' Test 7: Mixed Case Is with comma-separated values
x = 8
Select Case x
    Case 1, Is > 5, 15
        Response.Write "mixed|"
    Case Else
        Response.Write "else7|"
End Select

' Test 8: Multiple Is clauses with comma
x = 2
Select Case x
    Case Is > 10, Is < 3
        Response.Write "multiis|"
    Case Else
        Response.Write "else8|"
End Select

' Test 9: Case Is with range and Is mixed
x = 25
Select Case x
    Case 1 To 10, Is > 20
        Response.Write "rangeis|"
    Case Else
        Response.Write "else9|"
End Select

' Test 10: Case Else fallback when no Is matches
x = 50
Select Case x
    Case Is > 100
        Response.Write "nope|"
    Case Else
        Response.Write "fallback|"
End Select
%>
<%@ Language=VBScript %>
<%
Response.Write "SELECT_CASE_MULTI_START<br>"

val = 3
Select Case val
    Case 1, 2
        Response.Write "fail-multi<br>"
    Case 3, 4
        Response.Write "match-multi<br>"
    Case Else
        Response.Write "fail-else<br>"
End Select

val2 = 7
Select Case val2
    Case 1 To 5
        Response.Write "fail-range<br>"
    Case 6 To 8
        Response.Write "match-range<br>"
    Case Else
        Response.Write "fail-else<br>"
End Select

Response.Write "SELECT_CASE_MULTI_END"
%>

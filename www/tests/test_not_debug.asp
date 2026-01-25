<%@ Language=VBScript %>
<%
    Dim obj
    Set obj = Nothing
    
    Response.Write "Test 1: obj Is Nothing = " & (obj Is Nothing) & "<br>"
    Response.Write "Test 2: Not (obj Is Nothing) = " & Not (obj Is Nothing) & "<br>"
    Response.Write "Test 3: CBool(Not (obj Is Nothing)) = " & CBool(Not (obj Is Nothing)) & "<br>"
    
    Dim result1
    result1 = obj Is Nothing
    Response.Write "Test 4: result1 (obj Is Nothing) = " & result1 & "<br>"
    
    Dim result2
    result2 = Not (obj Is Nothing)
    Response.Write "Test 5: result2 (Not (obj Is Nothing)) = " & result2 & "<br>"
    
    Dim result3
    result3 = Not obj Is Nothing
    Response.Write "Test 6: result3 (Not obj Is Nothing) = " & result3 & " (should be same as Test 2)<br>"
    
    Response.Write "<br>---IF Tests---<br>"
    
    If Not obj Is Nothing Then
        Response.Write "Test 7: Not obj Is Nothing in If: True (EXPECTED)<br>"
    Else
        Response.Write "Test 7: Not obj Is Nothing in If: False (BUG!)<br>"
    End If
    
    If Not (obj Is Nothing) Then
        Response.Write "Test 8: Not (obj Is Nothing) in If: True (EXPECTED)<br>"
    Else
        Response.Write "Test 8: Not (obj Is Nothing) in If: False (BUG!)<br>"
    End If
%>

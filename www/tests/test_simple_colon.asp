<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<body>
    <h1>Minimal Dim Colon Test</h1>
    
    <%
    ' Test WITHOUT colon
    Dim a
    a = 10
    Response.Write "1. Without colon: a = " & a & "<br/>"
    
    ' Test WITH colon
    Dim b : b = 20
    Response.Write "2. With colon: b = " & b & "<br/>"
    
    ' Test with function call WITHOUT colon
    Function getVal()
        getVal = 30
    End Function
    Dim c
    c = getVal()
    Response.Write "3. Function, no colon: c = " & c & "<br/>"
    
    ' Test with function call WITH colon
    Dim d : d = getVal()
    Response.Write "4. Function, with colon: d = " & d & "<br/>"
    %>
    
</body>
</html>

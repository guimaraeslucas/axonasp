<%@ Language=VBScript %>
<html>
<head><title>Eval Test</title></head>
<body>
    <h3>Testing Eval Function</h3>
    <%
    Dim x, y, res
    x = 10
    y = 20
    
    Response.Write "<p>x = " & x & ", y = " & y & "</p>"
    
    res = Eval("x + y")
    Response.Write "<p>Eval('x + y') = " & res & " (Expected: 30)</p>"
    
    res = Eval("x * 5")
    Response.Write "<p>Eval('x * 5') = " & res & " (Expected: 50)</p>"
    
    res = Eval("y / 2")
    Response.Write "<p>Eval('y / 2') = " & res & " (Expected: 10)</p>"
    
    Dim str
    str = "Hello"
    res = Eval("str & ' World'")
    Response.Write "<p>Eval('str & '' World''') = " & res & " (Expected: Hello World)</p>"
    %>
</body>
</html>

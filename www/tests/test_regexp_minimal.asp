<%@ Language="VBScript" %>
<html>
<head><title>Regex Test - Absolute Minimal</title></head>
<body>

<%
    ' Test object creation
    Set regex = Server.CreateObject("G3REGEXP")
    Response.Write("<p>Object created: " & TypeName(regex) & "</p>")
    
    ' Direct assignment (VBScript style)
    regex.Pattern = "\d+"
    Response.Write("<p>Pattern set via assignment</p>")
    
    ' Test the method
    result = regex.Test("abc123def")
    Response.Write("<p>Test result: " & result & " (type: " & TypeName(result) & ")</p>")
    
    if result then
        Response.Write("<p><strong>SUCCESS: Pattern matched!</strong></p>")
    else
        Response.Write("<p><strong>FAIL: Pattern did not match</strong></p>")
    end if
%>

</body>
</html>

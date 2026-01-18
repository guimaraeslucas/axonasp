<%@ Language="VBScript" %>
<!--
Minimal test with direct error checking
-->
<html>
<body>

<%
    response.write "Test started<br>"
    
    On Error Resume Next
    Set regex = Server.CreateObject("G3REGEXP")
    response.write "Object created, error: " & Err.Number & " - " & Err.Description & "<br>"
    
    On Error Resume Next
    regex.Pattern = "\d+"
    response.write "Pattern set, error: " & Err.Number & " - " & Err.Description & "<br>"
    
    On Error Resume Next
    result = regex.Test("123")
    response.write "Test called, result: " & result & ", error: " & Err.Number & "<br>"

%>

</body>
</html>

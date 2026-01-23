<%@ Language="VBScript" %>
<%
    Option Explicit
    Response.Write "Start of script<br>"
    
    Dim result
    result = MyFunction("Hello")
    Response.Write "Result: " & result & "<br>"

    Function MyFunction(val)
        MyFunction = "Processed: " & val
    End Function
%>

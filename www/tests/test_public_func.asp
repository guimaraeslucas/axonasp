<%@ Language="VBScript" CodePage=65001 %>
<% 
    ' Output list of known functions in engine
    Response.Write("<h2>Debug: Checking function execution</h2>")
    
    ' Define a very simple function
    function Add(a, b)
        Add = a + b
    end function

    ' Try to call
    Dim x
    x = Add(10, 20)
    Response.Write("<p>Add(10, 20) = " & x & "</p>")
%>

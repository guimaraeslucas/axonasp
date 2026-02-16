<%@ Language="VBScript" CodePage=65001 %>
<% 
    ' Simple test: function without byref
    private function Multiply(x, y)
        Multiply = x * y
    end function

    Dim result
    result = Multiply(6, 7)
    Response.Write("<p>Multiply(6, 7) = " & result & "</p>")
%>

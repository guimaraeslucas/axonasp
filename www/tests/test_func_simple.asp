<%@ Language="VBScript" CodePage=65001 %>
<% 
    ' Debug: check if function is in engine
    Response.Write("<h2>Function Execution Test</h2>")
    
    ' Simple function to test
    private function Add(a, b)
        Add = a + b
    end function

    ' Try to call it
    Dim result
    result = Add(5, 3)
    Response.Write("<p>Add(5, 3) = " & result & "</p>")
%>

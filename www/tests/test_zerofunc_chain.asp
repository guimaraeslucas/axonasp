<%
Option Explicit

class Helper
    public sub greet(name)
        response.write "Hello " & name & vbCrLf
    end sub
end class

class Container
    private p_helper
    
    Private Sub Class_Initialize()
        set p_helper = nothing
    End Sub
    
    ' Zero-arg function returns object (like json() in asplite)
    public function helper()
        if p_helper is nothing then
            set p_helper = new Helper
        end if
        set helper = p_helper
    end function
    
    ' Test calling helper.greet("X") where helper resolves via zero-arg function
    public sub doTest()
        helper.greet("World")
    end sub
End Class

dim c : set c = new Container
c.doTest()
response.write "Test passed" & vbCrLf
%>

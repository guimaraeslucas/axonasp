<%
class TestCompare
    dim i_count, i_capacity
    
    private sub class_initialize()
        i_count = 5
        i_capacity = 10
    end sub
    
    public sub TestMethod()
        if i_count >= i_capacity then
            Response.Write "Count is less than capacity<br>"
        else
            Response.Write "Count is greater than or equal to capacity<br>"
        end if
    end sub
end class

dim obj
set obj = new TestCompare
obj.TestMethod()
%>

<%
class TestRedim
    dim i_items, i_count, i_capacity
    
    private sub class_initialize()
        redim i_items(-1)
        i_count = 0
        i_capacity = 0
    end sub
    
    public sub TestMethod()
        Response.Write "Before check<br>"
        if i_count >= i_capacity then
            Response.Write "Need to resize<br>"
            redim preserve i_items(i_capacity * 1.2 + 1)
            i_capacity = ubound(i_items) + 1
            Response.Write "Resized<br>"
        end if
        Response.Write "After method<br>"
    end sub
end class

dim obj
set obj = new TestRedim
obj.TestMethod()
%>

<%
class JSONpair
    dim name, value
end class

class TestAddComplex
    dim i_properties, i_count, i_capacity
    
    private sub class_initialize()
        redim i_properties(-1)
        i_count = 0
        i_capacity = 0
    end sub
    
    public sub add(byval prop, byval obj)
        dim item
        set item = new JSONpair
        item.name = prop
        item.value = obj

        if i_count >= i_capacity then
            Response.Write "Resizing<br>"
            redim preserve i_properties(i_capacity * 1.2 + 1)
            i_capacity = ubound(i_properties) + 1
        end if

        set i_properties(i_count) = item
        i_count = i_count + 1
        Response.Write "Added<br>"
    end sub
end class

dim obj
set obj = new TestAddComplex
obj.add "test", "value"
%>

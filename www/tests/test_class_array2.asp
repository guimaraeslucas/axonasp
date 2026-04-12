<%
class TestArray2
	dim i_items, i_count, i_capacity
	
	private sub class_initialize()
		redim i_items(-1)
		i_count = 0
		i_capacity = 0
	end sub
	
	public property get items()
		dim tmp
		tmp = i_items
		if i_count < i_capacity then
			redim preserve tmp(i_count - 1)
		end if
		items = tmp
	end property
end class

dim obj
set obj = new TestArray2
dim result
result = obj.items
Response.Write "Got items: " & TypeName(result) & "<br>"
%>

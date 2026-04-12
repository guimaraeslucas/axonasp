<%
class TestArray3
	dim i_items, i_count, i_capacity
	
	private sub class_initialize()
		redim i_items(-1)
		i_count = 0
		i_capacity = 0
	end sub
	
	public property get pairs()
		dim tmp
		tmp = i_items
		if i_count < i_capacity then
			redim preserve tmp(i_count - 1)
		end if
		pairs = tmp
	end property
end class

dim obj
set obj = new TestArray3
dim result
result = obj.pairs()
Response.Write "Got pairs: " & TypeName(result) & "<br>"
%>

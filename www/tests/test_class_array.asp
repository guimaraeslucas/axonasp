<%
class TestArray
	dim i_items
	
	private sub class_initialize()
		redim i_items(-1)
	end sub
	
	public property get items()
		dim tmp
		tmp = i_items
		items = tmp
	end property
end class

dim obj
set obj = new TestArray
dim result
result = obj.items
Response.Write "Got items: " & TypeName(result) & "<br>"
%>

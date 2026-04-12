<%
class TestClass
	dim i_value
	
	private sub class_initialize()
		i_value = 42
	end sub
	
	public property get myvalue()
		dim result
		result = i_value
		myvalue = result
	end property
end class

dim obj
set obj = new TestClass
Response.Write "Result: " & obj.myvalue & "<br>"
%>

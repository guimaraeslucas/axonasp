<%
class SimpleClass
    dim myname
end class

dim obj
set obj = new SimpleClass
obj.myname = "test"
Response.Write "Set name to: " & obj.myname & "<br>"
%>

<%
    Response.Write("Hello World<br>")
    Response.Write("Test 1<br>")
    
    Dim json
    Set json = Server.CreateObject("G3JSON")
    Response.Write("Created JSON Object<br>")
    
    Dim user
    Set user = json.NewObject()
    Response.Write("Created User Object<br>")
    
    user("name") = "Lucas"
    Response.Write("Set user name<br>")
    Response.Write("User Name: " & user("name") & "<br>")
%>

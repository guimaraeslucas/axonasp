
    <%
    ' 0. Create JSON Library
    Dim json
    Set json = Server.CreateObject("G3JSON")

    ' 1. Create a new object
    Dim user
    Set user = json.NewObject()
    
    ' 2. Edit properties
    user("name") = "Lucas"
    user("role") = "Developer"
    user("age") = 32
    
    Response.Write("<h2>Created User Object</h2>")
    Response.Write("User Name: " & user("name") & "<br>")
    Response.Write("User Role: " & user("role") & "<br>")
    Response.Write("User Age: " & user("age") & "<br>")
    
    ' 3. Serialize to JSON
    Dim jsonStr
    jsonStr = json.Stringify(user)
    Response.Write("<h2>Serialized JSON</h2>")
    Response.Write("<div class='code-block'>" & jsonStr & "</div>")
    
    ' 4. Parse a JSON string
    Dim data
    data = json.Parse("{""clients"": [""Google"", ""Microsoft""], ""active"": true}")
    
    Response.Write("<h2>Parsed JSON Data</h2>")
    Response.Write("Active: " & data("active") & "<br>")
    
    ' 5. Iterate over array
    Response.Write("<h2>Clients List</h2>")
    
    Dim clientList
    clientList = data("clients") 
    
    For Each client In clientList
        Response.Write("- " & client & "<br>")
    Next
    %>


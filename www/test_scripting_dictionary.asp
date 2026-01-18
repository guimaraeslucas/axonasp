<%@ Language="VBScript" %>
<%
' Test Scripting.Dictionary

Response.ContentType = "text/html; charset=utf-8"
%>

<!DOCTYPE html>
<html>
<head>
    <title>Test Scripting.Dictionary</title>
</head>
<body>
    <h1>Scripting.Dictionary Tests</h1>
    
    <%
    ' Create a dictionary object
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    ' Test 1: Add items
    Response.Write("<h2>Test 1: Add items</h2>")
    dict.Add "name", "John"
    dict.Add "age", 30
    dict.Add "city", "New York"
    Response.Write("Added 3 items<br>")
    
    ' Test 2: Count
    Response.Write("<h2>Test 2: Count</h2>")
    Response.Write("Dictionary count: " & dict.Count & "<br>")
    
    ' Test 3: Exists
    Response.Write("<h2>Test 3: Exists</h2>")
    Response.Write("Key 'name' exists: " & dict.Exists("name") & "<br>")
    Response.Write("Key 'notexist' exists: " & dict.Exists("notexist") & "<br>")
    
    ' Test 4: Item
    Response.Write("<h2>Test 4: Item retrieval</h2>")
    Response.Write("dict.Item('name'): " & dict.Item("name") & "<br>")
    Response.Write("dict.Item('age'): " & dict.Item("age") & "<br>")
    Response.Write("dict.Item('city'): " & dict.Item("city") & "<br>")
    
    ' Test 5: Keys
    Response.Write("<h2>Test 5: Keys</h2>")
    keys = dict.Keys()
    Response.Write("Keys: " & Join(keys, ", ") & "<br>")
    
    ' Test 6: Items
    Response.Write("<h2>Test 6: Items</h2>")
    items = dict.Items()
    Response.Write("Items: " & Join(items, ", ") & "<br>")
    
    ' Test 7: Remove
    Response.Write("<h2>Test 7: Remove</h2>")
    dict.Remove "city"
    Response.Write("Removed 'city' key<br>")
    Response.Write("Dictionary count after remove: " & dict.Count & "<br>")
    Response.Write("'city' exists: " & dict.Exists("city") & "<br>")
    
    ' Test 8: For Each
    Response.Write("<h2>Test 8: For Each Loop</h2>")
    Response.Write("Items via For Each:<br>")
    For Each key In dict
        Response.Write("  Key: " & key & " = " & dict.Item(key) & "<br>")
    Next
    
    ' Test 9: RemoveAll
    Response.Write("<h2>Test 9: RemoveAll</h2>")
    dict.RemoveAll()
    Response.Write("Dictionary count after RemoveAll: " & dict.Count & "<br>")
    
    ' Test 10: Multiple types
    Response.Write("<h2>Test 10: Multiple data types</h2>")
    dict.Add "string", "hello"
    dict.Add "number", 42
    dict.Add "float", 3.14
    dict.Add "boolean", True
    Response.Write("Added mixed types<br>")
    Response.Write("Count: " & dict.Count & "<br>")
    For Each key In dict
        Response.Write("  " & key & " = " & dict.Item(key) & "<br>")
    Next
    %>
    
    <p>All tests completed!</p>
</body>
</html>

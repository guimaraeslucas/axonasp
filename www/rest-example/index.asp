<%
  ' ====================================================================
  ' G3pix AxonASP REST API Front Controller
  ' ====================================================================
  ' This demonstrates a RESTful API implementation using AxonASP
  ' Features: GET, POST, PUT, DELETE methods
  ' Response formats: JSON, HTML
  ' ====================================================================
  
  Option Explicit
  
  Dim JSON, Response_Data
  Dim method, route, format, version
  Dim resource, resource_id
  
  Set JSON = Server.CreateObject("G3JSON")
  
  ' Get HTTP method
  method = Request.ServerVariables("REQUEST_METHOD")
  
  ' Get Accept header to determine response format (JSON or HTML)
  ' Default to JSON
  format = "json"
  If InStr(Request.ServerVariables("HTTP_ACCEPT"), "text/html") > 0 And _
     InStr(Request.ServerVariables("HTTP_ACCEPT"), "application/json") = 0 Then
    format = "html"
  End If
  
  ' Get response format from query string if provided
  If Not AxEmpty(Request.QueryString("format")) Then
    format = Request.QueryString("format")
  End If
  
  ' Parse route from rewrite rule (passed as query parameter)
  route = Request.QueryString("route")
  version = Request.QueryString("v")
  
  ' Parse resource and ID from route
  ' Format: /rest/users or /rest/users/123 or /rest/users/123/orders
  ParseRoute route, resource, resource_id
  
  ' Route to appropriate handler
  Select Case LCase(resource)
    Case "users"
      HandleUsers method, resource_id, format
    Case "products"
      HandleProducts method, resource_id, format
    Case "items"
      HandleItems method, resource_id, format
    Case "status"
      HandleStatus format
    Case Else
      SendErrorResponse 404, "Resource not found: " & resource, format
  End Select
  
  ' ====================================================================
  ' Handler: Users Resource
  ' ====================================================================
  Sub HandleUsers(httpMethod, id, outputFormat)
    Dim response_obj, users_data, user_obj, i
    
    Select Case UCase(httpMethod)
      Case "GET"
        If AxEmpty(id) Then
          ' GET /rest/users - List all users
          Set response_obj = JSON.NewObject()
          response_obj("status") = "success"
          response_obj("count") = 3
          
          Dim users_array
          Set users_array = JSON.NewArray()
          
          Dim user1
          Set user1 = JSON.NewObject()
          user1("id") = 1
          user1("name") = "Alice Johnson"
          user1("email") = "alice@example.com"
          user1("role") = "admin"
          user1("created_at") = AxDate("Y-m-d H:i:s")
          users_array.Add user1
          
          Dim user2
          Set user2 = JSON.NewObject()
          user2("id") = 2
          user2("name") = "Bob Smith"
          user2("email") = "bob@example.com"
          user2("role") = "user"
          user2("created_at") = AxDate("Y-m-d H:i:s")
          users_array.Add user2
          
          Dim user3
          Set user3 = JSON.NewObject()
          user3("id") = 3
          user3("name") = "Carol Williams"
          user3("email") = "carol@example.com"
          user3("role") = "moderator"
          user3("created_at") = AxDate("Y-m-d H:i:s")
          users_array.Add user3
          
          response_obj("data") = users_array
          
          SendResponse 200, response_obj, outputFormat
        Else
          ' GET /rest/users/123 - Get specific user
          Set response_obj = JSON.NewObject()
          response_obj("status") = "success"
          
          Dim user
          Set user = JSON.NewObject()
          user("id") = CLng(id)
          user("name") = "User #" & id
          user("email") = "user" & id & "@example.com"
          user("role") = "user"
          user("phone") = "+55 11 98765-4321"
          user("address") = "123 Main Street, São Paulo, Brasil"
          user("created_at") = AxDate("Y-m-d H:i:s")
          user("last_login") = AxDate("Y-m-d H:i:s", CDate(Now) - 2)
          
          response_obj("data") = user
          
          SendResponse 200, response_obj, outputFormat
        End If
        
      Case "POST"
        ' POST /rest/users - Create new user
        Dim post_data, new_user
        post_data = GetJSONBody()
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "User created"
        
        Set new_user = JSON.NewObject()
        new_user("id") = AxRand(100, 999)
        new_user("name") = post_data("name")
        new_user("email") = post_data("email")
        new_user("role") = post_data("role")
        new_user("created_at") = AxDate("Y-m-d H:i:s")
        
        response_obj("data") = new_user
        
        Response.Status = "201 Created"
        SendResponse 201, response_obj, outputFormat
        
      Case "PUT"
        ' PUT /rest/users/123 - Update user
        If AxEmpty(id) Then
          SendErrorResponse 400, "User ID required for PUT", outputFormat
          Exit Sub
        End If
        
        Dim put_data, updated_user
        put_data = GetJSONBody()
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "User updated"
        
        Set updated_user = JSON.NewObject()
        updated_user("id") = CLng(id)
        updated_user("name") = put_data("name")
        updated_user("email") = put_data("email")
        updated_user("role") = put_data("role")
        updated_user("updated_at") = AxDate("Y-m-d H:i:s")
        
        response_obj("data") = updated_user
        
        SendResponse 200, response_obj, outputFormat
        
      Case "DELETE"
        ' DELETE /rest/users/123 - Delete user
        If AxEmpty(id) Then
          SendErrorResponse 400, "User ID required for DELETE", outputFormat
          Exit Sub
        End If
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "User with ID " & id & " deleted successfully"
        response_obj("deleted_id") = CLng(id)
        
        SendResponse 200, response_obj, outputFormat
        
      Case Else
        SendErrorResponse 405, "Method " & httpMethod & " not allowed", outputFormat
    End Select
  End Sub
  
  ' ====================================================================
  ' Handler: Products Resource
  ' ====================================================================
  Sub HandleProducts(httpMethod, id, outputFormat)
    Dim response_obj, product
    
    Select Case UCase(httpMethod)
      Case "GET"
        If AxEmpty(id) Then
          ' GET /rest/products - List all products
          Set response_obj = JSON.NewObject()
          response_obj("status") = "success"
          
          Dim products_array
          Set products_array = JSON.NewArray()
          
          Dim p1
          Set p1 = JSON.NewObject()
          p1("id") = 1
          p1("name") = "MacBook Pro 16 inch"
          p1("price") = 2499.99
          p1("currency") = "USD"
          p1("stock") = 15
          p1("sku") = "APPLE-MBP-16-2024"
          p1("category") = "Electronics"
          products_array.Add p1
          
          Dim p2
          Set p2 = JSON.NewObject()
          p2("id") = 2
          p2("name") = "USB-C Hub"
          p2("price") = 79.99
          p2("currency") = "USD"
          p2("stock") = 150
          p2("sku") = "ADAPTOR-USB-C-HUB"
          p2("category") = "Accessories"
          products_array.Add p2
          
          Dim p3
          Set p3 = JSON.NewObject()
          p3("id") = 3
          p3("name") = "Wireless Keyboard"
          p3("price") = 149.99
          p3("currency") = "USD"
          p3("stock") = 45
          p3("sku") = "KEYS-WIRELESS-PRO"
          p3("category") = "Peripherals"
          products_array.Add p3
          
          response_obj("data") = products_array
          response_obj("count") = 3
          response_obj("formatted_total") = AxNumberFormat(2499.99 + 79.99 + 149.99, 2, ".", ",")
          
          SendResponse 200, response_obj, outputFormat
        Else
          ' GET /rest/products/123 - Get specific product
          Set response_obj = JSON.NewObject()
          response_obj("status") = "success"
          
          Set product = JSON.NewObject()
          product("id") = CLng(id)
          product("name") = "Product #" & id
          product("price") = AxRand(10, 500)
          product("discount_percent") = AxRand(0, 50)
          product("stock") = AxRand(0, 200)
          product("in_stock") = (product("stock") > 0)
          product("description") = "High-quality product with excellent features and warranty"
          product("rating") = 4.5
          product("reviews") = 127
          
          response_obj("data") = product
          
          SendResponse 200, response_obj, outputFormat
        End If
        
      Case "POST"
        ' POST /rest/products - Create new product
        Dim post_product_data
        post_product_data = GetJSONBody()
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "Product created"
        
        Set product = JSON.NewObject()
        product("id") = AxRand(1000, 9999)
        product("name") = post_product_data("name")
        product("price") = post_product_data("price")
        product("stock") = post_product_data("stock")
        product("created_at") = AxDate("Y-m-d H:i:s")
        
        response_obj("data") = product
        
        Response.Status = "201 Created"
        SendResponse 201, response_obj, outputFormat
        
      Case "PUT"
        ' PUT /rest/products/123 - Update product
        If AxEmpty(id) Then
          SendErrorResponse 400, "Product ID required for PUT", outputFormat
          Exit Sub
        End If
        
        Dim put_product_data
        put_product_data = GetJSONBody()
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "Product updated"
        
        Set product = JSON.NewObject()
        product("id") = CLng(id)
        product("name") = put_product_data("name")
        product("price") = put_product_data("price")
        product("stock") = put_product_data("stock")
        product("updated_at") = AxDate("Y-m-d H:i:s")
        
        response_obj("data") = product
        
        SendResponse 200, response_obj, outputFormat
        
      Case "DELETE"
        ' DELETE /rest/products/123 - Delete product
        If AxEmpty(id) Then
          SendErrorResponse 400, "Product ID required for DELETE", outputFormat
          Exit Sub
        End If
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "Product " & id & " deleted successfully"
        response_obj("deleted_id") = CLng(id)
        
        SendResponse 200, response_obj, outputFormat
        
      Case Else
        SendErrorResponse 405, "Method " & httpMethod & " not allowed", outputFormat
    End Select
  End Sub
  
  ' ====================================================================
  ' Handler: Items Resource (with advanced examples)
  ' ====================================================================
  Sub HandleItems(httpMethod, id, outputFormat)
    Dim response_obj, items_array, i, item
    
    Select Case UCase(httpMethod)
      Case "GET"
        If AxEmpty(id) Then
          ' GET /rest/items - List items with advanced Ax functions
          Set response_obj = JSON.NewObject()
          response_obj("status") = "success"
          response_obj("timestamp") = AxTime()
          response_obj("formatted_date") = AxDate("d/m/Y H:i:s")
          
          Set items_array = JSON.NewArray()
          
          ' Create sample items using Ax functions
          Dim item_names, descriptions
          item_names = AxExplode("|", "Notebook|Monitor|Mouse|Keyboard|Headset")
          descriptions = Array("Powerful laptop", "4K display", "Ergonomic mouse", "Mechanical keyboard", "Noise-cancelling headset")
          
          For i = 0 To AxCount(item_names) - 1
            Set item = JSON.NewObject()
            item("id") = i + 1
            item("name") = AxTrim(item_names(i))
            item("description") = descriptions(i)
            item("priority") = AxRand(1, 5)
            item("is_active") = (AxRand(0, 1) = 1)
            item("word_count") = AxWordCount(descriptions(i), 0)
            item("created_at") = AxDate("Y-m-d")
            items_array.Add item
          Next
          
          response_obj("data") = items_array
          response_obj("total_items") = AxCount(items_array)
          
          SendResponse 200, response_obj, outputFormat
        Else
          ' GET /rest/items/123 - Get specific item with encoding examples
          Set response_obj = JSON.NewObject()
          response_obj("status") = "success"
          
          Set item = JSON.NewObject()
          item("id") = CLng(id)
          item("name") = "Item #" & id & " - Special Characters: <>&"
          item("description") = "Item with HTML encoded: " & AxHtmlSpecialChars("<script>alert('xss')</script>")
          item("base64_encoded") = AxBase64Encode("Secret data for item " & id)
          item("hash_md5") = AxMd5("item_" & id)
          item("random_token") = AxHash("sha256", "token_" & AxTime())
          item("formatted_number") = AxNumberFormat(1234567.89, 2, ".", ",")
          
          response_obj("data") = item
          
          SendResponse 200, response_obj, outputFormat
        End If
        
      Case "POST"
        ' POST /rest/items - Create item with validation
        Dim post_item_data, error_obj
        post_item_data = GetJSONBody()
        
        ' Validate using Ax functions
        If AxEmpty(post_item_data("name")) Or Not AxCTypeAlpha(Left(post_item_data("name"), 1)) Then
          SendErrorResponse 422, "Item name is required and must start with letter", outputFormat
          Exit Sub
        End If
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "Item created"
        
        Set item = JSON.NewObject()
        item("id") = AxRand(10000, 99999)
        item("name") = AxUcFirst(post_item_data("name"))
        item("created_at") = AxDate("Y-m-d H:i:s")
        item("validation_passed") = True
        
        response_obj("data") = item
        
        Response.Status = "201 Created"
        SendResponse 201, response_obj, outputFormat
        
      Case "PUT"
        ' PUT /rest/items/123 - Update item
        If AxEmpty(id) Then
          SendErrorResponse 400, "Item ID required for PUT", outputFormat
          Exit Sub
        End If
        
        Dim put_item_data
        put_item_data = GetJSONBody()
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "Item updated"
        
        Set item = JSON.NewObject()
        item("id") = CLng(id)
        item("name") = put_item_data("name")
        item("updated_at") = AxDate("Y-m-d H:i:s")
        item("padded_id") = AxPad(CStr(id), 5, "0", 0)
        
        response_obj("data") = item
        
        SendResponse 200, response_obj, outputFormat
        
      Case "DELETE"
        ' DELETE /rest/items/123 - Delete item
        If AxEmpty(id) Then
          SendErrorResponse 400, "Item ID required for DELETE", outputFormat
          Exit Sub
        End If
        
        Set response_obj = JSON.NewObject()
        response_obj("status") = "success"
        response_obj("message") = "Item " & id & " successfully deleted"
        response_obj("deleted_id") = CLng(id)
        response_obj("deletion_time") = AxDate("Y-m-d H:i:s")
        
        SendResponse 200, response_obj, outputFormat
        
      Case Else
        SendErrorResponse 405, "Method " & httpMethod & " not allowed", outputFormat
    End Select
  End Sub
  
  ' ====================================================================
  ' Handler: Status/Health Check
  ' ====================================================================
  Sub HandleStatus(outputFormat)
    Dim response_obj, system_info
    
    Set response_obj = JSON.NewObject()
    response_obj("status") = "ok"
    response_obj("timestamp") = AxDate("Y-m-d H:i:s")
    response_obj("service") = "G3pix AxonASP REST API"
    response_obj("version") = "1.0.0"
    response_obj("uptime_seconds") = CLng(AxTime())
    
    Set system_info = JSON.NewObject()
    system_info("server_name") = Request.ServerVariables("SERVER_NAME")
    system_info("server_port") = Request.ServerVariables("SERVER_PORT")
    system_info("timestamp_unix") = AxTime()
    system_info("date_formatted") = AxDate("d/m/Y H:i:s")
    
    response_obj("system") = system_info
    
    SendResponse 200, response_obj, outputFormat
  End Sub
  
  ' ====================================================================
  ' Utility Functions
  ' ====================================================================
  
  Sub ParseRoute(routeString, byRef resource, byRef resourceId)
    Dim parts
    
    ' Remove leading/trailing slashes
    routeString = AxTrim(routeString, "/")
    
    If AxEmpty(routeString) Then
      resource = ""
      resourceId = ""
      Exit Sub
    End If
    
    ' Split by /
    parts = AxExplode("/", routeString)
    
    If AxCount(parts) > 0 Then
      resource = parts(0)
    End If
    
    If AxCount(parts) > 1 Then
      resourceId = parts(1)
    End If
  End Sub
  
  Sub SendResponse(statusCode, responseObj, outputFormat)
    Response.Status = statusCode & " OK"
    
    If LCase(outputFormat) = "html" Then
      SendHTMLResponse statusCode, responseObj
    Else
      SendJSONResponse statusCode, responseObj
    End If
  End Sub
  
  Sub SendJSONResponse(statusCode, jsonData)
    Response.ContentType = "application/json"
    Response.Charset = "utf-8"
    Response.Write JSON.Stringify(jsonData)
  End Sub
  
  Sub SendHTMLResponse(statusCode, dataObj)
    Response.ContentType = "text/html"
    Response.Charset = "utf-8"
    
    Response.Write "<!DOCTYPE html>" & vbCrLf
    Response.Write "<html>" & vbCrLf
    Response.Write "<head>" & vbCrLf
    Response.Write "  <meta charset='utf-8'>" & vbCrLf
    Response.Write "  <title>G3pix AxonASP REST API</title>" & vbCrLf
    Response.Write "  <style>" & vbCrLf
    Response.Write "    body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }" & vbCrLf
    Response.Write "    .container { max-width: 900px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }" & vbCrLf
    Response.Write "    .status-success { color: #28a745; font-weight: bold; }" & vbCrLf
    Response.Write "    .status-error { color: #dc3545; font-weight: bold; }" & vbCrLf
    Response.Write "    pre { background: #f4f4f4; padding: 15px; border-radius: 4px; overflow-x: auto; border-left: 4px solid #007bff; }" & vbCrLf
    Response.Write "    .header { border-bottom: 2px solid #007bff; padding-bottom: 10px; margin-bottom: 20px; }" & vbCrLf
    Response.Write "    h1 { margin: 0; color: #007bff; }" & vbCrLf
    Response.Write "    .info { color: #666; font-size: 12px; }" & vbCrLf
    Response.Write "  </style>" & vbCrLf
    Response.Write "</head>" & vbCrLf
    Response.Write "<body>" & vbCrLf
    Response.Write "  <div class='container'>" & vbCrLf
    Response.Write "    <div class='header'>" & vbCrLf
    Response.Write "      <h1>G3pix AxonASP REST API</h1>" & vbCrLf
    Response.Write "      <p class='info'>Status Code: " & statusCode & " | " & AxDate("d/m/Y H:i:s") & "</p>" & vbCrLf
    Response.Write "    </div>" & vbCrLf
    Response.Write "    <h2>Response Data</h2>" & vbCrLf
    Response.Write "    <pre>" & AxHtmlSpecialChars(JSON.Stringify(dataObj)) & "</pre>" & vbCrLf
    Response.Write "    <hr>" & vbCrLf
    Response.Write "    <h3>Request Information</h3>" & vbCrLf
    Response.Write "    <ul>" & vbCrLf
    Response.Write "      <li><strong>Method:</strong> " & Request.ServerVariables("REQUEST_METHOD") & "</li>" & vbCrLf
    Response.Write "      <li><strong>Path:</strong> " & Request.ServerVariables("REQUEST_URI") & "</li>" & vbCrLf
    Response.Write "      <li><strong>IP Address:</strong> " & Request.ServerVariables("REMOTE_ADDR") & "</li>" & vbCrLf
    Response.Write "    </ul>" & vbCrLf
    Response.Write "  </div>" & vbCrLf
    Response.Write "</body>" & vbCrLf
    Response.Write "</html>" & vbCrLf
  End Sub
  
  Sub SendErrorResponse(statusCode, errorMessage, outputFormat)
    Dim error_obj
    
    Set error_obj = JSON.NewObject()
    error_obj("status") = "error"
    error_obj("code") = statusCode
    error_obj("message") = errorMessage
    error_obj("timestamp") = AxDate("Y-m-d H:i:s")
    
    Select Case statusCode
      Case 400
        Response.Status = "400 Bad Request"
      Case 404
        Response.Status = "404 Not Found"
      Case 405
        Response.Status = "405 Method Not Allowed"
      Case 422
        Response.Status = "422 Unprocessable Entity"
      Case Else
        Response.Status = statusCode & " Error"
    End Select
    
    SendResponse statusCode, error_obj, outputFormat
  End Sub
  
  Function GetJSONBody()
    Dim G3JSON, body_content, request_length
    Dim result_obj
    
    Set G3JSON = Server.CreateObject("G3JSON")
    Set result_obj = G3JSON.NewObject()
    
    ' Get content length
    request_length = Request.ContentLength
    
    If request_length = 0 Then
      Set GetJSONBody = result_obj
      Exit Function
    End If
    
    ' Read body content using Request.BinaryRead or similar
    ' For POST/PUT JSON, try to parse the Request body
    On Error Resume Next
    
    ' Attempt to get raw body - this depends on AxonASP implementation
    ' Alternative: use Application state or temporary storage
    
    ' For this example, we'll parse from POST form fields
    ' In a real scenario, you'd read the raw request body
    Dim key, value
    Dim form_items
    
    For Each key In Request.Form
      value = Request.Form(key)
      result_obj(key) = value
    Next
    
    On Error Goto 0
    
    Set GetJSONBody = result_obj
  End Function
%>

# Request Method

## Overview
Sends an HTTP request to a remote server and returns the response body, providing functionality identical to the **Fetch** method.

## Syntax
```asp
result = http.Request(url [, method] [, body])
```

## Parameters and Arguments
- **url** (String, Required): The absolute URL for the request.
- **method** (String, Optional): The HTTP verb (e.g., "GET", "POST", "PUT", "DELETE"). The default is "GET".
- **body** (String, Optional): The request body payload. If provided, the library automatically sets the `Content-Type` header to `application/json`.

## Return Values
Returns a **Variant** containing the response from the remote server. 
- Returns a **Scripting.Dictionary** if the response `Content-Type` is `application/json` and the root element is an object.
- Returns an **Array** if the response `Content-Type` is `application/json` and the root element is an array.
- Returns a **String** for all other response types.
- Returns **Empty** if the request fails (e.g., network error or timeout).

## Remarks
The **Request** method is an alias for the **Fetch** method. For additional details on its operation, please refer to the documentation for the **Fetch** method.

## Code Example
The following example demonstrates how to use the **Request** method to query a remote API.

```asp
<%
Dim http, apiUrl, result
Set http = Server.CreateObject("G3HTTP")

apiUrl = "https://api.example.com/health"

' Perform an HTTP GET request
result = http.Request(apiUrl)

If Not IsEmpty(result) Then
    Response.Write "API Health: " & result
End If

Set http = Nothing
%>
```

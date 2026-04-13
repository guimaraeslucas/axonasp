# Fetch Method

## Overview
Sends an HTTP request to a remote server and returns the response body, automatically parsing JSON content into native AxonASP objects or arrays.

## Syntax
```asp
result = http.Fetch(url [, method] [, body])
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
- The G3HTTP library uses a default timeout of 10 seconds.
- JSON parsing is performed automatically for all `application/json` responses.

## Code Example
The following example demonstrates how to send a POST request with a JSON body.

```asp
<%
Dim http, postData, responseBody, apiUrl
Set http = Server.CreateObject("G3HTTP")

apiUrl = "https://api.example.com/v1/update"
postData = "{""id"": 123, ""status"": ""active""}"

' Perform a POST request
responseBody = http.Fetch(apiUrl, "POST", postData)

If IsObject(responseBody) Then
    Response.Write "Response ID: " & responseBody("id")
ElseIf Not IsEmpty(responseBody) Then
    Response.Write "Raw Response: " & responseBody
Else
    Response.Write "Request failed."
End If

Set http = Nothing
%>
```

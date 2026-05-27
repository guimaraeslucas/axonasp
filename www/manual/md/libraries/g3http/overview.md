# Use the G3HTTP Library

## Overview
Use **G3HTTP** to execute outbound HTTP requests from G3Pix AxonASP scripts. The library supports configurable HTTP methods, optional request bodies, and automatic JSON parsing when the response content type is JSON.

## Prerequisites
- Outbound network access to the destination host.
- Valid TLS trust configuration for HTTPS endpoints.
- Create the object with the primary ProgID:

```asp
Dim http
Set http = Server.CreateObject("G3HTTP")
```
```javascript
var http = Server.CreateObject("G3HTTP");
```

## How It Works
- `Fetch` and `Request` share the same implementation.
- Default HTTP method is `GET` when not provided.
- When a request body is provided, `Content-Type` is set to `application/json`.
- Requests use a 10-second timeout.
- If response `Content-Type` contains `application/json`, the body is parsed through G3JSON.
- If JSON parsing fails, the raw response body is returned as text.

## API Reference

### Methods
- **Fetch(url[, method][, body])**: Sends an HTTP request and returns parsed JSON or text.
- **Request(url[, method][, body])**: Alias of `Fetch`.

### Properties
G3HTTP does not expose public properties.

## Example
```asp
<%
Dim http, result
Set http = Server.CreateObject("G3HTTP")

result = http.Fetch("https://api.example.com/status")

If IsObject(result) Then
    Response.Write result("status")
ElseIf IsArray(result) Then
    Response.Write "Array response"
ElseIf Not IsEmpty(result) Then
    Response.Write result
Else
    Response.Write "Request failed"
End If

Set http = Nothing
%>
```

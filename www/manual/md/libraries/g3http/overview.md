# Use the G3HTTP Library

## Overview
The **G3HTTP** library provides a streamlined interface for making outbound HTTP requests from G3Pix AxonASP applications. It enables the consumption of remote REST APIs and web services by abstracting the complexities of network communication. The library automatically detects JSON responses and parses them into native AxonASP data structures, such as Dictionaries and Arrays, for immediate use in server-side scripts.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set http = Server.CreateObject("G3HTTP")
```

## Prerequisites
- **Network Connectivity**: The server hosting AxonASP must have outbound access to the target URLs.
- **SSL/TLS Certificates**: For HTTPS requests, ensure the server trusts the certificate authority of the remote endpoint.

## How it Works
The G3HTTP object utilizes the high-performance network stack of the AxonASP engine. When a request is initiated via the **Fetch** or **Request** methods, the library performs a synchronous operation with a default timeout of 10 seconds. 

If the remote server responds with a `Content-Type` of `application/json`, the library automatically passes the payload to the internal JSON engine. This results in a **Scripting.Dictionary** (for JSON objects) or a standard **Array** (for JSON arrays). All other content types are returned as a plain **String**.

## API Reference

### Methods
- **Fetch**: Performs an HTTP request and returns the response body, automatically parsing JSON content.
- **Request**: An alias for the **Fetch** method, providing identical functionality.

## Code Example
The following example demonstrates how to perform a GET request to a JSON API and access its data.

```asp
<%
Dim http, apiData, url
Set http = Server.CreateObject("G3HTTP")

url = "https://api.example.com/v1/status"

' Perform the request
Set apiData = http.Fetch(url)

' Check if the request returned a valid object
If IsObject(apiData) Then
    Response.Write "API Status: " & apiData("status") & "<br>"
    Response.Write "Server Time: " & apiData("server_time")
ElseIf Not IsEmpty(apiData) Then
    Response.Write "Raw Response: " & apiData
Else
    Response.Write "Request failed."
End If

Set http = Nothing
%>
```

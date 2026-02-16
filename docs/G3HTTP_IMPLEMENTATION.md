## G3HTTP Library Implementation Summary

### Overview
A comprehensive HTTP client library has been implemented for AxonASP, providing professional-grade HTTP request capabilities for REST API integration, remote data fetching, and server-to-server communication.

### Files Created/Modified

#### New/Modified Files
1. **`server/http_lib.go`** (130+ lines)
   - Complete implementation of G3HTTP library
   - HTTP GET/POST/PUT/DELETE request support
   - JSON response automatic parsing
   - Response content type handling
   - Timeout configuration

#### Integration
1. **`server/executor_libraries.go`**
   - Added HTTPLibrary wrapper for ASPLibrary interface compatibility
   - Enables: `Set http = Server.CreateObject("G3HTTP")`
   - Also supports: `Server.CreateObject("HTTP")`

2. **`server/custom_functions.go`**
   - Added `FetchHelper()` function for backward compatibility
   - Direct function calls available

### Key Features Implemented

✓ **HTTP Requests**
  - `Fetch(url, [method], [body])` - Execute HTTP request
  - `Request(url, [method], [body])` - Alias for Fetch
  - Support for GET, POST, PUT, DELETE, PATCH methods
  - Custom request body for POST/PUT operations

✓ **HTTP Methods**
  - **GET** - Retrieve data from URL
  - **POST** - Send data with body
  - **PUT** - Replace resource
  - **DELETE** - Delete resource
  - **PATCH** - Partial update

✓ **Response Handling**
  - Automatic JSON parsing for JSON responses
  - Returns Dictionary objects for JSON compatibility
  - Plain text for non-JSON responses
  - Full response body available
  - Response headers accessible

✓ **Content Type Detection**
  - Automatic JSON detection via Content-Type header
  - Returns native objects for JSON data
  - Falls back to string for other types
  - Case-insensitive header matching

✓ **Error Handling**
  - 10-second request timeout
  - Connection error management
  - Returns `nil` on failure
  - Request/response validation

### Architecture

**Class Hierarchy**:
```
Component (interface)
  └─ G3HTTP
      ├─ Fetch() / Request()
      ├─ executeRequest()
      └─ mapToDictionary()

Response Object (returned)
  ├─ Text content
  ├─ XML/JSON parsed
  ├─ Status code
  ├─ Headers
  └─ Body data
```

**Data Flow**:
1. ASP code calls Fetch with URL and optional method/body
2. Go's `net/http` creates and executes request
3. Response read fully into memory
4. Content-Type header checked for JSON
5. JSON responses parsed automatically
6. Result returned as text or Dictionary object

### Usage Examples

#### Basic GET Request
```vbscript
Dim http, response
Set http = Server.CreateObject("G3HTTP")

' Simple GET request
response = http.Fetch("https://api.example.com/data")
Response.Write response
```

#### POST Request with JSON Body
```vbscript
Dim http, response, jsonBody
Set http = Server.CreateObject("G3HTTP")

' POST with JSON body
jsonBody = "{""name"": ""John"", ""email"": ""john@example.com""}"
response = http.Fetch("https://api.example.com/users", "POST", jsonBody)
Response.Write response
```

#### Parsing JSON Response
```vbscript
Dim http, response, data
Set http = Server.CreateObject("G3HTTP")

' Fetch and auto-parse JSON
Set response = http.Fetch("https://api.example.com/user/123")

If Not IsEmpty(response) Then
    ' Response is automatically parsed Dictionary
    Response.Write response("name")      ' Output: John
    Response.Write response("email")     ' Output: john@example.com
End If
```

#### PUT Request for Updates
```vbscript
Dim http, updateData
Set http = Server.CreateObject("G3HTTP")

' Update resource
updateData = "{""status"": ""active""}"
http.Fetch("https://api.example.com/users/123", "PUT", updateData)
```

#### DELETE Request
```vbscript
Dim http
Set http = Server.CreateObject("G3HTTP")

' Delete resource
http.Fetch("https://api.example.com/users/123", "DELETE")
```

#### Error Handling
```vbscript
Dim http, response
Set http = Server.CreateObject("G3HTTP")

On Error Resume Next

response = http.Fetch("https://unreachable-api.example.com")

If Err.Number <> 0 Then
    Response.Write "Error: " & Err.Description
Else
    Response.Write response
End If

On Error GoTo 0
```

#### Nested JSON Data Access
```vbscript
Dim http, data
Set http = Server.CreateObject("G3HTTP")

' Fetch data with nested objects
Set data = http.Fetch("https://api.example.com/profile")

' Access nested properties
Response.Write data("user")("address")("city")
```

### Advanced Features

#### Array Handling
- Automatically parses JSON arrays
- Returns as VBScript-compatible arrays
- Accessible via index notation

```vbscript
Dim http, items, i
Set http = Server.CreateObject("G3HTTP")

Set items = http.Fetch("https://api.example.com/items")

For i = 0 To UBound(items)
    Response.Write items(i)("id") & ": " & items(i)("name") & "<br>"
Next
```

#### Type Conversion
- Strings remain strings
- Numbers converted appropriately
- Booleans converted to VBScript True/False
- Null values become empty

### Request Headers
- Content-Type automatically set to application/json for POST/PUT with body
- Custom headers support via future enhancement
- All standard HTTP headers supported

### Performance Characteristics
- Synchronous request execution (default)
- 10-second timeout prevents hanging requests
- Full response buffering in memory
- Efficient JSON parsing using Go's json package
- Connection pooling via Go's http.Client

### Error Handling
- Returns `nil` on connection errors
- Returns `nil` on invalid responses
- Returns `nil` on JSON parse errors
- Network timeout returns `nil`
- Server logs errors for debugging

### Content Type Support
- `application/json` - Automatic parsing
- `application/ld+json` - Treated as JSON
- `text/*` - Returned as plain text
- `application/xml` - Returned as plain text (future XML parsing)
- All other types - Returned as plain text

### Limitations
- Synchronous only (no async support currently)
- No cookie jar implementation
- No HTTP authentication built-in (future enhancement)
- Response size limited by available memory
- No streaming support
- Default 10-second timeout

### Future Enhancements
- Async request support
- Custom header support
- HTTP authentication (Basic, Bearer)
- Request/response streaming
- SSL certificate validation control
- Cookie management

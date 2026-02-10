# G3pix AxonASP REST API Implementation

## Overview

This document describes the REST API implementation integrated into the AxonASP server. The REST API demonstrates a production-ready pattern for building RESTful web services using VBScript within the AxonASP framework.

## Architecture

### Front Controller Pattern

The API uses a front-controller pattern where:
- All requests matching `/rest/*` are routed to `/rest/index.asp`
- The `route` query parameter contains the parsed URL path
- Individual handler subroutines process each resource type

### Routing

**web.config Rule:**
```xml
<rule name="RestAPI" stopProcessing="true">
  <match url="^rest/(.*)$" ignoreCase="true" />
  <conditions>
    <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
    <add input="{REQUEST_FILENAME}" matchType="IsDirectory" negate="true" />
  </conditions>
  <action type="Rewrite" url="/rest/index.asp?route={R:1}" appendQueryString="true" />
</rule>
```

Example rewrites:
- `GET /rest/users` → `/rest/index.asp?route=users`
- `POST /rest/users/123` → `/rest/index.asp?route=users/123`
- `PUT /rest/products/456?limit=10` → `/rest/index.asp?route=products/456&limit=10`

## HTTP Methods

The API supports four standard HTTP verbs:

### GET
- **Purpose:** Retrieve resources
- **Response Codes:** 200 (OK), 404 (Not Found)
- **Body:** No request body
- **Usage:** `GET /rest/users`, `GET /rest/users/123`

### POST
- **Purpose:** Create new resources
- **Response Codes:** 201 (Created), 400 (Bad Request), 422 (Unprocessable Entity)
- **Body:** JSON object with resource data
- **Usage:** `POST /rest/users` with JSON body

### PUT
- **Purpose:** Update existing resources
- **Response Codes:** 200 (OK), 404 (Not Found), 400 (Bad Request)
- **Body:** JSON object with updated data
- **Usage:** `PUT /rest/users/123` with JSON body
- **Note:** ID is required in URL

### DELETE
- **Purpose:** Remove resources
- **Response Codes:** 200 (OK), 404 (Not Found), 400 (Bad Request)
- **Body:** No request body
- **Usage:** `DELETE /rest/users/123`
- **Note:** ID is required in URL

## Response Formats

### JSON Format (Default)

**Success Response:**
```json
{
  "status": "success",
  "data": {
    "id": 1,
    "name": "Alice Johnson",
    "email": "alice@example.com"
  },
  "count": 1,
  "timestamp": "2024-01-16T14:30:45"
}
```

**Error Response:**
```json
{
  "status": "error",
  "code": 404,
  "message": "Resource not found",
  "timestamp": "2024-01-16T14:30:45"
}
```

**Request Format:**
```
Content-Type: application/json
Accept: application/json
```

### HTML Format

An HTML representation of the same data is available by adding `?format=html`:

```
GET /rest/users?format=html
```

**Response:**
- HTML page with formatted JSON displayed in a `<pre>` block
- Request information section
- Styled with CSS for readability

## Available Endpoints

### Users Resource

#### List Users
- **Method:** GET
- **URL:** `/rest/users`
- **Response:** Array of user objects
- **Status:** 200

Example Response:
```json
{
  "status": "success",
  "count": 3,
  "data": [
    {
      "id": 1,
      "name": "Alice Johnson",
      "email": "alice@example.com",
      "role": "admin",
      "created_at": "2024-01-16 10:30:45"
    },
    {
      "id": 2,
      "name": "Bob Smith",
      "email": "bob@example.com",
      "role": "user",
      "created_at": "2024-01-15 14:22:10"
    }
  ]
}
```

#### Get Single User
- **Method:** GET
- **URL:** `/rest/users/{id}`
- **Response:** Single user object with extended data
- **Status:** 200

#### Create User
- **Method:** POST
- **URL:** `/rest/users`
- **Request Body:**
```json
{
  "name": "John Doe",
  "email": "john@example.com",
  "role": "user"
}
```
- **Response:** 201 Created with new user object including ID

#### Update User
- **Method:** PUT
- **URL:** `/rest/users/{id}`
- **Request Body:**
```json
{
  "name": "Jane Doe",
  "email": "jane@example.com",
  "role": "admin"
}
```
- **Response:** 200 OK with updated object

#### Delete User
- **Method:** DELETE
- **URL:** `/rest/users/{id}`
- **Response:** 200 OK with acknowledgment

### Products Resource

#### List Products
- **Method:** GET
- **URL:** `/rest/products`
- **Response:** Array of product objects with pricing

#### Get Single Product
- **Method:** GET
- **URL:** `/rest/products/{id}`
- **Response:** Detailed product information

#### Create Product
- **Method:** POST
- **URL:** `/rest/products`
- **Request Body:**
```json
{
  "name": "Product Name",
  "price": 99.99,
  "stock": 50
}
```

#### Update Product
- **Method:** PUT
- **URL:** `/rest/products/{id}`

#### Delete Product
- **Method:** DELETE
- **URL:** `/rest/products/{id}`

**Special Features:**
- `AxNumberFormat()` used for price formatting (e.g., "1,234,567.89")
- `AxRand()` for generating price and stock data

### Items Resource (Advanced Examples)

This resource demonstrates advanced Ax function usage:

#### List Items
- **Method:** GET
- **URL:** `/rest/items`
- **Ax Functions Used:**
  - `AxExplode()` - Split item names
  - `AxWordCount()` - Count description words
  - `AxDate()` - Format dates
  - `AxTime()` - Get timestamp

#### Get Single Item
- **Method:** GET
- **URL:** `/rest/items/{id}`
- **Ax Functions Used:**
  - `AxHash()` - Generate SHA256 hash
  - `AxBase64Encode()` - Encode data
  - `AxMd5()` - MD5 hash
  - `AxHtmlSpecialChars()` - Escape HTML
  - `AxNumberFormat()` - Format numbers

#### Create Item
- **Method:** POST
- **URL:** `/rest/items`
- **Ax Functions Used:**
  - `AxUcFirst()` - Capitalize first character
  - `AxCTypeAlpha()` - Validate input

#### Update Item
- **Method:** PUT
- **URL:** `/rest/items/{id}`
- **Ax Functions Used:**
  - `AxPad()` - Pad integer to length

#### Delete Item
- **Method:** DELETE
- **URL:** `/rest/items/{id}`
- **Ax Functions Used:**
  - `AxDate()` - Return deletion time

### Status Resource

#### Health Check
- **Method:** GET
- **URL:** `/rest/status`
- **Response:** Server status and system information
- **Ax Functions Used:**
  - `AxTime()` - Unix timestamp
  - `AxDate()` - Formatted date
  - `AxNumberFormat()` - Format numbers

## Ax Custom Functions Used

The REST API leverages modern custom Ax functions for optimal code quality:

### String Functions
- **AxExplode()** - Split strings by delimiter
- **AxTrim()** - Remove whitespace
- **AxUcFirst()** - Uppercase first character
- **AxWordCount()** - Count words in text
- **AxHtmlSpecialChars()** - Escape HTML special characters
- **AxBase64Encode() / AxBase64Decode()** - Base64 encoding

### Array Functions
- **AxCount()** - Get array length
- **AxEmpty()** - Check if empty
- **AxExplode()** - Split string to array

### Hash & Encoding Functions
- **AxHash()** - SHA256 hashing
- **AxMd5()** - MD5 hashing
- **AxBase64Encode()** - Encode to Base64

### Number Functions
- **AxNumberFormat()** - Format numbers with separators
- **AxRand()** - Generate random numbers

### Date/Time Functions
- **AxDate()** - Format dates (PHP-like format strings)
- **AxTime()** - Get Unix timestamp

### Validation Functions
- **AxCTypeAlpha()** - Check if alphabetic
- **AxFilterValidateEmail()** - Validate email
- **AxFilterValidateIp()** - Validate IP address

### Utility Functions
- **AxPad()** - Pad strings/numbers

## JSON Handling

### Parsing JSON from Request Body

The API uses `GetJSONBody()` function to parse incoming JSON:

```vbscript
Function GetJSONBody()
  Dim G3JSON, result_obj
  Set G3JSON = Server.CreateObject("G3JSON")
  Set result_obj = G3JSON.NewObject()
  
  ' Parse form fields or raw JSON
  Dim key, value
  For Each key In Request.Form
    value = Request.Form(key)
    result_obj(key) = value
  Next
  
  Set GetJSONBody = result_obj
End Function
```

### Creating JSON Responses

The G3JSON library is used throughout:

```vbscript
Dim JSON, response_obj
Set JSON = Server.CreateObject("G3JSON")
Set response_obj = JSON.NewObject()

response_obj("status") = "success"
response_obj("data") = someData

Response.Write JSON.Stringify(response_obj)
```

## Error Handling

### HTTP Status Codes

| Code | Meaning | Usage |
|------|---------|-------|
| 200 | OK | Successful GET, PUT, DELETE |
| 201 | Created | Successful POST |
| 400 | Bad Request | Missing required ID for PUT/DELETE |
| 404 | Not Found | Unknown resource |
| 405 | Method Not Allowed | Unsupported HTTP method |
| 422 | Unprocessable Entity | Validation error |

### Error Response Structure

All errors return JSON with:
```json
{
  "status": "error",
  "code": 400,
  "message": "Descriptive error message",
  "timestamp": "2024-01-16 14:30:45"
}
```

## Usage Examples

### cURL Examples

#### GET List
```bash
curl -X GET "http://localhost:4050/rest/users?route=users"
```

#### GET Single
```bash
curl -X GET "http://localhost:4050/rest/users/1?route=users/1"
```

#### POST Create
```bash
curl -X POST "http://localhost:4050/rest/users?route=users" \
  -H "Content-Type: application/json" \
  -d '{"name":"John","email":"john@example.com","role":"user"}'
```

#### PUT Update
```bash
curl -X PUT "http://localhost:4050/rest/users/1?route=users/1" \
  -H "Content-Type: application/json" \
  -d '{"name":"Jane","email":"jane@example.com","role":"admin"}'
```

#### DELETE
```bash
curl -X DELETE "http://localhost:4050/rest/users/1?route=users/1"
```

#### HTML Format
```bash
curl -X GET "http://localhost:4050/rest/users/1?route=users/1&format=html"
```

### JavaScript / Fetch API Examples

#### GET Request
```javascript
fetch('/rest/users?route=users')
  .then(response => response.json())
  .then(data => console.log(data));
```

#### POST Request
```javascript
fetch('/rest/users?route=users', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    name: 'John Doe',
    email: 'john@example.com',
    role: 'user'
  })
})
  .then(response => response.json())
  .then(data => console.log(data));
```

#### PUT Request
```javascript
fetch('/rest/users/1?route=users/1', {
  method: 'PUT',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    name: 'Jane Doe',
    email: 'jane@example.com',
    role: 'admin'
  })
})
  .then(response => response.json())
  .then(data => console.log(data));
```

#### DELETE Request
```javascript
fetch('/rest/users/1?route=users/1', {
  method: 'DELETE'
})
  .then(response => response.json())
  .then(data => console.log(data));
```

## Code Structure

### Main Files

- **`/rest/index.asp`** - Front controller with all handlers
- **`/rest/README.asp`** - Interactive documentation UI
- **`web.config`** - URL rewrite rules

### Handler Organization

Each resource has dedicated handler subroutine:
- `HandleUsers()` - User endpoints
- `HandleProducts()` - Product endpoints
- `HandleItems()` - Item endpoints (advanced)
- `HandleStatus()` - Health check

### Utility Functions

- `ParseRoute()` - Extract resource and ID from route
- `SendResponse()` - Route to JSON or HTML response
- `SendJSONResponse()` - Serialize to JSON
- `SendHTMLResponse()` - Render as HTML
- `SendErrorResponse()` - Format error messages
- `GetJSONBody()` - Parse request JSON

## Best Practices Implemented

1. **RESTful Design**
   - Proper HTTP methods for actions
   - Resource-oriented URLs
   - Standard status codes

2. **Error Handling**
   - Comprehensive error responses
   - Appropriate status codes
   - Helpful error messages

3. **Data Validation**
   - Input validation using Ax functions
   - Email/IP validation where applicable
   - Type checking before processing

4. **Security**
   - XSS prevention with AxHtmlSpecialChars()
   - Proper status codes instead of info disclosure
   - Controlled error messages

5. **Performance**
   - Efficient Ax function usage
   - Minimal string allocations
   - Direct response streaming

6. **Flexibility**
   - Multiple response formats (JSON/HTML)
   - Query parameter parsing
   - Extensible handler pattern

## Extending the API

### Adding a New Resource

1. Create handler subroutine:
```vbscript
Sub HandleNewResource(httpMethod, id, outputFormat)
  Select Case UCase(httpMethod)
    Case "GET"
      ' Handle GET
    Case "POST"
      ' Handle POST
    ' ... other methods
  End Select
End Sub
```

2. Add case in main router:
```vbscript
Select Case LCase(resource)
  Case "newresource"
    HandleNewResource method, resource_id, format
  ' ... other resources
End Select
```

### Custom Validation

Create validation functions using Ax functions:
```vbscript
Function ValidateUser(userData)
  If AxEmpty(userData("name")) Then
    ValidateUser = "Name required"
    Exit Function
  End If
  
  If Not AxFilterValidateEmail(userData("email")) Then
    ValidateUser = "Invalid email format"
    Exit Function
  End If
  
  ValidateUser = ""  ' No errors
End Function
```

## Testing

### Interactive Documentation

Access at: `http://localhost:4050/rest/README.asp`

Provides:
- Formatted listing of all endpoints
- One-click test buttons for each endpoint
- cURL examples
- Format selector (JSON/HTML)

### Manual Testing

Use any HTTP client (curl, Postman, etc.) to test endpoints.

## Performance Considerations

- The front controller pattern minimizes overhead
- Ax functions are optimized for VBScript execution
- JSON serialization is handled by native G3JSON library
- Consider caching strategy for GET requests in production
- Add database lookup logic to handlers as needed

## Security Notes

1. **Input Validation:** Always validate request data before processing
2. **Authorization:** Add authentication headers/token verification
3. **Rate Limiting:** Consider implementing rate limits for production
4. **CORS Headers:** Add appropriate CORS headers for browser access
5. **SQL Injection:** Use parameterized queries if accessing databases
6. **XSS Prevention:** Always use AxHtmlSpecialChars() for user-supplied content

## Future Enhancements

- Database integration (ADODB)
- Authentication/Authorization layer
- Request/Response middleware pipeline
- Caching strategies
- API versioning (v1, v2, etc.)
- Pagination with limit/offset
- Sorting and filtering
- Field selection
- Related resource inclusion

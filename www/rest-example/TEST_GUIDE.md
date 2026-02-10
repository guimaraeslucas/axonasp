# G3pix AxonASP REST API - Test Guide

## Quick Start

The REST API is available at `/rest/` and provides a complete CRUD interface for demo resources.

### Access Points

1. **Interactive Documentation**: [/rest/README.asp](/rest/README.asp)
2. **Testing Console**: [/rest/test-console.html](/rest/test-console.html)
3. **Code Examples**: [/rest/code-examples.html](/rest/code-examples.html)
4. **API Front Controller**: [/rest/index.asp](/rest/index.asp)

---

## API Endpoints

### Users Resource

#### List All Users
```bash
curl -X GET "http://localhost:4050/rest/users?route=users"
```

#### Get Single User
```bash
curl -X GET "http://localhost:4050/rest/users/1?route=users/1"
```

#### Create User
```bash
curl -X POST "http://localhost:4050/rest/users?route=users" \
  -H "Content-Type: application/json" \
  -d '{
    "name": "John Doe",
    "email": "john@example.com",
    "role": "user"
  }'
```

#### Update User
```bash
curl -X PUT "http://localhost:4050/rest/users/1?route=users/1" \
  -H "Content-Type: application/json" \
  -d '{
    "name": "Jane Doe",
    "email": "jane@example.com",
    "role": "admin"
  }'
```

#### Delete User
```bash
curl -X DELETE "http://localhost:4050/rest/users/1?route=users/1"
```

---

### Products Resource

#### List All Products
```bash
curl "http://localhost:4050/rest/products?route=products"
```

#### Get Specific Product
```bash
curl "http://localhost:4050/rest/products/1?route=products/1"
```

#### Create Product
```bash
curl -X POST "http://localhost:4050/rest/products?route=products" \
  -H "Content-Type: application/json" \
  -d '{
    "name": "MacBook Pro 16",
    "price": 2499.99,
    "stock": 25
  }'
```

#### Update Product
```bash
curl -X PUT "http://localhost:4050/rest/products/1?route=products/1" \
  -H "Content-Type: application/json" \
  -d '{
    "name": "MacBook Pro 15",
    "price": 1999.99,
    "stock": 15
  }'
```

#### Delete Product
```bash
curl -X DELETE "http://localhost:4050/rest/products/1?route=products/1"
```

---

### Items Resource (Advanced Examples)

#### List All Items
```bash
curl "http://localhost:4050/rest/items?route=items"
```

**Ax Functions Demonstrated:**
- `AxExplode()` - Split item names
- `AxCount()` - Count array elements
- `AxWordCount()` - Count description words
- `AxDate()` - Format dates
- `AxTime()` - Get timestamps

#### Get Single Item with Encoding
```bash
curl "http://localhost:4050/rest/items/1?route=items/1"
```

**Ax Functions Demonstrated:**
- `AxHash()` - SHA256 hashing
- `AxBase64Encode()` - Base64 encoding
- `AxMd5()` - MD5 hashing
- `AxHtmlSpecialChars()` - Escape HTML
- `AxNumberFormat()` - Format numbers

#### Create Item with Validation
```bash
curl -X POST "http://localhost:4050/rest/items?route=items" \
  -H "Content-Type: application/json" \
  -d '{
    "name": "New Item"
  }'
```

**Ax Functions Demonstrated:**
- `AxUcFirst()` - Capitalize first character
- `AxCTypeAlpha()` - Validate alphabetic input

---

### Status/Health Check

```bash
curl "http://localhost:4050/rest/status?route=status"
```

**Response includes:**
- Server status and uptime
- Current timestamp (Unix and formatted)
- System information

**Ax Functions Demonstrated:**
- `AxTime()` - Unix timestamp
- `AxDate()` - Formatted date

---

## Response Formats

### JSON (Default)

```bash
curl "http://localhost:4050/rest/users?route=users"

# Response:
{
  "status": "success",
  "data": [ ... ],
  "count": 3,
  "timestamp": "2024-01-16T14:30:45"
}
```

### HTML Format

```bash
curl "http://localhost:4050/rest/users?route=users&format=html"

# Returns HTML page with formatted JSON display
```

### Testing with Different Formats

```bash
# JSON format (default)
curl "http://localhost:4050/rest/users/1?route=users/1"

# HTML format
curl "http://localhost:4050/rest/users/1?route=users/1&format=html"

# Or use format parameter
curl "http://localhost:4050/rest/users/1?route=users/1&format=json"
```

---

## JavaScript / Fetch API Examples

### GET Request
```javascript
// List all users
fetch('/rest/users?route=users')
  .then(response => response.json())
  .then(data => console.log(data))
  .catch(error => console.error('Error:', error));

// Get single user
fetch('/rest/users/1?route=users/1')
  .then(response => response.json())
  .then(data => console.log(data));
```

### POST Request (Create)
```javascript
fetch('/rest/users?route=users', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    name: 'Alice Johnson',
    email: 'alice@example.com',
    role: 'admin'
  })
})
  .then(response => response.json())
  .then(data => console.log(data))
  .catch(error => console.error('Error:', error));
```

### PUT Request (Update)
```javascript
fetch('/rest/users/1?route=users/1', {
  method: 'PUT',
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    name: 'Bob Smith',
    email: 'bob@example.com',
    role: 'moderator'
  })
})
  .then(response => response.json())
  .then(data => console.log(data));
```

### DELETE Request
```javascript
fetch('/rest/users/1?route=users/1', {
  method: 'DELETE'
})
  .then(response => response.json())
  .then(data => console.log(data));
```

### Handle Different Status Codes
```javascript
fetch('/rest/users?route=users')
  .then(response => {
    if (!response.ok) {
      throw new Error(`HTTP Error: ${response.status}`);
    }
    return response.json();
  })
  .then(data => {
    if (data.status === 'success') {
      console.log('Request successful:', data.data);
    } else {
      console.error('API Error:', data.message);
    }
  })
  .catch(error => console.error('Network Error:', error));
```

---

## Ax Functions Used in API

The API extensively uses modern Ax custom functions for data processing:

### String Functions
- **AxExplode()** - Split string by delimiter
- **AxImplode()** - Join array with glue string
- **AxTrim()** - Remove whitespace from string
- **AxUcFirst()** - Uppercase first character
- **AxWordCount()** - Count words in text
- **AxHtmlSpecialChars()** - Escape HTML special characters
- **AxBase64Encode()** - Encode to Base64
- **AxBase64Decode()** - Decode from Base64

### Array Functions
- **AxCount()** - Count array elements
- **AxEmpty()** - Check if value is empty
- **AxRange()** - Generate number range as array

### Hashing & Encoding
- **AxHash()** - Generate hash (SHA256, SHA1, etc.)
- **AxMd5()** - Generate MD5 hash
- **AxBase64Encode()** - Encode to Base64

### Number Functions
- **AxNumberFormat()** - Format numbers with separators
- **AxRand()** - Generate random number
- **AxMax()** - Find maximum value
- **AxMin()** - Find minimum value

### Date/Time Functions
- **AxDate()** - Format date/time (PHP-like)
- **AxTime()** - Get Unix timestamp

### Validation Functions
- **AxCTypeAlpha()** - Check if string is alphabetic
- **AxFilterValidateEmail()** - Validate email address
- **AxFilterValidateIp()** - Validate IP address

### Utility Functions
- **AxPad()** - Pad string/number to length
- **AxRepeat()** - Repeat string multiple times
- **AxIsset()** - Check if variable is set
- **AxEmpty()** - Check if empty

---

## HTTP Status Codes

| Code | Status | Usage |
|------|--------|-------|
| 200 | OK | Successful GET, PUT, DELETE |
| 201 | Created | Successful POST (resource created) |
| 400 | Bad Request | Missing required parameters |
| 404 | Not Found | Resource doesn't exist |
| 405 | Method Not Allowed | HTTP method not supported for resource |
| 422 | Unprocessable Entity | Validation error in request data |

---

## Error Response Format

All errors follow this format:

```json
{
  "status": "error",
  "code": 400,
  "message": "Descriptive error message",
  "timestamp": "2024-01-16 14:30:45"
}
```

### Common Error Cases

**Missing Required ID (for PUT/DELETE):**
```json
{
  "status": "error",
  "code": 400,
  "message": "User ID required for PUT",
  "timestamp": "2024-01-16 14:32:10"
}
```

**Invalid Resource:**
```json
{
  "status": "error",
  "code": 404,
  "message": "Resource not found: unknownresource",
  "timestamp": "2024-01-16 14:33:25"
}
```

**Unsupported Method:**
```json
{
  "status": "error",
  "code": 405,
  "message": "Method PATCH not allowed",
  "timestamp": "2024-01-16 14:34:40"
}
```

**Validation Error:**
```json
{
  "status": "error",
  "code": 422,
  "message": "Invalid email format",
  "timestamp": "2024-01-16 14:35:55"
}
```

---

## Testing Tools

### Using Postman

1. Create new HTTP request
2. Set method (GET, POST, PUT, DELETE)
3. Enter URL: `http://localhost:4050/rest/users?route=users`
4. For POST/PUT, set Body as JSON
5. Send and view response

### Using Thunder Client (VS Code)

1. Install Thunder Client extension
2. Create new request collection
3. Set up requests for each endpoint
4. Test with different methods and data

### Using Web Browser Console

Open browser developer tools (F12) and test:

```javascript
// List users
fetch('/rest/users?route=users')
  .then(r => r.json())
  .then(d => console.log(d));

// Create user
fetch('/rest/users?route=users', {
  method: 'POST',
  headers: {'Content-Type': 'application/json'},
  body: JSON.stringify({name: 'Test', email: 'test@example.com', role: 'user'})
})
  .then(r => r.json())
  .then(d => console.log(d));
```

---

## Performance Notes

- All requests are processed synchronously
- JSON serialization is handled by native G3JSON library
- Consider caching for repeated GET requests in production
- Response times typically < 100ms per request

---

## Next Steps

1. Review the [Interactive Documentation](/rest/README.asp) for visual endpoint listing
2. Test endpoints with the [Testing Console](/rest/test-console.html)
3. Study [Code Examples](/rest/code-examples.html) for extending the API
4. Check [main implementation](/rest/index.asp) to understand the pattern
5. Read the [full REST API documentation](../docs/REST_API_IMPLEMENTATION.md)

---

## Need Help?

- All endpoints support `?format=html` for debugging
- Check the interactive console for live testing
- Review code examples for extending with new resources
- Consult custom functions documentation for Ax function usage specs

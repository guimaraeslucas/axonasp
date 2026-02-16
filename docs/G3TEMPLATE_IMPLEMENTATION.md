## G3TEMPLATE Library Implementation Summary

### Overview
A comprehensive template rendering library has been implemented for AxonASP, providing professional-grade template processing capabilities using Go's powerful templating engine for dynamic content generation.

### Files Created/Modified

#### New/Modified Files
1. **`server/template_lib.go`** (95 lines)
   - Complete implementation of G3TEMPLATE library
   - Go template file parsing and rendering
   - Data binding to templates
   - HTML/text template support

#### Integration
1. **`server/executor_libraries.go`**
   - Added TemplateLibrary wrapper for ASPLibrary interface compatibility
   - Enables: `Set template = Server.CreateObject("G3TEMPLATE")`
   - Also supports: `Server.CreateObject("TEMPLATE")`

2. **`server/custom_functions.go`**
   - Added `TemplateHelper()` function for backward compatibility

### Key Features Implemented

✓ **Template Rendering**
  - `Render(path, [data])` - Parse and render template file
  - Support for both relative and virtual paths
  - Data binding via objects/dictionaries
  - Full Go template syntax support

✓ **Template Syntax**
  - Variable substitution: `{{.VariableName}}`
  - Conditionals: `{{if .Condition}}...{{end}}`
  - Loops: `{{range .Items}}...{{end}}`
  - Functions and pipelines
  - Nested objects: `{{.User.Name}}`

✓ **Data Types**
  - Map[string]interface{} from G3JSON
  - G3Dictionary objects
  - ASP objects and collections
  - Nested data structures
  - Arrays and slices

✓ **Path Security**
  - Template path validation
  - Directory traversal prevention
  - Access restricted to web root
  - Secure file system isolation

### Architecture

**Class Hierarchy**:
```
Component (interface)
  └─ G3TEMPLATE
      ├─ Render()
      ├─ Path validation
      └─ Template execution
```

**Template Engine**:
- Go's standard `html/template` package
- Full support for Go template language
- Automatic HTML escaping for security
- Custom delimiters support

**Data Flow**:
1. ASP code calls Render with template path
2. Path validated against root directory
3. Template file loaded and parsed
4. Data object passed to execution context
5. Template rendered to string
6. Result returned to ASP

### Usage Examples

#### Basic Template Rendering
```vbscript
Dim template, output
Set template = Server.CreateObject("G3TEMPLATE")

' Render template without data
output = template.Render("email.tmpl")
Response.Write output
```

#### Template with Data
```vbscript
Dim template, data, output
Set template = Server.CreateObject("G3TEMPLATE")
Set data = Server.CreateObject("G3JSON").NewObject()

' Prepare data
data("name") = "John"
data("email") = "john@example.com"

' Render with data
output = template.Render("user_profile.tmpl", data)
Response.Write output
```

#### Simple Variable Substitution
```html
<!-- user_profile.tmpl -->
<h1>Welcome, {{.name}}</h1>
<p>Email: {{.email}}</p>
<p>Registration: {{.joined_date}}</p>
```

```vbscript
Dim template, data
Set template = Server.CreateObject("G3TEMPLATE")
Set data = Server.CreateObject("G3JSON").NewObject()

data("name") = "Alice"
data("email") = "alice@example.com"
data("joined_date") = "2025-01-01"

Response.Write template.Render("user_profile.tmpl", data)
```

#### Conditional Rendering
```html
<!-- membership.tmpl -->
{{if .is_premium}}
  <p>Premium Member Benefits:</p>
  <ul>
    <li>Priority Support</li>
    <li>Advanced Features</li>
  </ul>
{{else}}
  <p><a href="/upgrade">Upgrade to Premium</a></p>
{{end}}
```

#### Loop/Range Processing
```html
<!-- order_list.tmpl -->
<h2>Your Orders</h2>
{{if .orders}}
  <table>
    {{range .orders}}
    <tr>
      <td>{{.id}}</td>
      <td>{{.product_name}}</td>
      <td>${{.price}}</td>
      <td>{{.date}}</td>
    </tr>
    {{end}}
  </table>
{{else}}
  <p>No orders yet</p>
{{end}}
```

```vbscript
Dim template, data, json, orders, order
Set template = Server.CreateObject("G3TEMPLATE")
Set json = Server.CreateObject("G3JSON")
Set data = json.NewObject()

' Create orders array
Set orders = json.NewArray()

' Add sample orders
Set order = json.NewObject()
order("id") = "ORD001"
order("product_name") = "Laptop"
order("price") = "999.99"
order("date") = "2025-01-15"
ReDim Preserve orders(0)
Set orders(0) = order

data("orders") = orders

' Render template
Response.Write template.Render("order_list.tmpl", data)
```

#### Nested Object Access
```html
<!-- company_profile.tmpl -->
<h1>{{.company.name}}</h1>
<p>{{.company.description}}</p>

<h2>Contact Information</h2>
<p>Email: {{.contact.email}}</p>
<p>Phone: {{.contact.phone}}</p>
<p>Address: {{.contact.address.city}}, {{.contact.address.state}}</p>
```

#### Arithmetic and Formatting
```html
<!-- invoice.tmpl -->
<table>
  {{range .items}}
  <tr>
    <td>{{.name}}</td>
    <td>${{.price}}</td>
    <td>{{.quantity}}</td>
  </tr>
  {{end}}
</table>
<p><strong>Total: ${{.total}}</strong></p>
```

#### Template Composition
```html
<!-- base.tmpl -->
<!DOCTYPE html>
<html>
<head>
  <title>{{.title}}</title>
</head>
<body>
  {{.content}}
</body>
</html>
```

```vbscript
Dim template, data
Set template = Server.CreateObject("G3TEMPLATE")
Set data = Server.CreateObject("G3JSON").NewObject()

data("title") = "My Website"
data("content") = "<h1>Welcome</h1><p>Content here</p>"

Response.Write template.Render("base.tmpl", data)
```

### Template Syntax Reference

#### Variables
```
{{.VariableName}}           ' Output variable
{{.Parent.Child}}           ' Nested access
{{.Array 0}}               ' Array indexing (if supported)
```

#### Conditionals
```
{{if .Condition}}...{{end}}
{{if .A}}...{{else if .B}}...{{else}}...{{end}}
{{with .Variable}}...{{end}}
```

#### Loops
```
{{range .Items}}...{{end}}
{{range .Items}}...{{else}}No items{{end}}
```

#### Functions
```
{{printf "%d" .Number}}
{{len .Array}}
{{index .Array 0}}
```

#### Pipelines
```
{{.Name | printf "Hello %s"}}
{{.Items | len}}
```

### Advanced Features

#### Custom Functions
- Use Go template functions for advanced processing
- String manipulation
- Mathematical operations
- Date/time formatting
- Type conversions

#### Error Handling
```vbscript
Dim template, output
Set template = Server.CreateObject("G3TEMPLATE")

On Error Resume Next

output = template.Render("template.tmpl", data)

If output Like "Error*" Then
    Response.Write "Template error: " & output
Else
    Response.Write output
End If

On Error GoTo 0
```

### Performance Characteristics
- Templates parsed on each render (cacheable in future)
- Fast rendering using Go's template engine
- Minimal memory overhead for data binding
- Efficient string building
- Direct output streaming

### Security Features
- Automatic HTML escaping
- Path validation prevents directory traversal
- Access restricted to web root directory
- Script injection protection via escaping
- Safe data binding

### Limitations
- Templates must be in web root or accessible paths
- No template caching (currently parsed per render)
- No custom delimiters (uses {{ }})
- Limited to Go template syntax
- No template inheritance built-in (but can be simulated)

### MIME Type Support
- HTML templates with auto-escaping
- Text templates (can be explicitly requested)
- XML templates (treated as text, manual escaping)
- JSON templates (string output only)

### Future Enhancements
- Template caching for performance
- Custom template delimiters
- Template inheritance/nesting
- Partial template support
- Block/block override syntax
- Preprocessor directives
- Template debugging information

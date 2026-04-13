# Use the G3TEMPLATE Library

## Overview
The **G3TEMPLATE** library provides a high-performance template rendering engine for G3Pix AxonASP. It allows developers to separate application logic from presentation by using external template files. The library leverages the powerful Go `html/template` engine, providing secure, context-aware HTML escaping and data binding.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set template = Server.CreateObject("G3TEMPLATE")
```

## Prerequisites
- **Template Files**: Ensure your template files (typically `.html` or `.tmpl`) are accessible within the server's file system.
- **Data Structures**: Data passed to templates should ideally be in the form of a **Scripting.Dictionary** or a standard **Array**.

## How it Works
The G3TEMPLATE object parses an external file and executes it using a provided data object. 
- **Data Binding**: You can pass complex nested structures (Dictionaries within Arrays, or vice-versa) to the template.
- **Path Resolution**: The library automatically resolves relative paths using `Server.MapPath`.
- **Security**: The underlying engine provides automatic contextual escaping, protecting your application against Cross-Site Scripting (XSS) attacks by default.

## API Reference

### Methods
- **Render**: Parses a template file and returns a string containing the rendered output.

## Code Example
The following example demonstrates how to pass a dictionary of data to a template and render the result.

```asp
<%
Dim template, data, html
Set template = Server.CreateObject("G3TEMPLATE")

' Prepare data for the template
Set data = Server.CreateObject("Scripting.Dictionary")
data.Add "Title", "Welcome to AxonASP"
data.Add "User", "Lucas"

' Render the template located at /views/index.html
html = template.Render("/views/index.html", data)

' Output the rendered HTML
Response.Write html

Set template = Nothing
%>
```

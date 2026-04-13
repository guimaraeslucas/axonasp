# Render Method

## Overview
Parses an external template file and returns a string containing the rendered output using an optional data object in the G3Pix AxonASP G3TEMPLATE library.

## Syntax
```asp
renderedOutput = obj.Render(templatePath [, data])
```

## Parameters and Arguments
- **templatePath** (String, Required): The relative or absolute path to the template file on the server.
- **data** (Variant, Optional): An object or array containing the data to be bound to the template. This is typically a **Scripting.Dictionary** or a standard **Array**.

## Return Values
Returns a **String** containing the fully rendered template output. If an error occurs during parsing or execution (e.g., file not found or invalid template syntax), the method returns a **String** prefixed with "Error: " followed by the error description.

## Remarks
- The template engine uses standard Go `html/template` syntax (e.g., `{{ .ValueName }}`).
- The method automatically resolves relative paths using `Server.MapPath`.
- Context-aware escaping is automatically applied to prevent XSS.

## Code Example
```asp
<%
Dim template, user, output
Set template = Server.CreateObject("G3TEMPLATE")

' Define a data dictionary
Set user = Server.CreateObject("Scripting.Dictionary")
user.Add "Name", "Lucas"
user.Add "IsAdmin", True

' Render the template
output = template.Render("/templates/profile.html", user)

' Check for errors and display result
If Left(output, 6) = "Error:" Then
    Response.Write "Rendering Failed: " & output
Else
    Response.Write output
End If

Set template = Nothing
%>
```

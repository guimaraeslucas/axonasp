# Unsafe Property

## Overview
Gets or sets a Boolean value that determines whether the G3Pix AxonASP G3MD library renders raw HTML and potentially dangerous URLs in the Markdown source.

## Syntax
```asp
' Get the current value
isUnsafe = md.Unsafe

' Set a new value
md.Unsafe = True
```

## Return Values
Returns a **Boolean**. The default value is **False**.

## Remarks
- When set to **False** (default), the processor sanitizes the output, omitting raw HTML tags and dangerous link protocols (like `javascript:`) to protect against cross-site scripting (XSS).
- When set to **True**, the processor renders the Markdown exactly as provided, including any embedded HTML.
- Use **Unsafe = True** only when processing content from a trusted source or when your application requires specific HTML embedding.

## Code Example
```asp
<%
Dim md, content, html
Set md = Server.CreateObject("G3MD")

' Securely render user-provided content
md.Unsafe = False

content = "Click [here](javascript:alert('XSS')) or <script>doBad()</script>"
html = md.Process(content)

' The dangerous script and link will be neutralized in the output
Response.Write html
Set md = Nothing
%>
```

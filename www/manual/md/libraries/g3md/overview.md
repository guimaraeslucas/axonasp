# Use the G3MD Library

## Overview
The **G3MD** library provides high-performance Markdown-to-HTML conversion services for G3Pix AxonASP applications. Built on the industry-standard Goldmark engine, it supports GitHub Flavored Markdown (GFM) extensions, including tables, task lists, and strikethrough. The library is optimized for zero-allocation performance and is ideal for rendering technical documentation, blog posts, or user-generated content directly on the server.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set md = Server.CreateObject("G3MD")
```

## Prerequisites
No external dependencies are required. The G3MD library is a built-in native component of the G3Pix AxonASP environment.

## How it Works
The G3MD object operates as a stateful processor. You can configure conversion options such as **HardWraps** and **Unsafe** before calling the **Process** method. 
- **GitHub Flavored Markdown**: By default, the library uses the GFM extension set, ensuring compatibility with modern Markdown standards.
- **Safety and Security**: The **Unsafe** property allows you to control whether raw HTML and potentially dangerous links are rendered, providing a layer of protection against cross-site scripting (XSS) when processing untrusted input.
- **Line Break Handling**: The **HardWraps** property determines how soft line breaks in the Markdown source are translated into the final HTML output.

## API Reference

### Methods
- **Process**: Converts a Markdown string into a string of formatted HTML.

### Properties
- **HardWraps**: Gets or sets a Boolean value indicating whether soft line breaks should be converted to `<br>` tags.
- **Unsafe**: Gets or sets a Boolean value indicating whether raw HTML and dangerous URLs should be rendered.

## Code Example
The following example demonstrates how to configure the library and convert a Markdown string into HTML.

```asp
<%
Dim md, markdownText, htmlOutput
Set md = Server.CreateObject("G3MD")

' Configure the processor
md.HardWraps = True
md.Unsafe = False

' Markdown source with GFM features
markdownText = "# Welcome to AxonASP" & vbCrLf & _
               "This is a **high-performance** server." & vbCrLf & _
               "- Feature A" & vbCrLf & _
               "- Feature B"

' Convert to HTML
htmlOutput = md.Process(markdownText)

' Output the result
Response.Write htmlOutput

Set md = Nothing
%>
```

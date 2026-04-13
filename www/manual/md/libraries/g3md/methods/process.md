# Process Method

## Overview
Converts a Markdown-formatted string into a string of formatted HTML in the G3Pix AxonASP G3MD library.

## Syntax
```asp
htmlString = md.Process(markdownText)
```

## Parameters and Arguments
- **markdownText** (String, Required): The source text in Markdown format to be converted.

## Return Values
Returns a **String** containing the rendered HTML. If the input is empty or an error occurs during conversion, it returns an empty string.

## Remarks
- The output follows the GitHub Flavored Markdown (GFM) specification, including support for tables and task lists.
- The conversion process respects the current settings of the **HardWraps** and **Unsafe** properties.
- This method is highly optimized for performance and is suitable for high-traffic backend rendering.

## Code Example
```asp
<%
Dim md, markdown, html
Set md = Server.CreateObject("G3MD")

markdown = "### System Status" & vbCrLf & "- **CPU**: 10%" & vbCrLf & "- **Memory**: 2GB"
html = md.Process(markdown)

Response.Write html
Set md = Nothing
%>
```

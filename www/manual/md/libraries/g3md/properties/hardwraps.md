# HardWraps Property

## Overview
Gets or sets a Boolean value that determines how the G3Pix AxonASP G3MD library handles soft line breaks in the Markdown source.

## Syntax
```asp
' Get the current value
isHardWrap = md.HardWraps

' Set a new value
md.HardWraps = True
```

## Return Values
Returns a **Boolean**. The default value is **False**.

## Remarks
- When set to **True**, every soft line break (a single newline) in the Markdown source is converted to an HTML `<br>` tag.
- When set to **False** (default), Markdown follows standard behavior where soft line breaks are treated as spaces unless they are preceded by two spaces or followed by a blank line.
- This property must be set before calling the **Process** method.

## Code Example
```asp
<%
Dim md, result
Set md = Server.CreateObject("G3MD")

' Enable hard wraps for comment-style rendering
md.HardWraps = True

result = md.Process("Line 1" & vbCrLf & "Line 2")
' Result will contain "Line 1<br>Line 2"

Response.Write result
Set md = Nothing
%>
```

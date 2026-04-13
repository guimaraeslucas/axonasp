# LoadFile Method

## Overview
Reads a JSON file from the disk and parses it into a native AxonASP object or array.

## Syntax
```asp
result = json.LoadFile(path)
```

## Parameters and Arguments
- **path** (String, Required): The file path to the JSON file. This can be a relative virtual path or a physical path. Virtual paths are automatically resolved via `Server.MapPath`.

## Return Values
Returns a **Variant** containing the parsed data. It returns a **Scripting.Dictionary** object if the root JSON element is an object, or an **Array** if the root element is a JSON array. If the file cannot be read or contains invalid JSON, it returns **Empty**.

## Remarks
This method is the most efficient way to load configuration files or static JSON datasets into your application. It leverages high-performance I/O and the built-in Go JSON unmarshaller.

## Code Example
```asp
<%
Dim json, config
Set json = Server.CreateObject("G3JSON")
Set config = json.LoadFile("/config/settings.json")

If Not IsEmpty(config) Then
    Response.Write "App Name: " & config("appName")
End If

Set json = Nothing
%>
```

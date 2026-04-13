# Use the G3JSON Library

## Overview
The **G3JSON** library provides high-performance JSON (JavaScript Object Notation) processing for G3Pix AxonASP applications. It enables seamless conversion between JSON strings and native AxonASP data structures, such as Dictionaries and Arrays. The library is optimized for backend data interchange, supporting deep nested structures and file-based JSON loading.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set json = Server.CreateObject("G3JSON")
```

## Prerequisites
No external dependencies are required. The G3JSON library is a built-in native component of the AxonASP environment.

## How it Works
The G3JSON object acts as a bridge between the JSON data format and the AxonASP Virtual Machine. 
- When **Parsing** a JSON string, the library recursively converts JSON objects into AxonASP **Scripting.Dictionary** objects and JSON arrays into standard VBScript-compatible **Arrays**.
- When **Stringifying** data, it traverses native Dictionaries and Arrays to produce a valid, compact JSON string.
- The **LoadFile** method integrates with the server's file system, automatically mapping virtual paths via `Server.MapPath` to retrieve and parse JSON configuration or data files efficiently.

## API Reference

### Methods
- **LoadFile**: Reads a JSON file from the disk and parses it into a native object or array.
- **NewArray**: Creates a new, empty VBScript-compatible array.
- **NewObject**: Creates a new, empty **Scripting.Dictionary** object.
- **Parse**: Converts a JSON-formatted string into a native AxonASP structure.
- **Stringify**: Serializes an AxonASP structure (Dictionary, Array, or primitive) into a JSON string.

## Code Example
The following example demonstrates how to parse a JSON string, modify the data, and serialize it back to JSON.

```asp
<%
Dim json, data, output
Set json = Server.CreateObject("G3JSON")

' Parse a JSON string into a Dictionary
Set data = json.Parse("{""name"": ""AxonASP"", ""version"": 2.0, ""features"": [""Fast"", ""Secure""]}")

' Access and modify data
Response.Write "Name: " & data("name") & "<br>"
data("version") = 2.1
data("features").Add "Scalable"

' Stringify back to JSON
output = json.Stringify(data)
Response.Write "Updated JSON: " & output

Set data = Nothing
Set json = Nothing
%>
```

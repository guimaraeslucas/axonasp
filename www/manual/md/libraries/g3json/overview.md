# Use the G3JSON Library

## Overview
Use **G3JSON** to parse JSON text into native G3Pix AxonASP values and serialize native values back to JSON text. The library is built into the runtime and supports object, array, and primitive JSON payloads. The Javascript implementation supports the full JSON specification, including Unicode characters and special number formats, this way you can rely on it for robust JSON handling in your server-side scripts without needing this library.

## Prerequisites
- Use a running G3Pix AxonASP environment.
- Create the object with the primary ProgID:

```asp
Dim json
Set json = Server.CreateObject("G3JSON")
```

## How It Works
- **Parse** and **LoadFile** convert JSON into AxonASP values.
- JSON objects become **Scripting.Dictionary** instances.
- JSON arrays become VBScript arrays.
- JSON primitives become scalar values (String, Integer/Double, Boolean, or Null).
- **Stringify** converts native values to a JSON string.

## API Reference

### Methods
- **LoadFile(path)**: Returns parsed JSON from a file, or **Empty** on read/parse failure.
- **NewArray()**: Returns an empty VBScript array.
- **NewObject()**: Returns a new **Scripting.Dictionary** object.
- **Parse(jsonText)**: Returns parsed JSON value, or **Empty** when input is missing, empty, or invalid.
- **Stringify(value)**: Returns a JSON string, or an empty string when no argument is provided or serialization fails.

### Properties
G3JSON does not expose public properties.

## Example
```asp
<%
Dim json, data, payload
Set json = Server.CreateObject("G3JSON")

Set data = json.Parse("{""name"": ""G3Pix AxonASP"", ""version"": 2}")
If IsObject(data) Then
	data("version") = data("version") + 1
	payload = json.Stringify(data)
	Response.Write payload
End If

Set data = Nothing
Set json = Nothing
%>
```

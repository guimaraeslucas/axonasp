# Parse Method

## Overview
Converts a JSON-formatted string into a native AxonASP structure (Dictionary or Array).

## Syntax
```asp
result = json.Parse(jsonString)
```

## Parameters and Arguments
- **jsonString** (String, Required): The JSON string to be parsed.

## Return Values
Returns a **Variant** containing the parsed JSON structure. 
- Returns a **Scripting.Dictionary** object for JSON objects (`{}`).
- Returns a standard **Array** for JSON arrays (`[]`).
- Returns **Empty** if the input is invalid or cannot be parsed.

## Remarks
The **Parse** method supports deep nesting and various JSON data types, including numbers, booleans, strings, and nulls. JSON nulls are converted to VBScript **Null**.

## Code Example
```asp
<%
Dim json, data
Set json = Server.CreateObject("G3JSON")
Set data = json.Parse("{""id"": 101, ""active"": true}")

If IsObject(data) Then
    Response.Write "ID: " & data("id")
End If

Set json = Nothing
%>
```

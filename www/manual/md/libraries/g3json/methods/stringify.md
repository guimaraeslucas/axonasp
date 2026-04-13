# Stringify Method

## Overview
Serializes an AxonASP data structure into a JSON-formatted string.

## Syntax
```asp
jsonStr = json.Stringify(data)
```

## Parameters and Arguments
- **data** (Variant, Required): The data to be serialized. This can be a **Scripting.Dictionary**, an **Array**, or a primitive value (String, Number, Boolean, Date, Null).

## Return Values
Returns a **String** containing the serialized JSON data. If the input cannot be serialized, it returns an empty string or the JSON representation of an empty object.

## Remarks
The **Stringify** method is essential for sending data from the server to web clients or external APIs. It correctly handles complex nested dictionaries and arrays.

## Code Example
```asp
<%
Dim json, dict, jsonStr
Set json = Server.CreateObject("G3JSON")
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "status", "success"
dict.Add "code", 200

jsonStr = json.Stringify(dict)
Response.Write "Serialized JSON: " & jsonStr

Set json = Nothing
%>
```

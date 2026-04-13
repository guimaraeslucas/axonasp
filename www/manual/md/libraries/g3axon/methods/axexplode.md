# Split a String into an Array

Splits a string by a specified separator and returns an array of strings.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Array of Strings = obj.AxExplode(separator, string)
```

## Return Value
Returns a Array of Strings.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxExplode(separator, string)

Response.Write result

Set obj = Nothing
```

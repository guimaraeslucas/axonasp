# Convert Newlines to HTML Line Breaks

Inserts HTML line breaks before all newlines in a string.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
String = obj.AxNl2br(input)
```

## Return Value
Returns a String.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxNl2br(input)

Response.Write result

Set obj = Nothing
```

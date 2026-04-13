# Make the First Character Uppercase

Makes a string's first character uppercase.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
String = obj.AxUcfirst(input)
```

## Return Value
Returns a String.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxUcfirst(input)

Response.Write result

Set obj = Nothing
```

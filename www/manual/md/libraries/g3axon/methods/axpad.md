# Pad a String to a Certain Length

Pads a string to a certain length with another string.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
String = obj.AxPad(input, padLength, padString, padType)
```

## Return Value
Returns a String.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxPad(input, padLength, padString, padType)

Response.Write result

Set obj = Nothing
```

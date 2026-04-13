# Repeat a String

Repeats a string a specified number of times.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
String = obj.AxRepeat(input, multiplier)
```

## Return Value
Returns a String.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxRepeat(input, multiplier)

Response.Write result

Set obj = Nothing
```

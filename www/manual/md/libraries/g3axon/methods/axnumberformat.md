# Format a Number

Formats a number with grouped thousands and a specified decimal point.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
String = obj.AxNumberFormat(value, decimals, decPoint, thousandsSep)
```

## Return Value
Returns a String.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxNumberFormat(value, decimals, decPoint, thousandsSep)

Response.Write result

Set obj = Nothing
```

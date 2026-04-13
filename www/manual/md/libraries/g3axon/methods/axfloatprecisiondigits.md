# Get Float Precision Digits

Returns the number of digits of precision for float values.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Integer) = obj.AxFloatPrecisionDigits()
```

## Return Value
Returns a Number (Integer).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxFloatPrecisionDigits()

Response.Write result

Set obj = Nothing
```

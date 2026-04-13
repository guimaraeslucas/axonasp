# Create an Array with a Range of Elements

Creates an array containing a range of elements progressing by step.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Array of Numbers = obj.AxRange(start, end, step)
```

## Return Value
Returns a Array of Numbers.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxRange(start, end, step)

Response.Write result

Set obj = Nothing
```

# Count Words in a String

Returns the number of words in a string.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Integer) = obj.AxWordCount(input)
```

## Return Value
Returns a Number (Integer).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxWordCount(input)

Response.Write result

Set obj = Nothing
```

# VBScript Built-in Functions Implementation

## New Features Implemented

This update adds support for essential VBScript type checking functions, literals, and operators.

### Hexadecimal and Octal Literals

VBScript supports special numeric literal formats:

- **Hexadecimal**: `&h` or `&H` prefix (e.g., `&h5C` = 92, `&hFF` = 255)
- **Octal**: `&o` or `&O` prefix (e.g., `&o10` = 8, `&o77` = 63)

**Examples:**
```vbscript
Dim hexValue
hexValue = &hFF        ' 255 in decimal
hexValue = &h5C        ' 92 in decimal

Dim octValue
octValue = &o10        ' 8 in decimal
octValue = &o77        ' 63 in decimal

' Can be used in expressions
Dim sum
sum = &h10 + &h20      ' 16 + 32 = 48
```

### Empty, Nothing, and Null Values

- **Empty**: Represents an uninitialized variable
- **Nothing**: Represents a null object reference
- **Null**: Represents no valid data

**Examples:**
```vbscript
Dim x
' x is Empty when first declared
If IsEmpty(x) Then
    Response.Write "x is empty"
End If

x = Null
If IsNull(x) Then
    Response.Write "x is null"
End If

Dim obj
Set obj = Nothing
If obj Is Nothing Then
    Response.Write "obj is nothing"
End If
```

### TypeName Function

Returns the VBScript type name as a string.

**Syntax:** `TypeName(variable)`

**Return Values:**
- `"Empty"` - Uninitialized variable
- `"Nothing"` - Null object reference
- `"Boolean"` - Boolean value
- `"Integer"` - Integer number
- `"Double"` - Floating-point number
- `"String"` - String value
- `"Variant()"` - Array
- `"Dictionary"` - Dictionary/Map object
- `"Object"` - Custom object

**Examples:**
```vbscript
Response.Write TypeName("Hello")      ' "String"
Response.Write TypeName(42)           ' "Integer"
Response.Write TypeName(3.14)         ' "Double"
Response.Write TypeName(True)         ' "Boolean"
Response.Write TypeName(Array(1,2))   ' "Variant()"
```

### VarType Function

Returns a numeric constant indicating the variable type.

**Syntax:** `VarType(variable)`

**Return Values:**
- `0` (vbEmpty) - Uninitialized
- `1` (vbNull) - Null
- `2` (vbInteger) - Integer
- `3` (vbLong) - Long integer
- `4` (vbSingle) - Single-precision float
- `5` (vbDouble) - Double-precision float
- `8` (vbString) - String
- `9` (vbObject) - Object
- `11` (vbBoolean) - Boolean
- `8204` (vbArray + vbVariant) - Array

**Examples:**
```vbscript
Response.Write VarType("Hello")     ' 8 (vbString)
Response.Write VarType(42)          ' 2 (vbInteger)
Response.Write VarType(3.14)        ' 5 (vbDouble)
Response.Write VarType(True)        ' 11 (vbBoolean)
Response.Write VarType(Array(1,2))  ' 8204 (vbArray)
```

### RGB Function

Returns an integer representing a color value from red, green, and blue components.

**Syntax:** `RGB(red, green, blue)`

**Parameters:**
- `red` - Red component (0-255)
- `green` - Green component (0-255)
- `blue` - Blue component (0-255)

**Returns:** Integer in BGR format (Blue in low byte)

**Examples:**
```vbscript
Dim red, green, blue, gray
red = RGB(255, 0, 0)        ' 255
green = RGB(0, 255, 0)      ' 65280
blue = RGB(0, 0, 255)       ' 16711680
gray = RGB(128, 128, 128)   ' 8421504
```

### IsObject Function

Checks if a variable contains an object reference.

**Syntax:** `IsObject(variable)`

**Returns:** Boolean (True if object, False otherwise)

**Examples:**
```vbscript
Dim obj, num
Set obj = Server.CreateObject("Scripting.Dictionary")
num = 42

Response.Write IsObject(obj)    ' True
Response.Write IsObject(num)    ' False
Response.Write IsObject("Hi")   ' False
```

### IsEmpty Function

Checks if a variable is Empty (uninitialized).

**Syntax:** `IsEmpty(variable)`

**Returns:** Boolean (True if empty, False otherwise)

**Examples:**
```vbscript
Dim x, y
y = "Test"

Response.Write IsEmpty(x)       ' True
Response.Write IsEmpty(y)       ' False
Response.Write IsEmpty(Empty)   ' True
```

### IsNull Function

Checks if a variable contains Null.

**Syntax:** `IsNull(variable)`

**Returns:** Boolean (True if null, False otherwise)

**Examples:**
```vbscript
Dim x
x = Null

Response.Write IsNull(x)        ' True
Response.Write IsNull("Test")   ' False
Response.Write IsNull(42)       ' False
```

### Is Nothing Operator

Checks if an object reference is Nothing.

**Syntax:** `object Is Nothing`

**Returns:** Boolean (True if nothing, False otherwise)

**Examples:**
```vbscript
Dim obj1, obj2
Set obj1 = Nothing
Set obj2 = Server.CreateObject("Scripting.Dictionary")

If obj1 Is Nothing Then
    Response.Write "obj1 is nothing"
End If

If Not (obj2 Is Nothing) Then
    Response.Write "obj2 is not nothing"
End If
```

## Implementation Details

### Files Modified/Created

1. **server/builtin_functions.go** - New file containing all built-in function implementations
2. **server/executor.go** - Updated to support:
   - Built-in function calls
   - Empty and Nothing special types
   - Is/Is Not operators for object comparison
   - Hex/octal literal parsing in numeric conversions

### Technical Notes

- Hex and octal literals are parsed during numeric conversion operations
- Empty is represented by the `EmptyValue` type
- Nothing is represented by the `NothingValue` type
- Null is represented as `nil`
- The `Is` operator compares object references
- RGB returns BGR-formatted integer (VBScript standard)

## Testing

Run the comprehensive test page:
```
http://localhost:4050/test_type_functions.asp
```

This page tests all new functionality including:
- Hexadecimal literals (&h)
- Octal literals (&o)
- TypeName with all types
- VarType with all types
- RGB color calculations
- IsObject checks
- IsEmpty checks
- IsNull checks
- Is Nothing operator
- Combined tests with arithmetic

## Example Usage

```vbscript
<%@ Language=VBScript %>
<%
' Hex and octal literals
Dim hexColor
hexColor = &hFF0000  ' Red in hex
Response.Write "Color: " & hexColor & "<br>"

' Type checking
Dim myVar
If IsEmpty(myVar) Then
    Response.Write "TypeName: " & TypeName(myVar) & "<br>"
    Response.Write "VarType: " & VarType(myVar) & "<br>"
End If

' RGB colors
Dim bgColor
bgColor = RGB(240, 240, 240)
Response.Write "Background: " & bgColor & "<br>"

' Object checking
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")
If IsObject(dict) Then
    Response.Write "It's an object!<br>"
End If

' Nothing checking
Dim obj
Set obj = Nothing
If obj Is Nothing Then
    Response.Write "Object is Nothing<br>"
End If
%>
```

## Compatibility

These implementations follow VBScript specifications strictly to ensure compatibility with classic ASP applications.

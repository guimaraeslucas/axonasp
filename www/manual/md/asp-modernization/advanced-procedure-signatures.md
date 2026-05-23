# Advanced Procedure Signatures

## Overview

AxonASP extends the standard VBScript procedure declaration syntax with three Visual Basic 6.0 (VB6) features that provide greater control over parameter passing and function signatures: **Optional parameters** with default values, **ParamArray** for variable-length argument lists, and **explicit ByVal** for pass-by-value semantics. These features are fully backward compatible with existing Classic ASP code.

## Optional Parameters

Optional parameters allow callers to omit arguments when invoking a Sub or Function. When an optional parameter is omitted, its declared default value is used. If no default value is specified, the parameter receives the Empty value.

### Syntax

```vb
Sub MySub(Optional paramName As Type = defaultValue)
Function MyFunc(Optional paramName As Type = defaultValue) As Type
```

- **paramName**: The parameter name.
- **As Type** (optional): Declares the parameter data type (Integer, Long, Single, Double, String, Boolean, Byte, Object, Variant, or a User-Defined Type).
- **= defaultValue**: A constant expression used when the caller omits the argument. If omitted, the parameter defaults to Empty.

### Parameters and Arguments

| Parameter | Required | Type | Description |
|-----------|----------|------|-------------|
| paramName | Yes | Identifier | The name of the optional parameter |
| As Type | No | Type keyword | Declares the parameter data type (Variant if omitted) |
| defaultValue | No | Constant expression | Default value used when argument is not supplied |

### Remarks

- Optional parameters must appear after all required parameters in the parameter list.
- The default value must be a compile-time constant expression (literal, enum constant, or folded expression).
- If a caller provides an argument for an optional parameter, the provided value is used instead of the default.
- Multiple optional parameters can be declared in the same procedure.
- When an optional parameter without an explicit default is omitted, it receives the VBScript Empty value.
- Optional parameters can be used together with the **As Type** clause to enforce type constraints.

### Code Example

```vb
<%@ Language="VBScript" %>
<%
' Function with optional parameters
Function FormatMessage(msg, Optional prefix As String = "INFO", Optional timestamp As String = "now")
    Dim result
    result = prefix & ": " & msg
    If timestamp <> "now" Then
        result = result & " [" & timestamp & "]"
    End If
    FormatMessage = result
End Function

' Call with all arguments
Response.Write FormatMessage("Server started", "DEBUG", "12:00") & "<br>"

' Call with omitted optional arguments (uses defaults)
Response.Write FormatMessage("Server started") & "<br>"

' Call with only the first optional argument
Response.Write FormatMessage("High memory", "WARN")
%>
```

## ParamArray

ParamArray allows a procedure to accept a variable number of arguments. All remaining arguments are collected into a zero-based array. ParamArray must be the last parameter in the declaration.

### Syntax

```vb
Sub MySub(ParamArray paramName())
Function MyFunc(ParamArray paramName())
```

- **paramName**: The name of the array parameter. Must be declared with empty parentheses `()`.

### Parameters and Arguments

| Parameter | Required | Type | Description |
|-----------|----------|------|-------------|
| paramName | Yes | Array | A zero-based array containing all remaining arguments passed by the caller |

### Remarks

- ParamArray must be the last parameter in the procedure declaration.
- The parameter is always passed by value (ByVal).
- When no arguments are provided for the ParamArray, the parameter receives an empty array (LBound > UBound).
- Use LBound and UBound to iterate over the array elements.
- ParamArray cannot be used with the Optional keyword or the As Type clause.
- ParamArray is ideal for utility functions like logging, formatting, and aggregation.

### Code Example

```vb
<%@ Language="VBScript" %>
<%
' Function that sums all provided numbers
Function Sum(ParamArray numbers())
    Dim total, i
    total = 0
    For i = LBound(numbers) To UBound(numbers)
        total = total + numbers(i)
    Next
    Sum = total
End Function

' Function that joins strings with a separator
Function JoinStrings(sep, ParamArray items())
    Dim result, i
    result = ""
    For i = LBound(items) To UBound(items)
        If result <> "" Then result = result & sep
        result = result & items(i)
    Next
    JoinStrings = result
End Function

' Usage
Response.Write "Sum of 1,2,3,4,5 = " & Sum(1, 2, 3, 4, 5) & "<br>"
Response.Write "Joined: " & JoinStrings(", ", "apple", "banana", "cherry")
%>
```

## Explicit ByVal

In standard VBScript, procedure parameters default to ByRef (passed by reference), meaning modifications to the parameter inside the procedure affect the caller's variable. The ByVal keyword changes this behavior: the procedure receives a local copy of the value, and modifications do not affect the caller's variable.

### Syntax

```vb
Sub MySub(ByVal paramName As Type)
Function MyFunc(ByVal paramName As Type) As Type
```

- **paramName**: The parameter name.
- **As Type** (optional): Declares the parameter data type.

### Parameters and Arguments

| Parameter | Required | Type | Description |
|-----------|----------|------|-------------|
| paramName | Yes | Identifier | The name of the parameter |
| As Type | No | Type keyword | Declares the parameter data type (Variant if omitted) |

### Remarks

- ByVal creates a local copy of the argument value at procedure entry. Any modifications to the parameter inside the procedure do not affect the caller's variable.
- ByRef (or the default when no modifier is specified) passes a reference to the caller's variable. Modifications to the parameter affect the caller's variable.
- Using ByVal is recommended for simple data types (Integer, String, Double) when the procedure does not need to modify the caller's variable. This prevents unintended side effects.
- ByRef is necessary when the procedure must return a value through a parameter.
- Use the **As Type** clause to enforce type safety for ByVal parameters.

### Code Example

```vb
<%@ Language="VBScript" %>
<%
Dim value

' Example 1: ByVal protects the caller's variable
value = 10
Sub IncrementByVal(ByVal x)
    x = x + 1
End Sub
Call IncrementByVal(value)
Response.Write "After ByVal: " & value & " (unchanged, still 10)<br>"

' Example 2: ByRef allows modification
Sub IncrementByRef(ByRef x)
    x = x + 1
End Sub
Call IncrementByRef(value)
Response.Write "After ByRef: " & value & " (changed to 11)<br>"

' Example 3: Default (no modifier) is ByRef
Sub IncrementDefault(x)
    x = x + 1
End Sub
Call IncrementDefault(value)
Response.Write "After default: " & value & " (changed to 12)"
%>
```

## Combined Usage

Optional parameters, ParamArray, and ByVal can be combined in a single procedure declaration, following these ordering rules:

1. Required parameters (with or without ByRef/ByVal)
2. Optional parameters (with or without default values)
3. ParamArray (must be last)

### Code Example

```vb
<%@ Language="VBScript" %>
<%
' Advanced logging function combining all features
Function LogMessage(ByVal level As String, _
                    ByVal msg As String, _
                    Optional timestamp As String = "now", _
                    ParamArray tags())
    Dim result, i
    result = "[" & level & "] " & msg
    
    If timestamp <> "now" Then
        result = result & " @" & timestamp
    End If
    
    If IsArray(tags) And UBound(tags) >= LBound(tags) Then
        result = result & " {"
        For i = LBound(tags) To UBound(tags)
            If i > LBound(tags) Then result = result & ", "
            result = result & tags(i)
        Next
        result = result & "}"
    End If
    
    LogMessage = result
End Function

' Examples
Response.Write LogMessage("INFO", "Application started") & "<br>"
Response.Write LogMessage("WARN", "High memory usage", "now", "server1", "memory") & "<br>"
Response.Write LogMessage("ERROR", "Connection failed", "12:00:00", "critical", "db", "timeout")
%>
```

## Compatibility Notes

- These extensions are specific to the G3Pix AxonASP implementation and are designed to match the Visual Basic 6.0 language specification wherever possible.
- Classic ASP applications that do not use these features continue to work without modification.
- The default parameter passing convention remains ByRef, matching standard VBScript behavior.
- ParamArray parameters are always zero-based arrays (LBound returns 0), matching VB6 conventions.
- Optional parameter default values support all VBScript literal types: Integer, Long, Double, String, Boolean, and Empty.

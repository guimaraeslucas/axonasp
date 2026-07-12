# Declare Variables with Strong Typing (VB6 As Type)

## Overview

AxonASP extends the standard VBScript language with VB6-style strong typing through the `As Type` clause on variable declarations. This feature lets you constrain a variable to a specific data type, enabling earlier error detection and clearer code intent while maintaining full backward compatibility with legacy ASP VBScript code.

When a variable is declared with `As Type`, AxonASP enforces the type constraint on every subsequent assignment. If a value cannot be coerced to the declared type, a **Type mismatch (error 13)** is raised at runtime. Variables declared without `As Type` continue to behave as standard VBScript Variants.

## Syntax

```vb
Dim variableName As Type
Public variableName As Type
Private variableName As Type

' Multiple declarations on one line:
Dim variable1 As Type1, variable2 As Type2
```

Supported `Type` values:

| Type      | Initial Value | Description                                                  |
|-----------|---------------|--------------------------------------------------------------|
| Integer   | 0             | 64-bit signed integer (same storage as Long in this engine)  |
| Long      | 0             | 64-bit signed integer (same as Integer in this engine)       |
| Single    | 0             | Double-precision floating-point (same storage as Double)      |
| Double    | 0             | Double-precision floating-point                              |
| String    | "" (empty)    | Variable-length string                                       |
| Boolean   | False         | Boolean (True/False)                                         |
| Byte      | 0             | 64-bit signed integer (range checking not enforced in Phase 1) |
| Object    | Nothing       | Generic Object reference                                     |
| ClassName | Nothing       | Concrete/Custom Class or Interface reference (e.g. `As Dog`) |
| Variant   | Empty         | No type constraint (standard VBScript behavior)              |

**Important:** `Variant` is the default when no `As` clause is present. Existing ASP code is unaffected.

## Parameters and Arguments

- **variableName** (Required): The name of the variable to declare.
- **Type** (Optional): One of the supported type names listed above. If omitted, the variable is treated as Variant (standard VBScript behavior).

The `As` keyword is only recognized in declaration contexts (`Dim`, `Public`, `Private`). It is not a reserved keyword and can still be used as an identifier in other contexts, preserving backward compatibility.

## Return Values

When a typed variable is created with `Dim x As Type`, it is automatically initialized to the default value for that type:
- Numeric types: 0
- String: empty string ""
- Boolean: False
- Object: Nothing
- Variant / no As clause: Empty

On assignment, the runtime attempts to coerce the value to the declared type:
- String to Integer: parses the string as a number
- Integer to String: converts using standard string representation
- Numeric to Boolean: 0 becomes False, non-zero becomes True
- Boolean to Integer: False is 0, True is 1
- Empty/Null: becomes the type's zero value

If coercion fails (e.g., assigning `"hello"` to an `Integer`), a Type mismatch error (13) is raised.

## Remarks

- **Backward compatibility:** This feature is additive. Existing ASP pages that do not use `As Type` are completely unaffected.
- **The `As` keyword** is context-sensitive. It is only parsed as a type clause inside `Dim`, `Public`, and `Private` statements. Outside these contexts, `As` is treated as a regular identifier.
- **Performance:** Type enforcement is implemented via direct type tag checks with minimal overhead. No reflection or heap allocation is used.
- **Coercion rules** follow VBScript semantics. Numeric strings are parsed with standard integer/float conversion. Non-numeric strings assigned to numeric types raise Type mismatch.
- **Object type:** `As Object` constrains the variable to hold object references (VTObject, VTNativeObject, or Nothing). Assigning a non-object value raises Type mismatch.
- **Custom Classes:** You can specify any user-defined class name (e.g., `As Dog`) or interface name as the type constraint. Variables declared with a class or interface type will be initialized to `Nothing`, and assignments will validate that the value is an instance of the class (or implements the interface) at runtime. When calling methods/properties via a concrete class variable, dispatch resolves directly to the class's own implementation.
- **No array support:** `As Type` cannot be combined with array dimension parentheses on the same variable declaration.
- **Scope modifiers:** `Public` and `Private` are only valid at page/module level, not inside procedures. Inside procedures, use `Dim` with `As Type`.

## Code Example

```asp
<%
Option Explicit

' Strongly typed declarations
Dim counter As Integer
Dim name As String
Dim isValid As Boolean
Dim price As Double

counter = 42
name = "Widget"
isValid = True
price = 19.99

Response.Write "Counter: " & counter & "<br>"
Response.Write "Name: " & name & "<br>"
Response.Write "Valid: " & isValid & "<br>"
Response.Write "Price: " & price & "<br>"

' Type coercion
Dim qty As Integer
qty = "10"           ' String coerces to Integer
Response.Write "Qty: " & qty & "<br>"

Dim label As String
label = 200           ' Integer coerces to String
Response.Write "Label: " & label & "<br>"

' Type mismatch handling
On Error Resume Next
Dim value As Integer
value = "not-a-number"
If Err.Number <> 0 Then
    Response.Write "Type mismatch caught: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

' Standard Variant (backward compatible)
Dim anything
anything = "mixed"
anything = 42
Response.Write "<br>Variant works as before: " & anything
%>
```



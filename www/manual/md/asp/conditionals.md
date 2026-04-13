# ASP Conditionals

## Overview
Conditionals in G3Pix AxonASP allow your script to make decisions and execute different blocks of code based on specific criteria. The AxonASP Virtual Machine supports traditional boolean logic and comparison operations, providing the primary flow control mechanisms for dynamic application logic.

## Syntax
AxonASP supports two main conditional structures: **If...Then...Else** and **Select Case**.

### If Statement
```asp
If condition Then
    ' Code for True
ElseIf otherCondition Then
    ' Code for other True
Else
    ' Default code
End If
```

### Select Case Statement
```asp
Select Case variable
    Case value1
        ' Code for value1
    Case value2
        ' Code for value2
    Case Else
        ' Default code
End Select
```

## How it Works
Conditionals evaluate expressions into boolean values (**True** or **False**).
- **Comparison Operators**: Includes `=`, `<>`, `<`, `>`, `<=`, and `>=`.
- **Logical Operators**: Includes `And`, `Or`, `Not`, and `Xor`.
- **Short-Circuit Evaluation**: **WARNING**: VBScript does *not* support short-circuit evaluation. In an expression like `If A And B Then`, both `A` and `B` are always evaluated. If `B` relies on `A` being true (e.g., checking an object property after checking if it exists), you must use nested **If** statements to avoid runtime errors.
- **Select Case Semantics**: The `Select Case` statement performs a series of equality comparisons. Unlike some other languages, once a matching case is executed, the entire block terminates; there is no "fall-through" behavior.

## API Reference

### Keywords
- **If...Then**: Starts a conditional block.
- **ElseIf**: Provides an alternative condition if previous ones are false.
- **Else**: Provides a catch-all block if no conditions are met.
- **End If**: Required to close a block **If** statement.
- **Select Case**: Starts a multi-way branch block.
- **Case**: Defines a value to match against the expression.
- **Case Else**: Optional catch-all for the **Select Case** block.

## Remarks
- **Mandatory Block Syntax**: In AxonASP, you should always use the block `If...Then...End If` syntax. Single-line `If` statements are discouraged as they are prone to parsing errors in complex scripts.
- **Coercion**: Non-boolean values are coerced during evaluation. For example, the integer `0` is treated as **False**, and non-zero integers are treated as **True**.
- **Case Insensitivity**: String comparisons using `=` are case-insensitive by default in AxonASP.

## Code Example
The following example demonstrates defensive nesting to handle the lack of short-circuiting, and a `Select Case` for status handling.

```asp
<%
Option Explicit

Dim user, status
Set user = Nothing ' Simulate uninitialized object
status = "PENDING"

' Defensive nesting (VBScript evaluates both sides of 'And')
If Not user Is Nothing Then
    If user.IsActive Then
        Response.Write "User is active.<br>"
    End If
Else
    Response.Write "No user object found.<br>"
End If

' Status handling with Select Case
Select Case status
    Case "ACTIVE"
        Response.Write "Status: Green"
    Case "PENDING", "WAITING"
        Response.Write "Status: Amber"
    Case Else
        Response.Write "Status: Red"
End Select
%>
```

# ASP Procedures

## Overview
Procedures in G3Pix AxonASP are discrete blocks of code that perform specific tasks. They allow for code modularity, better organization, and the implementation of reusable logic. AxonASP supports two types of procedures: **Sub** (Subroutines), which perform actions without returning a value, and **Function**, which perform calculations and return a result to the caller.

## Syntax
Procedures are defined using the **Sub** or **Function** keywords and must be explicitly closed with **End Sub** or **End Function**.

### Subroutine
```asp
Sub SubName(arg1, arg2)
    ' logic here
End Sub
```

### Function
```asp
Function FuncName(arg1, arg2)
    ' logic here
    FuncName = result ' Set return value
End Function
```

## How it Works
The AxonASP Virtual Machine executes procedures using a localized stack frame.
- **Parameters**: Arguments are passed by reference (**ByRef**) by default, meaning changes to the parameter within the procedure affect the original variable. Use the **ByVal** keyword to pass a copy of the value instead.
- **Return Values**: To return a value from a function, assign the result to the function's name within its body.
- **Call Syntax**:
    - **Subs**: Called without parentheses, or by using the **Call** keyword with parentheses.
    - **Functions**: Parentheses are required when the return value is being assigned to a variable.
- **Recursion**: AxonASP supports recursive procedure calls with a safe stack limit.

## API Reference

### Procedure Types
- **Sub**: A block of code that performs an action. It cannot be used in expressions.
- **Function**: A block of code that returns a value. It can be used in expressions and assignments.

### Keywords
- **ByRef**: (Default) Passes the memory address of the argument.
- **ByVal**: Passes a copy of the argument's value.
- **Exit Sub / Exit Function**: Immediately terminates execution of the procedure and returns control to the caller.

## Remarks
- **Scope**: Procedures defined in the main page or its includes are globally accessible throughout that page's execution.
- **Parentheses Trap**: A common error is using parentheses when calling a Sub without the `Call` keyword (e.g., `MySub(x, y)`). This can cause VBScript to treat the arguments as an expression, leading to unexpected behavior or syntax errors.
- **Case Insensitivity**: Procedure names are case-insensitive.

## Code Example
The following example demonstrates the definition and invocation of both a Sub and a Function.

```asp
<%
Option Explicit

' Define a Function to calculate a total with tax
Function GetTotal(price, taxRate)
    GetTotal = price * (1 + taxRate)
End Function

' Define a Sub to log an operation
Sub LogOperation(msg)
    Response.Write "[LOG]: " & msg & "<br>"
End Sub

' --- Main Execution ---
Dim netPrice, total
netPrice = 100.00

' Function call (parentheses required for assignment)
total = GetTotal(netPrice, 0.15)

' Sub calls (different valid syntaxes)
LogOperation "Calculation complete."
Call LogOperation("Final total: " & total)

Response.Write "Amount Due: " & FormatCurrency(total)
%>
```

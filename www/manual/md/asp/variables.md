# ASP Variables

## Overview
Variables in G3Pix AxonASP are used to store and manipulate data during script execution. In the AxonASP Virtual Machine, every variable is a **Variant**, a special data type that can hold different kinds of information, including strings, integers, floating-point numbers, booleans, dates, and object references.

## Syntax
Variables are declared using the **Dim** statement and assigned values using the equals operator (=). Object references require the **Set** keyword.

```asp
Dim variableName
variableName = value

Set objectVariable = Server.CreateObject("ProgID")
```

## How it Works
The AxonASP VM manages variables using an efficient, stack-based architecture. 
- **Variant Type**: The engine automatically handles type conversion (coercion) based on the operation performed. For example, adding a string that looks like a number to an integer will result in a numeric addition.
- **Option Explicit**: This directive forces the explicit declaration of all variables. Using it is a foundational mandate in AxonASP to prevent bugs caused by mistyped variable names.
- **Scope**:
    - **Local**: Declared within a procedure (Sub or Function); available only within that procedure.
    - **Global**: Declared outside procedures; available to all scripts in the same page and its includes.
- **Object Lifecycle**: Variables holding object references must be explicitly cleared by setting them to **Nothing** to ensure immediate resource cleanup.

## API Reference

### Declarations
- **Dim**: The standard way to declare one or more variables.
- **Public**: Declares variables with global scope (used in classes or global modules).
- **Private**: Declares variables accessible only within the script or class where they are defined.
- **Const**: Declares a constant value that cannot be changed during execution.

### Assignment
- **=**: Assigns a scalar value (String, Number, Boolean) to a variable.
- **Set**: Assigns an object reference to a variable.

## Remarks
- **Naming Conventions**: Variable names must begin with an alphabetic character and cannot contain periods or spaces.
- **Initialization**: Variables are initialized to **Empty** by default. Use `IsEmpty()` to check this state.
- **Case Insensitivity**: Variable names in AxonASP are case-insensitive (`myVar` is the same as `MYVAR`).
- **No Inline Initialization**: You cannot declare and assign a variable on the same line (e.g., `Dim x = 10` is invalid).

## Code Example
The following example demonstrates variable declaration, assignment of different types, and proper object handling.

```asp
<%
Option Explicit

' Scalar variables
Dim userName, userAge, isActive
userName = "Lucas"
userAge = 30
isActive = True

' Constant declaration
Const MAX_ATTEMPTS = 5

' Object variable
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "ID", 101
dict.Add "Status", "Active"

' Output values
Response.Write "User: " & userName & " (Age: " & userAge & ")<br>"
Response.Write "Dictionary Count: " & dict.Count

' Cleanup
Set dict = Nothing
%>
```

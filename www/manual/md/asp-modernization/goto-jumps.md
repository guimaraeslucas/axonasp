# VB6 GoTo and Absolute Jumps

## Overview
AxonASP modernizes Classic ASP by introducing support for VB6-style `GoTo` statements and line labels. This allows for absolute execution flow control within procedures, enabling patterns such as structured error handling or legacy code migration from Visual Basic 6.0.

## Labels
A label identifies a specific location in your code that can be targeted by a `GoTo` statement. Labels can be identifiers followed by a colon or numeric line numbers.

### Identifier Labels
An identifier label consists of a name followed immediately by a colon (`:`).
```vbscript
Sub Test()
    Response.Write "Step 1"
    GoTo FinalStep
    
    Response.Write "This will be skipped"
    
FinalStep:
    Response.Write "Step 2"
End Sub
```

## GoTo Statement
The `GoTo` statement unconditionally transfers execution to the specified label or line number within the **same procedure**.

### Syntax
```vbscript
GoTo label
```

### Constraints
- **Procedure Scoping**: `GoTo` can only jump to labels defined within the current Sub, Function, or Property. Jumps across procedure boundaries are strictly prohibited and will result in a compile-time error.
- **Single-Pass Compilation**: AxonASP's compiler handles forward references automatically. You can jump to a label defined later in the same procedure.

## Forward and Backward Jumps
AxonASP supports both forward jumps (skipping code) and backward jumps (creating custom loops).

```vbscript
Dim count
count = 1

StartLoop:
Response.Write count & " "
count = count + 1
If count <= 5 Then GoTo StartLoop

' Output: 1 2 3 4 5 
```

## Best Practices
While `GoTo` is a powerful tool for compatibility and specific patterns, modern structured control flow (`If...Then...Else`, `For...Next`, `Do...Loop`) is generally preferred for readability and maintainability. Use `GoTo` judiciously, primarily for:
1. Migrating legacy VB6 codebases.
2. Implementing complex error handling logic where `On Error Resume Next` is insufficient.
3. Optimizing performance-critical state machines.

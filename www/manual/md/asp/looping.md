# ASP Looping

## Overview
Looping structures in G3Pix AxonASP allow you to execute a block of code multiple times. This is essential for iterating over arrays, processing records from a database, or performing repetitive calculations. AxonASP provides several types of loops to handle different scenarios, including fixed-count iterations and conditional repetitions.

## Syntax
AxonASP supports **For...Next**, **For Each...Next**, **Do...Loop**, and **While...Wend** structures.

### For Loop
```asp
For i = start To end [Step increment]
    ' code
Next
```

### For Each Loop
```asp
For Each item In collection
    ' code
Next
```

### Do Loop
```asp
Do While condition
    ' code
Loop

Do
    ' code
Loop Until condition
```

## How it Works
The AxonASP Virtual Machine executes loop bodies repeatedly until a termination condition is met.
- **For...Next**: Best for a known number of iterations. The `Step` keyword allows for increments other than 1.
- **For Each...Next**: Specifically designed for iterating over collections (like the **ASP Dictionary** or **G3JSON** objects) and arrays.
- **Do...Loop**: Highly flexible. Can check the condition at the beginning (`Do While`) or at the end (`Loop Until`), ensuring the code runs at least once in the latter case.
- **Infinite Loops**: AxonASP includes safety mechanisms to terminate scripts that exceed the configured execution timeout, preventing infinite loops from hanging the server.

## API Reference

### Keywords
- **For / To / Next**: Defines a counter-based loop.
- **Each / In**: Defines a collection-based loop.
- **Do / While / Until / Loop**: Defines a condition-based loop.
- **While / Wend**: A legacy conditional loop (equivalent to `Do While...Loop`).
- **Exit For / Exit Do**: Immediately terminates the loop and moves execution to the statement following the loop closure.

## Remarks
- **Loop Closures**: In AxonASP, `For` loops must end with `Next`, `Do` loops must end with `Loop`, and `While` loops must end with `Wend`.
- **Collection Performance**: `For Each` is generally faster and more readable when working with objects like the **Scripting.Dictionary**.
- **Counter Variable**: In `Next [variable]`, the variable name is optional but can improve readability in nested loops.

## Code Example
The following example iterates over an array using a `For` loop and a dictionary using a `For Each` loop.

```asp
<%
Option Explicit

' Iterating over an Array
Dim colors, i
colors = Array("Red", "Green", "Blue")

Response.Write "<b>Colors:</b><br>"
For i = 0 To UBound(colors)
    Response.Write i & ": " & colors(i) & "<br>"
Next

' Iterating over a Dictionary
Dim dict, key
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "App", "AxonASP"
dict.Add "Engine", "Go"

Response.Write "<br><b>System Info:</b><br>"
For Each key In dict.Keys
    Response.Write key & " = " & dict(key) & "<br>"
Next

' Using Do Until
Dim count
count = 1
Do Until count > 3
    Response.Write "Loop count: " & count & "<br>"
    count = count + 1
Loop

Set dict = Nothing
%>
```

# Advanced Arrays & Option Base

AxonASP now supports advanced array declarations from VB6, including non-zero lower bounds and the `Option Base` directive.

## Non-Zero Lower Bounds

In Classic VBScript, arrays always start at index 0. AxonASP now supports the `A To B` syntax to specify both the lower and upper bounds of an array dimension.

### Syntax

```vbscript
Dim name(lower To upper)
```

- **lower**: The smallest index in the dimension.
- **upper**: The largest index in the dimension.

### Example

```vbscript
Dim years(2000 To 2025)
years(2000) = "Millennium"
years(2025) = "Future"

Response.Write "LBound: " & LBound(years) ' Output: 2000
Response.Write "UBound: " & UBound(years) ' Output: 2025
```

## Option Base

The `Option Base` statement is used at the module level to declare the default lower bound for array subscripts.

### Syntax

```vbscript
Option Base {0 | 1}
```

- **0**: (Default) Arrays start at index 0.
- **1**: Arrays start at index 1.

### Example

```vbscript
Option Base 1
Dim items(10) ' Equivalent to Dim items(1 To 10)

Response.Write "LBound: " & LBound(items) ' Output: 1
```

## ReDim Preserve Support

The `ReDim Preserve` statement now respects lower bounds. However, as per VB6 rules, you can only resize the last dimension, and you cannot change its lower bound.

### Example

```vbscript
Dim arr()
ReDim arr(1 To 10)
' ... some code ...
ReDim Preserve arr(1 To 20) ' Valid: resizing last dimension, same lower bound
```

## Performance
AxonASP implements non-zero lower bounds with O(1) performance using an offset calculation (`index - lower`). This mechanism is allocation-free and integrated into the VM's core array opcodes.

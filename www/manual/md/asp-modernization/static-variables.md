# Static Variables (Static)

The `Static` statement is used at the procedure level to declare variables that retain their values between procedure calls.

## Syntax

```vbscript
Static varname[([subscripts])] [As type][, varname[([subscripts])] [As type]] . . .
```

- **varname**: Required. Name of the variable.
- **subscripts**: Optional. Dimensions of an array variable.
- **type**: Optional. Data type of the variable (e.g., Integer, String, or a User-Defined Type).

## How it Works

Unlike regular `Dim` variables which are re-initialized every time a procedure is called, `Static` variables are initialized only once (the first time the procedure is executed during a request) and preserve their values until the script finishes execution.

In AxonASP, `Static` variables are mapped to hidden global slots, ensuring they persist across calls within the same request while maintaining zero-allocation performance.

## Performance
AxonASP compiles `Static` variable access directly to `OpGetGlobal` and `OpSetGlobal` using the mapped hidden slots. Initialization is guarded by a fast `IsEmpty` check, ensuring it only runs once per request.

## Example

```vbscript
Function GetNextID()
    ' id will persist its value between calls
    Static id
    id = id + 1
    GetNextID = id
End Function

Response.Write GetNextID() ' Output: 1
Response.Write GetNextID() ' Output: 2
Response.Write GetNextID() ' Output: 3
```

### With Static Arrays and Types

```vbscript
Sub LogMessage(msg)
    Static history(100)
    Static count As Integer
    
    If count <= 100 Then
        history(count) = msg
        count = count + 1
    End If
End Sub
```

## Remarks
- `Static` variables are only valid within `Sub`, `Function`, or `Property` blocks.
- In a `Class` method, `Static` variables are shared across all instances of that class within the same request.
- All `Static` variables are cleared when the request ends, preventing state bleeding between different users or requests.

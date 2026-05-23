# Event Orientation (Event, RaiseEvent, WithEvents)

AxonASP modernization brings Visual Basic 6 (VB6) event-driven programming patterns to Classic ASP. This allows for cleaner decoupling between components by enabling objects to notify observers when specific actions occur.

## Event Statement

Declares a user-defined event within a Class.

### Syntax

```vbscript
[Public] Event procedurename [(arglist)]
```

- **procedurename**: Required. Name of the event.
- **arglist**: Optional. List of variables representing arguments passed to the event when it is raised.

### Example

```vbscript
Class MyComponent
    Event OnComplete(status)
End Class
```

## RaiseEvent Statement

Fires a user-defined event.

### Syntax

```vbscript
RaiseEvent eventname [(arglist)]
```

- **eventname**: Required. Name of the event to raise.
- **arglist**: Optional. Comma-separated list of variables, expressions, or values to pass to the event handlers.

### Example

```vbscript
Class MyComponent
    Event OnComplete(status)
    
    Sub DoWork()
        ' ... processing ...
        RaiseEvent OnComplete("Success")
    End Sub
End Class
```

## WithEvents Keyword

Specifies that a variable will be used to handle events raised by an object.

### Syntax

```vbscript
[Public | Private] WithEvents varname [As typename]
```

- **varname**: Required. Name of the variable that will hold the object reference.
- **typename**: Optional. Name of the class that raises the events.

### Usage Rules

1.  **Naming Convention**: Event handlers must be named using the pattern `varname_eventname`.
2.  **Scope**: `WithEvents` can be used for class members or global variables. It is not supported for local variables within procedures (matching VB6 behavior).
3.  **Assignment**: Binding happens automatically when an object instance is assigned to the `WithEvents` variable using the `Set` statement.

### Example

```vbscript
Class EventSource
    Event OnNotify(msg)
    
    Sub Trigger()
        RaiseEvent OnNotify("Hello from Source!")
    End Sub
End Class

Dim WithEvents objSource

Sub objSource_OnNotify(msg)
    Response.Write "Received: " & msg
End Sub

Set objSource = New EventSource
objSource.Trigger()
' Output: Received: Hello from Source!
```

## Performance and Memory

AxonASP implements events using an efficient internal observer table. Circular references between the event source and the sink are handled carefully to prevent memory leaks during `Class_Terminate`.

- **Zero-Allocation Dispatch**: Event raising is optimized to minimize heap allocations.
- **Deterministic Cleanup**: Observers are automatically unbound when the `WithEvents` variable is reassigned or the containing object is destroyed.

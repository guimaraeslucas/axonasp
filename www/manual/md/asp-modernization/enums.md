# Enumerations (Enum)

Enumerations provide a way to work with sets of related constants. They improve code readability by replacing "magic numbers" with descriptive names.

## Syntax

```vbscript
[Public | Private] Enum name
    membername [= constantexpression]
    membername [= constantexpression]
    ...
End Enum
```

- **Public**: Optional. Used to declare enumerations that are available to all procedures in all scripts.
- **Private**: Optional. Used to declare enumerations that are available only within the script where the declaration is made.
- **name**: Required. Name of the enumeration type.
- **membername**: Required. Name of the constituent members.
- **constantexpression**: Optional. The value assigned to the member. If omitted, the first member is 0, and subsequent members are incremented by 1.

## Performance
AxonASP compiles `Enum` members directly into `OpConstant` opcodes. The `Enum` wrapper itself does not exist at runtime, ensuring zero-overhead for enumeration lookups.

## Example

```vbscript
Enum Colors
    Red      ' Value: 0
    Green    ' Value: 1
    Blue = 5 ' Value: 5
    Yellow   ' Value: 6
End Enum

Response.Write "The value of Green is: " & Green
Response.Write "The value of Yellow is: " & Yellow
```

## Remarks
- Enumerations must be declared at the module level (outside of Sub or Function blocks).
- Enumeration members are global in scope within the script or application (if Public).
- AxonASP supports using previously declared constants or other enumeration members as `constantexpression`.

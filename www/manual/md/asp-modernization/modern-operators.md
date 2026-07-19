# Modern VBScript Operators

## Overview
AxonASP extends the Classic ASP VBScript engine with a set of modern operators inspired by VB.NET. These operators enhance code clarity, enable bitwise data manipulation, and provide short-circuit evaluation for conditional expressions. Some of them are inspired by the twinBASIC project evolution of VBScript.

The following operators are available:

- **Bitshift Operators** (`<<` and `>>`) for logical bit shifting.
- **IsNot Operator** as the logical opposite of `Is` for object reference comparison.
- **AndAlso Operator** for short-circuit logical conjunction.
- **OrElse Operator** for short-circuit logical disjunction.

---

## Bitshift Operators (`<<` and `>>`)

### Overview
The bitshift operators perform a logical left shift (`<<`) or logical right shift (`>>`) on integer values. Vacated bits are filled with zero. If the shift amount equals or exceeds the operand's bit width (64 bits), the result is zero.

### Syntax

```vbscript
result = value << shiftAmount
result = value >> shiftAmount
```

### Parameters and Arguments
- **value** (Integer): The value to shift. Treated as a 64-bit unsigned integer.
- **shiftAmount** (Integer): The number of bit positions to shift. Must be non-negative.

### Return Values
- An **Integer** representing the shifted result.
- Returns **Null** if either operand is **Null**.

### Remarks
- Bitshift operators treat both operands as 64-bit integers. Floating-point values are coerced to integers before the operation.
- A shift amount of zero returns the original value unchanged.
- Shift amounts of 64 or greater always yield **0**.
- The right shift (`>>`) is a **logical** shift: vacated high-order bits are filled with zero.

### Code Example

```asp
<%
Dim a, b, c

' Left shift: 1 << 3 = 8
a = 1 << 3
Response.Write "1 << 3 = " & a & "<br>"

' Right shift: 16 >> 2 = 4
b = 16 >> 2
Response.Write "16 >> 2 = " & b & "<br>"

' Shift beyond bit width yields 0
c = 1 << 64
Response.Write "1 << 64 = " & c & "<br>"
%>
```

---

## IsNot Operator

### Overview
The `IsNot` operator is the logical opposite of the `Is` operator. It compares two object references and returns **True** if they refer to different objects, and **False** if they refer to the same object. It is syntactic sugar for `Not (obj1 Is obj2)`.

### Syntax

```vbscript
result = object1 IsNot object2
```

### Parameters and Arguments
- **object1** (Object): The first object reference to compare.
- **object2** (Object): The second object reference to compare.

### Return Values
- **True** if the two references point to different objects.
- **False** if the two references point to the same object.
- Raises a **Type mismatch** error if either operand is not a valid object reference.

### Remarks
- `IsNot` performs reference equality, not value equality.
- `obj IsNot Nothing` is **True** when `obj` holds a valid object reference, and **False** when `obj` is **Nothing**.
- Both `IsNot` (single keyword) and `Is Not` (two keywords, space-separated) are supported.

### Code Example

```asp
<%
Dim obj1, obj2, obj3

Set obj1 = Server.CreateObject("Scripting.Dictionary")
Set obj2 = obj1
Set obj3 = Server.CreateObject("Scripting.Dictionary")

' Same reference: IsNot returns False
Response.Write "obj1 IsNot obj2 = " & (obj1 IsNot obj2) & "<br>"

' Different references: IsNot returns True
Response.Write "obj1 IsNot obj3 = " & (obj1 IsNot obj3) & "<br>"

' Compared to Nothing
Set obj1 = Nothing
Response.Write "obj1 IsNot Nothing = " & (obj1 IsNot Nothing) & "<br>"
%>
```

---

## AndAlso Operator

### Overview
The `AndAlso` operator performs a short-circuit logical conjunction. The right-hand side (RHS) expression is evaluated only if the left-hand side (LHS) evaluates to **True**. If the LHS is **False**, the result is **False** and the RHS is never evaluated.

### Syntax

```vbscript
result = expression1 AndAlso expression2
```

### Parameters and Arguments
- **expression1** (Boolean): The left-hand side expression.
- **expression2** (Boolean): The right-hand side expression. Evaluated only when `expression1` is **True**.

### Return Values
- **True** if both expressions evaluate to **True**.
- **False** if either expression evaluates to **False**.

### Remarks
- Unlike the standard `And` operator, `AndAlso` guarantees that the RHS is not evaluated when the LHS is **False**. This prevents side effects and runtime errors from the RHS when the LHS already determines the result.
- The standard `And` operator always evaluates both sides and performs bitwise AND on numeric operands. `AndAlso` is purely logical and always returns a **Boolean**.
- Multiple `AndAlso` operators can be chained: `a AndAlso b AndAlso c`.

### Code Example

```asp
<%
Function IsValidUser(id)
    ' Simulate a check that might fail
    If id <= 0 Then
        IsValidUser = False
        Exit Function
    End If
    IsValidUser = True
End Function

Function GetUserName(id)
    ' This function is only safe to call when id > 0
    Dim names
    names = Array("Alice", "Bob", "Charlie")
    GetUserName = names(id - 1)
End Function

Dim userId, result

' Safe short-circuit: if userId is invalid, GetUserName is never called
userId = 0
result = IsValidUser(userId) AndAlso (GetUserName(userId) <> "")
Response.Write "Result: " & result & "<br>"
%>
```

---

## OrElse Operator

### Overview
The `OrElse` operator performs a short-circuit logical disjunction. The right-hand side (RHS) expression is evaluated only if the left-hand side (LHS) evaluates to **False**. If the LHS is **True**, the result is **True** and the RHS is never evaluated.

### Syntax

```vbscript
result = expression1 OrElse expression2
```

### Parameters and Arguments
- **expression1** (Boolean): The left-hand side expression.
- **expression2** (Boolean): The right-hand side expression. Evaluated only when `expression1` is **False**.

### Return Values
- **True** if either expression evaluates to **True**.
- **False** if both expressions evaluate to **False**.

### Remarks
- Unlike the standard `Or` operator, `OrElse` guarantees that the RHS is not evaluated when the LHS is **True**. This prevents unnecessary computation and avoids runtime errors from the RHS when the LHS already determines the result.
- The standard `Or` operator always evaluates both sides and performs bitwise OR on numeric operands. `OrElse` is purely logical and always returns a **Boolean**.
- Multiple `OrElse` operators can be chained: `a OrElse b OrElse c`.

### Code Example

```asp
<%
Function TryGetFromCache(key)
    ' Returns the value if found, otherwise Empty
    Dim cache
    Set cache = Server.CreateObject("Scripting.Dictionary")
    cache.Add "name", "Alice"
    
    If cache.Exists(key) Then
        TryGetFromCache = cache(key)
    Else
        TryGetFromCache = Empty
    End If
End Function

Function GetFromDatabase(key)
    ' Expensive database lookup
    Dim dbValue
    dbValue = "Alice" ' Simulated DB result
    GetFromDatabase = dbValue
End Function

Dim result

' Short-circuit: if value is found in cache, database is never queried
result = TryGetFromCache("name") OrElse GetFromDatabase("name")
Response.Write "Result: " & result & "<br>"
%>
```

---

## Compatibility Notes

- These operators are **AxonASP-specific extensions**. They are not available in standard Microsoft IIS / Classic ASP environments.
- The standard `And`, `Or`, and `Is` operators continue to work exactly as before. Code written for IIS remains fully compatible.
- All operators can be used inside expressions with other VBScript operators following the standard precedence rules.

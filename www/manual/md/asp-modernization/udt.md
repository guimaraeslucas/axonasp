# Define and Use VB6 User-Defined Types in ASP

## Overview
This page documents the AxonASP VB6 modernization support for User-Defined Types (UDT) declared with Type...End Type. It explains declaration, typed allocation with Dim ... As <UDT>, field access, and field assignment while preserving legacy ASP VBScript behavior for untyped Variant variables.

## Syntax
```vbscript
Type TypeName
    MemberName [As MemberType]
    MemberName2 [As MemberType]
End Type

Dim variableName As TypeName

variableName.MemberName = expression
result = variableName.MemberName
```

## Parameters and Arguments
- TypeName: Required. Name of the UDT declaration. The name is case-insensitive.
- MemberName: Required. Name of one field in the UDT layout. The name is case-insensitive.
- MemberType: Optional. Built-in VB6 supported types include Integer, Long, Single, Double, String, Boolean, Byte, Object, Variant, or another declared UDT name.
- variableName: Required. Variable declared with Dim ... As TypeName to allocate one UDT record instance.
- expression: Required for assignment statements. Value assigned to a UDT field.

## Return Values
- Type...End Type declaration: Returns no value. It registers one compile-time UDT layout.
- Dim variableName As TypeName: Returns no value. It allocates one UDT record in the typed variable slot.
- variableName.MemberName read: Returns the current value stored in the field.
- variableName.MemberName write: Returns no value. It updates the target field in place.

## Remarks
- UDT declarations are global-scope declarations.
- AxonASP stores UDTs as VTRecord values and executes member operations with fixed-index opcodes for fast access.
- Member reads use OpGetRecordMember and member writes use OpSetRecordMember.
- Record allocation uses OpInitRecord when a typed variable is declared as a UDT.
- Legacy ASP compatibility is preserved: untyped Dim variable declarations still behave as Variant.
- Type names and member names are resolved case-insensitively.
- Nested UDT fields are supported when a member is declared as another UDT.
- **Copy Semantics**: In alignment with VB6/VBA/VBScript specifications, UDTs are value types. Assigning a UDT to another UDT (or UDT member/array element) performs a deep memory copy of the data, rather than copying a reference. Modifying fields in the copy does not mutate fields in the original UDT structure.
- **Class Fields**: UDTs can be declared as class fields (e.g., `Private m_Pt As Point`). Within class methods, reading or writing members of class-encapsulated UDTs (e.g., `m_Pt.X = x`) is supported and compiled directly to record member offsets instead of invoking runtime object member dispatch.
- **Parameters and Return Types**: UDTs are supported as parameters and return types in Functions, Subs, and Class Methods (e.g., `Public Function Create() As Point`). If a function returns a UDT, its implicit return variable is pre-initialized as an empty UDT instance, allowing direct member assignments inside the body (e.g., `Create.X = x`).

## Code Example
```asp
<%
Type Address
    City As String
End Type

Type Person
    Name As String
    Age As Integer
    Home As Address
End Type

Dim addr As Address
Dim p As Person

addr.City = "Sao Paulo"
p.Name = "Maya"
p.Age = 29
p.Home = addr

Response.Write p.Name & "|" & p.Age & "|" & p.Home.City
%>
```

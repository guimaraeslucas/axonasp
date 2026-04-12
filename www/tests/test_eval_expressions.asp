<%
Option Explicit
Response.Write("Testing Eval() with Complex Expressions" & vbCrLf & vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 1: SIMPLE ARITHMETIC EXPRESSIONS" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

' Test 1.1: Basic arithmetic
Response.Write("1.1 - Basic arithmetic:" & vbCrLf)
Response.Write("  Eval('1 + 2') = " & Eval("1 + 2") & " (expected: 3)" & vbCrLf)
Response.Write("  Eval('10 - 3') = " & Eval("10 - 3") & " (expected: 7)" & vbCrLf)
Response.Write("  Eval('4 * 5') = " & Eval("4 * 5") & " (expected: 20)" & vbCrLf)
Response.Write("  Eval('20 / 4') = " & Eval("20 / 4") & " (expected: 5)" & vbCrLf)
Response.Write(vbCrLf)

' Test 1.2: Operator precedence
Response.Write("1.2 - Operator precedence:" & vbCrLf)
Response.Write("  Eval('2 + 3 * 4') = " & Eval("2 + 3 * 4") & " (expected: 14)" & vbCrLf)
Response.Write("  Eval('10 - 2 * 3') = " & Eval("10 - 2 * 3") & " (expected: 4)" & vbCrLf)
Response.Write("  Eval('(2 + 3) * 4') = " & Eval("(2 + 3) * 4") & " (expected: 20)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 2: STRING EXPRESSIONS" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Response.Write("2.1 - String literals:" & vbCrLf)
Response.Write("  Eval('""Hello""') = " & Eval("""Hello""") & " (expected: Hello)" & vbCrLf)
Response.Write("  Eval('""Test"" & "" String""') = " & Eval("""Test"" & "" String""") & " (expected: Test String)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 3: VARIABLE REFERENCES" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Dim x, y, z, name
x = 10
y = 20
z = 30
name = "John"

Response.Write("3.1 - Simple variables:" & vbCrLf)
Response.Write("  x = " & x & ", y = " & y & ", z = " & z & vbCrLf)
Response.Write("  Eval('x') = " & Eval("x") & " (expected: 10)" & vbCrLf)
Response.Write("  Eval('y') = " & Eval("y") & " (expected: 20)" & vbCrLf)
Response.Write("  Eval('x + y') = " & Eval("x + y") & " (expected: 30)" & vbCrLf)
Response.Write("  Eval('x + y + z') = " & Eval("x + y + z") & " (expected: 60)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("3.2 - String variables:" & vbCrLf)
Response.Write("  name = " & name & vbCrLf)
Response.Write("  Eval('name') = " & Eval("name") & " (expected: John)" & vbCrLf)
Response.Write("  Eval('name & "" Doe""') = " & Eval("name & "" Doe""") & " (expected: John Doe)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 4: ARRAY ACCESS IN EVAL" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Dim arr, numbers
arr = Array("apple", "banana", "cherry")
numbers = Array(10, 20, 30, 40, 50)

Response.Write("4.1 - Array indexing:" & vbCrLf)
Response.Write("  arr(0) = " & arr(0) & vbCrLf)
Response.Write("  Eval('arr(0)') = " & Eval("arr(0)") & " (expected: apple)" & vbCrLf)
Response.Write("  Eval('arr(1)') = " & Eval("arr(1)") & " (expected: banana)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("4.2 - Numeric arrays:" & vbCrLf)
Response.Write("  numbers(0) = " & numbers(0) & vbCrLf)
Response.Write("  Eval('numbers(0)') = " & Eval("numbers(0)") & " (expected: 10)" & vbCrLf)
Response.Write("  Eval('numbers(0) + numbers(1)') = " & Eval("numbers(0) + numbers(1)") & " (expected: 30)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 5: FUNCTION CALLS IN EVAL" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Response.Write("5.1 - Built-in functions:" & vbCrLf)
Response.Write("  Eval('Len(""Hello"")') = " & Eval("Len(""Hello"")") & " (expected: 5)" & vbCrLf)
Response.Write("  Eval('UCase(""hello"")') = " & Eval("UCase(""hello"")") & " (expected: HELLO)" & vbCrLf)
Response.Write("  Eval('LCase(""HELLO"")') = " & Eval("LCase(""HELLO"")") & " (expected: hello)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("5.2 - Nested functions:" & vbCrLf)
Response.Write("  Eval('Len(UCase(""test""))') = " & Eval("Len(UCase(""test""))") & " (expected: 4)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 6: COMPLEX EXPRESSIONS COMBINING EVERYTHING" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Dim price, quantity, tax_rate
price = 100
quantity = 3
tax_rate = 0.1

Response.Write("6.1 - Complex arithmetic:" & vbCrLf)
Response.Write("  price = " & price & ", quantity = " & quantity & ", tax_rate = " & tax_rate & vbCrLf)
Response.Write("  Eval('price * quantity') = " & Eval("price * quantity") & " (expected: 300)" & vbCrLf)
Response.Write("  Eval('price * quantity * (1 + tax_rate)') = " & Eval("price * quantity * (1 + tax_rate)") & " (expected: 330)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("6.2 - Mixed operations:" & vbCrLf)
Response.Write("  Eval('Len(name) + Len(""Doe"")') = " & Eval("Len(name) + Len(""Doe"")") & " (expected: 7)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 7: CONDITIONAL EXPRESSIONS" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Response.Write("7.1 - Comparison operators:" & vbCrLf)
Response.Write("  Eval('10 > 5') = " & Eval("10 > 5") & " (expected: True)" & vbCrLf)
Response.Write("  Eval('10 < 5') = " & Eval("10 < 5") & " (expected: False)" & vbCrLf)
Response.Write("  Eval('10 = 10') = " & Eval("10 = 10") & " (expected: True)" & vbCrLf)
Response.Write(vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("EVAL COMPLEX TESTS COMPLETED" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf)
%>

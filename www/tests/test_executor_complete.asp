<% 
  ' Test file for complete executor implementation
  ' Tests all major features: variables, loops, conditionals, functions

  ' Test 1: Basic variables and output
  Response.Write "<h2>Test 1: Basic Variables and Output</h2>"
  Dim x
  x = 42
  Response.Write "x = " & x & "<br>"

  ' Test 2: For loop
  Response.Write "<h2>Test 2: For Loop</h2>"
  Dim i
  For i = 1 To 5
    Response.Write "i = " & i & "<br>"
  Next

  ' Test 3: If-Else statement
  Response.Write "<h2>Test 3: If-Else Statement</h2>"
  Dim age
  age = 25
  If age >= 18 Then
    Response.Write "Age is " & age & " - Adult<br>"
  Else
    Response.Write "Age is " & age & " - Minor<br>"
  End If

  ' Test 4: Function definition and call
  Response.Write "<h2>Test 4: Function Definition and Call</h2>"
  Function Add(a, b)
    Add = a + b
  End Function
  
  Dim result
  result = Add(10, 20)
  Response.Write "Add(10, 20) = " & result & "<br>"

  ' Test 5: While loop
  Response.Write "<h2>Test 5: While Loop</h2>"
  Dim counter
  counter = 1
  While counter <= 3
    Response.Write "Counter: " & counter & "<br>"
    counter = counter + 1
  Wend

  ' Test 6: Do-Loop
  Response.Write "<h2>Test 6: Do-Loop</h2>"
  Dim num
  num = 1
  Do While num <= 3
    Response.Write "Number: " & num & "<br>"
    num = num + 1
  Loop

  ' Test 7: Select Case
  Response.Write "<h2>Test 7: Select Case</h2>"
  Dim grade
  grade = 85
  Select Case grade
    Case 90 To 100
      Response.Write "Grade A<br>"
    Case 80 To 89
      Response.Write "Grade B<br>"
    Case Else
      Response.Write "Grade F<br>"
  End Select

  ' Test 8: String concatenation and operations
  Response.Write "<h2>Test 8: String Concatenation and Operations</h2>"
  Dim firstName, lastName
  firstName = "John"
  lastName = "Doe"
  Response.Write "Full Name: " & firstName & " " & lastName & "<br>"

  ' Test 9: Array operations
  Response.Write "<h2>Test 9: Array Operations</h2>"
  Dim arr(2)
  arr(0) = "Apple"
  arr(1) = "Banana"
  arr(2) = "Cherry"
  Response.Write "arr(0) = " & arr(0) & "<br>"
  Response.Write "arr(1) = " & arr(1) & "<br>"
  Response.Write "arr(2) = " & arr(2) & "<br>"

  ' Test 10: Mathematical operations
  Response.Write "<h2>Test 10: Mathematical Operations</h2>"
  Dim a, b, c
  a = 10
  b = 3
  c = a + b
  Response.Write "10 + 3 = " & c & "<br>"
  c = a - b
  Response.Write "10 - 3 = " & c & "<br>"
  c = a * b
  Response.Write "10 * 3 = " & c & "<br>"
  c = a / b
  Response.Write "10 / 3 = " & c & "<br>"

  Response.Write "<h2>All Tests Completed Successfully!</h2>"
%>

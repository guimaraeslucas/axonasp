<%
Dim t
Dim emptyValue
Dim nullValue
Dim values
Dim zip

Set t = Server.CreateObject("G3TestSuite")
nullValue = Null
values = Array("a", "b", "c")
Set zip = Nothing

t.Describe "Core assertion primitives"
t.AssertEqual 4, 2 + 2, "2 + 2 should equal 4"
t.AssertNotEqual 4, 2 + 3, "2 + 3 should not equal 4"
t.AssertTrue Len("axon") = 4, "Len should return 4"
t.AssertFalse IsEmpty("ready"), "Literal text should not be Empty"
t.AssertEmpty emptyValue, "Undeclared Variant values should be Empty"
t.AssertNull nullValue, "Explicit Null values should be Null"
t.AssertNothing zip, "Set Nothing references should pass AssertNothing"
t.AssertTypeName "String", "hello", "TypeName(String) should match"
t.AssertLength 3, values, "Array length should be 3"
t.AssertCount 5, "Axon!", "String length should be 5"
%>
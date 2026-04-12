<%
Dim t

Set t = Server.CreateObject("G3TestSuite")

t.Describe "Error assertions"
t.AssertRaises "Err.Raise 13, ""testsuite"", ""type mismatch""", 13, "Err.Raise should preserve explicit error numbers"
t.AssertRaises "Err.Raise 13, ""testsuite"", ""type mismatch""", "type mismatch", "AssertRaises should match error description text"
t.AssertRaises "Function Broken(", "AssertRaises should trap dynamic syntax errors"
%>
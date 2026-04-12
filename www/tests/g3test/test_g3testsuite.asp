<%
Dim t
Set t = Server.CreateObject("G3TestSuite")

t.Describe "G3TestSuite smoke test"
t.AssertEqual 4, 2 + 2, "2 + 2 should equal 4"
t.AssertTrue Len("axon") = 4, "Len should report the expected size"
t.AssertFalse IsEmpty("ready"), "Literal text should not be Empty"
%>
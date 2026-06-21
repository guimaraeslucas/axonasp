<%@ Language=VBScript %>
<%
Option Compare Text

' Test 1: Case-insensitive equality
Dim test1 : test1 = ("ABC" = "abc")
Response.Write "Test1 (Eq Text): " & test1 & "<br>"

' Test 2: Case-insensitive inequality
Dim test2 : test2 = ("ABC" <> "def")
Response.Write "Test2 (Neq Text): " & test2 & "<br>"

' Test 3: Less than with text compare
Dim test3 : test3 = ("a" < "b")
Response.Write "Test3 (Lt Text): " & test3 & "<br>"

' Test 4: StrComp with default (should use Option Compare Text)
Dim test4 : test4 = (StrComp("HELLO", "hello") = 0)
Response.Write "Test4 (StrComp default=Text): " & test4 & "<br>"

' Test 5: InStr with default compare
Dim test5 : test5 = InStr("Hello World", "WORLD")
Response.Write "Test5 (InStr default=Text): " & test5 & "<br>"

' Test6: InStr with explicit vbTextCompare
Dim test6 : test6 = InStr(1, "Hello World", "WORLD", 1)
Response.Write "Test6 (InStr explicit Text): " & test6 & "<br>"

' Test7: Replace with default compare
Dim test7 : test7 = Replace("Hello hello", "HELLO", "X")
Response.Write "Test7 (Replace default=Text): " & test7 & "<br>"

' Test8: Replace with explicit vbTextCompare
Dim test8 : test8 = Replace("Hello hello", "hello", "X", 1, 1, 1)
Response.Write "Test8 (Replace explicit Text count=1): " & test8 & "<br>"

' Test9: Filter with default compare
Dim arr(2)
arr(0) = "Hello"
arr(1) = "WORLD"
arr(2) = "test"
Dim filtered : filtered = Filter(arr, "h", True, -1)
Response.Write "Test9 (Filter default=Text): " & Join(filtered, ",") & "<br>"

' Test10: InStrRev with default compare
Dim test10 : test10 = InStrRev("Hello hello HELLO", "hello")
Response.Write "Test10 (InStrRev default=Text): " & test10 & "<br>"

' Test11: Explicit binary compare still works
Dim test11 : test11 = InStr(1, "Hello World", "WORLD", 0)
Response.Write "Test11 (InStr explicit Binary): " & test11 & "<br>"

' Test12: Explicit binary StrComp still works
Dim test12 : test12 = StrComp("ABC", "abc", 0)
Response.Write "Test12 (StrComp explicit Binary): " & test12 & "<br>"

' All tests passed marker
Response.Write "ALL_TESTS_COMPLETED"
%>
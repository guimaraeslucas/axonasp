<%
Response.Write "Start test<br/>"

' First, test simple dim with colon
Response.Write "1. Testing: Dim a : a = 10<br/>"
Dim a : a = 10
Response.Write "a = " & a & "<br/><br/>"

' Then test function call
Response.Write "2. Testing function call<br/>"
Function getNum()
    getNum = 42
End Function

Dim b
b = getNum()
Response.Write "b = " & b & " (via getNum without colon)<br/>"

' Now test function call WITH colon
Response.Write "3. Testing function call with Dim colon<br/>"
Dim c : c = getNum()
Response.Write "c = " & c & " (via getNum WITH colon)<br/><br/>"

' Now test function with Dim colon inside
Response.Write "4. Testing function WITH Dim colon inside<br/>"
Function compute(n)
    Response.Write "  Inside compute: n = " & n & "<br/>"
    Dim x : x = n * 2
    Response.Write "  Inside compute: x = " & x & "<br/>"
    Dim y : y = n + 10
    Response.Write "  Inside compute: y = " & y & "<br/>"
    compute = x + y
    Response.Write "  Inside compute: returning " & (x + y) & "<br/>"
End Function

Dim result : result = compute(5)
Response.Write "result = " & result & " (expected 25)<br/>"
%>

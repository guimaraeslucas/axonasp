<%
Class MyClass
    Public Prop
End Class

Dim o
Set o = New MyClass
o.Prop = "hello"

on error resume next

' Test 1: Accessing property of object returned by Eval
Response.Write "Test 1: " & Eval("o").Prop & vbCrLf
if err.number <> 0 then
    Response.Write "Test 1 Failed: " & err.description & " (0x" & hex(err.number) & ")" & vbCrLf
    err.clear
end if

' Test 2: Assigning result of Eval to a variable using Set
Dim result
Set result = Eval("o")
if err.number <> 0 then
    Response.Write "Test 2 Failed: " & err.description & " (0x" & hex(err.number) & ")" & vbCrLf
    err.clear
else
    Response.Write "Test 2 Result Prop: " & result.Prop & vbCrLf
end if

' Test 3: New object in Eval
Response.Write "Test 3: " & Eval("new MyClass").Prop & vbCrLf
if err.number <> 0 then
    Response.Write "Test 3 Failed: " & err.description & " (0x" & hex(err.number) & ")" & vbCrLf
    err.clear
end if
%>

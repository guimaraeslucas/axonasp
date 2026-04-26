<%
' VBScript console test
console.Log("Hello from VBScript log")
console.info("Info message from VBScript")
console.warn("Warning from VBScript")
console.Error("Error from VBScript")

' Test with an array
Dim arr(2)
arr(0) = "one"
arr(1) = "two"
arr(2) = 3
console.Log(arr)
%>

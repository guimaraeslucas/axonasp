<%
' Test Class Lifecycle with Reference Counting
Option Explicit

Dim initializeCount, terminateCount

Class TestClass
    Public Sub Class_Initialize
        initializeCount = initializeCount + 1
        Response.Write "Class_Initialize called. Count: " & initializeCount & vbCrLf
    End Sub

    Public Sub Class_Terminate
        terminateCount = terminateCount + 1
        Response.Write "Class_Terminate called. Count: " & terminateCount & vbCrLf
    End Sub
End Class

' Initialize counters
initializeCount = 0
terminateCount = 0

Response.Write "=== Test 1: Simple Create and Destroy ===" & vbCrLf
Dim obj1
Set obj1 = New TestClass
Set obj1 = Nothing
Response.Write vbCrLf

Response.Write "=== Test 2: Multiple References ===" & vbCrLf
Dim obj2, obj3
Set obj2 = New TestClass
Set obj3 = obj2
Set obj2 = Nothing
Response.Write "obj2 set to Nothing" & vbCrLf
Set obj3 = Nothing
Response.Write "obj3 set to Nothing" & vbCrLf
Response.Write vbCrLf

Response.Write "=== Test 3: Array Assignment ===" & vbCrLf
Dim arr(2)
Set arr(0) = New TestClass
Set arr(1) = arr(0)
Set arr(0) = Nothing
Response.Write "arr(0) set to Nothing" & vbCrLf
Set arr(1) = Nothing
Response.Write "arr(1) set to Nothing" & vbCrLf
Response.Write vbCrLf

Response.Write "=== Test Results ===" & vbCrLf
Response.Write "Total Class_Initialize calls: " & initializeCount & vbCrLf
Response.Write "Total Class_Terminate calls: " & terminateCount & vbCrLf
Response.Write "Test completed successfully!" & vbCrLf
%>

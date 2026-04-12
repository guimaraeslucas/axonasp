<%
Option Explicit

' Test Global Variable Hoisting
x = 100
Response.Write "Global x before Dim: " & x & "<br>"
Dim x

' Test Global Sub Hoisting
Call GlobalSub()

Sub GlobalSub()
    Response.Write "Inside GlobalSub<br>"
    
    ' Test Local Variable Hoisting
    y = 200
    Response.Write "Local y before Dim: " & y & "<br>"
    Dim y
End Sub

' Test Global Function Hoisting
Response.Write "GlobalFunc result: " & GlobalFunc() & "<br>"

Function GlobalFunc()
    GlobalFunc = "Hello from Function"
End Function
%>

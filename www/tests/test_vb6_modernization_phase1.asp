<!--
    AxonASP Modernization Test Suite
    Phase 1: Advanced Arrays & Enumerations
-->
<%
Option Base 1

' Test Enumerations
Enum Colors
    Red
    Green
    Blue = 10
    Yellow
End Enum

Response.Write "<h3>VB6 Phase 1: Enumerations</h3>"
Response.Write "Red (0): " & Red & "<br>"
Response.Write "Green (1): " & Green & "<br>"
Response.Write "Blue (10): " & Blue & "<br>"
Response.Write "Yellow (11): " & Yellow & "<br>"

' Test Advanced Arrays with Option Base 1
Dim arr(5)
arr(1) = "First"
arr(5) = "Last"

Response.Write "<h3>VB6 Phase 1: Option Base 1 Arrays</h3>"
Response.Write "LBound (1): " & LBound(arr) & "<br>"
Response.Write "UBound (5): " & UBound(arr) & "<br>"
Response.Write "arr(1): " & arr(1) & "<br>"
Response.Write "arr(5): " & arr(5) & "<br>"

' Test Explicit Lower Bounds
Dim custom(10 To 20)
custom(10) = "Start"
custom(20) = "End"

Response.Write "<h3>VB6 Phase 1: Custom Lower Bound Arrays</h3>"
Response.Write "LBound (10): " & LBound(custom) & "<br>"
Response.Write "UBound (20): " & UBound(custom) & "<br>"
Response.Write "custom(10): " & custom(10) & "<br>"
Response.Write "custom(20): " & custom(20) & "<br>"

' Test Multi-dimensional with mixed bounds
Dim multi(1 To 2, 10 To 11)
multi(1, 10) = "A"
multi(2, 11) = "B"

Response.Write "<h3>VB6 Phase 1: Multi-dimensional Arrays</h3>"
Response.Write "LBound(1) (1): " & LBound(multi, 1) & "<br>"
Response.Write "UBound(1) (2): " & UBound(multi, 1) & "<br>"
Response.Write "LBound(2) (10): " & LBound(multi, 2) & "<br>"
Response.Write "UBound(2) (11): " & UBound(multi, 2) & "<br>"

' Test ReDim Preserve with Lower Bounds
Dim dyn()
ReDim dyn(1 To 2)
dyn(1) = "D1"
dyn(2) = "D2"
ReDim Preserve dyn(1 To 3)
dyn(3) = "D3"

Response.Write "<h3>VB6 Phase 1: ReDim Preserve</h3>"
Response.Write "LBound (1): " & LBound(dyn) & "<br>"
Response.Write "UBound (3): " & UBound(dyn) & "<br>"
Response.Write "dyn(1): " & dyn(1) & "<br>"
Response.Write "dyn(3): " & dyn(3) & "<br>"
%>

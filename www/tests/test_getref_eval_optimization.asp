<%
Option Explicit
Response.Write("Testing GetRef and Eval repeated execution" & vbCrLf & vbCrLf)

Function AddTwoNumbers(a, b)
    AddTwoNumbers = a + b
End Function

Sub EmitMessage(message)
    Response.Write("EmitMessage called with: " & message & vbCrLf)
End Sub

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 1: GETREF FUNCTION CALL" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf)

Dim fnRef
Set fnRef = GetRef("AddTwoNumbers")
Response.Write("GetRef('AddTwoNumbers')(7, 5) = " & fnRef(7, 5) & " (expected: 12)" & vbCrLf & vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 2: GETREF SUB INVOCATION" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf)

Dim subRef
Set subRef = GetRef("EmitMessage")
Call subRef("OK")
Response.Write("Sub reference invocation completed (expected: line above printed)" & vbCrLf & vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 3: EVAL REPEATED SAME EXPRESSION" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf)

Dim i, sumExpected, sumEval
sumExpected = 0
sumEval = 0

For i = 1 To 500
    sumExpected = sumExpected + 7
    sumEval = sumEval + Eval("7")
Next

Response.Write("Repeated Eval aggregate = " & sumEval & vbCrLf)
Response.Write("Expected aggregate      = " & sumExpected & vbCrLf)
Response.Write("Match                   = " & (sumEval = sumExpected) & " (expected: True)" & vbCrLf & vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("GETREF/EVAL TEST COMPLETED" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf)
%>

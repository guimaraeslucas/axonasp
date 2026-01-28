<%
Option Explicit

Response.Write "Testing Bitwise Semantics and Hex/Oct Formatting (32-bit vs 64-bit)" & "<br>"

' Test 1: Hex of -1 (Should be FFFFFFFF, not -1 or 64-bit FFFFFFFFFFFFFFFF)
Dim hexVal
hexVal = Hex(-1)
Response.Write "Hex(-1): Expecting 'FFFFFFFF', Got '" & hexVal & "'" 
If hexVal = "FFFFFFFF" Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
Response.Write "<br>"

' Test 2: Not 0 (Should be -1 in decimal)
Dim notVal
notVal = Not 0
Response.Write "Not 0: Expecting -1, Got " & notVal
If notVal = -1 Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
Response.Write "<br>"

' Test 3: Standard MD5-like rotate check
' RotateLeft(x, n) is usually (x << n) Or (x >> (32-n))
' Let's check a simple overflow case.
' 0x80000000 is -2147483648
Dim minInt
minInt = &H80000000
Response.Write "minInt (&H80000000): " & minInt & "<br>"

' Check And behavior
Dim resultAnd
resultAnd = minInt And &HFFFFFFFF
Response.Write "minInt And &HFFFFFFFF: Expecting " & minInt & ", Got " & resultAnd
If resultAnd = minInt Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
Response.Write "<br>"

' Check Or behavior
Dim resultOr
resultOr = minInt Or 0
Response.Write "minInt Or 0: Expecting " & minInt & ", Got " & resultOr
If resultOr = minInt Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
Response.Write "<br>"

' Check Hex of Negative Integer
Dim checkHexNeg
checkHexNeg = Hex(-2147483648)
Response.Write "Hex(-2147483648): Expecting '80000000', Got '" & checkHexNeg & "'"
If checkHexNeg = "80000000" Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
Response.Write "<br>"

' Check XOR
Dim xorRes
xorRes = -1 Xor -1
Response.Write "-1 Xor -1: Expecting 0, Got " & xorRes
If xorRes = 0 Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
Response.Write "<br>"

%>

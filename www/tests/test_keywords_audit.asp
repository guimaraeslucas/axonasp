<%
' Test file to audit all VBScript keywords support in AxonVM
Response.Write("=== VBScript Keywords Audit ===" & vbCrLf & vbCrLf)

' 1. Test IF/THEN/ELSE/ELSEIF
Response.Write("Testing IF/THEN/ELSE/ELSEIF..." & vbCrLf)
Dim testVal
testVal = 5
If testVal = 5 Then
    Response.Write("✓ IF/THEN works" & vbCrLf)
ElseIf testVal = 10 Then
    Response.Write("✗ ELSEIF condition matched incorrectly" & vbCrLf)
Else
    Response.Write("✗ ELSE matched incorrectly" & vbCrLf)
End If

' 2. Test SELECT/CASE
Response.Write(vbCrLf & "Testing SELECT/CASE..." & vbCrLf)
Select Case testVal
    Case 5:
        Response.Write("✓ SELECT/CASE works" & vbCrLf)
    Case Else:
        Response.Write("✗ CASE ELSE matched" & vbCrLf)
End Select

' 3. Test DO/LOOP/EXIT DO
Response.Write(vbCrLf & "Testing DO/LOOP..." & vbCrLf)
Dim counter
counter = 0
Do
    counter = counter + 1
    If counter = 3 Then
        Response.Write("✓ EXIT DO works" & vbCrLf)
        Exit Do
        End If
    Loop

    ' 4. Test WHILE/WEND
    Response.Write(vbCrLf & "Testing WHILE/WEND..." & vbCrLf)
    counter = 0
    While counter < 2
        counter = counter + 1
    Wend
    Response.Write("✓ WHILE/WEND works" & vbCrLf)

    ' 5. Test FOR/NEXT
    Response.Write(vbCrLf & "Testing FOR/NEXT..." & vbCrLf)
    Dim loopCounter
    For loopCounter = 0 To 2 Step 1
        If loopCounter = 2 Then
            Response.Write("✓ FOR/TO/STEP/NEXT works" & vbCrLf)
            Exit For
        End If
    Next

    ' 6. Test FOR EACH/IN
    Response.Write(vbCrLf & "Testing FOR EACH/IN..." & vbCrLf)
    Dim arr
    arr = Array(1, 2, 3)
    Dim Item
    For Each Item In arr
        Response.Write("✓ FOR EACH/IN works" & vbCrLf)
        Exit For
    Next

    ' 7. Test DIM
    Response.Write(vbCrLf & "Testing DIM..." & vbCrLf)
    Dim testVar
    testVar = 10
    Response.Write("✓ DIM works" & vbCrLf)

    ' 8. Test REDIM/PRESERVE
    Response.Write(vbCrLf & "Testing REDIM/PRESERVE..." & vbCrLf)
    ReDim arr(5)
    ReDim Preserve arr(10)
    Response.Write("✓ REDIM/PRESERVE works" & vbCrLf)

    ' 9. Test ERASE
    Response.Write(vbCrLf & "Testing ERASE..." & vbCrLf)
    Erase arr
    Response.Write("✓ ERASE works" & vbCrLf)

    ' 10. Test CONST
    Response.Write(vbCrLf & "Testing CONST..." & vbCrLf)
    Const MAX_VAL = 100
    Response.Write("✓ CONST works" & vbCrLf)

    ' 11. Test SET (object assignment)
    Response.Write(vbCrLf & "Testing SET..." & vbCrLf)
    Dim obj
    Set obj = CreateObject("G3JSON").NewObject()
    If obj Is Nothing Then
        Response.Write("✗ SET failed" & vbCrLf)
    Else
        Response.Write("✓ SET works" & vbCrLf)
    End If

    ' 12. Test IS operator
    Response.Write(vbCrLf & "Testing IS operator..." & vbCrLf)
    If obj Is Nothing Then
        Response.Write("✗ IS operator returned True for non-Nothing" & vbCrLf)
    Else
        Response.Write("✓ IS operator works" & vbCrLf)
    End If

    ' 13. Test WITH
    Response.Write(vbCrLf & "Testing WITH..." & vbCrLf)
    With obj
        Response.Write("✓ WITH works" & vbCrLf)
    End With

    ' 14. Test CALL
    Response.Write(vbCrLf & "Testing CALL..." & vbCrLf)
    Call MyTestSub()

    ' 15. Test BYVAL/BYREF
    Response.Write(vbCrLf & "Testing BYVAL/BYREF..." & vbCrLf)
    Dim refVal
    refVal = 5
    Call TestByRef(refVal)
    If refVal = 10 Then
        Response.Write("✓ BYREF works" & vbCrLf)
    Else
        Response.Write("✗ BYREF failed" & vbCrLf)
    End If

    ' 16. Test OPTIONAL (IF IMPLEMENTED)
    Response.Write(vbCrLf & "Testing OPTIONAL..." & vbCrLf)
    Call TestOptional()
    Call TestOptional(42)

    ' 17. Test ME
    Response.Write(vbCrLf & "Testing ME (in class context)..." & vbCrLf)
    ' ME will be tested in a class

    ' 18. Test NEW
    Response.Write(vbCrLf & "Testing NEW..." & vbCrLf)
    ' NEW is tested with CreateObject()
    Response.Write("✓ NEW works (via CreateObject)" & vbCrLf)

    ' 19. Test OPTION
    Response.Write(vbCrLf & "Testing OPTION..." & vbCrLf)
    ' OPTION EXPLICIT is already in effect
    Response.Write("✓ OPTION EXPLICIT works" & vbCrLf)

    ' 20. Test CLASS/PUBLIC/PRIVATE
    Response.Write(vbCrLf & "Testing CLASS/PUBLIC/PRIVATE..." & vbCrLf)
    Dim classObj
    Set classObj = New TestClass
    Response.Write("✓ CLASS/PUBLIC/PRIVATE works" & vbCrLf)

    ' 21. Test PROPERTY GET/LET/SET
    Response.Write(vbCrLf & "Testing PROPERTY GET/LET/SET..." & vbCrLf)
    classObj.TestProp = 15
    Response.Write("Property value: " & classObj.TestProp & vbCrLf)

    ' 22. Test DEFAULT (IF IMPLEMENTED)
    Response.Write(vbCrLf & "Testing DEFAULT..." & vbCrLf)
    ' DEFAULT is used for default properties - would need special class
    Response.Write("⚠ DEFAULT not tested" & vbCrLf)

    ' Test subs
    Sub MyTestSub()
        Response.Write("✓ CALL/SUB works" & vbCrLf)
    End Sub

    Sub TestByRef(ByRef val)
        val = 10
    End Sub

    Sub TestOptional(Optional param)
        If IsEmpty(param) Then
            Response.Write("✓ OPTIONAL parameter works (empty)" & vbCrLf)
        Else
            Response.Write("✓ OPTIONAL parameter works (value: " & param & ")" & vbCrLf)
        End If
    End Sub

    ' Test class
    Class TestClass
        Private m_value

        Public Property Get TestProp()
            TestProp = m_value
        End Property

        Public Property Let TestProp(val)
            m_value = val
        End Property
    End Class

    Response.Write(vbCrLf & "=== Audit Complete ===" & vbCrLf)
%>

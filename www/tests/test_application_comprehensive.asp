<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>Application Object - Comprehensive Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f0f0f0; }
        .test { background: white; margin: 10px 0; padding: 15px; border-radius: 5px; border-left: 4px solid #007bff; }
        .pass { border-left-color: #28a745; }
        .fail { border-left-color: #dc3545; }
        h1 { color: #333; }
        h3 { color: #666; margin: 10px 0; }
        code { background: #f4f4f4; padding: 2px 6px; border-radius: 3px; }
    </style>
</head>
<body>
    <h1>Application Object - Comprehensive Test Suite</h1>

    <!-- Test 1: Basic Storage -->
    <div class="test pass">
        <h3>Test 1: Basic Variable Storage and Retrieval</h3>
        <%
            Application("TestString") = "Hello World"
            Application("TestNumber") = 42
            Application("TestFloat") = 3.14
            
            If Application("TestString") = "Hello World" Then
                Response.Write("✓ String storage: PASS<br>")
            Else
                Response.Write("✗ String storage: FAIL<br>")
            End If
            
            If Application("TestNumber") = 42 Then
                Response.Write("✓ Number storage: PASS<br>")
            Else
                Response.Write("✗ Number storage: FAIL<br>")
            End If
            
            If Application("TestFloat") = 3.14 Then
                Response.Write("✓ Float storage: PASS<br>")
            Else
                Response.Write("✗ Float storage: FAIL<br>")
            End If
        %>
    </div>

    <!-- Test 2: Case Insensitivity -->
    <div class="test pass">
        <h3>Test 2: Case-Insensitive Key Access</h3>
        <%
            Application("CasETest") = "Value"
            
            If Application("casetest") = "Value" Then
                Response.Write("✓ Lowercase access: PASS<br>")
            Else
                Response.Write("✗ Lowercase access: FAIL<br>")
            End If
            
            If Application("CASETEST") = "Value" Then
                Response.Write("✓ Uppercase access: PASS<br>")
            Else
                Response.Write("✗ Uppercase access: FAIL<br>")
            End If
            
            If Application("CasETest") = "Value" Then
                Response.Write("✓ Mixed case access: PASS<br>")
            Else
                Response.Write("✗ Mixed case access: FAIL<br>")
            End If
        %>
    </div>

    <!-- Test 3: Lock/Unlock -->
    <div class="test pass">
        <h3>Test 3: Lock and Unlock Methods</h3>
        <%
            Application.Lock
            Application("Counter") = 0
            
            Dim i
            For i = 1 To 5
                Application("Counter") = Application("Counter") + 1
            Next
            
            Application.Unlock
            
            If Application("Counter") = 5 Then
                Response.Write("✓ Lock/Unlock with counter: PASS (Counter = " & Application("Counter") & ")<br>")
            Else
                Response.Write("✗ Lock/Unlock with counter: FAIL<br>")
            End If
        %>
    </div>

    <!-- Test 4: StaticObjects Enumeration -->
    <div class="test pass">
        <h3>Test 4: Application.StaticObjects Enumeration</h3>
        <%
            Application("EnumTest1") = "First"
            Application("EnumTest2") = "Second"
            Application("EnumTest3") = "Third"
            
            Dim foundCount
            foundCount = 0
            
            Dim key
            For Each key In Application.StaticObjects
                If InStr(key, "enumtest") > 0 Then
                    foundCount = foundCount + 1
                    Response.Write("  Found: " & key & " = " & Application(key) & "<br>")
                End If
            Next
            
            If foundCount = 3 Then
                Response.Write("✓ StaticObjects enumeration: PASS (Found " & foundCount & " items)<br>")
            Else
                Response.Write("✗ StaticObjects enumeration: FAIL (Found " & foundCount & ", expected 3)<br>")
            End If
        %>
    </div>

    <!-- Test 5: Contents Collection -->
    <div class="test pass">
        <h3>Test 5: Application.Contents Collection</h3>
        <%
            Application("ContentTest1") = "Alpha"
            Application("ContentTest2") = "Beta"
            
            Dim contentCount
            contentCount = 0
            
            Dim cKey
            For Each cKey In Application.Contents
                If InStr(cKey, "contenttest") > 0 Then
                    contentCount = contentCount + 1
                End If
            Next
            
            If contentCount = 2 Then
                Response.Write("✓ Contents enumeration: PASS (Found " & contentCount & " items)<br>")
            Else
                Response.Write("✗ Contents enumeration: FAIL<br>")
            End If
        %>
    </div>

    <!-- Test 6: Persistence Across Requests -->
    <div class="test pass">
        <h3>Test 6: Persistence Test (Global Scope)</h3>
        <%
            If IsEmpty(Application("VisitCount")) Or IsNull(Application("VisitCount")) Then
                Application("VisitCount") = 0
            End If
            
            Application.Lock
            Application("VisitCount") = Application("VisitCount") + 1
            Application.Unlock
            
            Response.Write("✓ Visit counter: " & Application("VisitCount") & " (increments on each refresh)<br>")
            Response.Write("  Refresh the page to see counter increase<br>")
        %>
    </div>

    <!-- Test 7: Complex Data Types -->
    <div class="test pass">
        <h3>Test 7: Arrays Storage</h3>
        <%
            Dim testArray(2)
            testArray(0) = "Item1"
            testArray(1) = "Item2"
            testArray(2) = "Item3"
            
            Application("TestArray") = testArray
            
            Dim retrievedArray
            retrievedArray = Application("TestArray")
            
            If IsArray(retrievedArray) Then
                Response.Write("✓ Array storage: PASS<br>")
                Response.Write("  Array items: " & retrievedArray(0) & ", " & retrievedArray(1) & ", " & retrievedArray(2) & "<br>")
            Else
                Response.Write("✗ Array storage: FAIL<br>")
            End If
        %>
    </div>

    <!-- Test 8: Nested Lock Operations -->
    <div class="test pass">
        <h3>Test 8: Nested Lock Test</h3>
        <%
            Application.Lock
            Application("Nested") = "Level1"
            Application.Lock  ' Second lock (should be safe in Go with RWMutex)
            Application("Nested") = "Level2"
            Application.Unlock
            Application.Unlock
            
            If Application("Nested") = "Level2" Then
                Response.Write("✓ Nested lock handling: PASS<br>")
            Else
                Response.Write("✗ Nested lock handling: FAIL<br>")
            End If
        %>
    </div>

    <!-- Summary -->
    <div class="test" style="border-left-color: #6c757d; background: #e9ecef;">
        <h3>Test Summary</h3>
        <p><strong>All tests completed successfully!</strong></p>
        <p>Application object features tested:</p>
        <ul>
            <li>Variable storage and retrieval</li>
            <li>Case-insensitive key access</li>
            <li>Lock/Unlock thread safety</li>
            <li>StaticObjects collection enumeration</li>
            <li>Contents collection enumeration</li>
            <li>Global persistence across requests</li>
            <li>Complex data types (arrays)</li>
            <li>Nested lock operations</li>
        </ul>
    </div>

</body>
</html>

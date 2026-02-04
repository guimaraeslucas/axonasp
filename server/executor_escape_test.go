package server

import (
	"fmt"
	"net/http/httptest"
	"strings"
	"testing"
)

// Helper function to execute ASP code and return output
func executeASP(code string) (string, error) {
	recorder := httptest.NewRecorder()
	req := httptest.NewRequest("GET", "/test.asp", nil)

	config := &ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 5,
		DebugASP:      false,
	}

	executor := NewASPExecutor(config)

	err := executor.Execute(code, "test.asp", recorder, req, "TESTSESSION")

	if err != nil {
		return "", err
	}

	return recorder.Body.String(), nil
}

// TestFunctionReturnAssignment tests that `funcName = value` inside a function
// assigns the return value and does NOT call the function recursively
func TestFunctionReturnAssignment(t *testing.T) {
	tests := []struct {
		name     string
		code     string
		expected string
	}{
		{
			name: "Simple function return",
			code: `<%
Function TestFunc()
    TestFunc = "hello"
End Function
Response.Write TestFunc()
%>`,
			expected: "hello",
		},
		{
			name: "Function with parameter return",
			code: `<%
Function Echo(val)
    Echo = val
End Function
Response.Write Echo("world")
%>`,
			expected: "world",
		},
		{
			name: "Function return with concatenation",
			code: `<%
Function Wrap(val)
    Wrap = "[" & val & "]"
End Function
Response.Write Wrap("test")
%>`,
			expected: "[test]",
		},
		{
			name: "Nested function calls",
			code: `<%
Function Inner(val)
    Inner = val & "!"
End Function
Function Outer(val)
    Outer = Inner(val)
End Function
Response.Write Outer("hi")
%>`,
			expected: "hi!",
		},
		{
			name: "Function return self-reference should NOT recurse",
			code: `<%
Dim callCount
callCount = 0

Function Escape(val)
    callCount = callCount + 1
    If callCount > 5 Then
        Escape = "ERROR:RECURSION"
        Exit Function
    End If
    Escape = val & "-escaped"
End Function

Response.Write Escape("test")
Response.Write "|count:" & callCount
%>`,
			expected: "test-escaped|count:1",
		},
		{
			name: "Class method return assignment",
			code: `<%
Class JsonHelper
    Public Function Escape(val)
        Escape = "[" & val & "]"
    End Function
End Class

Dim helper
Set helper = New JsonHelper
Response.Write helper.Escape("data")
%>`,
			expected: "[data]",
		},
		{
			name: "Class method with self-reference",
			code: `<%
Dim methodCallCount
methodCallCount = 0

Class JsonHelper
    Public Function Escape(val)
        methodCallCount = methodCallCount + 1
        If methodCallCount > 5 Then
            Escape = "ERROR:RECURSION"
            Exit Function
        End If
        Escape = val & "-escaped"
    End Function
End Class

Dim helper
Set helper = New JsonHelper
Response.Write helper.Escape("test")
Response.Write "|count:" & methodCallCount
%>`,
			expected: "test-escaped|count:1",
		},
		{
			name: "Nested class with escape pattern like asplite",
			code: `<%
Dim escapeCallCount
escapeCallCount = 0

Class ASPL
    Public json
    
    Private Sub Class_Initialize()
        Set json = New JsonClass
    End Sub
End Class

Class JsonClass
    Public Function Escape(val)
        escapeCallCount = escapeCallCount + 1
        If escapeCallCount > 5 Then
            Escape = "ERROR:RECURSION:" & escapeCallCount
            Exit Function
        End If
        ' Simulate some processing
        Escape = "[" & val & "]"
    End Function
End Class

Dim aspl
Set aspl = New ASPL
Response.Write aspl.json.Escape("hello")
Response.Write "|count:" & escapeCallCount
%>`,
			expected: "[hello]|count:1",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := executeASP(tt.code)

			if err != nil {
				t.Errorf("Execute error: %v", err)
				return
			}

			if result != tt.expected {
				t.Errorf("Expected %q, got %q", tt.expected, result)
			}
		})
	}
}

// TestEscapeFunctionRecursionDetection specifically tests if the escape pattern
// in asplite causes infinite recursion
func TestEscapeFunctionRecursionDetection(t *testing.T) {
	code := `<%
Dim globalCallCount
globalCallCount = 0

Class ASPL
    Public json
    
    Private Sub Class_Initialize()
        Set json = New ASPLJSON
    End Sub
End Class

Class ASPLJSON
    Public Function Escape(val)
        globalCallCount = globalCallCount + 1
        
        ' Detect recursion
        If globalCallCount > 10 Then
            Escape = "RECURSION_DETECTED"
            Exit Function
        End If
        
        ' Mimic the real escape function pattern
        If IsNull(val) Or IsEmpty(val) Then
            Escape = ""
            Exit Function
        End If
        
        Dim result
        result = CStr(val)
        
        ' Simple replacement simulation
        result = Replace(result, "\", "\\")
        result = Replace(result, """", "\""")
        
        ' This is the critical line - assigning to function name
        Escape = result
    End Function
End Class

Dim aspl
Set aspl = New ASPL

' Test multiple calls
Dim output
output = ""
output = output & aspl.json.Escape("test1") & "|"
output = output & aspl.json.Escape("test2") & "|"
output = output & aspl.json.Escape("test3")

Response.Write output
Response.Write "|totalCalls:" & globalCallCount
%>`

	result, err := executeASP(code)

	if err != nil {
		t.Errorf("Execute error: %v", err)
		return
	}

	fmt.Printf("Result: %s\n", result)

	// Should have exactly 3 calls
	if !strings.Contains(result, "totalCalls:3") {
		t.Errorf("Expected exactly 3 calls to Escape, got result: %s", result)
	}

	// Should NOT contain recursion detection
	if strings.Contains(result, "RECURSION_DETECTED") {
		t.Errorf("Recursion detected! Result: %s", result)
	}
}

// TestFunctionNameAssignmentVsCall tests the specific case where
// the parser/executor might confuse function return assignment with a call
func TestFunctionNameAssignmentVsCall(t *testing.T) {
	code := `<%
Dim assignmentCount, callCount

assignmentCount = 0
callCount = 0

Function MyFunc(val)
    callCount = callCount + 1
    
    ' When we write "MyFunc = ...", this should be an assignment, not a call
    ' The executor should recognize this is inside MyFunc and treat it as return value
    assignmentCount = assignmentCount + 1
    MyFunc = val & "-processed"
End Function

Dim result
result = MyFunc("input")

Response.Write "result:" & result
Response.Write "|calls:" & callCount
Response.Write "|assignments:" & assignmentCount
%>`

	result, err := executeASP(code)

	if err != nil {
		t.Errorf("Execute error: %v", err)
		return
	}

	fmt.Printf("Result: %s\n", result)

	// Should have exactly 1 call
	if !strings.Contains(result, "calls:1") {
		t.Errorf("Expected exactly 1 call, got: %s", result)
	}

	// Result should be correct
	if !strings.Contains(result, "result:input-processed") {
		t.Errorf("Expected correct result, got: %s", result)
	}
}

// TestClassMethodReturnAssignment tests class method return assignment
func TestClassMethodReturnAssignment(t *testing.T) {
	code := `<%
Dim methodCallCount
methodCallCount = 0

Class MyClass
    Public Function Process(val)
        methodCallCount = methodCallCount + 1
        
        ' This should be return value assignment, not recursive call
        Process = "processed:" & val
    End Function
End Class

Dim obj
Set obj = New MyClass

Dim r1, r2
r1 = obj.Process("a")
r2 = obj.Process("b")

Response.Write r1 & "|" & r2 & "|calls:" & methodCallCount
%>`

	result, err := executeASP(code)

	if err != nil {
		t.Errorf("Execute error: %v", err)
		return
	}

	fmt.Printf("Result: %s\n", result)

	expected := "processed:a|processed:b|calls:2"
	if result != expected {
		t.Errorf("Expected %q, got %q", expected, result)
	}
}

// TestFunctionSelfConcatenation tests the pattern "funcName = funcName & value"
// which is commonly used to build strings inside functions (like asplite's escape function)
// This MUST NOT cause recursion - funcName on the right side should read the current return value
func TestFunctionSelfConcatenation(t *testing.T) {
	code := `<%
Dim callCount
callCount = 0

Function BuildString(val)
    callCount = callCount + 1
    If callCount > 5 Then
        BuildString = "ERROR:RECURSION"
        Exit Function
    End If
    
    Dim i
    For i = 1 to Len(val)
        ' This is the critical pattern! BuildString = BuildString & char
        BuildString = BuildString & Mid(val, i, 1) & "-"
    Next
End Function

Response.Write BuildString("ABC")
Response.Write "|count:" & callCount
%>`

	result, err := executeASP(code)

	if err != nil {
		t.Errorf("Execute error: %v", err)
		return
	}

	fmt.Printf("Result: %s\n", result)

	// Should have exactly 1 call
	if !strings.Contains(result, "count:1") {
		t.Errorf("Expected exactly 1 call, got: %s", result)
	}

	// Result should be "A-B-C-"
	if !strings.Contains(result, "A-B-C-") {
		t.Errorf("Expected 'A-B-C-' result, got: %s", result)
	}

	// Should NOT contain recursion error
	if strings.Contains(result, "RECURSION") {
		t.Errorf("Recursion detected! Result: %s", result)
	}
}

// TestClassMethodSelfConcatenation tests the pattern inside a class method
// This mimics asplite's JSON.escape() function which does: escape = escape & currentDigit
func TestClassMethodSelfConcatenation(t *testing.T) {
	code := `<%
Dim callCount
callCount = 0

Class JsonHelper
    Public Function Escape(val)
        callCount = callCount + 1
        If callCount > 5 Then
            Escape = "ERROR:RECURSION"
            Exit Function
        End If
        
        Dim i, currentDigit
        For i = 1 to Len(val)
            currentDigit = Mid(val, i, 1)
            ' This is the critical pattern from asplite!
            Escape = Escape & "[" & currentDigit & "]"
        Next
    End Function
End Class

Dim helper
Set helper = New JsonHelper

Response.Write helper.Escape("AB")
Response.Write "|count:" & callCount
%>`

	result, err := executeASP(code)

	if err != nil {
		t.Errorf("Execute error: %v", err)
		return
	}

	fmt.Printf("Result: %s\n", result)

	// Should have exactly 1 call
	if !strings.Contains(result, "count:1") {
		t.Errorf("Expected exactly 1 call, got: %s", result)
	}

	// Result should be "[A][B]"
	if !strings.Contains(result, "[A][B]") {
		t.Errorf("Expected '[A][B]' result, got: %s", result)
	}

	// Should NOT contain recursion error
	if strings.Contains(result, "RECURSION") {
		t.Errorf("Recursion detected! Result: %s", result)
	}
}

// TestEscapeWithNilArgument tests calling escape on a nil/Nothing value
// This simulates calling aspl.json.escape(rs("field")) when rs is Nothing/EOF
func TestEscapeWithNilArgument(t *testing.T) {
	code := `<%
Dim callCount
callCount = 0

Class JsonHelper
    Public Function Escape(val)
        callCount = callCount + 1
        If callCount > 5 Then
            Escape = "ERROR:RECURSION"
            Exit Function
        End If
        
        ' Handle nil/empty - this is what asplite does
        If IsNull(val) Or IsEmpty(val) Then
            Escape = ""
            Exit Function
        End If
        
        Escape = "[" & CStr(val) & "]"
    End Function
End Class

Dim helper
Set helper = New JsonHelper

' Test with a nil variable (simulates rs("field") returning nil when rs is EOF)
Dim nilVar
Response.Write helper.Escape(nilVar)
Response.Write "|count:" & callCount
%>`

	result, err := executeASP(code)

	if err != nil {
		t.Errorf("Execute error: %v", err)
		return
	}

	fmt.Printf("Result: %s\n", result)

	// Should have exactly 1 call
	if !strings.Contains(result, "count:1") {
		t.Errorf("Expected exactly 1 call, got: %s", result)
	}

	// Should NOT contain recursion error
	if strings.Contains(result, "RECURSION") {
		t.Errorf("Recursion detected! Result: %s", result)
	}
}

// TestDoWhileWithNilObject tests a Do While loop checking an object's property
// when the object is nil/Nothing
func TestDoWhileWithNilObject(t *testing.T) {
	code := `<%
Dim loopCount
loopCount = 0

' Simulate a nil recordset
Dim rs
rs = Nothing

' This should NOT cause an infinite loop
' When rs is Nothing, "Not rs.eof" should either:
' 1. Throw an error (can't access property of Nothing)
' 2. Or evaluate to False (since Nothing has no properties)
Do While Not rs Is Nothing
    loopCount = loopCount + 1
    If loopCount > 5 Then
        Response.Write "ERROR:INFINITE_LOOP"
        Exit Do
    End If
Loop

Response.Write "loopCount:" & loopCount
%>`

	result, err := executeASP(code)

	if err != nil {
		// An error is acceptable - accessing property of Nothing should fail
		fmt.Printf("Got expected error: %v\n", err)
		return
	}

	fmt.Printf("Result: %s\n", result)

	// Should have 0 iterations - rs is Nothing so the loop shouldn't run
	if !strings.Contains(result, "loopCount:0") {
		t.Errorf("Expected 0 loop iterations, got: %s", result)
	}

	// Should NOT contain infinite loop error
	if strings.Contains(result, "INFINITE_LOOP") {
		t.Errorf("Infinite loop detected! Result: %s", result)
	}
}

// TestDoWhileNotRsEof tests the classic "Do While Not rs.eof" pattern
// with a mock recordset object
func TestDoWhileNotRsEof(t *testing.T) {
	code := `<%
Dim loopCount
loopCount = 0

' Create a mock recordset class that simulates EOF
Class MockRecordset
    Private mEOF
    Private mCallCount
    
    Private Sub Class_Initialize()
        mEOF = False
        mCallCount = 0
    End Sub
    
    Public Property Get EOF()
        mCallCount = mCallCount + 1
        If mCallCount > 3 Then
            mEOF = True
        End If
        EOF = mEOF
    End Property
    
    Public Sub MoveNext()
    End Sub
End Class

Dim rs
Set rs = New MockRecordset

Do While Not rs.EOF
    loopCount = loopCount + 1
    Response.Write "row" & loopCount & "|"
    
    If loopCount > 10 Then
        Response.Write "ERROR:INFINITE_LOOP"
        Exit Do
    End If
    
    rs.MoveNext
Loop

Response.Write "total:" & loopCount
%>`

	result, err := executeASP(code)

	if err != nil {
		t.Errorf("Execute error: %v", err)
		return
	}

	fmt.Printf("Result: %s\n", result)

	// Should NOT contain infinite loop error
	if strings.Contains(result, "INFINITE_LOOP") {
		t.Errorf("Infinite loop detected! Result: %s", result)
	}

	// Should eventually exit (3 iterations based on our mock)
	if !strings.Contains(result, "total:") {
		t.Errorf("Loop didn't complete properly, got: %s", result)
	}
}

// TestNotNilExpression tests how "Not Nothing" is evaluated
func TestNotNilExpression(t *testing.T) {
	code := `<%
Dim result

If Not Nothing Then
    result = "Not Nothing is truthy"
Else
    result = "Not Nothing is falsy"
End If

Response.Write result
%>`

	result, err := executeASP(code)

	if err != nil {
		t.Errorf("Execute error: %v", err)
		return
	}

	fmt.Printf("Result for 'Not Nothing': %s\n", result)
}

// TestPropertyAccessOnNothing tests what happens when accessing a property on Nothing
func TestPropertyAccessOnNothing(t *testing.T) {
	code := `<%
On Error Resume Next

Dim rs
Set rs = Nothing

' This should either error or return something falsy
Dim result
result = rs.EOF

If Err.Number <> 0 Then
    Response.Write "Error:" & Err.Description
Else
    Response.Write "Value:" & result
End If
%>`

	result, err := executeASP(code)
	fmt.Printf("Property access on Nothing: %s, err=%v\n", result, err)
}

// TestDoWhileNotRsEofWithNothing tests the exact scenario from ebook.asp
// where rs becomes Nothing (from a failed GetRows) and the loop keeps running
func TestDoWhileNotRsEofWithNothing(t *testing.T) {
	code := `<%
Dim loopCount
loopCount = 0

' Simulate rs being Nothing (what happens when GetRows fails)
Dim rs
Set rs = Nothing

' This is the actual loop from ebook.asp:
' Do While Not rs.eof
' The problem: when rs is Nothing, "rs.eof" returns nil/Nothing
' And "Not Nothing" returns -1 (truthy), causing infinite loop!

On Error Resume Next

Do While Not rs.eof
    loopCount = loopCount + 1
    Response.Write "iter:" & loopCount & "|"
    
    If loopCount > 5 Then
        Response.Write "ERROR:INFINITE_LOOP"
        Exit Do
    End If
    
    rs.MoveNext
Loop

Response.Write "total:" & loopCount
%>`

	result, err := executeASP(code)

	fmt.Printf("Result: %s, err=%v\n", result, err)

	// The test should either:
	// 1. Error (can't access property of Nothing)
	// 2. Or NOT loop infinitely
	if strings.Contains(result, "INFINITE_LOOP") {
		t.Errorf("FOUND THE BUG! Infinite loop when rs is Nothing! Result: %s", result)
	}
}

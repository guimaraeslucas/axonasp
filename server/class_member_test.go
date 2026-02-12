package server

import (
	"net/http/httptest"
	"strings"
	"testing"
)

// TestClassMemberAssignment tests that class public member variables retain values
// assigned inside class methods. This is the core issue with sOrderBY being lost.
func TestClassMemberAssignment(t *testing.T) {
	aspCode := `<%
Class cls_test
    Public sName
    Public sValue
    Public iCount

    Public Sub SetValues()
        sName = "TestName"
        sValue = "TestValue"
        iCount = 42
    End Sub
End Class

Dim obj
Set obj = New cls_test
obj.SetValues

If obj.sName = "TestName" Then
    Response.Write "sName=OK "
Else
    Response.Write "sName=FAIL(" & TypeName(obj.sName) & ":" & obj.sName & ") "
End If

If obj.sValue = "TestValue" Then
    Response.Write "sValue=OK "
Else
    Response.Write "sValue=FAIL(" & TypeName(obj.sValue) & ":" & obj.sValue & ") "
End If

If obj.iCount = 42 Then
    Response.Write "iCount=OK"
Else
    Response.Write "iCount=FAIL(" & TypeName(obj.iCount) & ":" & obj.iCount & ")"
End If
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-1")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "sName=OK") {
		t.Errorf("sName not retained after SetValues(): %s", body)
	}
	if !strings.Contains(body, "sValue=OK") {
		t.Errorf("sValue not retained after SetValues(): %s", body)
	}
	if !strings.Contains(body, "iCount=OK") {
		t.Errorf("iCount not retained after SetValues(): %s", body)
	}
}

// TestClassMemberAssignmentWithFunction tests with a Function (not Sub)
// and with ON Error Resume Next like the real pick() function.
func TestClassMemberAssignmentWithFunction(t *testing.T) {
	aspCode := `<%
Class cls_page
    Public iId
    Public sTitle
    Public sOrderBY

    Public Function Pick(id)
        ON Error Resume Next
        iId = id
        sTitle = "MyTitle"
        sOrderBY = "sTitle"
    End Function
End Class

Dim selectedPage
Set selectedPage = New cls_page
selectedPage.Pick(492)

If selectedPage.iId = 492 Then
    Response.Write "iId=OK "
Else
    Response.Write "iId=FAIL(" & TypeName(selectedPage.iId) & ":" & selectedPage.iId & ") "
End If

If selectedPage.sTitle = "MyTitle" Then
    Response.Write "sTitle=OK "
Else
    Response.Write "sTitle=FAIL(" & TypeName(selectedPage.sTitle) & ":" & selectedPage.sTitle & ") "
End If

If selectedPage.sOrderBY = "sTitle" Then
    Response.Write "sOrderBY=OK"
Else
    Response.Write "sOrderBY=FAIL(" & TypeName(selectedPage.sOrderBY) & ":" & selectedPage.sOrderBY & ")"
End If
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-2")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "iId=OK") {
		t.Errorf("iId not retained: %s", body)
	}
	if !strings.Contains(body, "sTitle=OK") {
		t.Errorf("sTitle not retained: %s", body)
	}
	if !strings.Contains(body, "sOrderBY=OK") {
		t.Errorf("sOrderBY not retained: %s", body)
	}
}

// TestClassMemberAssignmentWithDimCollision tests whether Dim at module level
// interferes with class member assignment in a method.
func TestClassMemberAssignmentWithDimCollision(t *testing.T) {
	aspCode := `<%
' Simulate the real scenario: dim variables at module level that share names
' with class members (e.g., process.asp has "dim listitems" and cls_page has "listitems" function)
Dim sTitle
sTitle = "GlobalTitle"

Class cls_page
    Public iId
    Public sTitle
    Public sOrderBY

    Public Function Pick(id)
        ON Error Resume Next
        iId = id
        sTitle = "PageTitle"
        sOrderBY = "sTitle"
    End Function
End Class

Dim selectedPage
Set selectedPage = New cls_page
selectedPage.Pick(492)

Response.Write "globalSTitle=" & sTitle & " "

If selectedPage.sTitle = "PageTitle" Then
    Response.Write "sTitle=OK "
Else
    Response.Write "sTitle=FAIL(" & TypeName(selectedPage.sTitle) & ":" & selectedPage.sTitle & ") "
End If

If selectedPage.sOrderBY = "sTitle" Then
    Response.Write "sOrderBY=OK"
Else
    Response.Write "sOrderBY=FAIL(" & TypeName(selectedPage.sOrderBY) & ":" & selectedPage.sOrderBY & ")"
End If
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-3")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)

	// The global sTitle should NOT be modified by the class method
	if !strings.Contains(body, "globalSTitle=GlobalTitle") {
		t.Errorf("Global sTitle was modified by class method! Output: %s", body)
	}
	if !strings.Contains(body, "sTitle=OK") {
		t.Errorf("Class sTitle not retained: %s", body)
	}
	if !strings.Contains(body, "sOrderBY=OK") {
		t.Errorf("Class sOrderBY not retained: %s", body)
	}
}

// TestClassMemberIsNull tests that isNull works correctly on class members
func TestClassMemberIsNull(t *testing.T) {
	aspCode := `<%
Class cls_page
    Public iId
    Public sOrderBY

    Public Function Pick(id)
        iId = id
        sOrderBY = "sTitle"
    End Function
End Class

Dim selectedPage
Set selectedPage = New cls_page
selectedPage.Pick(492)

If IsNull(selectedPage.iId) Then
    Response.Write "iId=NULL "
Else
    Response.Write "iId=NOTNULL(" & selectedPage.iId & ") "
End If

If IsNull(selectedPage.sOrderBY) Then
    Response.Write "sOrderBY=NULL"
Else
    Response.Write "sOrderBY=NOTNULL(" & selectedPage.sOrderBY & ")"
End If
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-4")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "iId=NOTNULL(492)") {
		t.Errorf("iId should not be null: %s", body)
	}
	if !strings.Contains(body, "sOrderBY=NOTNULL(sTitle)") {
		t.Errorf("sOrderBY should not be null: %s", body)
	}
}

// TestClassMemberWithCallMethodOnRecordsetLikeObject tests the real pattern:
// class method assigns class member from an object's CallMethod result
func TestClassMemberWithCallMethodOnRecordsetLikeObject(t *testing.T) {
	aspCode := `<%
Class cls_fakeRS
    Public Function Item(fieldName)
        Select Case LCase(fieldName)
            Case "iid"
                Item = 492
            Case "stitle"
                Item = "ListPage"
            Case "sorderby"
                Item = "sTitle"
        End Select
    End Function
End Class

Class cls_page
    Public iId
    Public sTitle
    Public sOrderBY

    Public Function Pick(id)
        ON Error Resume Next
        Dim rs
        Set rs = New cls_fakeRS
        iId = rs("iId")
        sTitle = rs("sTitle")
        sOrderBY = rs("sOrderBY")
        Set rs = Nothing
    End Function
End Class

Dim selectedPage
Set selectedPage = New cls_page
selectedPage.Pick(492)

Response.Write "iId=" & selectedPage.iId & " "
Response.Write "sTitle=" & selectedPage.sTitle & " "
Response.Write "sOrderBY=" & selectedPage.sOrderBY
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-5")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "iId=492") {
		t.Errorf("iId not set correctly: %s", body)
	}
	if !strings.Contains(body, "sTitle=ListPage") {
		t.Errorf("sTitle not set correctly: %s", body)
	}
	if !strings.Contains(body, "sOrderBY=sTitle") {
		t.Errorf("sOrderBY not set correctly: %s", body)
	}
}

// TestClassMemberWithIsLeeg tests the isLeeg pattern used to check sOrderBY
func TestClassMemberWithIsLeeg(t *testing.T) {
	aspCode := `<%
Function isLeeg(ByVal value)
    isLeeg = False
    If IsNull(value) Then
        isLeeg = True
    Else
        If IsEmpty(value) Or Trim(value) = "" Then isLeeg = True
    End If
End Function

Class cls_page
    Public iId
    Public sOrderBY

    Public Function Pick(id)
        iId = id
        sOrderBY = "sTitle"
    End Function
End Class

Dim selectedPage
Set selectedPage = New cls_page
selectedPage.Pick(492)

If Not isLeeg(selectedPage.sOrderBY) Then
    Response.Write "LIST_ITEMS_SHOWN"
Else
    Response.Write "LIST_ITEMS_HIDDEN"
End If
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-6")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "LIST_ITEMS_SHOWN") {
		t.Errorf("isLeeg should return false for 'sTitle', but list items are hidden: %s", body)
	}
}

// TestClassMemberCalledFromFunction tests if calling pick() from within another
// function that has a scope variable matching a class member name causes the value
// to be stored in the wrong scope.
func TestClassMemberCalledFromFunction(t *testing.T) {
	aspCode := `<%
Function isLeeg(ByVal value)
    isLeeg = False
    If IsNull(value) Then
        isLeeg = True
    Else
        If IsEmpty(value) Or Trim(value) = "" Then isLeeg = True
    End If
End Function

Class cls_page
    Public iId
    Public sTitle
    Public sOrderBY

    Public Function Pick(id)
        ON Error Resume Next
        iId = id
        sTitle = "PageTitle"
        sOrderBY = "sTitle"
    End Function
End Class

' This function simulates a wrapper that has a local variable with the same
' name as a class member
Function SetupPage()
    Dim sTitle
    Dim selectedPage
    Set selectedPage = New cls_page
    selectedPage.Pick(492)
    
    Response.Write "Inside: sOrderBY=" & selectedPage.sOrderBY & " "
    
    Set SetupPage = selectedPage
End Function

Dim page
Set page = SetupPage()

If Not isLeeg(page.sOrderBY) Then
    Response.Write "OUTSIDE: sOrderBY=" & page.sOrderBY
Else
    Response.Write "OUTSIDE: sOrderBY=EMPTY"
End If
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-7")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "sOrderBY=sTitle") {
		t.Errorf("sOrderBY should be 'sTitle' but got: %s", body)
	}
}

// TestClassMemberWith60Fields tests the real scenario with many field assignments
// in the pick method, similar to the real cls_page
func TestClassMemberWith60Fields(t *testing.T) {
	aspCode := `<%
Class cls_page
    Public iId
    Public iParentID
    Public iListPageID
    Public iCustomerID
    Public sTitle
    Public sValue
    Public sExternalURL
    Public bOnline
    Public bDeleted
    Public bHomepage
    Public sOrderBY
    Public sCode
    Public bIntranet
    Public iFormID
    Public iTemplateID
    Public bPushRSS
    Public sUserFriendlyURL

    Public Function Pick(id)
        ON Error Resume Next
        iId = 492
        iParentID = 0
        iListPageID = 492
        iCustomerID = 73
        sTitle = "List Page"
        sValue = "<p>Content</p>"
        sExternalURL = ""
        bOnline = True
        bDeleted = False
        bHomepage = False
        sOrderBY = "sTitle"
        sCode = ""
        bIntranet = False
        iFormID = 0
        iTemplateID = 1
        bPushRSS = True
        sUserFriendlyURL = ""
    End Function
End Class

Dim selectedPage
Set selectedPage = New cls_page
selectedPage.Pick(492)

Response.Write "iId=" & selectedPage.iId & "|"
Response.Write "sTitle=" & selectedPage.sTitle & "|"
Response.Write "sOrderBY=" & selectedPage.sOrderBY & "|"
Response.Write "bPushRSS=" & selectedPage.bPushRSS
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-8")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "sOrderBY=sTitle") {
		t.Errorf("sOrderBY not retained with many fields: %s", body)
	}
	if !strings.Contains(body, "iId=492") {
		t.Errorf("iId not retained: %s", body)
	}
	if !strings.Contains(body, "bPushRSS=True") {
		t.Errorf("bPushRSS not retained: %s", body)
	}
}

// TestDimHoistingInClassMethod tests the exact bug scenario:
// A class method uses a variable BEFORE its Dim statement.
// VBScript hoists Dim declarations, so this should work.
func TestDimHoistingInClassMethod(t *testing.T) {
	aspCode := `<%
Class cls_page
    Public iId
    Public sOrderBY

    Public Function fastlistitems()
        Set fastlistitems = Server.CreateObject("Scripting.Dictionary")
        If True Then
            If True Then
                sql = "SELECT * FROM tblPage WHERE iId=492"
                sql = sql & " ORDER BY sTitle"
                Dim rs, sql, page
                ' In real VBScript, Dim is hoisted so sql should
                ' still have its value here
                Response.Write "sql=" & sql
            End If
        End If
    End Function
End Class

Dim selectedPage
Set selectedPage = New cls_page
selectedPage.fastlistitems()
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-9")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "sql=SELECT * FROM tblPage WHERE iId=492 ORDER BY sTitle") {
		t.Errorf("Dim hoisting failed - sql was reset by late Dim statement: %s", body)
	}
}

// TestDimHoistingInFunction tests Dim hoisting in a regular function
func TestDimHoistingInFunction(t *testing.T) {
	aspCode := `<%
Function BuildQuery()
    sql = "SELECT * FROM table"
    sql = sql & " WHERE x=1"
    Dim sql
    BuildQuery = sql
End Function

Response.Write "result=" & BuildQuery()
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-10")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "result=SELECT * FROM table WHERE x=1") {
		t.Errorf("Dim hoisting in function failed: %s", body)
	}
}

// TestDimHoistingInSub tests Dim hoisting in a Sub
func TestDimHoistingInSub(t *testing.T) {
	aspCode := `<%
Sub TestSub()
    x = 42
    Dim x
    Response.Write "x=" & x
End Sub

TestSub()
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-11")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	t.Logf("Output: %s", body)
	if !strings.Contains(body, "x=42") {
		t.Errorf("Dim hoisting in Sub failed: %s", body)
	}
}

package axonvm

import (
	"bytes"
	"strings"
	"testing"
)

// TestIISCompatiblePrivateInitialize ensures that 'Private Sub Class_Initialize'
// does NOT block external instantiation, matching legacy IIS/VBScript behavior.
func TestIISCompatiblePrivateInitialize(t *testing.T) {
	script := `<%
	Class MyTestClass
		Private initCount

		' In IIS, Private here is effectively ignored for external 'New' calls.
		Private Sub Class_Initialize()
			initCount = 42
		End Sub

		Public Function GetInitCount()
			GetInitCount = initCount
		End Function
	End Class

	Dim obj
	Set obj = New MyTestClass
	Response.Write obj.GetInitCount()
	%>`

	compiler := NewASPCompiler(script)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("Compilation failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("Expected successful instantiation, got error: %v", err)
	}

	if output.String() != "42" {
		t.Errorf("Expected '42', got %q. Class_Initialize was not executed properly.", output.String())
	}
}

// personLifetimeClass declares the class used by the object lifetime tests. Class_Terminate
// writes "T" so the tests can observe exactly when an instance is released.
const personLifetimeClass = `Class Person
    Public Id
    Private Sub Class_Initialize()
        Id = 41
    End Sub
    Private Sub Class_Terminate()
        Response.Write "T"
    End Sub
End Class
`

// assertLifetimeOutput compares the output with the Class_Terminate markers removed and verifies
// how many instances were terminated. Use it for the scenarios where the exact position of a
// termination is implementation defined but the payload and the termination count are not.
func assertLifetimeOutput(t *testing.T, output, wantPayload string, wantTerminations int) {
	t.Helper()
	if payload := strings.ReplaceAll(output, "T", ""); payload != wantPayload {
		t.Fatalf("unexpected payload: got %q want %q (raw output %q)", payload, wantPayload, output)
	}
	if count := strings.Count(output, "T"); count != wantTerminations {
		t.Fatalf("unexpected Class_Terminate count: got %d want %d (raw output %q)", count, wantTerminations, output)
	}
}

// TestClassLifetimeDictionaryAddKeepsInstanceAlive verifies an instance stored in a
// Scripting.Dictionary with Add survives the function scope that created it, including
// Item() access and For Each iteration over Items().
func TestClassLifetimeDictionaryAddKeepsInstanceAlive(t *testing.T) {
	source := `<%` + personLifetimeClass + `Function BuildModel()
    Dim model, p
    Set model = Server.CreateObject("Scripting.Dictionary")
    Set p = New Person
    model.Add "person", p
    Set BuildModel = model
End Function

Dim m, item, items
Set m = BuildModel()
Response.Write m.Item("person").Id
items = m.Items()
For Each item In items
    Response.Write "|" & item.Id
Next
Response.Write "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	assertLifetimeOutput(t, output, "41|41|end", 1)
}

// TestClassLifetimeDictionarySetForms verifies both "Set dict.Item(k) = obj" and
// "Set dict(k) = obj" keep the instance alive after the creating scope exits.
func TestClassLifetimeDictionarySetForms(t *testing.T) {
	source := `<%` + personLifetimeClass + `Function BuildViaItem()
    Dim dict, p
    Set dict = Server.CreateObject("Scripting.Dictionary")
    Set p = New Person
    Set dict.Item("person") = p
    Set BuildViaItem = dict
End Function

Function BuildViaDefault()
    Dim dict, p
    Set dict = Server.CreateObject("Scripting.Dictionary")
    Set p = New Person
    Set dict("person") = p
    Set BuildViaDefault = dict
End Function

Dim m
Set m = BuildViaItem()
Response.Write m.Item("person").Id & "|"
Set m = BuildViaDefault()
Response.Write m.Item("person").Id & "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	assertLifetimeOutput(t, output, "41|41|end", 2)
}

// TestClassLifetimeHeldByAnotherObject verifies "Set holder.Item = obj" keeps the instance
// alive both for a public field and for a Property Set accessor.
func TestClassLifetimeHeldByAnotherObject(t *testing.T) {
	source := `<%` + personLifetimeClass + `Class FieldHolder
    Public Item
End Class

Class PropertyHolder
    Private m_Item
    Public Property Set Item(value)
        Set m_Item = value
    End Property
    Public Property Get Item()
        Set Item = m_Item
    End Property
End Class

Function BuildFieldHolder()
    Dim h, p
    Set h = New FieldHolder
    Set p = New Person
    Set h.Item = p
    Set BuildFieldHolder = h
End Function

Function BuildPropertyHolder()
    Dim h, p
    Set h = New PropertyHolder
    Set p = New Person
    Set h.Item = p
    Set BuildPropertyHolder = h
End Function

Dim holder
Set holder = BuildFieldHolder()
Response.Write holder.Item.Id & "|"
Set holder = BuildPropertyHolder()
Response.Write holder.Item.Id & "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	assertLifetimeOutput(t, output, "41|41|end", 2)
}

// TestClassLifetimeCollectionAndArray verifies instances stored in a Collection and in a
// Variant array returned from a function survive the creating scope.
func TestClassLifetimeCollectionAndArray(t *testing.T) {
	source := `<%` + personLifetimeClass + `Function BuildCollection()
    Dim col, p
    Set col = Server.CreateObject("Collection")
    Set p = New Person
    col.Add p
    Set BuildCollection = col
End Function

Function BuildArray()
    Dim p
    Set p = New Person
    BuildArray = Array(p)
End Function

Dim col, x, arr, first
Set col = BuildCollection()
For Each x In col
    Response.Write x.Id
Next
arr = BuildArray()
first = arr(0)
Response.Write "|" & first.Id & "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	assertLifetimeOutput(t, output, "41|41|end", 2)
}

// TestClassLifetimeTransitiveHolderChain verifies a Dictionary -> holder -> instance chain
// keeps the innermost instance alive after the creating function returns.
func TestClassLifetimeTransitiveHolderChain(t *testing.T) {
	source := `<%` + personLifetimeClass + `Class ChildHolder
    Public Item
End Class

Function BuildNested()
    Dim dict, child, p
    Set dict = Server.CreateObject("Scripting.Dictionary")
    Set child = New ChildHolder
    Set p = New Person
    Set child.Item = p
    dict.Add "child", child
    Set BuildNested = dict
End Function

Dim dict, child
Set dict = BuildNested()
Set child = dict.Item("child")
Response.Write child.Item.Id & "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	assertLifetimeOutput(t, output, "41|end", 1)
}

// TestClassTerminateDeterministicLocalScope verifies a purely local instance is still
// terminated as soon as the declaring procedure returns instead of waiting for the end of
// the page, and that the termination marker is written before the next statement runs.
func TestClassTerminateDeterministicLocalScope(t *testing.T) {
	source := `<%` + personLifetimeClass + `Sub ShowLocal()
    Dim p
    Set p = New Person
    Response.Write p.Id
End Sub

ShowLocal()
Response.Write "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	if output != "41T|end" {
		t.Fatalf("unexpected output: got %q want %q", output, "41T|end")
	}
}

// TestClassTerminateOnDictionaryRemoval verifies removing an entry from a Scripting.Dictionary
// terminates the instance it held, and that the termination happens before the following
// statement runs.
func TestClassTerminateOnDictionaryRemoval(t *testing.T) {
	source := `<%` + personLifetimeClass + `Dim d, p
Set d = Server.CreateObject("Scripting.Dictionary")
Set p = New Person
d.Add "k", p
Set p = Nothing
Response.Write "a"
d.Remove "k"
Response.Write "b"
Set p = New Person
d.Add "j", p
Set p = Nothing
d.RemoveAll
Response.Write "c"
%>`

	output := runVBSAndGetOutput(t, source)
	if output != "aTbTc" {
		t.Fatalf("unexpected output: got %q want %q", output, "aTbTc")
	}
}

// TestClassLifetimeObjectParameters verifies instances passed to procedures as ByVal and ByRef
// parameters are neither terminated early nor lost by the reference-count bookkeeping.
func TestClassLifetimeObjectParameters(t *testing.T) {
	source := `<%` + personLifetimeClass + `Function ReadId(ByVal o)
    Dim snapshot
    Set snapshot = o
    ReadId = snapshot.Id
End Function

Sub AssignByRef(ByRef target, ByVal o)
    Set target = o
End Sub

Dim p, q
Set p = New Person
Response.Write ReadId(p) & "|"
AssignByRef q, p
Response.Write q.Id & "|"
Set q = Nothing
Response.Write ReadId(p) & "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	if output != "41|41|41|endT" {
		t.Fatalf("unexpected output: got %q want %q", output, "41|41|41|endT")
	}
}

// TestClassLifetimePageLevelVariable verifies an instance created inside a function and returned
// to a page-level variable keeps its state after the creating scope exits.
func TestClassLifetimePageLevelVariable(t *testing.T) {
	source := `<%` + personLifetimeClass + `Function MakePerson()
    Dim p
    Set p = New Person
    Set MakePerson = p
End Function

Dim kept
Set kept = MakePerson()
Response.Write kept.Id & "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	if output != "41|endT" {
		t.Fatalf("unexpected output: got %q want %q", output, "41|endT")
	}
}

// TestClassLifetimePageLevelDictionary verifies an instance created at page level and stored in a
// Scripting.Dictionary survives clearing the variable that created it.
func TestClassLifetimePageLevelDictionary(t *testing.T) {
	source := `<%` + personLifetimeClass + `Dim d, p
Set d = Server.CreateObject("Scripting.Dictionary")
Set p = New Person
d.Add "person", p
Set p = Nothing
Response.Write d.Item("person").Id & "|end"
%>`

	output := runVBSAndGetOutput(t, source)
	if output != "41|endT" {
		t.Fatalf("unexpected output: got %q want %q", output, "41|endT")
	}
}

// TestClassTerminateSelfAssignment verifies "Set x = x" does not drop the reference count to
// zero and terminate the instance while it is still assigned.
func TestClassTerminateSelfAssignment(t *testing.T) {
	source := `<%` + personLifetimeClass + `Dim p
Set p = New Person
Set p = p
Response.Write p.Id & "|"
Set p = Nothing
Response.Write "end"
%>`

	output := runVBSAndGetOutput(t, source)
	if output != "41|Tend" {
		t.Fatalf("unexpected output: got %q want %q", output, "41|Tend")
	}
}

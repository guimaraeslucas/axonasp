package axonvm

import (
	"bytes"
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

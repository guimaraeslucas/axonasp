package axonvm

import (
	"testing"
)

// TestIISCompatiblePrivateInitialize ensures that 'Private Sub Class_Initialize'
// does NOT block external instantiation, matching legacy IIS/VBScript behavior.
func TestIISCompatiblePrivateInitialize(t *testing.T) {
	script := `
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
	Dim result
	result = obj.GetInitCount()
	`

	vm := NewVM()
	err := vm.Run(script)

	if err != nil {
		t.Fatalf("Expected successful instantiation, got error: %v", err)
	}

	result, err := vm.GetValue("result")
	if err != nil {
		t.Fatalf("Failed to retrieve result: %v", err)
	}

	if result.Value() != 42 {
		t.Errorf("Expected initCount to be 42, got %v. Class_Initialize was not executed properly.", result.Value())
	}
}

package axonvm

import (
	"bytes"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/vbscript"
)

func TestASPPrivateConstructorEncapsulation(t *testing.T) {
	source := `
<%
Class Singleton
	Private Sub Class_Initialize()
		Response.Write "Init;"
	End Sub

	Public Function CreateInternal()
		Set CreateInternal = New Singleton
	End Function
End Class

Dim obj
On Error Resume Next
Set obj = New Singleton
If Err.Number <> 0 Then
	Response.Write "Err:" & Err.Number & ";"
End If
If IsEmpty(obj) Then
	Response.Write "IsEmpty;"
End If
%>
`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	var output bytes.Buffer
	vm := NewVMFromCompiler(compiler)
	vm.SetOutput(&output)

	// Run the script to test external instantiation failure
	err := vm.Run()
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	outStr := output.String()
	if !strings.Contains(outStr, "Init;Err:91;IsEmpty;") {
		t.Fatalf("Expected output to contain 'Init;Err:91;IsEmpty;', got: %q", outStr)
	}

	// Test internal instantiation by bootstrapping an instance from Go
	instance := vm.newRuntimeClassInstance("Singleton")
	if instance.Type != VTObject {
		t.Fatalf("Failed to create bootstrap instance")
	}

	// Call the public factory method CreateInternal
	method, ok := vm.resolveRuntimeClassMethod(instance, "CreateInternal", true)
	if !ok {
		t.Fatalf("Could not find CreateInternal method")
	}

	// Setup call frame for CreateInternal
	vm.callStack = nil
	vm.sp = -1
	vm.fp = 0

	if vm.beginUserSubCall(method, nil, false, instance.Num) {
		// execute the method body
		err := vm.Run()
		if err != nil {
			t.Fatalf("CreateInternal failed: %v", err)
		}
	} else {
		t.Fatalf("Failed to begin user sub call")
	}

	// Verify the returned value is a new Singleton object
	if vm.sp < 0 {
		t.Fatalf("Expected return value on stack")
	}
	retVal := vm.pop()
	if retVal.Type != VTObject {
		t.Fatalf("Expected VTObject returned by CreateInternal, got type %d", retVal.Type)
	}

	createdInst, exists := vm.runtimeClassItems[retVal.Num]
	if !exists || !strings.EqualFold(createdInst.ClassName, "Singleton") {
		t.Fatalf("Expected newly created Singleton instance, got %v", createdInst)
	}
}

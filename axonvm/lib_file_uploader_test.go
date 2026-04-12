package axonvm

import (
	"testing"
)

func TestFileUploader(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	lib := vm.newG3FileUploaderObject()
	if lib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", lib.Type)
	}

	obj := vm.fileUploaderItems[lib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	obj.DispatchPropertySet("MaxFileSize", []Value{NewInteger(5000)})
	size := obj.DispatchPropertyGet("MaxFileSize")
	if size.Num != 5000 {
		t.Errorf("expected 5000, got %d", size.Num)
	}

	obj.DispatchMethod("BlockExtension", []Value{NewString(".exe")})
	blocked := obj.DispatchPropertyGet("BlockedExtensions")
	if blocked.Type != VTArray || blocked.Arr == nil {
		t.Fatal("expected array")
	}

	if len(blocked.Arr.Values) != 1 || blocked.Arr.Values[0].String() != ".exe" {
		t.Errorf("expected .exe in blocked list")
	}
}

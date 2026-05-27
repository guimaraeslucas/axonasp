package axonvm

import (
	"bytes"
	"io"
	"mime/multipart"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestG3FileUploader_Constants(t *testing.T) {
	if VTNativeObject != 9 {
		t.Errorf("VTNativeObject constant is %d, expected 9", VTNativeObject)
	}
}

func TestG3FileUploader_FormFields(t *testing.T) {
	vm := NewVM(nil, nil, 4096)

	body := &bytes.Buffer{}
	writer := multipart.NewWriter(body)
	_ = writer.WriteField("field1", "value1")
	_ = writer.WriteField("field2", "value2")
	_ = writer.Close()

	req := httptest.NewRequest("POST", "/", body)
	req.Header.Set("Content-Type", writer.FormDataContentType())

	host := NewMockHost()
	host.Request().SetHTTPRequest(req)
	vm.host = host

	uploaderVal := vm.newG3FileUploaderObject()
	if uploaderVal.Type != VTNativeObject {
		t.Fatalf("newG3FileUploaderObject returned type %d, expected %d", uploaderVal.Type, VTNativeObject)
	}

	uploader := vm.fileUploaderItems[uploaderVal.Num]
	if uploader == nil {
		t.Fatalf("Uploader not found in vm.fileUploaderItems for ID %d", uploaderVal.Num)
	}

	fieldsVal := uploader.DispatchPropertyGet("FormFields")
	if fieldsVal.Type != VTNativeObject {
		t.Fatalf("Expected VTNativeObject (9) for FormFields, got type %d (Value: %+v)", fieldsVal.Type, fieldsVal)
	}

	// Verify field1
	val1, ok1 := vm.dispatchDictionaryMethod(fieldsVal.Num, "Item", []Value{NewString("field1")})
	if !ok1 {
		t.Fatal("dispatchDictionaryMethod failed for field1")
	}
	if val1.String() != "value1" {
		t.Errorf("Expected value1 for field1, got %s", val1.String())
	}

	// Verify field2
	val2, ok2 := vm.dispatchDictionaryMethod(fieldsVal.Num, "Item", []Value{NewString("field2")})
	if !ok2 {
		t.Fatal("dispatchDictionaryMethod failed for field2")
	}
	if val2.String() != "value2" {
		t.Errorf("Expected value2 for field2, got %s", val2.String())
	}
}

func TestG3FileUploader_AbsolutePaths(t *testing.T) {
	vm := NewVM(nil, nil, 4096)

	tempDir, err := os.MkdirTemp("", "axon_upload_test_*")
	if err != nil {
		t.Fatalf("Failed to create temp dir: %v", err)
	}
	defer os.RemoveAll(tempDir)

	body := &bytes.Buffer{}
	writer := multipart.NewWriter(body)
	part, err := writer.CreateFormFile("file1", "test.txt")
	if err != nil {
		t.Fatalf("Failed to create form file: %v", err)
	}
	_, _ = io.WriteString(part, "test content")
	_ = writer.Close()

	req := httptest.NewRequest("POST", "/", body)
	req.Header.Set("Content-Type", writer.FormDataContentType())

	host := NewMockHost()
	host.Request().SetHTTPRequest(req)
	host.Server().SetRootDir(tempDir)
	vm.host = host

	uploaderVal := vm.newG3FileUploaderObject()
	uploader := vm.fileUploaderItems[uploaderVal.Num]

	// Test default behavior (restricted to sandbox)
	result := uploader.DispatchMethod("Process", []Value{NewString("file1"), NewString("/uploads")})
	isSuccess, _ := vm.dispatchDictionaryPropertyGet(result.Num, "IsSuccess")
	if isSuccess.Num == 0 {
		errMsg, _ := vm.dispatchDictionaryPropertyGet(result.Num, "ErrorMessage")
		t.Errorf("Process failed: %s", errMsg.String())
	}

	finalPath, _ := vm.dispatchDictionaryPropertyGet(result.Num, "FinalPath")
	expectedPrefix := filepath.Join(tempDir, "uploads")
	if !strings.HasPrefix(filepath.Clean(finalPath.String()), filepath.Clean(expectedPrefix)) {
		t.Errorf("Expected path to be inside %s, got %s", expectedPrefix, finalPath.String())
	}

	// Test Absolute Path Toggle
	uploader.DispatchPropertySet("AllowAbsolutePaths", []Value{NewBool(true)})

	absTarget := filepath.Join(tempDir, "absolute_dir")

	resultAbs := uploader.DispatchMethod("Process", []Value{NewString("file1"), NewString(absTarget)})
	isSuccessAbs, _ := vm.dispatchDictionaryPropertyGet(resultAbs.Num, "IsSuccess")
	if isSuccessAbs.Num == 0 {
		errMsgAbs, _ := vm.dispatchDictionaryPropertyGet(resultAbs.Num, "ErrorMessage")
		t.Errorf("Process with absolute path failed: %s", errMsgAbs.String())
	} else {
		finalPathAbs, _ := vm.dispatchDictionaryPropertyGet(resultAbs.Num, "FinalPath")
		if !strings.HasPrefix(filepath.Clean(finalPathAbs.String()), filepath.Clean(absTarget)) {
			t.Errorf("Expected path to be inside %s, got %s", absTarget, finalPathAbs.String())
		}
	}
}

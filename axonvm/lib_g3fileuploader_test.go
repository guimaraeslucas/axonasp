//go:build !wasm

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

func TestG3FileUploader_Compatibility(t *testing.T) {
	vm := NewVM(nil, nil, 4096)

	// Create a mock multipart body
	body := &bytes.Buffer{}
	writer := multipart.NewWriter(body)
	_ = writer.WriteField("text_field", "hello_world")
	part, _ := writer.CreateFormFile("file_field", "hello.txt")
	_, _ = io.WriteString(part, "file content")
	_ = writer.Close()

	req := httptest.NewRequest("POST", "/", body)
	req.Header.Set("Content-Type", writer.FormDataContentType())

	host := NewMockHost()
	host.Request().SetHTTPRequest(req)
	tempDir, err := os.MkdirTemp("", "axon_compat_test_*")
	if err != nil {
		t.Fatalf("failed to create temp dir: %v", err)
	}
	defer os.RemoveAll(tempDir)
	host.Server().SetRootDir(tempDir)
	vm.host = host

	// Test 1: Persits.Upload Alias & Method/Properties
	persitsVal := vm.newG3FileUploaderObjectWithProgID("Persits.Upload")
	persitsObj := vm.fileUploaderItems[persitsVal.Num]
	if persitsObj == nil {
		t.Fatal("expected Persits.Upload instance")
	}
	if persitsObj.mode != ModePersitsUpload {
		t.Errorf("expected ModePersitsUpload, got %v", persitsObj.mode)
	}

	// Read TotalBytes (should parse request and return length > 0)
	totalBytesVal := persitsObj.DispatchPropertyGet("TotalBytes")
	if totalBytesVal.Type != VTInteger || totalBytesVal.Num <= 0 {
		t.Errorf("expected TotalBytes > 0, got %+v", totalBytesVal)
	}

	// Test Save(Path) method collision (Persits returns count of files)
	targetDir := filepath.Join(tempDir, "persits_save")
	saveResult := persitsObj.DispatchMethod("Save", []Value{NewString(targetDir)})
	if saveResult.Type != VTInteger || saveResult.Num != 1 {
		t.Errorf("expected Save to return 1 saved file, got %+v", saveResult)
	}

	// Get Files collection
	filesVal := persitsObj.DispatchPropertyGet("Files")
	if filesVal.Type != VTNativeObject {
		t.Fatalf("expected Files collection object, got %+v", filesVal)
	}
	filesCount := vm.fileUploaderItems[filesVal.Num].DispatchPropertyGet("Count")
	if filesCount.Num != 1 {
		t.Errorf("expected Files Count to be 1, got %d", filesCount.Num)
	}

	// Get file by index (1-based)
	fileItemVal := vm.fileUploaderItems[filesVal.Num].DispatchMethod("Item", []Value{NewInteger(1)})
	if fileItemVal.Type != VTNativeObject {
		t.Fatalf("expected file item, got %+v", fileItemVal)
	}
	fileItemObj := vm.fileUploaderItems[fileItemVal.Num]
	fileNameVal := fileItemObj.DispatchPropertyGet("FileName")
	if fileNameVal.String() != "hello.txt" {
		t.Errorf("expected filename hello.txt, got %q", fileNameVal.String())
	}
	fileSizeVal := fileItemObj.DispatchPropertyGet("Size")
	if fileSizeVal.Num != 12 {
		t.Errorf("expected size 12, got %d", fileSizeVal.Num)
	}

	// Test 2: SoftArtisans.FileUp Alias & Method/Properties
	// Re-create request for the next uploader
	body2 := &bytes.Buffer{}
	writer2 := multipart.NewWriter(body2)
	_ = writer2.WriteField("text_field", "hello_world")
	part2, _ := writer2.CreateFormFile("file_field", "hello.txt")
	_, _ = io.WriteString(part2, "file content")
	_ = writer2.Close()
	req2 := httptest.NewRequest("POST", "/", body2)
	req2.Header.Set("Content-Type", writer2.FormDataContentType())
	host2 := NewMockHost()
	host2.Request().SetHTTPRequest(req2)
	host2.Server().SetRootDir(tempDir)
	vm.host = host2

	saVal := vm.newG3FileUploaderObjectWithProgID("SoftArtisans.FileUp")
	saObj := vm.fileUploaderItems[saVal.Num]
	if saObj == nil {
		t.Fatal("expected SoftArtisans.FileUp instance")
	}
	if saObj.mode != ModeSAFileUp {
		t.Errorf("expected ModeSAFileUp, got %v", saObj.mode)
	}

	// Set path property
	saSaveDir := filepath.Join(tempDir, "sa_save")
	saObj.DispatchPropertySet("Path", []Value{NewString(saSaveDir)})

	// Test parameterless Save() for SA-FileUp
	saveResult2 := saObj.DispatchMethod("Save", nil)
	if saveResult2.Type != VTEmpty {
		t.Errorf("expected Save to return empty, got %+v", saveResult2)
	}

	// SA-FileUp Form mixes fields and files
	formVal := saObj.DispatchPropertyGet("Form")
	if formVal.Type != VTNativeObject {
		t.Fatalf("expected Form collection, got %+v", formVal)
	}

	formFieldVal := vm.fileUploaderItems[formVal.Num].DispatchMethod("", []Value{NewString("text_field")})
	if formFieldVal.String() != "hello_world" {
		t.Errorf("expected 'hello_world', got %q", formFieldVal.String())
	}

	formFileVal := vm.fileUploaderItems[formVal.Num].DispatchMethod("", []Value{NewString("file_field")})
	if formFileVal.Type != VTNativeObject {
		t.Fatalf("expected form file field to return file object, got %+v", formFileVal)
	}
}

// TestG3FileUploader_MoveUploadedFile verifies the cross-volume-safe move
// helper: the fast os.Rename path and the copy+delete fallback used when the
// temp directory and the final destination live on different drives.
func TestG3FileUploader_MoveUploadedFile(t *testing.T) {
	payload := []byte("upload payload content")

	// Same-volume fast path: os.Rename is used and must succeed.
	dirA, err := os.MkdirTemp("", "axon_move_a_*")
	if err != nil {
		t.Fatalf("failed to create source temp dir: %v", err)
	}
	defer os.RemoveAll(dirA)
	dirB, err := os.MkdirTemp("", "axon_move_b_*")
	if err != nil {
		t.Fatalf("failed to create dest temp dir: %v", err)
	}
	defer os.RemoveAll(dirB)

	srcPath := filepath.Join(dirA, "upload_*.tmp")
	srcFile, err := os.CreateTemp(dirA, "upload_*.tmp")
	if err != nil {
		t.Fatalf("failed to create temp source: %v", err)
	}
	srcPath = srcFile.Name()
	if _, err := srcFile.Write(payload); err != nil {
		t.Fatalf("failed to write temp source: %v", err)
	}
	if err := srcFile.Close(); err != nil {
		t.Fatalf("failed to close temp source: %v", err)
	}

	finalPath := filepath.Join(dirB, "renamed.bin")
	if err := moveUploadedFile(srcPath, finalPath); err != nil {
		t.Fatalf("moveUploadedFile (rename path) failed: %v", err)
	}
	if _, err := os.Stat(srcPath); !os.IsNotExist(err) {
		t.Errorf("expected source temp file to be removed, stat err = %v", err)
	}
	got, err := os.ReadFile(finalPath)
	if err != nil {
		t.Fatalf("failed to read moved file: %v", err)
	}
	if string(got) != string(payload) {
		t.Errorf("moved content mismatch: got %q, want %q", got, payload)
	}

	// Copy+delete fallback: the path used when rename crosses volumes.
	src2Path := filepath.Join(dirA, "upload_copy_*.tmp")
	src2, err := os.CreateTemp(dirA, "upload_copy_*.tmp")
	if err != nil {
		t.Fatalf("failed to create temp source: %v", err)
	}
	src2Path = src2.Name()
	if _, err := src2.Write(payload); err != nil {
		t.Fatalf("failed to write temp source: %v", err)
	}
	if err := src2.Close(); err != nil {
		t.Fatalf("failed to close temp source: %v", err)
	}

	fallbackPath := filepath.Join(dirB, "copied.bin")
	if err := copyUploadedFile(src2Path, fallbackPath); err != nil {
		t.Fatalf("copyUploadedFile failed: %v", err)
	}
	if _, err := os.Stat(src2Path); !os.IsNotExist(err) {
		t.Errorf("expected source temp file to be removed after copy, stat err = %v", err)
	}
	got2, err := os.ReadFile(fallbackPath)
	if err != nil {
		t.Fatalf("failed to read copied file: %v", err)
	}
	if string(got2) != string(payload) {
		t.Errorf("copied content mismatch: got %q, want %q", got2, payload)
	}
}

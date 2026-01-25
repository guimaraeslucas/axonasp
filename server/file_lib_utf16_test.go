package server

import (
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestADODBStreamLoadFromFileUTF16LE(t *testing.T) {
	recorder := httptest.NewRecorder()
	req := httptest.NewRequest("GET", "/", nil)
	ctx := NewExecutionContext(recorder, req, "sess", time.Second)

	tempDir := t.TempDir()
	ctx.RootDir = tempDir
	ctx.Server.SetRootDir(tempDir)

	data := []byte{0xFF, 0xFE, 'H', 0x00, 'i', 0x00}
	filePath := filepath.Join(tempDir, "utf16.txt")
	if err := os.WriteFile(filePath, data, 0644); err != nil {
		t.Fatalf("failed to write UTF-16 test file: %v", err)
	}

	stream := NewADODBStream(ctx)
	stream.CallMethod("Open")
	stream.SetProperty("Type", 2)
	stream.CallMethod("LoadFromFile", filePath)

	raw := stream.CallMethod("ReadText")
	text, ok := raw.(string)
	if !ok {
		t.Fatalf("expected string from ReadText, got %T", raw)
	}

	if text != "Hi" {
		t.Fatalf("expected decoded text 'Hi', got %q", text)
	}
}

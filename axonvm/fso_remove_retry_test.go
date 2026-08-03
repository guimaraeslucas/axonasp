package axonvm

import (
	"os"
	"path/filepath"
	"testing"
)

func TestFSORemoveWithRetry(t *testing.T) {
	vm := &VM{}

	t.Run("StandaloneFile", func(t *testing.T) {
		tempDir := t.TempDir()
		filePath := filepath.Join(tempDir, "standalone.txt")
		if err := os.WriteFile(filePath, []byte("standalone test content"), 0644); err != nil {
			t.Fatalf("failed to create standalone file: %v", err)
		}

		err := vm.fsoRemoveWithRetry(filePath, false)
		if err != nil {
			t.Fatalf("fsoRemoveWithRetry failed for standalone file: %v", err)
		}
		if _, statErr := os.Stat(filePath); !os.IsNotExist(statErr) {
			t.Errorf("expected standalone file to be deleted, but it still exists")
		}
	})

	t.Run("PopulatedDirectory", func(t *testing.T) {
		tempDir := t.TempDir()

		parentDir := filepath.Join(tempDir, "parentDir")
		if err := os.MkdirAll(parentDir, 0755); err != nil {
			t.Fatalf("failed to create parent directory: %v", err)
		}

		childFile := filepath.Join(parentDir, "childFile.txt")
		if err := os.WriteFile(childFile, []byte("nested test content"), 0644); err != nil {
			t.Fatalf("failed to create child file: %v", err)
		}

		err := vm.fsoRemoveWithRetry(parentDir, true)
		if err != nil {
			t.Fatalf("fsoRemoveWithRetry failed for populated directory: %v", err)
		}

		if _, statErr := os.Stat(parentDir); !os.IsNotExist(statErr) {
			t.Errorf("expected parent directory to be deleted post-execution, but it still exists")
		}
	})
}

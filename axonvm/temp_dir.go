package axonvm

import (
	"path/filepath"
	"strings"

	"g3pix.com.br/axonasp/v2/axonconfig"
)

// resolveConfiguredTempDir returns global.temp_dir with a safe project-local fallback.
func resolveConfiguredTempDir() string {
	tempDir := strings.TrimSpace(axonconfig.NewViper().GetString("global.temp_dir"))
	if tempDir == "" {
		tempDir = filepath.Join(".", "temp")
	}
	return filepath.Clean(tempDir)
}

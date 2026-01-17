package server

import (
	"fmt"
	"net/http"
	"time"
)

// ASPProcessorConfig contains configuration for ASP processing
type ASPProcessorConfig struct {
	RootDir       string
	ScriptTimeout int // in seconds
}

// ASPProcessor handles ASP file execution
// Delegates to ASPExecutor for actual code execution
type ASPProcessor struct {
	config *ASPProcessorConfig
}

// NewASPProcessor creates a new ASP processor
func NewASPProcessor(config *ASPProcessorConfig) *ASPProcessor {
	if config == nil {
		config = &ASPProcessorConfig{
			RootDir:       "./www",
			ScriptTimeout: 30,
		}
	}
	return &ASPProcessor{
		config: config,
	}
}

// ExecuteASPFile processes and executes an ASP file
// Takes the file content as string and returns the rendered output
// Delegates to ASPExecutor in executor.go
func (ap *ASPProcessor) ExecuteASPFile(fileContent string, w http.ResponseWriter, r *http.Request) error {
	// Create the executor with configuration
	executor := NewASPExecutor(ap.config)

	// Generate session ID from request cookie or create new one
	sessionID := generateSessionID(r)

	// Execute the ASP file using the full executor
	return executor.Execute(fileContent, w, r, sessionID)
}

// generateSessionID creates or retrieves a session ID from request cookies
func generateSessionID(r *http.Request) string {
	// Look for existing ASPSESSIONID cookie
	if cookie, err := r.Cookie("ASPSESSIONID"); err == nil {
		return cookie.Value
	}

	// Generate new session ID
	return generateUniqueID()
}

// generateUniqueID generates a unique identifier for sessions
func generateUniqueID() string {
	// Simple implementation - in production use crypto/rand with proper UUID
	return fmt.Sprintf("ASP%d", time.Now().UnixNano())
}

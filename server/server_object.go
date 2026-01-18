package server

import (
	"fmt"
	"html"
	"net/http"
	"net/url"
	"path/filepath"
	"strings"
	"time"
)

// ServerObject represents the ASP Classic Server Object
// Provides server utilities like path mapping, encoding, and object creation
type ServerObject struct {
	// Context
	context  *ExecutionContext
	executor interface {
		CreateObject(string) (interface{}, error)
	}

	// Properties
	scriptTimeout int // In seconds

	// Internal properties
	rootDir     string
	httpRequest *http.Request
	lastError   *ASPError
}

// ASPError represents an ASP error object returned by Server.GetLastError()
type ASPError struct {
	ASPCode        int
	ASPDescription string
	Category       string
	Column         int
	Description    string
	File           string
	Line           int
	Number         int
	Source         string
}

// NewServerObjectWithContext creates a new Server object with execution context
func NewServerObjectWithContext(ctx *ExecutionContext) *ServerObject {
	return &ServerObject{
		context:       ctx,
		scriptTimeout: 90, // Default ASP timeout is 90 seconds
		rootDir:       "./www",
	}
}

// SetContext sets the execution context for the Server object
func (s *ServerObject) SetContext(ctx *ExecutionContext) {
	s.context = ctx
}

// SetExecutor sets the executor for CreateObject calls
func (s *ServerObject) SetExecutor(executor interface {
	CreateObject(string) (interface{}, error)
}) {
	s.executor = executor
}

// SetRootDir sets the root directory for MapPath operations
func (s *ServerObject) SetRootDir(rootDir string) {
	s.rootDir = rootDir
}

// SetHttpRequest sets the HTTP request for context-aware operations
func (s *ServerObject) SetHttpRequest(r *http.Request) {
	s.httpRequest = r
}

// SetLastError stores an error for GetLastError()
func (s *ServerObject) SetLastError(err *ASPError) {
	s.lastError = err
}

// ----- Properties -----

// GetScriptTimeout returns the current script timeout in seconds
func (s *ServerObject) GetScriptTimeout() int {
	return s.scriptTimeout
}

// SetScriptTimeout sets the script timeout in seconds (max 2,147,483,647)
// Timeout must be between 1 and 2,147,483,647 seconds
func (s *ServerObject) SetScriptTimeout(timeout int) error {
	if timeout < 1 {
		return fmt.Errorf("ScriptTimeout must be at least 1 second")
	}
	if timeout > 2147483647 {
		return fmt.Errorf("ScriptTimeout exceeds maximum value")
	}

	s.scriptTimeout = timeout

	// Update context timeout if available
	if s.context != nil {
		s.context.timeout = time.Duration(timeout) * time.Second
	}

	return nil
}

// ----- Methods -----

// CreateObject creates an ASP COM-compatible object
// Maintains compatibility with custom G3 libraries
func (s *ServerObject) CreateObject(progID string) (interface{}, error) {
	if progID == "" {
		return nil, fmt.Errorf("CreateObject requires a ProgID")
	}

	// Use executor's CreateObject if available (for custom libraries)
	if s.executor != nil {
		return s.executor.CreateObject(progID)
	}

	return nil, fmt.Errorf("CreateObject not available: no executor configured")
}

// Execute executes an ASP file and returns control to the calling page
// The calling page continues execution after the included file finishes
func (s *ServerObject) Execute(path string) error {
	if path == "" {
		return fmt.Errorf("Execute requires a path")
	}

	// This is similar to Server-Side Include but programmatic
	// Implementation requires:
	// 1. Resolve path using MapPath
	// 2. Load and parse ASP file
	// 3. Execute in current context
	// 4. Return control to caller

	if s.context == nil {
		return fmt.Errorf("Execute not available: no execution context")
	}

	// TODO: Full implementation requires integration with ASP parser
	// For now, return error indicating not yet implemented
	return fmt.Errorf("Server.Execute is not yet implemented")
}

// GetLastError returns an ASPError object containing error details
// Returns nil if no error has occurred
func (s *ServerObject) GetLastError() *ASPError {
	return s.lastError
}

// HTMLEncode encodes a string for safe HTML output
// Converts special characters to HTML entities
func (s *ServerObject) HTMLEncode(str string) string {
	return html.EscapeString(str)
}

// MapPath converts a virtual or relative path to a physical path on the server
// Supports both absolute (virtual="/path") and relative (file="path") paths
func (s *ServerObject) MapPath(path string) string {
	// Handle empty path
	if path == "" {
		return s.rootDir
	}

	// Handle root path
	if path == "/" || path == "\\" {
		return s.rootDir
	}

	// Normalize path separators
	path = strings.ReplaceAll(path, "\\", "/")

	// Check if it's an absolute virtual path (starts with /)
	if strings.HasPrefix(path, "/") {
		// Remove leading slash and join with root
		path = strings.TrimPrefix(path, "/")
		fullPath := filepath.Join(s.rootDir, path)
		absPath, err := filepath.Abs(fullPath)
		if err != nil {
			return fullPath
		}
		return absPath
	}

	// Relative path - needs current script directory
	// In a full implementation, this would use the current script's directory
	// For now, treat as relative to root
	fullPath := filepath.Join(s.rootDir, path)
	absPath, err := filepath.Abs(fullPath)
	if err != nil {
		return fullPath
	}

	return absPath
}

// Transfer transfers control to another ASP page
// Unlike Execute, Transfer does not return to the calling page
func (s *ServerObject) Transfer(path string) error {
	if path == "" {
		return fmt.Errorf("Transfer requires a path")
	}

	if s.context == nil {
		return fmt.Errorf("Transfer not available: no execution context")
	}

	// TODO: Full implementation requires:
	// 1. Load and parse target ASP file
	// 2. Execute in current context
	// 3. Terminate current script execution
	// 4. Prevent return to caller

	// For now, return error indicating not yet implemented
	return fmt.Errorf("Server.Transfer is not yet implemented")
}

// URLEncode encodes a string for safe use in URLs
// Follows RFC 3986 percent-encoding standard
func (s *ServerObject) URLEncode(str string) string {
	// Use QueryEscape which follows RFC 3986
	// Converts spaces to + and special characters to %XX
	return url.QueryEscape(str)
}

// URLPathEncode encodes only the path portion of a URL
// Does not encode forward slashes
func (s *ServerObject) URLPathEncode(str string) string {
	// Encode each path segment separately
	parts := strings.Split(str, "/")
	for i, part := range parts {
		parts[i] = url.PathEscape(part)
	}
	return strings.Join(parts, "/")
}

// ----- Helper Methods for VBScript Integration -----

// GetProperty gets a property by name (for VBScript interop)
func (s *ServerObject) GetProperty(name string) interface{} {
	nameLower := strings.ToLower(name)

	switch nameLower {
	case "scripttimeout":
		return s.scriptTimeout
	default:
		return nil
	}
}

// SetProperty sets a property by name (for VBScript interop)
func (s *ServerObject) SetProperty(name string, value interface{}) error {
	nameLower := strings.ToLower(name)

	switch nameLower {
	case "scripttimeout":
		// Convert value to int
		var timeout int
		switch v := value.(type) {
		case int:
			timeout = v
		case int32:
			timeout = int(v)
		case int64:
			timeout = int(v)
		case float64:
			timeout = int(v)
		case string:
			// Try parsing string
			var err error
			fmt.Sscanf(v, "%d", &timeout)
			if err != nil {
				return fmt.Errorf("invalid ScriptTimeout value: %v", value)
			}
		default:
			return fmt.Errorf("invalid ScriptTimeout type: %T", value)
		}
		return s.SetScriptTimeout(timeout)
	default:
		// Allow setting internal properties (used by executor)
		// Store them in a map if needed
		return nil
	}
}

// CallMethod calls a method by name (for VBScript interop)
func (s *ServerObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	nameLower := strings.ToLower(name)

	switch nameLower {
	case "createobject":
		if len(args) == 0 {
			return nil, fmt.Errorf("CreateObject requires a ProgID argument")
		}
		progID := fmt.Sprintf("%v", args[0])
		return s.CreateObject(progID)

	case "execute":
		if len(args) == 0 {
			return nil, fmt.Errorf("Execute requires a path argument")
		}
		path := fmt.Sprintf("%v", args[0])
		err := s.Execute(path)
		return nil, err

	case "getlasterror":
		return s.GetLastError(), nil

	case "htmlencode":
		if len(args) == 0 {
			return "", nil
		}
		str := fmt.Sprintf("%v", args[0])
		return s.HTMLEncode(str), nil

	case "mappath":
		if len(args) == 0 {
			return s.rootDir, nil
		}
		path := fmt.Sprintf("%v", args[0])
		return s.MapPath(path), nil

	case "transfer":
		if len(args) == 0 {
			return nil, fmt.Errorf("Transfer requires a path argument")
		}
		path := fmt.Sprintf("%v", args[0])
		err := s.Transfer(path)
		return nil, err

	case "urlencode":
		if len(args) == 0 {
			return "", nil
		}
		str := fmt.Sprintf("%v", args[0])
		return s.URLEncode(str), nil

	default:
		return nil, fmt.Errorf("unknown Server method: %s", name)
	}
}

// ----- ASPError Methods -----

// NewASPError creates a new ASP error object
func NewASPError(number int, description string, source string, file string, line int, column int) *ASPError {
	return &ASPError{
		Number:      number,
		Description: description,
		Source:      source,
		File:        file,
		Line:        line,
		Column:      column,
		ASPCode:     0,
		Category:    "ASP",
	}
}

// GetProperty gets an ASPError property by name
func (e *ASPError) GetProperty(name string) interface{} {
	if e == nil {
		return nil
	}

	nameLower := strings.ToLower(name)
	switch nameLower {
	case "aspcode":
		return e.ASPCode
	case "aspdescription":
		return e.ASPDescription
	case "category":
		return e.Category
	case "column":
		return e.Column
	case "description":
		return e.Description
	case "file":
		return e.File
	case "line":
		return e.Line
	case "number":
		return e.Number
	case "source":
		return e.Source
	default:
		return nil
	}
}

// String returns a string representation of the error
func (e *ASPError) String() string {
	if e == nil {
		return "No error"
	}
	return fmt.Sprintf("ASP Error %d: %s (File: %s, Line: %d)", e.Number, e.Description, e.File, e.Line)
}

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
// Implements all properties from Classic ASP ASPError object
type ASPError struct {
	// Core ASP Error Properties
	ASPCode        int    // ASP-specific error code (0-500)
	ASPDescription string // Description of the ASP error
	Category       string // Error category ("ASP", "VBScript", "ADODB", etc.)
	Column         int    // Column number where error occurred
	Description    string // Full error description
	File           string // File where error occurred
	Line           int    // Line number where error occurred
	Number         int    // Error number (VBScript error code or HTTP status)
	Source         string // Source of the error (component name, function, etc.)

	// Extended properties for debugging
	Stack     []string  // Stack trace
	Context   string    // Code context where error occurred
	Timestamp time.Time // When error occurred
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

// GetExecutor returns the executor
func (s *ServerObject) GetExecutor() interface {
	CreateObject(string) (interface{}, error)
} {
	return s.executor
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

	if s.context == nil {
		return fmt.Errorf("Execute not available: no execution context")
	}

	// Use executor to execute the file
	if exec, ok := s.executor.(interface{ ExecuteASPPath(string) error }); ok {
		return exec.ExecuteASPPath(path)
	}

	return fmt.Errorf("executor does not support Execute")
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

// MapPath resolves a virtual/relative path to a physical path
// Delegates to the execution context for script-aware resolution
func (s *ServerObject) MapPath(path string) string {
	if s.context != nil {
		return s.context.Server_MapPath(path)
	}

	// Fallback: legacy root-based resolution
	if path == "" || path == "/" || path == "\\" {
		return s.rootDir
	}
	path = strings.ReplaceAll(path, "\\", "/")
	fullPath := filepath.Join(s.rootDir, strings.TrimPrefix(path, "/"))
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

	// Clear the current response buffer
	s.context.Response.Clear()

	// Use executor to execute the file
	if exec, ok := s.executor.(interface{ ExecuteASPPath(string) error }); ok {
		if err := exec.ExecuteASPPath(path); err != nil {
			return err
		}
		// Signal to stop execution of the current page
		// We return a special error that the executor handles
		return ErrServerTransfer
	}

	return fmt.Errorf("executor does not support Transfer")
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

// NewASPError creates a new ASP error object with all details
func NewASPError(number int, description string, source string, file string, line int, column int) *ASPError {
	category := DetermineErrorCategory(number)
	aspCode := DetermineASPCode(number)

	return &ASPError{
		Number:         number,
		Description:    description,
		Source:         source,
		File:           file,
		Line:           line,
		Column:         column,
		ASPCode:        aspCode,
		ASPDescription: description,
		Category:       category,
		Stack:          make([]string, 0),
		Timestamp:      time.Now(),
	}
}

// NewASPErrorFromVBScript creates an ASP error from VBScript parser error
func NewASPErrorFromVBScript(vbErrorCode int, description string, file string, line int, column int, context string) *ASPError {
	err := &ASPError{
		Number:         vbErrorCode,
		Description:    FormatVBScriptError(vbErrorCode, description),
		Source:         "VBScript",
		File:           file,
		Line:           line,
		Column:         column,
		ASPCode:        0,
		ASPDescription: "",
		Category:       "VBScript",
		Stack:          make([]string, 0),
		Context:        context,
		Timestamp:      time.Now(),
	}

	return err
}

// AddStackFrame adds a stack frame to the error
func (e *ASPError) AddStackFrame(frame string) {
	if e != nil {
		e.Stack = append(e.Stack, frame)
	}
}

// DetermineErrorCategory determines the category based on error number
func DetermineErrorCategory(number int) string {
	// VBScript errors: 1-65535
	if number >= 1 && number <= 65535 {
		// ASP-specific VBScript errors: 1000-1999
		if number >= 1000 && number <= 1999 {
			return "VBScript"
		}
		return "VBScript"
	}

	// HTTP errors: 400-599
	if number >= 400 && number <= 599 {
		return "HTTP"
	}

	// ADODB errors: -2147467259 to -2147467247
	if number >= -2147467259 && number <= -2147467247 {
		return "ADODB"
	}

	return "ASP"
}

// DetermineASPCode determines ASP-specific error code
func DetermineASPCode(number int) int {
	// Map common VBScript errors to ASP codes
	switch number {
	case 1002, 1003, 1005, 1006, 1007: // Syntax errors
		return 100 // ASP 0100: Out of Memory or similar
	case 424: // Object required
		return 115 // ASP 0115: Unexpected error
	case 91: // Object variable not set
		return 115
	default:
		if number >= 1000 && number <= 1999 {
			return 101 // ASP 0101: Unexpected error
		}
		return 0
	}
}

// FormatVBScriptError formats a VBScript error message with proper description
func FormatVBScriptError(errorCode int, customMsg string) string {
	// If custom message provided, use it
	if customMsg != "" {
		return customMsg
	}

	// Standard VBScript error messages
	switch errorCode {
	case 1002:
		return "Syntax error"
	case 1003:
		return "Expected ':'"
	case 1005:
		return "Expected '('"
	case 1006:
		return "Expected ')'"
	case 1007:
		return "Expected ']'"
	case 1010:
		return "Expected identifier"
	case 1011:
		return "Expected '='"
	case 1012:
		return "Expected 'If'"
	case 1013:
		return "Expected 'To'"
	case 1014:
		return "Expected 'End'"
	case 1015:
		return "Expected 'Function'"
	case 1016:
		return "Expected 'Sub'"
	case 1017:
		return "Expected 'Then'"
	case 1018:
		return "Expected 'Wend'"
	case 1019:
		return "Expected 'Loop'"
	case 1020:
		return "Expected 'Next'"
	case 1021:
		return "Expected 'Case'"
	case 1022:
		return "Expected 'Select'"
	case 1023:
		return "Expected expression"
	case 1024:
		return "Expected statement"
	case 1025:
		return "Expected end of statement"
	case 1026:
		return "Expected integer constant"
	case 1027:
		return "Expected 'While' or 'Until'"
	case 1028:
		return "Expected 'While', 'Until' or end of statement"
	case 1029:
		return "Expected 'With'"
	case 1030:
		return "Identifier too long"
	case 1031:
		return "Invalid number"
	case 1032:
		return "Invalid character"
	case 1033:
		return "Unterminated string constant"
	case 1034:
		return "Unterminated comment"
	case 1037:
		return "Invalid use of 'Me' keyword"
	case 1038:
		return "'Loop' without 'Do'"
	case 1039:
		return "Invalid 'Exit' statement"
	case 1040:
		return "Invalid 'for' loop control variable"
	case 1041:
		return "Name redefined"
	case 1042:
		return "Must be first statement on the line"
	case 1043:
		return "Cannot assign to non-ByVal argument"
	case 1044:
		return "Cannot use parentheses when calling a Sub"
	case 1045:
		return "Expected literal constant"
	case 1046:
		return "Expected 'In'"
	case 1047:
		return "Expected 'Class'"
	case 1048:
		return "Must be defined inside a Class"
	case 1049:
		return "Expected 'Let' or 'Get' or 'Set'"
	case 1050:
		return "Expected 'Property'"
	case 1051:
		return "Number of arguments must be consistent across properties specification"
	case 1052:
		return "Cannot have multiple default property/method in a Class"
	case 1053:
		return "'Class_Initialize' or 'Class_Terminate' do not have arguments"
	case 1054:
		return "'Property Set' or 'Property Let' must have at least one argument"
	case 1055:
		return "Unexpected 'Next'"
	case 1056:
		return "'Default' specification can only be on Property Get"
	case 1057:
		return "'Default' specification must also specify 'Public'"
	case 1058:
		return "'Default' specification can only be on Property Get"
	default:
		return fmt.Sprintf("VBScript error %d", errorCode)
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
	case "stack":
		return strings.Join(e.Stack, "\n")
	case "context":
		return e.Context
	case "timestamp":
		return e.Timestamp
	default:
		return nil
	}
}

// String returns a string representation of the error
func (e *ASPError) String() string {
	if e == nil {
		return "No error"
	}

	var builder strings.Builder

	// Main error message
	if e.Number != 0 {
		builder.WriteString(fmt.Sprintf("Error Number: %d\n", e.Number))
	}

	if e.ASPCode != 0 {
		builder.WriteString(fmt.Sprintf("ASP Code: %d\n", e.ASPCode))
	}

	if e.Description != "" {
		builder.WriteString(fmt.Sprintf("Description: %s\n", e.Description))
	}

	if e.Source != "" {
		builder.WriteString(fmt.Sprintf("Source: %s\n", e.Source))
	}

	if e.File != "" {
		builder.WriteString(fmt.Sprintf("File: %s\n", e.File))
	}

	if e.Line > 0 {
		builder.WriteString(fmt.Sprintf("Line: %d\n", e.Line))
	}

	if e.Column > 0 {
		builder.WriteString(fmt.Sprintf("Column: %d\n", e.Column))
	}

	if e.Category != "" {
		builder.WriteString(fmt.Sprintf("Category: %s\n", e.Category))
	}

	if e.Context != "" {
		builder.WriteString(fmt.Sprintf("\nContext:\n%s\n", e.Context))
	}

	if len(e.Stack) > 0 {
		builder.WriteString("\nStack Trace:\n")
		for i, frame := range e.Stack {
			builder.WriteString(fmt.Sprintf("  %d. %s\n", i+1, frame))
		}
	}

	return builder.String()
}

// GetHTMLFormattedError returns an HTML-formatted error message for debugging
func (e *ASPError) GetHTMLFormattedError() string {
	if e == nil {
		return ""
	}

	var builder strings.Builder
	builder.WriteString("<div style='border: 2px solid #dc3545; padding: 20px; margin: 20px; background: #fff; font-family: monospace;'>")
	builder.WriteString("<h2 style='color: #dc3545; margin-top: 0;'>ASP Error</h2>")

	if e.Number != 0 {
		builder.WriteString(fmt.Sprintf("<p><strong>Error Number:</strong> %d", e.Number))
		if e.ASPCode != 0 {
			builder.WriteString(fmt.Sprintf(" (ASP Code: %d)", e.ASPCode))
		}
		builder.WriteString("</p>")
	}

	if e.Description != "" {
		builder.WriteString(fmt.Sprintf("<p><strong>Description:</strong> %s</p>", html.EscapeString(e.Description)))
	}

	if e.Source != "" {
		builder.WriteString(fmt.Sprintf("<p><strong>Source:</strong> %s</p>", html.EscapeString(e.Source)))
	}

	if e.File != "" {
		builder.WriteString(fmt.Sprintf("<p><strong>File:</strong> %s", html.EscapeString(e.File)))
		if e.Line > 0 {
			builder.WriteString(fmt.Sprintf(", <strong>Line:</strong> %d", e.Line))
		}
		if e.Column > 0 {
			builder.WriteString(fmt.Sprintf(", <strong>Column:</strong> %d", e.Column))
		}
		builder.WriteString("</p>")
	}

	if e.Category != "" {
		builder.WriteString(fmt.Sprintf("<p><strong>Category:</strong> %s</p>", e.Category))
	}

	if e.Context != "" {
		builder.WriteString("<p><strong>Context:</strong></p>")
		builder.WriteString(fmt.Sprintf("<pre style='background: #f5f5f5; padding: 10px; overflow-x: auto;'>%s</pre>", html.EscapeString(e.Context)))
	}

	if len(e.Stack) > 0 {
		builder.WriteString("<p><strong>Stack Trace:</strong></p>")
		builder.WriteString("<ol style='background: #f5f5f5; padding: 20px;'>")
		for _, frame := range e.Stack {
			builder.WriteString(fmt.Sprintf("<li>%s</li>", html.EscapeString(frame)))
		}
		builder.WriteString("</ol>")
	}

	builder.WriteString(fmt.Sprintf("<p style='color: #666; font-size: 0.9em;'><strong>Timestamp:</strong> %s</p>", e.Timestamp.Format("2006-01-02 15:04:05")))
	builder.WriteString("</div>")

	return builder.String()
}

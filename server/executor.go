package server

import (
	"fmt"
	"go-asp/asp"
	"math"
	"net/http"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/guimaraeslucas/vbscript-go/ast"
)

// LoopExitError represents a loop exit statement (Exit For, Exit Do, etc)
type LoopExitError struct {
	LoopType string // "for", "do", "while", "select"
}

func (e *LoopExitError) Error() string {
	return fmt.Sprintf("Exit %s", e.LoopType)
}

// Global Application singleton (shared across all requests)
var globalApplication *ApplicationObject
var globalAppOnce sync.Once

// GetGlobalApplication returns the singleton Application object
func GetGlobalApplication() *ApplicationObject {
	globalAppOnce.Do(func() {
		globalApplication = NewApplicationObject()
	})
	return globalApplication
}

// ExecutionContext holds all runtime state for ASP execution
type ExecutionContext struct {
	// ASP core objects
	Request     *RequestObject
	Response    *ResponseObject
	Server      *asp.ServerObject
	Session     *SessionObject
	Application *ApplicationObject

	// Variable storage (case-insensitive keys)
	variables map[string]interface{}

	// Constant storage (case-insensitive keys) - read-only values
	constants map[string]interface{}

	// HTTP context
	httpWriter  http.ResponseWriter
	httpRequest *http.Request

	// Execution state
	startTime time.Time
	timeout   time.Duration

	// Library instances
	libraries map[string]interface{}

	// Configuration
	RootDir string

	// Session management
	sessionID      string
	sessionManager *SessionManager

	// Scoping
	scopeStack    []map[string]interface{}
	contextObject interface{} // For Class Instance (Me)

	// Mutex for thread safety
	mu sync.RWMutex
}

// NewExecutionContext creates a new execution context
func NewExecutionContext(w http.ResponseWriter, r *http.Request, sessionID string, timeout time.Duration) *ExecutionContext {
	sessionManager := GetSessionManager()

	// Load or create session data
	sessionData, err := sessionManager.GetOrCreateSession(sessionID)
	if err != nil {
		fmt.Printf("Warning: Failed to load session: %v\n", err)
		// Create empty session data as fallback
		sessionData = &SessionData{
			ID:           sessionID,
			Data:         make(map[string]interface{}),
			CreatedAt:    time.Now(),
			LastAccessed: time.Now(),
			Timeout:      20,
		}
	}

	return &ExecutionContext{
		Request:        NewRequestObject(),
		Response:       NewResponseObject(w, r),
		Server:         asp.NewServerObject(),
		Session:        NewSessionObject(sessionID, sessionData.Data),
		Application:    GetGlobalApplication(),
		variables:      make(map[string]interface{}),
		constants:      make(map[string]interface{}),
		libraries:      make(map[string]interface{}),
		scopeStack:     make([]map[string]interface{}, 0),
		httpWriter:     w,
		httpRequest:    r,
		startTime:      time.Now(),
		timeout:        timeout,
		sessionID:      sessionID,
		sessionManager: sessionManager,
	}
}

// PushScope pushes a new local scope
func (ec *ExecutionContext) PushScope() {
	ec.mu.Lock()
	defer ec.mu.Unlock()
	ec.scopeStack = append(ec.scopeStack, make(map[string]interface{}))
}

// PopScope pops the last local scope
func (ec *ExecutionContext) PopScope() {
	ec.mu.Lock()
	defer ec.mu.Unlock()
	if len(ec.scopeStack) > 0 {
		ec.scopeStack = ec.scopeStack[:len(ec.scopeStack)-1]
	}
}

// SetContextObject sets the current context object (e.g. Class Instance)
func (ec *ExecutionContext) SetContextObject(obj interface{}) {
	ec.mu.Lock()
	defer ec.mu.Unlock()
	ec.contextObject = obj
}

// GetContextObject returns the current context object
func (ec *ExecutionContext) GetContextObject() interface{} {
	ec.mu.RLock()
	defer ec.mu.RUnlock()
	return ec.contextObject
}

// SetVariable sets a variable in the execution context (case-insensitive)
// Returns error if attempting to overwrite a constant
func (ec *ExecutionContext) SetVariable(name string, value interface{}) error {
	nameLower := strings.ToLower(name)

	ec.mu.Lock()
	defer ec.mu.Unlock()

	// Check if this is a constant
	if _, exists := ec.constants[nameLower]; exists {
		return fmt.Errorf("cannot reassign constant '%s'", name)
	}

	// 1. Check Local Scopes (Top to Bottom)
	for i := len(ec.scopeStack) - 1; i >= 0; i-- {
		if _, exists := ec.scopeStack[i][nameLower]; exists {
			ec.scopeStack[i][nameLower] = value
			return nil
		}
	}

	// 2. Check Context Object (Class Member)
	if ec.contextObject != nil {
		if internal, ok := ec.contextObject.(interface {
			SetMember(string, interface{}) bool
		}); ok {
			if internal.SetMember(nameLower, value) {
				return nil
			}
		}
	}

	// 3. Global Variables (Default)
	// If not found in locals, set in global (Variables)
	// Note: If we are in a scope (Function), and it's not declared with Dim,
	// VBScript behavior depends on Option Explicit.
	// Assuming Option Explicit is OFF (or we treat it loosely), it goes to Global?
	// Actually, if it's not Dim'ed locally, it searches up. If nowhere, it creates Global.

	// BUT, if we want to support 'Dim x' inside function, 'visitVariableDeclaration' calls 'SetVariable'.
	// We need 'DefineVariable' vs 'SetVariable'.
	// 'visitVariableDeclaration' should put it in the CURRENT scope.

	// For now, standard SetVariable puts in global if not found in locals.
	ec.variables[nameLower] = value
	return nil
}

// DefineVariable defines a variable in the current scope (Dim)
func (ec *ExecutionContext) DefineVariable(name string, value interface{}) error {
	nameLower := strings.ToLower(name)
	ec.mu.Lock()
	defer ec.mu.Unlock()

	if len(ec.scopeStack) > 0 {
		// Define in top scope
		ec.scopeStack[len(ec.scopeStack)-1][nameLower] = value
	} else {
		// Define global
		ec.variables[nameLower] = value
	}
	return nil
}

// GetVariable gets a variable from the execution context (case-insensitive)
func (ec *ExecutionContext) GetVariable(name string) (interface{}, bool) {
	nameLower := strings.ToLower(name)

	ec.mu.RLock()
	defer ec.mu.RUnlock()

	// 1. Check constants
	if val, exists := ec.constants[nameLower]; exists {
		return val, true
	}

	// 2. Check Local Scopes (Top to Bottom)
	for i := len(ec.scopeStack) - 1; i >= 0; i-- {
		if val, exists := ec.scopeStack[i][nameLower]; exists {
			return val, true
		}
	}

	// 3. Check Context Object (Class Member)
	if ec.contextObject != nil {
		// Try internal access first (allows Private members)
		if internal, ok := ec.contextObject.(interface {
			GetMember(string) (interface{}, bool)
		}); ok {
			if val, found := internal.GetMember(nameLower); found {
				return val, true
			}
		} else if getter, ok := ec.contextObject.(interface{ GetProperty(string) interface{} }); ok {
			val := getter.GetProperty(nameLower)
			if val != nil {
				return val, true
			}
		}
	}

	// 4. Check Global Variables
	val, exists := ec.variables[nameLower]
	return val, exists
}

// SetConstant sets a constant in the execution context (case-insensitive)
// Constants cannot be changed after initialization
func (ec *ExecutionContext) SetConstant(name string, value interface{}) error {
	nameLower := strings.ToLower(name)

	ec.mu.Lock()
	defer ec.mu.Unlock()

	// Check if constant already exists
	if _, exists := ec.constants[nameLower]; exists {
		return fmt.Errorf("constant '%s' already defined", name)
	}

	// Check if variable with same name exists
	if _, exists := ec.variables[nameLower]; exists {
		return fmt.Errorf("cannot define constant with name of existing variable '%s'", name)
	}

	ec.constants[nameLower] = value
	return nil
}

// CheckTimeout checks if execution has exceeded timeout
func (ec *ExecutionContext) CheckTimeout() error {
	if time.Since(ec.startTime) > ec.timeout {
		return fmt.Errorf("execution timeout exceeded (%v)", ec.timeout)
	}
	return nil
}

// Server_MapPath converts a virtual path to an absolute file system path
func (ec *ExecutionContext) Server_MapPath(path string) string {
	rootDir := ec.RootDir
	if rootDir == "" {
		rootDir = "./www"
	}

	// Handle different path formats
	if path == "/" || path == "" {
		return rootDir
	}

	// Remove leading slash if present
	if len(path) > 0 && (path[0] == '/' || path[0] == '\\') {
		path = path[1:]
	}

	// Join with root directory
	fullPath := fmt.Sprintf("%s%c%s", rootDir, '/', strings.ReplaceAll(path, "\\", "/"))

	return fullPath
}

// EvaluateExpression evaluates a simple expression (simplified version for legacy helpers)
// This is a simple wrapper that returns the value as-is or converts strings
// The full expression evaluation is handled by the VBScript parser
func EvaluateExpression(expr interface{}, ctx *ExecutionContext) interface{} {
	// If it's already a value, return it
	if expr == nil {
		return nil
	}

	// If it's a string, check if it's a variable name
	if strExpr, ok := expr.(string); ok {
		// Try to get as variable
		if val, exists := ctx.GetVariable(strExpr); exists {
			return val
		}
		// Otherwise return the string itself
		return strExpr
	}

	// Return as-is for other types
	return expr
}

// ASPExecutor handles execution of ASP code with VBScript programs
type ASPExecutor struct {
	config  *ASPProcessorConfig
	context *ExecutionContext
}

// NewASPExecutor creates a new ASP executor
func NewASPExecutor(config *ASPProcessorConfig) *ASPExecutor {
	if config == nil {
		config = &ASPProcessorConfig{
			RootDir:       "./www",
			ScriptTimeout: 30,
		}
	}

	return &ASPExecutor{
		config: config,
	}
}

// Execute processes ASP code and returns rendered output
func (ae *ASPExecutor) Execute(fileContent string, w http.ResponseWriter, r *http.Request, sessionID string) error {
	// Create execution context
	timeout := time.Duration(ae.config.ScriptTimeout) * time.Second
	ae.context = NewExecutionContext(w, r, sessionID, timeout)

	// Set RootDir in context
	ae.context.RootDir = ae.config.RootDir

	// Configure Server object with context
	ae.context.Server.SetProperty("_rootDir", ae.config.RootDir)
	ae.context.Server.SetProperty("_httpRequest", r)
	ae.context.Server.SetProperty("_executor", ae)

	// Populate Request object
	populateRequestData(ae.context.Request, r)

	// Parse ASP code
	parsingOptions := &asp.ASPParsingOptions{
		SaveComments:      false,
		StrictMode:        false,
		AllowImplicitVars: true,
		DebugMode:         ae.config.DebugASP,
	}
	parser := asp.NewASPParserWithOptions(fileContent, parsingOptions)
	result, err := parser.Parse()
	if err != nil {
		return fmt.Errorf("failed to parse ASP code: %w", err)
	}

	// Check for parse errors
	if len(result.Errors) > 0 {
		if ae.config.DebugASP {
			fmt.Println("[ASP Parse] Multiple errors found during parsing:")
			for _, parseErr := range result.Errors {
				fmt.Printf("  - %v\n", parseErr)
			}
		}
		return fmt.Errorf("ASP parse error: %v", result.Errors[0])
	}

	// Execute blocks in order with timeout protection
	done := make(chan error, 1)

	go func() {
		defer func() {
			if rec := recover(); rec != nil {
				done <- fmt.Errorf("runtime panic: %v", rec)
			}
		}()

		err := ae.executeBlocks(result)
		done <- err
	}()

	// Wait for execution or timeout
	select {
	case err := <-done:
		if err != nil {
			return err
		}
	case <-time.After(timeout):
		return fmt.Errorf("script execution timeout (>%d seconds)", ae.config.ScriptTimeout)
	}

	// Write response to HTTP ResponseWriter
	// Check if Response.End() was already called (which flushes headers/content)
	if !ae.context.Response.IsEnded() {
		// Flush the Response object (writes headers and buffer)
		if err := ae.context.Response.Flush(); err != nil {
			return fmt.Errorf("failed to flush response: %w", err)
		}
	}

	// Set session cookie (if not already set by Response.Flush)
	http.SetCookie(w, &http.Cookie{
		Name:     "ASPSESSIONID",
		Value:    ae.context.sessionID,
		Path:     "/",
		HttpOnly: false,
		Secure:   false,
		MaxAge:   20 * 60, // 20 minutes (matches ASP session timeout)
		SameSite: http.SameSiteStrictMode,
	})

	// Save session data to file after request completes
	if err := ae.saveSession(); err != nil {
		fmt.Printf("Warning: Failed to save session: %v\n", err)
	}

	return nil
}

// saveSession persists the current session data to file
func (ae *ASPExecutor) saveSession() error {
	if ae.context == nil || ae.context.sessionManager == nil {
		return fmt.Errorf("no session context available")
	}

	sessionData := &SessionData{
		ID:           ae.context.sessionID,
		Data:         ae.context.Session.Data,
		LastAccessed: time.Now(),
		Timeout:      20, // Default timeout in minutes
	}

	// Load existing session to preserve CreatedAt
	existingSession, err := ae.context.sessionManager.LoadSession(ae.context.sessionID)
	if err == nil {
		sessionData.CreatedAt = existingSession.CreatedAt
		sessionData.Timeout = existingSession.Timeout
	} else {
		sessionData.CreatedAt = time.Now()
	}

	return ae.context.sessionManager.SaveSession(sessionData)
}

// executeBlocks executes all blocks in order (HTML and ASP)
func (ae *ASPExecutor) executeBlocks(result *asp.ASPParserResult) error {
	for i, block := range result.Blocks {
		// Check timeout periodically
		if i%100 == 0 {
			if err := ae.context.CheckTimeout(); err != nil {
				return err
			}
		}

		switch block.Type {
		case "html":
			// Write HTML content directly
			if err := ae.context.Response.Write(block.Content); err != nil {
				return fmt.Errorf("failed to write HTML: %w", err)
			}

		case "asp":
			// Execute VBScript block if parsed
			if program, exists := result.VBPrograms[i]; exists && program != nil {
				if err := ae.executeVBProgram(program); err != nil {
					// Check if it's a Response.End() or Response.Redirect() signal
					if err.Error() == "RESPONSE_END" {
						// This is normal - stop execution and return
						return nil
					}
					return err
				}
			}
		}
	}

	return nil
}

// executeVBProgram executes a VBScript AST program
func (ae *ASPExecutor) executeVBProgram(program *ast.Program) error {
	if program == nil {
		return nil
	}

	// Create a visitor to traverse the AST
	v := NewASPVisitor(ae.context, ae)

	// Visit all statements in the program
	if program.Body != nil {
		for _, stmt := range program.Body {
			if stmt == nil {
				continue
			}

			// Check timeout
			if err := ae.context.CheckTimeout(); err != nil {
				return err
			}

			// Execute statement
			if err := v.VisitStatement(stmt); err != nil {
				// Check if it's a Response.End() or Response.Redirect() signal
				if err.Error() == "RESPONSE_END" {
					// This is normal - stop execution and return
					return err
				}
				return err
			}
		}
	}

	return nil
}

// CreateObject creates an ASP COM object (like Server.CreateObject)
func (ae *ASPExecutor) CreateObject(objType string) (interface{}, error) {
	objType = strings.ToUpper(objType)

	switch objType {
	case "G3JSON":
		return NewJSONLibrary(ae.context), nil
	case "G3FILES":
		return NewFileSystemLibrary(ae.context), nil
	case "G3HTTP":
		return NewHTTPLibrary(ae.context), nil
	case "G3TEMPLATE":
		return NewTemplateLibrary(ae.context), nil
	case "G3MAIL":
		return NewMailLibrary(ae.context), nil
	case "G3CRYPTO":
		return NewCryptoLibrary(ae.context), nil
	case "SCRIPTING.FILESYSTEMOBJECT":
		return NewFileSystemObjectLibrary(ae.context), nil
	case "MSXML2.SERVERXMLHTTP":
		return NewServerXMLHTTP(ae.context), nil
	case "MSXML2.DOMDOCUMENT":
		return NewDOMDocument(ae.context), nil
	case "ADODB.CONNECTION":
		return NewADOConnection(ae.context), nil
	case "ADODB.RECORDSET":
		return NewADORecordset(ae.context), nil
	case "ADODB.STREAM":
		return NewADOStream(ae.context), nil
	case "SCRIPTING.DICTIONARY":
		return NewDictionary(ae.context), nil
	default:
		return nil, fmt.Errorf("unsupported object type: %s", objType)
	}
}

// ASPVisitor traverses and executes the VBScript AST
type ASPVisitor struct {
	context   *ExecutionContext
	executor  *ASPExecutor
	depth     int
	withStack []interface{}
}

// NewASPVisitor creates a new ASP visitor for AST traversal
func NewASPVisitor(ctx *ExecutionContext, executor *ASPExecutor) *ASPVisitor {
	return &ASPVisitor{
		context:   ctx,
		executor:  executor,
		depth:     0,
		withStack: make([]interface{}, 0),
	}
}

// VisitStatement executes a single statement from the AST

func (v *ASPVisitor) VisitStatement(node ast.Statement) error {

	if node == nil {

		return nil

	}

	v.depth++

	if v.depth > 1000 {
		return fmt.Errorf("maximum call depth exceeded")
	}
	defer func() { v.depth-- }()

	switch stmt := node.(type) {
	case *ast.AssignmentStatement:
		return v.visitAssignment(stmt)

	case *ast.CallStatement:
		_, err := v.visitExpression(stmt.Callee)
		return err

	case *ast.CallSubStatement:
		return v.visitCallSubStatement(stmt)

	case *ast.ReDimStatement:
		return v.visitReDim(stmt)

	case *ast.IfStatement:
		return v.visitIf(stmt)

	case *ast.ForStatement:
		return v.visitFor(stmt)

	case *ast.ForEachStatement:
		return v.visitForEach(stmt)

	case *ast.DoStatement:
		return v.visitDo(stmt)

	case *ast.WhileStatement:
		return v.visitWhile(stmt)

	case *ast.SelectStatement:
		return v.visitSelect(stmt)

	case *ast.SubDeclaration:
		return v.visitSubDeclaration(stmt)

	case *ast.FunctionDeclaration:
		return v.visitFunctionDeclaration(stmt)

	case *ast.ClassDeclaration:
		return v.visitClassDeclaration(stmt)

	case *ast.WithStatement:
		return v.visitWithStatement(stmt)

	case *ast.VariableDeclaration:
		return v.visitVariableDeclaration(stmt)

	case *ast.VariablesDeclaration:
		return v.visitVariablesDeclaration(stmt)

	case *ast.ConstsDeclaration:
		return v.visitConstDeclaration(stmt)

	case *ast.StatementList:
		for _, s := range stmt.Statements {
			if err := v.VisitStatement(s); err != nil {
				return err
			}
		}

	case *ast.OnErrorResumeNextStatement:
		// Error handling - continue on error
		return nil

	case *ast.OnErrorGoTo0Statement:
		// Error handling - reset error
		return nil

	default:
		// Try to evaluate as expression for side effects
		if expr, ok := node.(ast.Expression); ok {
			_, err := v.visitExpression(expr)
			return err
		}
	}

	return nil
}

// visitAssignment handles variable assignment
func (v *ASPVisitor) visitAssignment(stmt *ast.AssignmentStatement) error {
	if stmt == nil || stmt.Right == nil {
		return nil
	}

	// Evaluate right-hand side
	value, err := v.visitExpression(stmt.Right)
	if err != nil {
		return err
	}

	// Handle different left-hand side patterns

	// Case 1: Simple variable assignment (Dim x = 5 or x = 5)
	if ident, ok := stmt.Left.(*ast.Identifier); ok {
		if err := v.context.SetVariable(ident.Name, value); err != nil {
			return err
		}
		return nil
	}

	// Case 2: Indexed/Property assignment (obj("key") = value or obj.prop = value)
	if indexCall, ok := stmt.Left.(*ast.IndexOrCallExpression); ok {
		// For array indexing, we need to get the variable name directly
		// to modify the original array, not a copy
		if ident, ok := indexCall.Object.(*ast.Identifier); ok {
			varName := ident.Name
			varNameLower := strings.ToLower(varName)

			// Check if it's a built-in ASP object first
			var obj interface{}
			switch varNameLower {
			case "session":
				obj = v.context.Session
			case "application":
				obj = v.context.Application
			case "request":
				obj = v.context.Request
			case "response":
				obj = v.context.Response
			case "server":
				obj = v.context.Server
			default:
				// Otherwise get from variables
				var exists bool
				obj, exists = v.context.GetVariable(varName)
				if !exists {
					return fmt.Errorf("variable '%s' is undefined", varName)
				}
			}

			// Handle array index assignment (arr(0) = value)
			if arrObj, ok := obj.([]interface{}); ok {
				if len(indexCall.Indexes) > 0 {
					// Get the index
					idx, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					index := toInt(idx)
					// Bounds check
					if index >= 0 && index < len(arrObj) {
						arrObj[index] = value
						// Re-set the variable to ensure it's updated
						_ = v.context.SetVariable(varName, arrObj)
						return nil
					}
					// Out of bounds - VBScript error
					return fmt.Errorf("subscript out of range")
				}
			}

			// Handle map assignment
			if mapObj, ok := obj.(map[string]interface{}); ok {
				if len(indexCall.Indexes) > 0 {
					key, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					mapObj[fmt.Sprintf("%v", key)] = value
					return nil
				}
			}

			// Handle SessionObject assignment (Session("key") = value)
			if sessionObj, ok := obj.(*SessionObject); ok {
				if len(indexCall.Indexes) > 0 {
					key, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					return sessionObj.SetIndex(key, value)
				}
			}

			// Handle ApplicationObject assignment (Application("key") = value)
			if appObj, ok := obj.(*ApplicationObject); ok {
				if len(indexCall.Indexes) > 0 {
					key, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					appObj.Set(fmt.Sprintf("%v", key), value)
					return nil
				}
			}

			// Handle ASP Library wrapper
			if lib, ok := obj.(interface {
				SetProperty(string, interface{}) error
			}); ok && len(indexCall.Indexes) > 0 {
				key, err := v.visitExpression(indexCall.Indexes[0])
				if err != nil {
					return err
				}
				return lib.SetProperty(fmt.Sprintf("%v", key), value)
			}
		} else {
			// For complex expressions, evaluate normally
			obj, err := v.visitExpression(indexCall.Object)
			if err != nil {
				return err
			}

			// If it's a map (dictionary-like object), set the indexed property
			if mapObj, ok := obj.(map[string]interface{}); ok {
				if len(indexCall.Indexes) > 0 {
					key, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					mapObj[fmt.Sprintf("%v", key)] = value
					return nil
				}
			}

			// If it's a SessionObject, set the indexed property
			if sessionObj, ok := obj.(*SessionObject); ok {
				if len(indexCall.Indexes) > 0 {
					key, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					return sessionObj.SetIndex(key, value)
				}
			}

			// If it's an ApplicationObject, set the indexed property
			if appObj, ok := obj.(*ApplicationObject); ok {
				if len(indexCall.Indexes) > 0 {
					key, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					appObj.Set(fmt.Sprintf("%v", key), value)
					return nil
				}
			}

			// If it's an ASP Library wrapper, try to call a setter
			if lib, ok := obj.(interface {
				SetProperty(string, interface{}) error
			}); ok && len(indexCall.Indexes) > 0 {
				key, err := v.visitExpression(indexCall.Indexes[0])
				if err != nil {
					return err
				}
				return lib.SetProperty(fmt.Sprintf("%v", key), value)
			}
		}
	}

	// Case 3: Member assignment (obj.prop = value)
	if member, ok := stmt.Left.(*ast.MemberExpression); ok {
		// Get the object
		obj, err := v.visitExpression(member.Object)
		if err != nil {
			return err
		}

		// Get property name
		propName := ""
		if member.Property != nil {
			propName = member.Property.Name
		}

		// If it's an ASP object, set the property
		if aspObj, ok := obj.(asp.ASPObject); ok {
			return aspObj.SetProperty(propName, value)
		}

		// If it's a ResponseObject, set the property
		if respObj, ok := obj.(*ResponseObject); ok {
			propNameLower := strings.ToLower(propName)
			switch propNameLower {
			case "buffer":
				if b, ok := value.(bool); ok {
					respObj.SetBuffer(b)
				}
				return nil
			case "cachecontrol":
				respObj.SetCacheControl(fmt.Sprintf("%v", value))
				return nil
			case "charset":
				respObj.SetCharset(fmt.Sprintf("%v", value))
				return nil
			case "contenttype":
				respObj.SetContentType(fmt.Sprintf("%v", value))
				return nil
			case "expires":
				respObj.SetExpires(toInt(value))
				return nil
			case "expiresabsolute":
				if t, ok := value.(time.Time); ok {
					respObj.SetExpiresAbsolute(t)
				}
				return nil
			case "pics":
				respObj.SetPICS(fmt.Sprintf("%v", value))
				return nil
			case "status":
				respObj.SetStatus(fmt.Sprintf("%v", value))
				return nil
			}
		}
	}

	return nil
}

// visitReDim handles ReDim statements
func (v *ASPVisitor) visitReDim(stmt *ast.ReDimStatement) error {
	if stmt == nil || stmt.ReDims == nil {
		return nil
	}

	for _, redim := range stmt.ReDims {
		if redim == nil || redim.Identifier == nil {
			continue
		}

		varName := redim.Identifier.Name

		// Evaluate array dimensions
		dims := make([]int, len(redim.ArrayDims))
		for i, dimExpr := range redim.ArrayDims {
			dimVal, err := v.visitExpression(dimExpr)
			if err != nil {
				return err
			}
			dims[i] = toInt(dimVal)
		}

		if stmt.Preserve {
			// ReDim Preserve - resize array keeping existing elements
			oldVal, _ := v.context.GetVariable(varName)
			var newArr []interface{}

			if oldArr, ok := oldVal.([]interface{}); ok {
				// Create new array with new dimensions
				newArr = v.makeNestedArray(dims)

				// Copy old values to new array
				v.preserveCopy(oldArr, newArr, dims)
			} else {
				// If variable doesn't exist or isn't an array, create new array
				newArr = v.makeNestedArray(dims)
			}

			_ = v.context.SetVariable(varName, newArr)
		} else {
			// ReDim without Preserve - create new array (old values lost)
			newArr := v.makeNestedArray(dims)
			_ = v.context.SetVariable(varName, newArr)
		}
	}

	return nil
}

// preserveCopy copies elements from old array to new array
func (v *ASPVisitor) preserveCopy(oldArr, newArr []interface{}, dims []int) {
	if len(oldArr) == 0 || len(newArr) == 0 {
		return
	}

	// Determine how many elements to copy (minimum of old and new lengths)
	copyLen := len(oldArr)
	if len(newArr) < copyLen {
		copyLen = len(newArr)
	}

	if len(dims) == 1 {
		// Single-dimensional array: direct copy
		copy(newArr, oldArr[:copyLen])
	} else {
		// Multi-dimensional array: recursive copy
		for i := 0; i < copyLen; i++ {
			if oldInner, ok := oldArr[i].([]interface{}); ok {
				if newInner, ok := newArr[i].([]interface{}); ok {
					v.preserveCopy(oldInner, newInner, dims[1:])
				}
			}
		}
	}
}

// visitIf handles if-else statements
func (v *ASPVisitor) visitIf(stmt *ast.IfStatement) error {
	if stmt == nil || stmt.Test == nil {
		return nil
	}

	condition, err := v.visitExpression(stmt.Test)
	if err != nil {
		return err
	}

	// Convert condition to boolean
	if isTruthy(condition) {
		// Execute consequent block
		if stmt.Consequent != nil {
			if err := v.VisitStatement(stmt.Consequent); err != nil {
				return err
			}
		}
	} else {
		// Execute alternate block
		if stmt.Alternate != nil {
			if err := v.VisitStatement(stmt.Alternate); err != nil {
				return err
			}
		}
	}

	return nil
}

// visitFor handles for loops
func (v *ASPVisitor) visitFor(stmt *ast.ForStatement) error {
	if stmt == nil || stmt.From == nil || stmt.To == nil {
		return nil
	}

	// Get variable name
	var varName string
	if stmt.Identifier != nil {
		varName = stmt.Identifier.Name
	}
	if varName == "" {
		return nil
	}

	// Evaluate From and To
	from, err := v.visitExpression(stmt.From)
	if err != nil {
		return err
	}

	to, err := v.visitExpression(stmt.To)
	if err != nil {
		return err
	}

	// Evaluate Step (default 1)
	step := 1.0
	if stmt.Step != nil {
		stepVal, err := v.visitExpression(stmt.Step)
		if err != nil {
			return err
		}
		step = toFloat(stepVal)
	}

	// Loop
	current := toFloat(from)
	end := toFloat(to)

	if step > 0 {
		for current <= end {
			_ = v.context.SetVariable(varName, current)

			// Execute body
			if stmt.Body != nil {
				for _, s := range stmt.Body {
					if err := v.VisitStatement(s); err != nil {
						return err
					}
				}
			}

			current += step
		}
	} else if step < 0 {
		for current >= end {
			_ = v.context.SetVariable(varName, current)

			// Execute body
			if stmt.Body != nil {
				for _, s := range stmt.Body {
					if err := v.VisitStatement(s); err != nil {
						return err
					}
				}
			}

			current += step
		}
	}

	return nil
}

// visitForEach handles for-each loops
func (v *ASPVisitor) visitForEach(stmt *ast.ForEachStatement) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	// Evaluate the collection expression
	collection, err := v.visitExpression(stmt.In)
	if err != nil {
		return err
	}

	// Handle different collection types
	switch col := collection.(type) {
	case []interface{}:
		// Iterate over array
		for _, item := range col {
			// Set loop variable
			_ = v.context.SetVariable(stmt.Identifier.Name, item)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "for" {
						break
					}
					return err
				}
			}
		}
	case map[string]interface{}:
		// Iterate over map (VBScript dictionary)
		for key := range col {
			// Set loop variable to key
			_ = v.context.SetVariable(stmt.Identifier.Name, key)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "for" {
						break
					}
					return err
				}
			}
		}
	case *Collection:
		// Iterate over Collection (Request.QueryString, Request.Form, etc.)
		keys := col.GetKeys()
		for _, key := range keys {
			// Set loop variable to key
			_ = v.context.SetVariable(stmt.Identifier.Name, key)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "for" {
						break
					}
					return err
				}
			}
		}
	default:
		// Try Enumeration interface
		if enumerable, ok := collection.(interface{ Enumeration() []interface{} }); ok {
			items := enumerable.Enumeration()
			for _, item := range items {
				// Set loop variable
				_ = v.context.SetVariable(stmt.Identifier.Name, item)

				// Execute loop body
				for _, body := range stmt.Body {
					if err := v.VisitStatement(body); err != nil {
						// Handle Exit For
						if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "for" {
							break
						}
						return err
					}
				}
			}
		}
	}

	return nil
}

// visitDo handles do-while loops
func (v *ASPVisitor) visitDo(stmt *ast.DoStatement) error {
	if stmt == nil {
		return nil
	}

	for {
		// Check pre-test condition if needed
		if stmt.TestType == ast.ConditionTestTypePreTest {
			condition, err := v.visitExpression(stmt.Condition)
			if err != nil {
				return err
			}

			// Handle loop type (While vs Until)
			shouldContinue := isTruthy(condition)
			if stmt.LoopType == ast.LoopTypeUntil {
				shouldContinue = !shouldContinue
			}

			if !shouldContinue {
				break
			}
		}

		// Execute loop body
		for _, body := range stmt.Body {
			if err := v.VisitStatement(body); err != nil {
				// Handle Exit Do
				if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "do" {
					return nil
				}
				return err
			}
		}

		// Check post-test condition if needed
		if stmt.TestType == ast.ConditionTestTypePostTest {
			condition, err := v.visitExpression(stmt.Condition)
			if err != nil {
				return err
			}

			// Handle loop type (While vs Until)
			shouldContinue := isTruthy(condition)
			if stmt.LoopType == ast.LoopTypeUntil {
				shouldContinue = !shouldContinue
			}

			if !shouldContinue {
				break
			}
		}
	}

	return nil
}

// visitWhile handles while loops
func (v *ASPVisitor) visitWhile(stmt *ast.WhileStatement) error {
	if stmt == nil {
		return nil
	}

	for {
		condition, err := v.visitExpression(stmt.Condition)
		if err != nil {
			return err
		}

		if !isTruthy(condition) {
			break
		}

		// Execute loop body
		for _, body := range stmt.Body {
			if err := v.VisitStatement(body); err != nil {
				// Handle Exit While
				if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "while" {
					return nil
				}
				return err
			}
		}
	}

	return nil
}

// visitSelect handles select-case statements
func (v *ASPVisitor) visitSelect(stmt *ast.SelectStatement) error {
	if stmt == nil {
		return nil
	}

	// Evaluate select expression
	selectValue, err := v.visitExpression(stmt.Condition)
	if err != nil {
		return err
	}

	// Check each case
	for _, caseStmt := range stmt.Cases {
		if caseStmt == nil {
			continue
		}

		// Check if case matches
		matched := false
		if len(caseStmt.Values) == 0 {
			// Case Else
			matched = true
		} else {
			for _, caseValue := range caseStmt.Values {
				// 1. Check for Range: __RANGE__(start, end)
				if call, ok := caseValue.(*ast.IndexOrCallExpression); ok {
					if ident, ok := call.Object.(*ast.Identifier); ok && ident.Name == "__RANGE__" {
						if len(call.Indexes) == 2 {
							startVal, err := v.visitExpression(call.Indexes[0])
							if err != nil {
								return err
							}
							endVal, err := v.visitExpression(call.Indexes[1])
							if err != nil {
								return err
							}

							// Check range: selectValue >= startVal AND selectValue <= endVal
							ge, _ := performBinaryOperation(ast.BinaryOperationGreaterOrEqual, selectValue, startVal)
							le, _ := performBinaryOperation(ast.BinaryOperationLessOrEqual, selectValue, endVal)

							if isTruthy(ge) && isTruthy(le) {
								matched = true
								break
							}
							continue // Handled as Range
						}
					}
				}

				// 2. Check for Comparison: BinaryExpression with Missing Left (from Case Is > 5)
				if bin, ok := caseValue.(*ast.BinaryExpression); ok {
					// Check if Left is MissingValueExpression
					if _, ok := bin.Left.(*ast.MissingValueExpression); ok {
						// Evaluate Right
						rightVal, err := v.visitExpression(bin.Right)
						if err != nil {
							return err
						}

						// Perform comparison with selectValue as Left
						res, err := performBinaryOperation(bin.Operation, selectValue, rightVal)
						if err != nil {
							return err
						}
						if isTruthy(res) {
							matched = true
							break
						}
						continue // Handled as Comparison
					}
				}

				val, err := v.visitExpression(caseValue)
				if err != nil {
					return err
				}

				// Compare case value with select value
				if compareEqual(selectValue, val) {
					matched = true
					break
				}
			}
		}

		// Execute case body if matched
		if matched {
			for _, body := range caseStmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit Select
					if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "select" {
						return nil
					}
					return err
				}
			}
			// Don't continue to next case (VBScript behavior)
			break
		}
	}

	return nil
}

// visitSubDeclaration handles sub declarations
func (v *ASPVisitor) visitSubDeclaration(stmt *ast.SubDeclaration) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	// Store sub in context for later calls
	_ = v.context.SetVariable(stmt.Identifier.Name, stmt)
	return nil
}

// visitFunctionDeclaration handles function declarations
func (v *ASPVisitor) visitFunctionDeclaration(stmt *ast.FunctionDeclaration) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	// Store function in context for later calls
	_ = v.context.SetVariable(stmt.Identifier.Name, stmt)
	return nil
}

// visitClassDeclaration handles class declarations
func (v *ASPVisitor) visitClassDeclaration(stmt *ast.ClassDeclaration) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	classDef := &ClassDef{
		Name:           stmt.Identifier.Name,
		Variables:      make(map[string]ClassMemberVar),
		Methods:        make(map[string]*ast.SubDeclaration),
		Functions:      make(map[string]*ast.FunctionDeclaration),
		Properties:     make(map[string][]PropertyDef),
		PrivateMethods: make(map[string]ast.Node),
	}

	for _, member := range stmt.Members {
		switch m := member.(type) {
		case *ast.FieldsDeclaration:
			// Private/Public variables
			vis := VisPublic
			if m.Modifier == ast.FieldAccessModifierPrivate {
				vis = VisPrivate
			}
			for _, field := range m.Fields {
				classDef.Variables[strings.ToLower(field.Identifier.Name)] = ClassMemberVar{
					Name:       field.Identifier.Name,
					Visibility: vis,
				}
			}
		case *ast.SubDeclaration:
			nameLower := strings.ToLower(m.Identifier.Name)
			if m.AccessModifier == ast.MethodAccessModifierPrivate {
				classDef.PrivateMethods[nameLower] = m
			} else {
				classDef.Methods[nameLower] = m
			}
		case *ast.FunctionDeclaration:
			nameLower := strings.ToLower(m.Identifier.Name)
			if m.AccessModifier == ast.MethodAccessModifierPrivate {
				classDef.PrivateMethods[nameLower] = m
			} else {
				classDef.Functions[nameLower] = m
			}
		case *ast.PropertyGetDeclaration:
			v.addPropertyDef(classDef, m.Identifier.Name, PropGet, m, m.AccessModifier)
		case *ast.PropertyLetDeclaration:
			v.addPropertyDef(classDef, m.Identifier.Name, PropLet, m, m.AccessModifier)
		case *ast.PropertySetDeclaration:
			v.addPropertyDef(classDef, m.Identifier.Name, PropSet, m, m.AccessModifier)
		}
	}

	// Store class definition in context
	return v.context.SetVariable(stmt.Identifier.Name, classDef)
}

func (v *ASPVisitor) addPropertyDef(classDef *ClassDef, name string, pType PropertyType, node ast.Node, access ast.MethodAccessModifier) {
	nameLower := strings.ToLower(name)
	vis := VisPublic
	if access == ast.MethodAccessModifierPrivate {
		vis = VisPrivate
	}

	def := PropertyDef{
		Name:       name,
		Type:       pType,
		Node:       node,
		Visibility: vis,
	}

	if _, ok := classDef.Properties[nameLower]; !ok {
		classDef.Properties[nameLower] = []PropertyDef{}
	}
	classDef.Properties[nameLower] = append(classDef.Properties[nameLower], def)
}

// visitVariableDeclaration handles variable declarations (Dim statement)
func (v *ASPVisitor) visitVariableDeclaration(stmt *ast.VariableDeclaration) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	varName := stmt.Identifier.Name

	// Check if it's an array declaration
	if len(stmt.ArrayDims) > 0 {
		// Fixed-size array: Dim arr(5) or Dim arr(2,3)
		dims := make([]int, len(stmt.ArrayDims))
		for i, dimExpr := range stmt.ArrayDims {
			dimVal, err := v.visitExpression(dimExpr)
			if err != nil {
				return err
			}
			dims[i] = toInt(dimVal)
		}

		// Create multi-dimensional array
		arr := v.makeNestedArray(dims)
		if err := v.context.DefineVariable(varName, arr); err != nil {
			return err
		}
	} else if stmt.IsDynamicArray {
		// Dynamic array: Dim arr() - initialize as empty array
		if err := v.context.DefineVariable(varName, []interface{}{}); err != nil {
			return err
		}
	} else {
		// Regular variable - initialize to nil (VBScript Empty)
		if err := v.context.DefineVariable(varName, nil); err != nil {
			return err
		}
	}

	return nil
}

// visitVariablesDeclaration handles Dim statements with multiple variables
func (v *ASPVisitor) visitVariablesDeclaration(stmt *ast.VariablesDeclaration) error {
	if stmt == nil || len(stmt.Variables) == 0 {
		return nil
	}

	for _, decl := range stmt.Variables {
		if decl == nil {
			continue
		}
		if err := v.visitVariableDeclaration(decl); err != nil {
			return err
		}
	}
	return nil
}

// makeNestedArray creates a nested array based on dimensions
// VBScript arrays are 0-indexed: Dim arr(5) creates array with indices 0-5 (6 elements)
func (v *ASPVisitor) makeNestedArray(dims []int) []interface{} {
	if len(dims) == 0 {
		return []interface{}{}
	}

	size := dims[0] + 1 // VBScript: Dim arr(5) means 0 to 5 (6 elements)
	arr := make([]interface{}, size)

	if len(dims) > 1 {
		// Recursive for multi-dimensional arrays
		innerDims := dims[1:]
		for i := 0; i < size; i++ {
			arr[i] = v.makeNestedArray(innerDims)
		}
	}

	return arr
}

// visitConstDeclaration handles constant declarations (Const statement)
func (v *ASPVisitor) visitConstDeclaration(stmt *ast.ConstsDeclaration) error {
	if stmt == nil {
		return nil
	}

	// Process all constant declarations
	for _, constDecl := range stmt.Declarations {
		if constDecl == nil || constDecl.Identifier == nil {
			continue
		}

		// Evaluate the constant expression
		val, err := v.visitExpression(constDecl.Init)
		if err != nil {
			return err
		}

		// Set the constant (this will check for conflicts)
		if err := v.context.SetConstant(constDecl.Identifier.Name, val); err != nil {
			return err
		}
	}

	return nil
}

// visitExpression evaluates an expression and returns its value
func (v *ASPVisitor) visitExpression(expr ast.Expression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	switch e := expr.(type) {
	case *ast.Identifier:
		varName := e.Name

		// Handle "Me" keyword
		if strings.EqualFold(varName, "me") {
			if obj := v.context.GetContextObject(); obj != nil {
				return obj, nil
			}
			return nil, fmt.Errorf("invalid use of 'Me' keyword outside of class")
		}

		if val, exists := v.context.GetVariable(varName); exists {
			if fn, ok := val.(*ast.FunctionDeclaration); ok {
				return v.executeFunction(fn, nil)
			}
			if sub, ok := val.(*ast.SubDeclaration); ok {
				return v.executeSub(sub, nil)
			}
			return val, nil
		}

		// Check built-in ASP objects (case-insensitive)
		switch strings.ToLower(varName) {
		case "response":
			return v.context.Response, nil
		case "request":
			return v.context.Request, nil
		case "server":
			return v.context.Server, nil
		case "session":
			return v.context.Session, nil
		case "application":
			return v.context.Application, nil
		}

		// Undefined variable returns nil in VBScript
		return nil, nil

	case *ast.StringLiteral:
		return e.Value, nil

	case *ast.IntegerLiteral:
		return int(e.Value), nil

	case *ast.FloatLiteral:
		return e.Value, nil

	case *ast.BooleanLiteral:
		return e.Value, nil

	case *ast.NullLiteral:
		// Null in VBScript represents a special value (no valid data)
		return nil, nil

	case *ast.EmptyLiteral:
		// Empty in VBScript represents an uninitialized variable
		return EmptyValue{}, nil

	case *ast.NothingLiteral:
		// Nothing in VBScript represents an empty object reference
		return NothingValue{}, nil

	case *ast.BinaryExpression:
		return v.visitBinaryExpression(e)

	case *ast.UnaryExpression:
		return v.visitUnaryExpression(e)

	case *ast.IndexOrCallExpression:
		return v.visitIndexOrCall(e)

	case *ast.MemberExpression:
		return v.visitMemberExpression(e)

	case *ast.NewExpression:
		return v.visitNewExpression(e)

	case *ast.WithMemberAccessExpression:
		return v.visitWithMemberAccess(e)

	default:
		return nil, nil
	}
}

// visitBinaryExpression evaluates binary operations
func (v *ASPVisitor) visitBinaryExpression(expr *ast.BinaryExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	left, err := v.visitExpression(expr.Left)
	if err != nil {
		return nil, err
	}

	// Handle short-circuit evaluation
	switch expr.Operation {
	case ast.BinaryOperationAnd:
		if !isTruthy(left) {
			return false, nil
		}
		right, err := v.visitExpression(expr.Right)
		if err != nil {
			return nil, err
		}
		return isTruthy(right), nil

	case ast.BinaryOperationOr:
		if isTruthy(left) {
			return true, nil
		}
		right, err := v.visitExpression(expr.Right)
		if err != nil {
			return nil, err
		}
		return isTruthy(right), nil
	}

	right, err := v.visitExpression(expr.Right)
	if err != nil {
		return nil, err
	}

	// Perform operation
	return performBinaryOperation(expr.Operation, left, right)
}

// visitUnaryExpression evaluates unary operations
func (v *ASPVisitor) visitUnaryExpression(expr *ast.UnaryExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	operand, err := v.visitExpression(expr.Argument)
	if err != nil {
		return nil, err
	}

	switch expr.Operation {
	case ast.UnaryOperationNot:
		// Check for boolean type to preserve it
		if b, ok := operand.(bool); ok {
			return !b, nil
		}
		// In VBScript, Not works as bitwise operator (invert all bits)
		operandInt := int(toFloat(operand))
		return ^operandInt, nil
	case ast.UnaryOperationMinus:
		return negateValue(operand), nil
	case ast.UnaryOperationPlus:
		return operand, nil
	default:
		return nil, fmt.Errorf("unknown unary operation: %v", expr.Operation)
	}
}

// visitCallSubStatement handles subroutine calls
func (v *ASPVisitor) visitCallSubStatement(stmt *ast.CallSubStatement) error {
	if stmt == nil {
		return nil
	}
	_, err := v.resolveCall(stmt.Callee, stmt.Arguments)
	return err
}

// visitIndexOrCall handles function calls and array indexing
func (v *ASPVisitor) visitIndexOrCall(expr *ast.IndexOrCallExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}
	return v.resolveCall(expr.Object, expr.Indexes)
}

// resolveCall evaluates a call or index expression
func (v *ASPVisitor) resolveCall(objectExpr ast.Expression, arguments []ast.Expression) (interface{}, error) {
	// 1. Check if objectExpr is an Identifier referring to a User Function/Sub (Manual Lookup)
	if ident, ok := objectExpr.(*ast.Identifier); ok {
		if val, exists := v.context.GetVariable(ident.Name); exists {
			if fn, ok := val.(*ast.FunctionDeclaration); ok {
				// Evaluate arguments
				args := make([]interface{}, 0)
				for _, arg := range arguments {
					val, err := v.visitExpression(arg)
					if err != nil {
						return nil, err
					}
					args = append(args, val)
				}
				return v.executeFunction(fn, args)
			}
			if sub, ok := val.(*ast.SubDeclaration); ok {
				// Evaluate arguments
				args := make([]interface{}, 0)
				for _, arg := range arguments {
					val, err := v.visitExpression(arg)
					if err != nil {
						return nil, err
					}
					args = append(args, val)
				}
				return v.executeSub(sub, args)
			}
		}
	}

	// Evaluate base expression
	base, err := v.visitExpression(objectExpr)
	if err != nil {
		return nil, err
	}

	// Evaluate indexes (arguments)
	args := make([]interface{}, 0)
	for _, arg := range arguments {
		val, err := v.visitExpression(arg)
		if err != nil {
			return nil, err
		}
		args = append(args, val)
	}

	// Check if this is a built-in function call
	if ident, ok := objectExpr.(*ast.Identifier); ok && base == nil {
		funcName := ident.Name
		if result, handled := evalBuiltInFunction(funcName, args, v.context); handled {
			return result, nil
		}
	}

	// Handle method calls on built-in objects (Default Property / Method)
	if obj, ok := base.(asp.ASPObject); ok {
		methodName := ""
		if ident, ok := objectExpr.(*ast.Identifier); ok {
			methodName = ident.Name
		}

		if methodName != "" {
			return obj.CallMethod(methodName, args...)
		}
	}

	// Handle Member Method Calls (obj.Method(args)) where Method is not a property
	if base == nil {
		if member, ok := objectExpr.(*ast.MemberExpression); ok {
			// Evaluate the object (e.g., Response in Response.Write)
			parentObj, err := v.visitExpression(member.Object)
			if err != nil {
				return nil, err
			}

			if parentObj != nil {
				propName := ""
				if member.Property != nil {
					propName = member.Property.Name
				}

				// Try ResponseObject methods
				if respObj, ok := parentObj.(*ResponseObject); ok {
					propNameLower := strings.ToLower(propName)
					switch propNameLower {
					case "write":
						for _, arg := range args {
							if err := respObj.Write(arg); err != nil {
								return nil, err
							}
						}
						return nil, nil
					case "binarywrite":
						if len(args) > 0 {
							if err := respObj.BinaryWrite(args[0]); err != nil {
								return nil, err
							}
						}
						return nil, nil
					case "addheader":
						if len(args) >= 2 {
							respObj.AddHeader(fmt.Sprintf("%v", args[0]), fmt.Sprintf("%v", args[1]))
						}
						return nil, nil
					case "appendtolog":
						if len(args) > 0 {
							respObj.AppendToLog(fmt.Sprintf("%v", args[0]))
						}
						return nil, nil
					case "clear":
						respObj.Clear()
						return nil, nil
					case "flush":
						if err := respObj.Flush(); err != nil {
							return nil, err
						}
						return nil, nil
					case "end":
						if err := respObj.End(); err != nil {
							return nil, err
						}
						// Signal that execution should stop
						return nil, fmt.Errorf("RESPONSE_END")
					case "redirect":
						if len(args) > 0 {
							if err := respObj.Redirect(fmt.Sprintf("%v", args[0])); err != nil {
								return nil, err
							}
							// Signal that execution should stop
							return nil, fmt.Errorf("RESPONSE_END")
						}
						return nil, nil
					}
				}

				// Try RequestObject methods
				if reqObj, ok := parentObj.(*RequestObject); ok {
					propNameLower := strings.ToLower(propName)
					switch propNameLower {
					case "binaryread":
						if len(args) > 0 {
							result, err := reqObj.CallMethod("binaryread", args...)
							if err != nil {
								return nil, err
							}
							return result, nil
						}
						return []byte{}, nil
					}
				}

				// Try ApplicationObject methods
				if appObj, ok := parentObj.(*ApplicationObject); ok {
					propNameLower := strings.ToLower(propName)
					switch propNameLower {
					case "lock":
						appObj.Lock()
						return nil, nil
					case "unlock":
						appObj.Unlock()
						return nil, nil
					}
				}

				// Try ASPObject
				if aspObj, ok := parentObj.(asp.ASPObject); ok {
					return aspObj.CallMethod(propName, args...)
				}

				// Try ASPLibrary / Generic interface with CallMethod
				if lib, ok := parentObj.(interface {
					CallMethod(string, ...interface{}) (interface{}, error)
				}); ok {
					return lib.CallMethod(propName, args...)
				}
			}
		}
	}

	// Handle array access
	if arr, ok := base.([]interface{}); ok && len(args) > 0 {
		idx := toInt(args[0])
		if idx >= 0 && idx < len(arr) {
			return arr[idx], nil
		}
		return nil, nil
	}

	// Handle Collection access (Request.QueryString("key"), Request.Form("key"), etc.)
	if collection, ok := base.(*Collection); ok && len(args) > 0 {
		key := fmt.Sprintf("%v", args[0])
		return collection.Get(key), nil
	}

	// Handle SessionObject index access (Session("key"))
	if sessionObj, ok := base.(*SessionObject); ok && len(args) > 0 {
		return sessionObj.GetIndex(args[0]), nil
	}

	// Handle ApplicationObject index access (Application("key"))
	if appObj, ok := base.(*ApplicationObject); ok && len(args) > 0 {
		key := fmt.Sprintf("%v", args[0])
		return appObj.Get(key), nil
	}

	// Handle map/dictionary access (for JSON objects, etc.)
	if mapObj, ok := base.(map[string]interface{}); ok && len(args) > 0 {
		key := fmt.Sprintf("%v", args[0])
		if val, exists := mapObj[key]; exists {
			return val, nil
		}
		return nil, nil
	}

	return nil, nil
}

// visitMemberExpression evaluates member access (obj.property)
func (v *ASPVisitor) visitMemberExpression(expr *ast.MemberExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	// Evaluate object
	obj, err := v.visitExpression(expr.Object)
	if err != nil {
		return nil, err
	}

	// Get property name
	propName := ""
	if expr.Property != nil {
		propName = expr.Property.Name
	}

	// Handle ASP built-in objects
	switch strings.ToLower(propName) {
	case "response":
		return v.context.Response, nil
	case "request":
		return v.context.Request, nil
	case "server":
		return v.context.Server, nil
	case "session":
		return v.context.Session, nil
	case "application":
		return v.context.Application, nil
	}

	// Handle generic property access
	if aspObj, ok := obj.(asp.ASPObject); ok {
		return aspObj.GetProperty(propName), nil
	}

	// Handle ResponseObject properties
	if respObj, ok := obj.(*ResponseObject); ok {
		propNameLower := strings.ToLower(propName)
		switch propNameLower {
		case "buffer":
			return respObj.GetBuffer(), nil
		case "cachecontrol":
			return respObj.GetCacheControl(), nil
		case "charset":
			return respObj.GetCharset(), nil
		case "contenttype":
			return respObj.GetContentType(), nil
		case "expires":
			return respObj.GetExpires(), nil
		case "expiresabsolute":
			return respObj.GetExpiresAbsolute(), nil
		case "isclientconnected":
			return respObj.IsClientConnected(), nil
		case "pics":
			return respObj.GetPICS(), nil
		case "status":
			return respObj.GetStatus(), nil
		case "cookies":
			return respObj.Cookies(), nil
		}
	}

	// Handle SessionObject
	if sessionObj, ok := obj.(*SessionObject); ok {
		return sessionObj.GetProperty(propName), nil
	}

	// Handle ApplicationObject properties and methods
	if appObj, ok := obj.(*ApplicationObject); ok {
		propNameLower := strings.ToLower(propName)
		switch propNameLower {
		case "lock":
			// Return a callable that will execute Lock when invoked
			return func() {
				appObj.Lock()
			}, nil
		case "unlock":
			// Return a callable that will execute Unlock when invoked
			return func() {
				appObj.Unlock()
			}, nil
		case "staticobjects":
			// Return the StaticObjects collection (as a map for enumeration)
			return appObj.GetStaticObjects(), nil
		case "contents":
			// Return the Contents collection (as a map for enumeration)
			return appObj.GetContents(), nil
		}
	}

	// Handle Collection properties (Count, etc.)
	if collection, ok := obj.(*Collection); ok {
		propNameLower := strings.ToLower(propName)
		switch propNameLower {
		case "count":
			return collection.Count(), nil
		}
	}

	// Handle generic GetProperty interface
	if getter, ok := obj.(interface{ GetProperty(string) interface{} }); ok {
		return getter.GetProperty(propName), nil
	}

	return nil, nil
}

// Helper functions

// isTruthy checks if a value is truthy in VBScript
func isTruthy(val interface{}) bool {
	if val == nil {
		return false
	}
	if b, ok := val.(bool); ok {
		return b
	}
	if i, ok := val.(int); ok {
		return i != 0
	}
	if i, ok := val.(float64); ok {
		return i != 0
	}
	if s, ok := val.(string); ok {
		return s != ""
	}
	return true
}

// isNothingValue checks if a value is Nothing or nil
func isNothingValue(val interface{}) bool {
	if val == nil {
		return true
	}
	if _, ok := val.(NothingValue); ok {
		return true
	}
	return false
}

// toString converts a value to string
func toString(val interface{}) string {
	if val == nil {
		return ""
	}
	switch v := val.(type) {
	case EmptyValue:
		return ""
	case NothingValue:
		return ""
	case string:
		return v
	case bool:
		if v {
			return "True"
		}
		return "False"
	case int:
		return fmt.Sprintf("%d", v)
	case float64:
		if v == float64(int(v)) {
			return fmt.Sprintf("%d", int(v))
		}
		return fmt.Sprintf("%g", v)
	default:
		return fmt.Sprintf("%v", v)
	}
}

// toInt converts a value to integer
func toInt(val interface{}) int {
	if val == nil {
		return 0
	}
	switch v := val.(type) {
	case EmptyValue, NothingValue:
		return 0
	case int:
		return v
	case float64:
		return int(v)
	case string:
		// Try to parse hex, octal, or decimal
		if parsed, ok := tryParseNumericLiteral(v); ok {
			if intVal, ok := parsed.(int); ok {
				return intVal
			}
			if floatVal, ok := parsed.(float64); ok {
				return int(floatVal)
			}
		}
		return 0
	case bool:
		if v {
			return -1
		}
		return 0
	default:
		return 0
	}
}

// toFloat converts a value to float64
func toFloat(val interface{}) float64 {
	if val == nil {
		return 0
	}
	switch v := val.(type) {
	case EmptyValue, NothingValue:
		return 0
	case int:
		return float64(v)
	case float64:
		return v
	case string:
		// Try to parse hex, octal, or decimal
		if parsed, ok := tryParseNumericLiteral(v); ok {
			if intVal, ok := parsed.(int); ok {
				return float64(intVal)
			}
			if floatVal, ok := parsed.(float64); ok {
				return floatVal
			}
		}
		return 0
	case bool:
		if v {
			return -1
		}
		return 0
	default:
		return 0
	}
}

// negateValue negates a value
func negateValue(val interface{}) interface{} {
	if val == nil {
		return 0
	}
	switch v := val.(type) {
	case int:
		return -v
	case float64:
		return -v
	default:
		return 0
	}
}

// performBinaryOperation performs a binary operation
func performBinaryOperation(op ast.BinaryOperation, left, right interface{}) (interface{}, error) {
	switch op {
	case ast.BinaryOperationAnd:
		// Check for boolean preservation
		if b1, ok1 := left.(bool); ok1 {
			if b2, ok2 := right.(bool); ok2 {
				return b1 && b2, nil
			}
		}
		// In VBScript, And works as bitwise operator
		// When used in conditional context, VBScript evaluates truthiness first
		leftInt := int(toFloat(left))
		rightInt := int(toFloat(right))
		return leftInt & rightInt, nil
	case ast.BinaryOperationOr:
		// Check for boolean preservation
		if b1, ok1 := left.(bool); ok1 {
			if b2, ok2 := right.(bool); ok2 {
				return b1 || b2, nil
			}
		}
		// In VBScript, Or works as bitwise operator
		leftInt := int(toFloat(left))
		rightInt := int(toFloat(right))
		return leftInt | rightInt, nil
	case ast.BinaryOperationAddition:
		leftNum := toFloat(left)
		rightNum := toFloat(right)
		return leftNum + rightNum, nil
	case ast.BinaryOperationSubtraction:
		leftNum := toFloat(left)
		rightNum := toFloat(right)
		return leftNum - rightNum, nil
	case ast.BinaryOperationMultiplication:
		leftNum := toFloat(left)
		rightNum := toFloat(right)
		return leftNum * rightNum, nil
	case ast.BinaryOperationDivision:
		leftNum := toFloat(left)
		rightNum := toFloat(right)
		if rightNum == 0 {
			return nil, fmt.Errorf("division by zero")
		}
		return leftNum / rightNum, nil
	case ast.BinaryOperationIntDivision:
		leftNum := int(toFloat(left))
		rightNum := int(toFloat(right))
		if rightNum == 0 {
			return nil, fmt.Errorf("division by zero")
		}
		return leftNum / rightNum, nil
	case ast.BinaryOperationMod:
		leftNum := int(toFloat(left))
		rightNum := int(toFloat(right))
		if rightNum == 0 {
			return nil, fmt.Errorf("division by zero")
		}
		return leftNum % rightNum, nil
	case ast.BinaryOperationExponentiation:
		return math.Pow(toFloat(left), toFloat(right)), nil
	case ast.BinaryOperationEqual:
		return compareEqual(left, right), nil
	case ast.BinaryOperationNotEqual:
		return !compareEqual(left, right), nil
	case ast.BinaryOperationLess:
		return compareLess(left, right), nil
	case ast.BinaryOperationGreater:
		return compareLess(right, left), nil
	case ast.BinaryOperationLessOrEqual:
		return !compareLess(right, left), nil
	case ast.BinaryOperationGreaterOrEqual:
		return !compareLess(left, right), nil
	case ast.BinaryOperationConcatenation:
		return toString(left) + toString(right), nil
	case ast.BinaryOperationIs:
		// Is operator for object comparison (checks if same reference)
		// Special handling for Nothing comparisons
		leftIsNothing := isNothingValue(left)
		rightIsNothing := isNothingValue(right)

		if leftIsNothing && rightIsNothing {
			return true, nil
		}
		if leftIsNothing || rightIsNothing {
			return false, nil
		}

		// For other objects, compare references
		return left == right, nil
	case ast.BinaryOperationXor, ast.BinaryOperationEqv, ast.BinaryOperationImp:
		// TODO: implement bitwise operations
		return nil, fmt.Errorf("binary operation %d not yet implemented", op)
	default:
		return nil, fmt.Errorf("unknown binary operator: %d", op)
	}
}

// compareEqual compares two values for equality
func compareEqual(left, right interface{}) bool {
	if left == nil && right == nil {
		return true
	}
	if left == nil || right == nil {
		return false
	}

	// Compare as strings first
	leftStr := toString(left)
	rightStr := toString(right)
	if leftStr == rightStr {
		return true
	}

	// Try numeric comparison
	if ln, lok := toNumeric(left); lok {
		if rn, rok := toNumeric(right); rok {
			return ln == rn
		}
	}

	return false
}

// compareLess compares if left is less than right
func compareLess(left, right interface{}) bool {
	leftNum, lok := toNumeric(left)
	rightNum, rok := toNumeric(right)

	if lok && rok {
		return leftNum < rightNum
	}

	// String comparison
	return toString(left) < toString(right)
}

// toNumeric attempts to convert a value to numeric type
func toNumeric(val interface{}) (float64, bool) {
	if val == nil {
		return 0, true
	}
	switch v := val.(type) {
	case int:
		return float64(v), true
	case float64:
		return v, true
	case string:
		if f, err := strconv.ParseFloat(v, 64); err == nil {
			return f, true
		}
		return 0, false
	case bool:
		if v {
			return -1, true
		}
		return 0, true
	default:
		return 0, false
	}
}

// isBoolType checks if a value is boolean type
func isBoolType(val interface{}) bool {
	_, ok := val.(bool)
	return ok
}

// populateRequestData fills a RequestObject with data from HTTP request
func populateRequestData(req *RequestObject, r *http.Request) {
	// Set the HTTP request for BinaryRead support
	req.SetHTTPRequest(r)

	// Parse form data
	r.ParseForm()

	// Set query string parameters (from URL only)
	for key, values := range r.URL.Query() {
		if len(values) > 0 {
			req.QueryString.Add(key, values[0])
		}
	}

	// Set form parameters (POST data only)
	// We need to check both PostForm and Form to handle different content types
	if r.Method == "POST" || r.Method == "PUT" {
		for key, values := range r.Form {
			// Only add if not in QueryString (to avoid duplicates)
			if !req.QueryString.Exists(key) && len(values) > 0 {
				req.Form.Add(key, values[0])
			}
		}
	}

	// Set cookies
	for _, cookie := range r.Cookies() {
		req.Cookies.Add(cookie.Name, cookie.Value)
	}

	// Set server variables
	req.ServerVariables.Add("REQUEST_METHOD", r.Method)
	req.ServerVariables.Add("REQUEST_PATH", r.URL.Path)
	req.ServerVariables.Add("QUERY_STRING", r.URL.RawQuery)
	req.ServerVariables.Add("REMOTE_ADDR", r.RemoteAddr)
	req.ServerVariables.Add("SERVER_NAME", r.Host)
	req.ServerVariables.Add("SERVER_PORT", r.URL.Port())
	req.ServerVariables.Add("HTTP_USER_AGENT", r.UserAgent())
	req.ServerVariables.Add("HTTP_REFERER", r.Referer())
	req.ServerVariables.Add("CONTENT_TYPE", r.Header.Get("Content-Type"))
	req.ServerVariables.Add("CONTENT_LENGTH", r.Header.Get("Content-Length"))

	// Add all HTTP headers as HTTP_* server variables
	for name, values := range r.Header {
		if len(values) > 0 {
			varName := "HTTP_" + strings.ToUpper(strings.ReplaceAll(name, "-", "_"))
			req.ServerVariables.Add(varName, values[0])
		}
	}

	// ClientCertificate collection (stub - SSL certificates not currently handled)
	// In a full implementation, this would parse TLS client certificate info
	req.ClientCertificate.Add("Subject", "")
	req.ClientCertificate.Add("Issuer", "")
	req.ClientCertificate.Add("ValidFrom", "")
	req.ClientCertificate.Add("ValidUntil", "")
	req.ClientCertificate.Add("SerialNumber", "")
}

// visitWithStatement handles With statements
func (v *ASPVisitor) visitWithStatement(stmt *ast.WithStatement) error {
	if stmt == nil || stmt.Expression == nil {
		return nil
	}

	// Evaluate the expression to get the object
	obj, err := v.visitExpression(stmt.Expression)
	if err != nil {
		return err
	}

	// Push object to stack
	v.withStack = append(v.withStack, obj)
	defer func() {
		// Pop object from stack
		if len(v.withStack) > 0 {
			v.withStack = v.withStack[:len(v.withStack)-1]
		}
	}()

	// Execute body
	for _, s := range stmt.Body {
		if err := v.VisitStatement(s); err != nil {
			return err
		}
	}

	return nil
}

// visitWithMemberAccess handles .property access inside With block
func (v *ASPVisitor) visitWithMemberAccess(expr *ast.WithMemberAccessExpression) (interface{}, error) {
	if len(v.withStack) == 0 {
		return nil, fmt.Errorf("invalid reference to .%s outside of With block", expr.Property.Name)
	}

	// Get current With object
	obj := v.withStack[len(v.withStack)-1]
	propName := expr.Property.Name

	// Handle ASP built-in objects
	switch strings.ToLower(propName) {
	case "response":
		return v.context.Response, nil
	case "request":
		return v.context.Request, nil
	case "server":
		return v.context.Server, nil
	case "session":
		return v.context.Session, nil
	case "application":
		return v.context.Application, nil
	}

	// Handle generic property access
	if aspObj, ok := obj.(asp.ASPObject); ok {
		return aspObj.GetProperty(propName), nil
	}

	// Handle SessionObject
	if sessionObj, ok := obj.(*SessionObject); ok {
		return sessionObj.GetProperty(propName), nil
	}

	// Handle ClassInstance
	if classInst, ok := obj.(*ClassInstance); ok {
		return classInst.GetProperty(propName)
	}

	return nil, nil
}

// visitNewExpression handles New ClassName
func (v *ASPVisitor) visitNewExpression(expr *ast.NewExpression) (interface{}, error) {
	if expr == nil || expr.Argument == nil {
		return nil, nil
	}

	// Arg should be Identifier
	if ident, ok := expr.Argument.(*ast.Identifier); ok {
		className := ident.Name

		// Lookup ClassDef
		classDefVal, exists := v.context.GetVariable(className)
		if !exists {
			// Maybe it's a built-in COM object (unlikely syntax 'New X', usually 'Server.CreateObject')
			if strings.ToLower(className) == "regexp" {
				// TODO: Implement RegExp
				return nil, fmt.Errorf("RegExp not implemented yet")
			}
			return nil, fmt.Errorf("class not defined: %s", className)
		}

		if classDef, ok := classDefVal.(*ClassDef); ok {
			return NewClassInstance(classDef, v.context)
		}

		return nil, fmt.Errorf("%s is not a class", className)
	}

	return nil, fmt.Errorf("invalid New expression")
}

// executeFunction executes a user defined function
func (v *ASPVisitor) executeFunction(fn *ast.FunctionDeclaration, args []interface{}) (interface{}, error) {
	v.context.PushScope()
	defer v.context.PopScope()

	// Bind parameters
	for i, param := range fn.Parameters {
		var val interface{}
		if i < len(args) {
			val = args[i]
		}
		_ = v.context.DefineVariable(param.Identifier.Name, val)
	}

	// Define return variable
	funcName := fn.Identifier.Name
	_ = v.context.DefineVariable(funcName, nil)

	// Execute body
	if fn.Body != nil {
		if list, ok := fn.Body.(*ast.StatementList); ok {
			for _, stmt := range list.Statements {
				if err := v.VisitStatement(stmt); err != nil {
					return nil, err
				}
			}
		} else if stmt, ok := fn.Body.(ast.Statement); ok {
			if err := v.VisitStatement(stmt); err != nil {
				return nil, err
			}
		}
	}

	// Get return value
	val, _ := v.context.GetVariable(funcName)
	return val, nil
}

// executeSub executes a user defined subroutine
func (v *ASPVisitor) executeSub(sub *ast.SubDeclaration, args []interface{}) (interface{}, error) {
	v.context.PushScope()
	defer v.context.PopScope()

	// Bind parameters
	for i, param := range sub.Parameters {
		var val interface{}
		if i < len(args) {
			val = args[i]
		}
		_ = v.context.DefineVariable(param.Identifier.Name, val)
	}

	// Execute body
	if sub.Body != nil {
		if list, ok := sub.Body.(*ast.StatementList); ok {
			for _, stmt := range list.Statements {
				if err := v.VisitStatement(stmt); err != nil {
					return nil, err
				}
			}
		} else if stmt, ok := sub.Body.(ast.Statement); ok {
			if err := v.VisitStatement(stmt); err != nil {
				return nil, err
			}
		}
	}

	return nil, nil
}

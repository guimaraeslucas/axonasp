/*
 * AxonASP Server - Version 1.0
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package server

import (
	"fmt"
	"go-asp/asp"
	"math"
	"math/rand"
	"net"
	"net/http"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"time"
	"go-asp/vbscript/ast"
)

// LoopExitError represents a loop exit statement (Exit For, Exit Do, etc)
type LoopExitError struct {
	LoopType string // "for", "do", "while", "select"
}

func (e *LoopExitError) Error() string {
	return fmt.Sprintf("Exit %s", e.LoopType)
}

// ProcedureExitError represents an Exit Sub/Function/Property control flow signal
type ProcedureExitError struct {
	Kind string // "sub", "function", "property"
}

func (e *ProcedureExitError) Error() string {
	return fmt.Sprintf("Exit %s", e.Kind)
}

// ErrServerTransfer indicates a Server.Transfer call, stopping current execution
var ErrServerTransfer = fmt.Errorf("Server.Transfer")

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
	Server      *ServerObject
	Session     *SessionObject
	Application *ApplicationObject
	Err         *ErrObject

	// Variable storage (case-insensitive keys)
	variables map[string]interface{}

	// Constant storage (case-insensitive keys) - read-only values
	constants map[string]interface{}

	// HTTP context
	httpWriter  http.ResponseWriter
	httpRequest *http.Request

	// Execution state
	startTime  time.Time
	timeout    time.Duration
	rng        *rand.Rand
	lastRnd    float64
	hasLastRnd bool

	// Library instances
	libraries map[string]interface{}

	// Configuration
	RootDir     string
	CurrentFile string
	CurrentDir  string
	compareMode ast.OptionCompareMode
	optionBase  int

	// Session management
	sessionID      string
	sessionManager *SessionManager

	// Scoping
	scopeStack      []map[string]interface{}
	contextObject   interface{}  // For Class Instance (Me)
	currentExecutor *ASPExecutor // Current executor for ExecuteGlobal/Execute

	// Mutex for thread safety
	mu sync.RWMutex

	// Session tracking
	isNewSession bool
}

// NewExecutionContext creates a new execution context
func NewExecutionContext(w http.ResponseWriter, r *http.Request, sessionID string, timeout time.Duration) *ExecutionContext {
	sessionManager := GetSessionManager()

	// Load or create session data
	sessionData, isNew, err := sessionManager.GetOrCreateSession(sessionID)
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
		isNew = true
	}

	ctx := &ExecutionContext{
		Request:        NewRequestObject(),
		Response:       NewResponseObject(w, r),
		Server:         nil, // Will be initialized below
		Session:        NewSessionObject(sessionID, sessionData.Data),
		Application:    GetGlobalApplication(),
		Err:            NewErrObject(),
		variables:      make(map[string]interface{}),
		constants:      make(map[string]interface{}),
		libraries:      make(map[string]interface{}),
		scopeStack:     make([]map[string]interface{}, 0),
		httpWriter:     w,
		httpRequest:    r,
		startTime:      time.Now(),
		timeout:        timeout,
		rng:            rand.New(rand.NewSource(1)),
		sessionID:      sessionID,
		sessionManager: sessionManager,
		isNewSession:   isNew,
		compareMode:    ast.OptionCompareBinary,
		optionBase:     0,
	}

	// Initialize Server object with context reference
	ctx.Server = NewServerObjectWithContext(ctx)
	ctx.Server.SetHttpRequest(r)

	// Built-in VBScript constants (case-insensitive)
	ctx.constants["vbcrlf"] = "\r\n"
	ctx.constants["vbcr"] = "\r"
	ctx.constants["vblf"] = "\n"
	ctx.constants["vbnewline"] = "\r\n"
	ctx.constants["vbtab"] = "\t"
	ctx.constants["vbnullchar"] = string(rune(0))
	ctx.constants["vbback"] = string(rune(8))
	ctx.constants["vbformfeed"] = string(rune(12))
	ctx.constants["vbverticaltab"] = string(rune(11))
	ctx.constants["vbnullstring"] = ""

	// We'll set the executor reference after creating it (circular dependency)
	// This is done in ASPExecutor.Execute()

	// Add Document object to variables (for Document.Write access)
	ctx.variables["document"] = NewDocumentObject(ctx)
	// Expose Err intrinsic object
	ctx.variables["err"] = ctx.Err

	return ctx
}

// SetOptionCompare updates string comparison mode for this context
func (ec *ExecutionContext) SetOptionCompare(mode ast.OptionCompareMode) {
	switch mode {
	case ast.OptionCompareText:
		ec.compareMode = ast.OptionCompareText
	default:
		ec.compareMode = ast.OptionCompareBinary
	}
}

// OptionCompareMode returns the active string comparison mode
func (ec *ExecutionContext) OptionCompareMode() ast.OptionCompareMode {
	return ec.compareMode
}

// SetOptionBase updates the default array lower bound for this context.
func (ec *ExecutionContext) SetOptionBase(base int) {
	if base == 1 {
		ec.optionBase = 1
		return
	}
	ec.optionBase = 0
}

// OptionBase returns the configured default array lower bound.
func (ec *ExecutionContext) OptionBase() int {
	return ec.optionBase
}

func (ec *ExecutionContext) ensureRandomSource() {
	ec.mu.Lock()
	defer ec.mu.Unlock()

	if ec.rng == nil {
		ec.rng = rand.New(rand.NewSource(1))
	}
}

func (ec *ExecutionContext) randomizeWithSeed(seed int64) {
	ec.ensureRandomSource()

	ec.mu.Lock()
	defer ec.mu.Unlock()
	ec.rng.Seed(seed)
	ec.hasLastRnd = false
}

func (ec *ExecutionContext) nextRandomValue(arg interface{}, hasArg bool) float64 {
	ec.ensureRandomSource()

	ec.mu.Lock()
	defer ec.mu.Unlock()

	if hasArg {
		seedVal := toFloat(arg)
		if seedVal < 0 {
			ec.rng.Seed(seedFromNumber(seedVal))
			ec.hasLastRnd = false
			ec.lastRnd = ec.rng.Float64()
			ec.hasLastRnd = true
			return ec.lastRnd
		}
		if seedVal == 0 {
			if ec.hasLastRnd {
				return ec.lastRnd
			}
			ec.lastRnd = ec.rng.Float64()
			ec.hasLastRnd = true
			return ec.lastRnd
		}
	}

	ec.lastRnd = ec.rng.Float64()
	ec.hasLastRnd = true
	return ec.lastRnd
}

func seedFromNumber(val float64) int64 {
	seed := int64(val * 1000000)
	if seed == 0 {
		seed = int64(val)
	}
	if seed == 0 {
		if val >= 0 {
			return 1
		}
		return -1
	}
	return seed
}

// lowerKey returns a lowercase key without allocating when the input is already lower ASCII
func lowerKey(name string) string {
	for i := 0; i < len(name); i++ {
		c := name[i]
		if c >= 'A' && c <= 'Z' {
			b := []byte(name)
			for j := i; j < len(b); j++ {
				if b[j] >= 'A' && b[j] <= 'Z' {
					b[j] += 'a' - 'A'
				}
			}
			return string(b)
		}
	}
	return name
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
	nameLower := lowerKey(name)

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

	// 3. Check Context Object (Class Member)
	if ec.contextObject != nil {
		// Try internal access first (allows Private members)
		if internal, ok := ec.contextObject.(interface {
			SetMember(string, interface{}) (bool, error)
		}); ok {
			if handled, err := internal.SetMember(nameLower, value); handled {
				return err
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
	nameLower := lowerKey(name)
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
	nameLower := lowerKey(name)

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
			GetMember(string) (interface{}, bool, error)
		}); ok {
			if val, found, _ := internal.GetMember(nameLower); found {
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

// GetVariableFromParentScope gets a variable from the parent scope (for ByRef parameters)
func (ec *ExecutionContext) GetVariableFromParentScope(name string) (interface{}, bool) {
	nameLower := lowerKey(name)

	ec.mu.RLock()
	defer ec.mu.RUnlock()

	// Get from parent scope (one level up from current)
	// We need at least 2 scopes: current and parent
	if len(ec.scopeStack) >= 2 {
		parentIndex := len(ec.scopeStack) - 2
		if val, exists := ec.scopeStack[parentIndex][nameLower]; exists {
			return val, true
		}
	}

	// Check globals as fallback (variable might be global)
	if val, exists := ec.variables[nameLower]; exists {
		return val, true
	}

	return nil, false
}

// SetVariableInParentScope sets a variable in the parent scope (for ByRef parameters)
func (ec *ExecutionContext) SetVariableInParentScope(name string, value interface{}) error {
	nameLower := lowerKey(name)

	ec.mu.Lock()
	defer ec.mu.Unlock()

	// Set in parent scope (one level up from current)
	if len(ec.scopeStack) >= 2 {
		parentIndex := len(ec.scopeStack) - 2
		ec.scopeStack[parentIndex][nameLower] = value
		return nil
	}

	// If no parent scope (at global level), set globally
	if len(ec.scopeStack) == 0 {
		ec.variables[nameLower] = value
		return nil
	}

	// If we have only one scope, treat parent as global
	ec.variables[nameLower] = value
	return nil
}

// SetConstant sets a constant in the execution context (case-insensitive)
// Constants cannot be changed after initialization
func (ec *ExecutionContext) SetConstant(name string, value interface{}) error {
	nameLower := lowerKey(name)

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

	cleanPath := strings.ReplaceAll(path, "\\", "/")

	// Absolute virtual path (starts with /) -> root-based
	if strings.HasPrefix(cleanPath, "/") {
		cleanPath = strings.TrimPrefix(cleanPath, "/")
		fullPath := filepath.Join(rootDir, cleanPath)
		return fullPath
	}

	// Relative path -> prefer current script directory when available
	baseDir := rootDir
	if ec.CurrentDir != "" {
		baseDir = ec.CurrentDir
	}
	fullPath := filepath.Join(baseDir, cleanPath)
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

// SetContext sets the execution context for the executor
// Useful for manual context manipulation when not using Execute()
func (ae *ASPExecutor) SetContext(ctx *ExecutionContext) {
	ae.context = ctx
}

// Execute processes ASP code and returns rendered output
func (ae *ASPExecutor) Execute(fileContent string, filePath string, w http.ResponseWriter, r *http.Request, sessionID string) error {
	return ae.executeInternal(fileContent, filePath, w, r, sessionID, nil, false)
}

// ExecuteWithParsed runs ASP using pre-resolved content and a cached parse result
func (ae *ASPExecutor) ExecuteWithParsed(resolvedContent string, parsedResult *asp.ASPParserResult, filePath string, w http.ResponseWriter, r *http.Request, sessionID string) error {
	return ae.executeInternal(resolvedContent, filePath, w, r, sessionID, parsedResult, true)
}

func (ae *ASPExecutor) executeInternal(fileContent string, filePath string, w http.ResponseWriter, r *http.Request, sessionID string, parsedResult *asp.ASPParserResult, alreadyResolved bool) error {
	// Create execution context
	timeout := time.Duration(ae.config.ScriptTimeout) * time.Second
	ae.context = NewExecutionContext(w, r, sessionID, timeout)
	ae.context.CurrentFile = filePath
	ae.context.CurrentDir = filepath.Dir(filePath)

	// Pre-process Includes if not already resolved
	if !alreadyResolved {
		resolvedContent, err := asp.ResolveIncludes(fileContent, filePath, ae.config.RootDir, nil)
		if err != nil {
			return fmt.Errorf("include error: %w", err)
		}
		fileContent = resolvedContent
	}

	// Fast path: if there's no executable ASP (only directives/comments), stream HTML directly
	if parsedResult == nil && !hasExecutableASP(fileContent) {
		cleanHTML := stripDirectivesAndComments(fileContent)
		if err := ae.context.Response.Write(cleanHTML); err != nil {
			return fmt.Errorf("failed to write static HTML: %w", err)
		}
		if err := ae.context.Response.Flush(); err != nil {
			return fmt.Errorf("failed to flush response: %w", err)
		}

		if err := ae.saveSession(); err != nil {
			fmt.Printf("Warning: Failed to save session: %v\n", err)
		}
		return nil
	}

	// Set RootDir in context
	ae.context.RootDir = ae.config.RootDir

	// Always set session cookie to ensure it persists and slides if needed
	// Standard ASP session cookie is a browser-session cookie (no MaxAge)
	// The server handles the timeout (20 mins default)
	http.SetCookie(w, &http.Cookie{
		Name:     "ASPSESSIONID",
		Value:    ae.context.sessionID,
		Path:     "/",
		HttpOnly: true,
		Secure:   false, // Set to true if using HTTPS
		SameSite: http.SameSiteLaxMode,
	})

	// Configure Server object with root directory, script timeout, and executor
	ae.context.Server.SetRootDir(ae.config.RootDir)
	ae.context.Server.SetScriptTimeout(ae.config.ScriptTimeout)
	ae.context.Server.SetExecutor(ae) // Set executor for CreateObject calls

	// Set current script directory for relative path mapping
	// This is critical for Server.MapPath("relative/file.asp")
	_ = ae.context.Server.SetProperty("_scriptDir", ae.context.CurrentDir)

	// Populate Request object
	populateRequestData(ae.context.Request, r, ae.context)

	// Call Session_OnStart if this is a new session
	if ae.context.isNewSession {
		globalASAManager := GetGlobalASAManager()
		if globalASAManager.HasSessionOnStart() {
			if err := globalASAManager.ExecuteSessionOnStart(ae, ae.context); err != nil {
				if ae.config.DebugASP {
					fmt.Printf("Warning: Error in Session_OnStart: %v\n", err)
				}
				// Don't return error, just log it - Session_OnStart errors shouldn't break the request
			}
		}
	}

	result := parsedResult
	if result == nil {
		parsingOptions := &asp.ASPParsingOptions{
			SaveComments:      false,
			StrictMode:        false,
			AllowImplicitVars: true,
			DebugMode:         ae.config.DebugASP,
		}
		parser := asp.NewASPParserWithOptions(fileContent, parsingOptions)
		parsed, err := parser.Parse()
		if err != nil {
			return fmt.Errorf("failed to parse ASP code: %w", err)
		}
		result = parsed
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
	execFunc := func() error {
		// Always execute per-block to avoid re-emitting large HTML through VB conversion
		return ae.executeBlocks(result)
	}

	go func() {
		defer func() {
			if rec := recover(); rec != nil {
				done <- fmt.Errorf("runtime panic: %v", rec)
			}
		}()

		err := execFunc()
		done <- err
	}()

	// Wait for execution or timeout
	select {
	case err := <-done:
		if err != nil {
			if err.Error() == "RESPONSE_END" || err == ErrServerTransfer {
				return nil
			}
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

// ExecuteASPPath executes an ASP file from a path within the current execution environment
// Used by Server.Execute and Server.Transfer
// Inherits Request, Response, Session, Application from current context
// But has its own variable scope (Variables, Constants)
func (ae *ASPExecutor) ExecuteASPPath(path string) error {
	// Map virtual path to physical path using the current context
	physicalPath := ae.context.Server_MapPath(path)

	// Read file content with proper UTF-8 decoding (handles BOM/UTF-16)
	content, err := asp.ReadFileText(physicalPath)
	if err != nil {
		return err
	}

	// Pre-process Includes
	resolvedContent, err := asp.ResolveIncludes(content, physicalPath, ae.config.RootDir, nil)
	if err != nil {
		return fmt.Errorf("include error: %w", err)
	}
	// It shares Request, Response, Session, Application
	// But has its own scope for variables and constants
	childCtx := &ExecutionContext{
		Request:        ae.context.Request,
		Response:       ae.context.Response,
		Server:         nil, // Set below
		Session:        ae.context.Session,
		Application:    ae.context.Application,
		variables:      make(map[string]interface{}),
		constants:      make(map[string]interface{}),
		libraries:      make(map[string]interface{}),
		scopeStack:     make([]map[string]interface{}, 0),
		httpWriter:     ae.context.httpWriter,
		httpRequest:    ae.context.httpRequest,
		startTime:      ae.context.startTime,
		timeout:        ae.context.timeout,
		sessionID:      ae.context.sessionID,
		sessionManager: ae.context.sessionManager,
		isNewSession:   false, // Session already exists
		RootDir:        ae.context.RootDir,
	}

	// Initialize Server object for child context
	childCtx.Server = NewServerObjectWithContext(childCtx)
	childCtx.Server.SetHttpRequest(ae.context.httpRequest)
	childCtx.Server.SetRootDir(ae.context.RootDir)
	childCtx.Server.SetScriptTimeout(ae.context.Server.GetScriptTimeout())

	// Create a new executor for the child context
	childExecutor := &ASPExecutor{
		config:  ae.config,
		context: childCtx,
	}

	childCtx.Server.SetExecutor(childExecutor)

	// Add Document object
	childCtx.variables["document"] = NewDocumentObject(childCtx)

	// Parse ASP code
	parsingOptions := &asp.ASPParsingOptions{
		SaveComments:      false,
		StrictMode:        false,
		AllowImplicitVars: true,
		DebugMode:         ae.config.DebugASP,
	}

	parser := asp.NewASPParserWithOptions(resolvedContent, parsingOptions)
	result, err := parser.Parse()
	if err != nil {
		return fmt.Errorf("failed to parse included ASP file: %w", err)
	}

	// Check for parse errors
	if len(result.Errors) > 0 {
		return fmt.Errorf("ASP parse error in included file: %v", result.Errors[0])
	}

	// Execute blocks (avoid CombinedProgram to keep HTML streaming fast)
	return childExecutor.executeBlocks(result)
}

// executeBlocks executes all blocks in order (HTML and ASP)
func (ae *ASPExecutor) executeBlocks(result *asp.ASPParserResult) error {
	// Hoist declarations across all parsed programs up-front so forward references
	// across ASP blocks (including class members calling later global functions)
	// resolve before any execution begins.
	ae.hoistAllPrograms(result)

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
					// Check for Server.Transfer signal
					if err == ErrServerTransfer {
						// Stop execution of current page (Transfer complete)
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

	// Apply Option Compare for this program before executing any statements
	ae.context.SetOptionCompare(program.OptionCompare)
	// Apply Option Base for array default lower bounds
	ae.context.SetOptionBase(program.OptionBase)

	// Create a visitor to traverse the AST
	v := NewASPVisitor(ae.context, ae)

	// Hoist all procedure and class declarations so forward references work
	ae.hoistDeclarations(v, program)

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
				// Propagate Server.Transfer signal
				if err == ErrServerTransfer {
					return err
				}
				return err
			}
		}
	}

	return nil
}

// hoistAllPrograms pre-registers declarations for every parsed VB program so
// that calls across blocks can resolve forward references.
func (ae *ASPExecutor) hoistAllPrograms(result *asp.ASPParserResult) {
	if result == nil || len(result.VBPrograms) == 0 {
		return
	}

	visitor := NewASPVisitor(ae.context, ae)
	for idx, program := range result.VBPrograms {
		// Skip synthetic combined program (-1) to avoid duplicate registrations
		if idx < 0 {
			continue
		}
		ae.hoistDeclarations(visitor, program)
	}
}

// hoistDeclarations pre-registers procedures and classes before execution
func (ae *ASPExecutor) hoistDeclarations(v *ASPVisitor, program *ast.Program) {
	if v == nil || program == nil || len(program.Body) == 0 {
		return
	}

	for _, stmt := range program.Body {
		ae.hoistStatement(v, stmt)
	}
}

// hoistStatement walks a statement and registers any declarations it finds
func (ae *ASPExecutor) hoistStatement(v *ASPVisitor, stmt ast.Statement) {
	switch s := stmt.(type) {
	case *ast.SubDeclaration:
		if s != nil && s.Identifier != nil {
			_ = v.context.SetVariable(s.Identifier.Name, s)
		}
	case *ast.FunctionDeclaration:
		if s != nil && s.Identifier != nil {
			_ = v.context.SetVariable(s.Identifier.Name, s)
		}
	case *ast.ClassDeclaration:
		_ = v.visitClassDeclaration(s)
	case *ast.StatementList:
		for _, inner := range s.Statements {
			ae.hoistStatement(v, inner)
		}
	}
}

// Detects whether the content has executable ASP code (non-directive/comment blocks)
func hasExecutableASP(content string) bool {
	pos := 0
	for {
		idx := strings.Index(content[pos:], "<%")
		if idx == -1 {
			return false
		}
		idx += pos
		tail := content[idx+2:]
		if strings.HasPrefix(tail, "@") || strings.HasPrefix(tail, "--") {
			end := strings.Index(tail, "%>")
			if end == -1 {
				return false
			}
			pos = idx + 2 + end + 2
			continue
		}
		return true
	}
}

// Removes directives/comments so pure HTML can be streamed without parsing
func stripDirectivesAndComments(content string) string {
	var b strings.Builder
	pos := 0
	for pos < len(content) {
		idx := strings.Index(content[pos:], "<%")
		if idx == -1 {
			b.WriteString(content[pos:])
			break
		}
		idx += pos
		b.WriteString(content[pos:idx])
		tail := content[idx+2:]
		if strings.HasPrefix(tail, "@") || strings.HasPrefix(tail, "--") {
			end := strings.Index(tail, "%>")
			if end == -1 {
				break
			}
			pos = idx + 2 + end + 2
			continue
		}
		// Found executable code; return original remainder
		b.WriteString(content[idx:])
		return b.String()
	}
	return b.String()
}

// CreateObject creates an ASP COM object (like Server.CreateObject)
func (ae *ASPExecutor) CreateObject(objType string) (interface{}, error) {
	objType = strings.ToUpper(objType)

	switch objType {
	// G3 Custom Libraries
	case "G3JSON", "JSON":
		return NewJSONLibrary(ae.context), nil
	case "G3FILES", "FILES":
		return NewFileSystemLibrary(ae.context), nil
	case "G3HTTP", "HTTP":
		return NewHTTPLibrary(ae.context), nil
	case "G3TEMPLATE", "TEMPLATE":
		return NewTemplateLibrary(ae.context), nil
	case "G3MAIL", "MAIL":
		return NewMailLibrary(ae.context), nil
	case "G3CRYPTO", "CRYPTO":
		return NewCryptoLibrary(ae.context), nil
	case "G3REGEXP", "REGEXP":
		return NewRegExpLibrary(ae.context), nil

	// Scripting Objects
	case "SCRIPTING.FILESYSTEMOBJECT", "FILESYSTEMOBJECT", "FSO":
		return NewFileSystemObjectLibrary(ae.context), nil
	case "SCRIPTING.DICTIONARY", "DICTIONARY":
		return NewDictionary(ae.context), nil

	// MSXML2 Objects
	case "MSXML2.SERVERXMLHTTP", "MSXML2.XMLHTTP", "SERVERXMLHTTP", "XMLHTTP":
		return NewServerXMLHTTP(ae.context), nil
	case "MSXML2.DOMDOCUMENT", "DOMDOCUMENT", "MSXML2.DOMDOCUMENT.6.0", "MSXML2.DOMDOCUMENT.3.0":
		return NewDOMDocument(ae.context), nil

	// ADODB Objects
	case "ADODB.CONNECTION", "ADODB", "CONNECTION":
		return NewADOConnection(ae.context), nil
	case "ADODB.RECORDSET", "RECORDSET":
		return NewADORecordset(ae.context), nil
	case "ADODB.STREAM", "STREAM":
		return NewADOStream(ae.context), nil

	default:
		return nil, fmt.Errorf("unsupported object type: %s", objType)
	}
}

// ASPVisitor traverses and executes the VBScript AST
type ASPVisitor struct {
	context       *ExecutionContext
	executor      *ASPExecutor
	depth         int
	withStack     []interface{}
	resumeOnError bool
}

// NewASPVisitor creates a new ASP visitor for AST traversal
func NewASPVisitor(ctx *ExecutionContext, executor *ASPExecutor) *ASPVisitor {
	// Store current executor in context for ExecuteGlobal/Execute
	if executor != nil {
		ctx.currentExecutor = executor
	}
	return &ASPVisitor{
		context:       ctx,
		executor:      executor,
		depth:         0,
		withStack:     make([]interface{}, 0),
		resumeOnError: false,
	}
}

// handleStatementError applies VBScript-style error handling semantics, respecting On Error Resume Next.
func (v *ASPVisitor) handleStatementError(err error) error {
	if err == nil {
		return nil
	}

	// Propagate control-flow signals regardless of resume mode
	if err == ErrServerTransfer || err.Error() == "RESPONSE_END" {
		return err
	}
	if _, ok := err.(*LoopExitError); ok {
		return err
	}
	if _, ok := err.(*ProcedureExitError); ok {
		return err
	}

	if v.resumeOnError {
		v.context.Err.SetError(err)
		return nil
	}

	return err
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
		return v.handleStatementError(v.visitAssignment(stmt))
	case *ast.CallStatement:
		_, err := v.visitExpression(stmt.Callee)
		return v.handleStatementError(err)
	case *ast.CallSubStatement:
		return v.handleStatementError(v.visitCallSubStatement(stmt))
	case *ast.ReDimStatement:
		return v.handleStatementError(v.visitReDim(stmt))
	case *ast.IfStatement:
		return v.handleStatementError(v.visitIf(stmt))
	case *ast.ElseIfStatement:
		return v.handleStatementError(v.visitElseIf(stmt))
	case *ast.ForStatement:
		return v.handleStatementError(v.visitFor(stmt))
	case *ast.ForEachStatement:
		return v.handleStatementError(v.visitForEach(stmt))
	case *ast.DoStatement:
		return v.handleStatementError(v.visitDo(stmt))
	case *ast.WhileStatement:
		return v.handleStatementError(v.visitWhile(stmt))
	case *ast.SelectStatement:
		return v.handleStatementError(v.visitSelect(stmt))
	case *ast.SubDeclaration:
		return v.handleStatementError(v.visitSubDeclaration(stmt))
	case *ast.FunctionDeclaration:
		return v.handleStatementError(v.visitFunctionDeclaration(stmt))
	case *ast.ClassDeclaration:
		return v.handleStatementError(v.visitClassDeclaration(stmt))
	case *ast.WithStatement:
		return v.handleStatementError(v.visitWithStatement(stmt))
	case *ast.VariableDeclaration:
		return v.handleStatementError(v.visitVariableDeclaration(stmt))
	case *ast.VariablesDeclaration:
		return v.handleStatementError(v.visitVariablesDeclaration(stmt))
	case *ast.ConstsDeclaration:
		return v.handleStatementError(v.visitConstDeclaration(stmt))
	case *ast.StatementList:
		for _, s := range stmt.Statements {
			if err := v.VisitStatement(s); err != nil {
				if handled := v.handleStatementError(err); handled != nil {
					return handled
				}
			}
		}
		return nil
	case *ast.OnErrorResumeNextStatement:
		v.resumeOnError = true
		v.context.Err.Clear()
		return nil
	case *ast.OnErrorGoTo0Statement:
		v.resumeOnError = false
		v.context.Err.Clear()
		return nil
	case *ast.ExitForStatement:
		return &LoopExitError{LoopType: "for"}
	case *ast.ExitDoStatement:
		return &LoopExitError{LoopType: "do"}
	case *ast.ExitSubStatement:
		return &ProcedureExitError{Kind: "sub"}
	case *ast.ExitFunctionStatement:
		return &ProcedureExitError{Kind: "function"}
	case *ast.ExitPropertyStatement:
		return &ProcedureExitError{Kind: "property"}
	default:
		// Try to evaluate as expression for side effects
		if expr, ok := node.(ast.Expression); ok {
			_, err := v.visitExpression(expr)
			return v.handleStatementError(err)
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
			if arrObj, ok := toVBArray(obj); ok {
				if len(indexCall.Indexes) > 0 {
					idx, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					index := toInt(idx)
					if arrObj.Set(index, value) {
						_ = v.context.SetVariable(varName, arrObj)
						return nil
					}
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

		// If it's a ServerObject, set the property
		if serverObj, ok := obj.(*ServerObject); ok {
			propNameLower := strings.ToLower(propName)
			switch propNameLower {
			case "scripttimeout":
				timeout := toInt(value)
				return serverObj.SetScriptTimeout(timeout)
			}
		}

		// If it's an ASPLibrary, set the property
		if lib, ok := obj.(interface {
			SetProperty(string, interface{}) error
		}); ok {
			return lib.SetProperty(propName, value)
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
			var newArr *VBArray

			if oldArr, ok := toVBArray(oldVal); ok {
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
func (v *ASPVisitor) preserveCopy(oldArr, newArr *VBArray, dims []int) {
	if oldArr == nil || newArr == nil || len(dims) == 0 {
		return
	}

	copyLen := len(oldArr.Values)
	if len(newArr.Values) < copyLen {
		copyLen = len(newArr.Values)
	}

	if len(dims) == 1 {
		for i := 0; i < copyLen; i++ {
			oldIdx := oldArr.Lower + i
			newIdx := newArr.Lower + i
			if val, ok := oldArr.Get(oldIdx); ok {
				_ = newArr.Set(newIdx, val)
			}
		}
		return
	}

	for i := 0; i < copyLen; i++ {
		oldInner, okOld := toVBArray(oldArr.Values[i])
		newInner, okNew := toVBArray(newArr.Values[i])
		if okOld && okNew {
			v.preserveCopy(oldInner, newInner, dims[1:])
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

// visitElseIf handles ElseIf statements (same structure as If but separate AST type)
func (v *ASPVisitor) visitElseIf(stmt *ast.ElseIfStatement) error {
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
	case *VBArray:
		// Iterate over VBArray (created by ReDim or Array function)
		for _, item := range col.Values {
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
							ge, _ := performBinaryOperation(ast.BinaryOperationGreaterOrEqual, selectValue, startVal, v.context.OptionCompareMode())
							le, _ := performBinaryOperation(ast.BinaryOperationLessOrEqual, selectValue, endVal, v.context.OptionCompareMode())

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
						res, err := performBinaryOperation(bin.Operation, selectValue, rightVal, v.context.OptionCompareMode())
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
				if compareEqual(selectValue, val, v.context.OptionCompareMode()) {
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
				if m.AccessModifier == ast.MethodAccessModifierPublicDefault {
					classDef.DefaultMethod = nameLower
				}
			}
		case *ast.FunctionDeclaration:
			nameLower := strings.ToLower(m.Identifier.Name)
			if m.AccessModifier == ast.MethodAccessModifierPrivate {
				classDef.PrivateMethods[nameLower] = m
			} else {
				classDef.Functions[nameLower] = m
				if m.AccessModifier == ast.MethodAccessModifierPublicDefault {
					classDef.DefaultMethod = nameLower
				}
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

	if access == ast.MethodAccessModifierPublicDefault {
		classDef.DefaultMethod = nameLower
	}
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
		if err := v.context.DefineVariable(varName, NewVBArray(v.context.OptionBase(), 0)); err != nil {
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

// makeNestedArray creates a nested array based on dimensions and Option Base.
func (v *ASPVisitor) makeNestedArray(dims []int) *VBArray {
	base := v.context.OptionBase()
	if len(dims) == 0 {
		return NewVBArray(base, 0)
	}

	lower := base
	upper := dims[0]
	size := upper - lower + 1
	if size < 0 {
		size = 0
	}

	arr := NewVBArray(lower, size)

	if len(dims) > 1 {
		innerDims := dims[1:]
		for i := 0; i < size; i++ {
			arr.Values[i] = v.makeNestedArray(innerDims)
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

		// Check for built-in functions (parameterless call)
		if val, handled := evalBuiltInFunction(varName, []interface{}{}, v.context); handled {
			return val, nil
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

	case *ast.DateLiteral:
		return e.Value, nil

	case *ast.NullLiteral:
		// Null in VBScript represents a special value (no valid data)
		return NullValue{}, nil

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

	right, err := v.visitExpression(expr.Right)
	if err != nil {
		return nil, err
	}

	// Perform operation
	return performBinaryOperation(expr.Operation, left, right, v.context.OptionCompareMode())
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
	// Evaluate arguments first (we'll need them for built-in function checks)
	args := make([]interface{}, 0)
	for _, arg := range arguments {
		val, err := v.visitExpression(arg)
		if err != nil {
			return nil, err
		}
		args = append(args, val)
	}

	// CRITICAL FIX: Handle built-in function vs class method resolution.
	// When an identifier is BRACKETED [isEmpty], it means the user explicitly wants to call
	// a local function/method, NOT the built-in function. This is VBScript's way of escaping
	// reserved words and built-in function names.
	// When NOT bracketed, built-in functions take precedence.
	if ident, ok := objectExpr.(*ast.Identifier); ok {
		// Only use built-in function if:
		// 1. The name matches a built-in function
		// 2. The identifier is NOT bracketed (i.e., user did NOT explicitly escape it)
		if isBuiltInFunctionName(ident.Name) && !ident.IsBracketed {
			if result, handled := evalBuiltInFunction(ident.Name, args, v.context); handled {
				return result, nil
			}
		}
	}

	// CRITICAL: Handle member-method dispatch BEFORE evaluating base for MemberExpression
	// This prevents visitMemberExpression from calling the method with 0 args when
	// we actually want to call it with the provided arguments.
	if member, ok := objectExpr.(*ast.MemberExpression); ok {
		// Evaluate just the parent object, not the full member expression
		parentObj, err := v.visitExpression(member.Object)
		if err != nil {
			return nil, err
		}
		if parentObj != nil {
			propName := ""
			if member.Property != nil {
				propName = member.Property.Name
			}

			// Class instance methods
			if classInst, ok := parentObj.(*ClassInstance); ok {
				return classInst.CallMethod(propName, args...)
			}

			// ASPObject / server-native CallMethod dispatch
			if aspObj, ok := parentObj.(asp.ASPObject); ok {
				return aspObj.CallMethod(propName, args...)
			}
			if caller, ok := parentObj.(interface {
				CallMethod(string, ...interface{}) (interface{}, error)
			}); ok {
				return caller.CallMethod(propName, args...)
			}
		}
	}

	// 1. Check if objectExpr is an Identifier referring to a User Function/Sub (Manual Lookup)
	if ident, ok := objectExpr.(*ast.Identifier); ok {
		if val, exists := v.context.GetVariable(ident.Name); exists {
			if fn, ok := val.(*ast.FunctionDeclaration); ok {
				// For ByRef support, we need to pass the original expressions, not evaluated values
				return v.executeFunctionWithRefs(fn, arguments, v)
			}
			if sub, ok := val.(*ast.SubDeclaration); ok {
				// For ByRef support, we need to pass the original expressions, not evaluated values
				return v.executeSubWithRefs(sub, arguments, v)
			}
		}
	}

	// Evaluate base expression
	var base interface{}
	var err error

	// Special handling for Identifier to avoid premature evaluation of built-in functions
	if ident, ok := objectExpr.(*ast.Identifier); ok {
		// Try to find variable or simple object (like Response)
		// But do NOT trigger parameterless built-in function fallback

		// 1. Check Me
		if strings.EqualFold(ident.Name, "me") {
			base = v.context.GetContextObject()
		} else {
			// 2. Check Variable
			if val, exists := v.context.GetVariable(ident.Name); exists {
				base = val
			} else {
				// 3. Check Built-in Objects
				switch strings.ToLower(ident.Name) {
				case "response":
					base = v.context.Response
				case "request":
					base = v.context.Request
				case "server":
					base = v.context.Server
				case "session":
					base = v.context.Session
				case "application":
					base = v.context.Application
				}
			}
		}
		// If base is still nil, we deliberately do NOT call visitExpression
		// This allows the "built-in function call" logic below to handle it
	} else {
		// Not an identifier, evaluate normally
		base, err = v.visitExpression(objectExpr)
		if err != nil {
			return nil, err
		}
	}

	// VBScript allows property access with optional parentheses when no arguments are provided.
	// Example: Request.QueryString.Count() should behave like Request.QueryString.Count
	// If the callee is a MemberExpression and base resolved to a value, return it when called with empty args.
	if base != nil {
		if _, ok := objectExpr.(*ast.MemberExpression); ok && len(args) == 0 {
			return base, nil
		}
	}

	// Check if this is a built-in function call
	if ident, ok := objectExpr.(*ast.Identifier); ok && base == nil {
		// 3. Fallback: Check Built-in Objects again (if variable lookup failed)
		switch strings.ToLower(ident.Name) {
		case "response":
			base = v.context.Response
		case "request":
			base = v.context.Request
		case "server":
			base = v.context.Server
		case "session":
			base = v.context.Session
		case "application":
			base = v.context.Application
		}

		if base == nil {
			funcName := ident.Name
			// Try custom functions first
			if result, handled := evalCustomFunction(funcName, args, v.context); handled {
				return result, nil
			}

			// CRITICAL FIX: When identifier is BRACKETED [name], it means the user wants to
			// call a local function/method, NOT the built-in function. Check class methods
			// BEFORE built-in functions when bracketed.
			if ident.IsBracketed {
				// For bracketed identifiers, prioritize class methods over built-ins
				if ctxObj := v.context.GetContextObject(); ctxObj != nil {
					if classInst, ok := ctxObj.(*ClassInstance); ok {
						funcNameLower := strings.ToLower(funcName)
						if _, exists := classInst.ClassDef.Functions[funcNameLower]; exists {
							return classInst.CallMethod(funcName, args...)
						}
						if _, exists := classInst.ClassDef.Methods[funcNameLower]; exists {
							return classInst.CallMethod(funcName, args...)
						}
						if _, exists := classInst.ClassDef.PrivateMethods[funcNameLower]; exists {
							return classInst.CallMethod(funcName, args...)
						}
					}
				}
			}

			// Then try built-in functions (but skip if identifier was bracketed - they wanted local)
			if !ident.IsBracketed {
				if result, handled := evalBuiltInFunction(funcName, args, v.context); handled {
					return result, nil
				}
			}

			// If we're inside a class context and didn't find a standalone function,
			// try calling this as a method on the current class instance (implicit Me.FunctionName())
			if ctxObj := v.context.GetContextObject(); ctxObj != nil {
				if classInst, ok := ctxObj.(*ClassInstance); ok {
					// Check if this is actually a method/function in the class
					funcNameLower := strings.ToLower(funcName)
					if _, exists := classInst.ClassDef.Functions[funcNameLower]; exists {
						return classInst.CallMethod(funcName, args...)
					}
					if _, exists := classInst.ClassDef.Methods[funcNameLower]; exists {
						return classInst.CallMethod(funcName, args...)
					}
					if _, exists := classInst.ClassDef.PrivateMethods[funcNameLower]; exists {
						return classInst.CallMethod(funcName, args...)
					}
				}
			}
		}
	}

	// Handle class instance default dispatch (calling the object directly)
	if classInst, ok := base.(*ClassInstance); ok {
		methodName := ""
		if _, ok := objectExpr.(*ast.Identifier); ok {
			// Invocation on the instance variable itself (t(...)) should use default member
			methodName = ""
		} else if ident, ok := objectExpr.(*ast.MemberExpression); ok && ident.Property != nil {
			// Member calls preserve explicit method name
			methodName = ident.Property.Name
		}

		if methodName == "" {
			methodName = classInst.ClassDef.DefaultMethod
		}
		return classInst.CallMethod(methodName, args...)
	}

	// Handle method calls on built-in objects (Default Property / Method)
	// Support both asp.ASPObject and server-native objects that expose CallMethod
	{
		// Determine explicit vs default dispatch
		// - Identifier with parentheses (e.g., Request("x")) should use default dispatch
		// - MemberExpression with property (e.g., Response.Write("x")) should use explicit method
		methodName := ""
		if member, ok := objectExpr.(*ast.MemberExpression); ok && member.Property != nil {
			methodName = member.Property.Name
		}

		// Generic interface for server-native objects
		type callMethoder interface {
			CallMethod(string, ...interface{}) (interface{}, error)
		}

		// First, try asp.ASPObject
		if obj, ok := base.(asp.ASPObject); ok {
			return obj.CallMethod(methodName, args...)
		}

		// Then, try server-native objects implementing CallMethod (e.g., *RequestObject)
		if obj, ok := base.(callMethoder); ok {
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

				// Handle ClassInstance methods (including default members declared in classes)
				if classInst, ok := parentObj.(*ClassInstance); ok {
					return classInst.CallMethod(propName, args...)
				}

				// Handle generic ASPObject method dispatch
				if aspObj, ok := parentObj.(asp.ASPObject); ok {
					return aspObj.CallMethod(propName, args...)
				}

				// Handle server-native CallMethod implementers
				if caller, ok := parentObj.(interface {
					CallMethod(string, ...interface{}) (interface{}, error)
				}); ok {
					return caller.CallMethod(propName, args...)
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

				// Handle ServerObject methods
				if serverObj, ok := parentObj.(*ServerObject); ok {
					propNameLower := strings.ToLower(propName)
					switch propNameLower {
					case "createobject":
						if len(args) > 0 {
							progID := fmt.Sprintf("%v", args[0])
							return serverObj.CreateObject(progID)
						}
						return nil, fmt.Errorf("CreateObject requires a ProgID argument")
					case "htmlencode":
						if len(args) > 0 {
							str := fmt.Sprintf("%v", args[0])
							return serverObj.HTMLEncode(str), nil
						}
						return "", nil
					case "urlencode":
						if len(args) > 0 {
							str := fmt.Sprintf("%v", args[0])
							return serverObj.URLEncode(str), nil
						}
						return "", nil
					case "mappath":
						if len(args) > 0 {
							path := fmt.Sprintf("%v", args[0])
							return serverObj.MapPath(path), nil
						}
						return serverObj.MapPath(""), nil
					case "execute":
						if len(args) > 0 {
							path := fmt.Sprintf("%v", args[0])
							err := serverObj.Execute(path)
							return nil, err
						}
						return nil, fmt.Errorf("Execute requires a path argument")
					case "transfer":
						if len(args) > 0 {
							path := fmt.Sprintf("%v", args[0])
							err := serverObj.Transfer(path)
							return nil, err
						}
						return nil, fmt.Errorf("Transfer requires a path argument")
					case "getlasterror":
						return serverObj.GetLastError(), nil
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
	if arr, ok := toVBArray(base); ok && len(args) > 0 {
		idx := toInt(args[0])
		if val, ok := arr.Get(idx); ok {
			return val, nil
		}
		return nil, nil
	}

	// Handle DictionaryLibrary access (dict("key"))
	if dictLib, ok := base.(*DictionaryLibrary); ok && len(args) > 0 {
		return dictLib.dict.Item(args), nil
	}

	// Handle Dictionary access (dict("key"))
	if dict, ok := base.(*Dictionary); ok && len(args) > 0 {
		return dict.Item(args), nil
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

	if obj != nil {
		// fmt.Printf("DEBUG: visitMemberExpression obj type: %T prop=%s\n", obj, propName)
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
		case "end":
			if err := respObj.End(); err != nil {
				return nil, err
			}
			return nil, fmt.Errorf("RESPONSE_END")
		case "clear":
			respObj.Clear()
			return nil, nil
		case "flush":
			if err := respObj.Flush(); err != nil {
				return nil, err
			}
			return nil, nil
		case "redirect":
			// Redirect usually requires an argument, but if called without one (unlikely for Redirect), we can't do much.
			// However, if parsed as MemberExpression, it implies no arguments were passed here.
			// If arguments were passed, it would likely be IndexOrCallExpression.
			return nil, nil
		}
	}

	// Handle ServerObject properties
	if serverObj, ok := obj.(*ServerObject); ok {
		propNameLower := strings.ToLower(propName)
		switch propNameLower {
		case "scripttimeout":
			return serverObj.GetScriptTimeout(), nil
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

	// Handle ClassInstance
	if classInst, ok := obj.(*ClassInstance); ok {
		val, err := classInst.GetProperty(propName)
		if err != nil {
			return nil, err
		}
		if val != nil {
			return val, nil
		}
		// Fallback to method call (parameterless)
		return classInst.CallMethod(propName)
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
	case time.Time:
		return formatVBDateDefault(v)
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
func performBinaryOperation(op ast.BinaryOperation, left, right interface{}, mode ast.OptionCompareMode) (interface{}, error) {
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
		// VBScript + operator behavior:
		// 1. If BOTH operands are strings, perform string concatenation
		// 2. Otherwise, attempt numeric addition
		// This is critical for code like: "\u00" + "2F" which expects string concatenation
		leftStr, leftIsStr := left.(string)
		rightStr, rightIsStr := right.(string)
		if leftIsStr && rightIsStr {
			return leftStr + rightStr, nil
		}
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
		return compareEqual(left, right, mode), nil
	case ast.BinaryOperationNotEqual:
		return !compareEqual(left, right, mode), nil
	case ast.BinaryOperationLess:
		return compareLess(left, right, mode), nil
	case ast.BinaryOperationGreater:
		return compareLess(right, left, mode), nil
	case ast.BinaryOperationLessOrEqual:
		return !compareLess(right, left, mode), nil
	case ast.BinaryOperationGreaterOrEqual:
		return !compareLess(left, right, mode), nil
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

// compareEqual compares two values for equality with VBScript rules
func compareEqual(left, right interface{}, mode ast.OptionCompareMode) bool {
	if left == nil && right == nil {
		return true
	}
	if left == nil || right == nil {
		return false
	}

	if ls, ok := left.(string); ok {
		if rs, ok2 := right.(string); ok2 {
			return compareStrings(ls, rs, mode) == 0
		}
	}

	if lb, ok := left.(bool); ok {
		if rb, ok2 := right.(bool); ok2 {
			return lb == rb
		}
	}

	if ln, lok := toNumeric(left); lok {
		if rn, rok := toNumeric(right); rok {
			return ln == rn
		}
	}

	return compareStrings(toString(left), toString(right), mode) == 0
}

// compareLess compares if left is less than right with VBScript rules
func compareLess(left, right interface{}, mode ast.OptionCompareMode) bool {
	if ls, ok := left.(string); ok {
		if rs, ok2 := right.(string); ok2 {
			return compareStrings(ls, rs, mode) < 0
		}
	}

	if ln, lok := toNumeric(left); lok {
		if rn, rok := toNumeric(right); rok {
			return ln < rn
		}
	}

	return compareStrings(toString(left), toString(right), mode) < 0
}

func compareStrings(left, right string, mode ast.OptionCompareMode) int {
	if mode == ast.OptionCompareText {
		left = strings.ToLower(left)
		right = strings.ToLower(right)
	}

	switch {
	case left == right:
		return 0
	case left < right:
		return -1
	default:
		return 1
	}
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
func populateRequestData(req *RequestObject, r *http.Request, ctx *ExecutionContext) {
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

	// Resolve core server info
	pathInfo := r.URL.Path
	serverName := r.Host
	serverPort := r.URL.Port()
	if hostOnly, portOnly, err := net.SplitHostPort(r.Host); err == nil {
		serverName = hostOnly
		serverPort = portOnly
	}

	if serverPort == "" {
		if r.TLS != nil {
			serverPort = "443"
		} else {
			serverPort = "80"
		}
	}

	remoteAddr := r.RemoteAddr
	if host, _, err := net.SplitHostPort(r.RemoteAddr); err == nil {
		remoteAddr = host
	}

	localAddr := ""
	if la, ok := r.Context().Value(http.LocalAddrContextKey).(net.Addr); ok {
		if host, port, err := net.SplitHostPort(la.String()); err == nil {
			localAddr = host
			// Prefer bound port if not parsed from host header
			if serverPort == "" {
				serverPort = port
			}
		}
	}

	rootDir := "./www"
	if ctx != nil && ctx.RootDir != "" {
		rootDir = ctx.RootDir
	}
	absRoot, _ := filepath.Abs(rootDir)
	pathTranslated := filepath.Join(absRoot, strings.TrimPrefix(pathInfo, "/"))

	httpsVal := "OFF"
	serverPortSecure := "0"
	if r.TLS != nil {
		httpsVal = "ON"
		serverPortSecure = "1"
	}

	// Authentication info (Basic only)
	authType := ""
	authUser := ""
	authPass := ""
	if user, pass, ok := r.BasicAuth(); ok {
		authType = "Basic"
		authUser = user
		authPass = pass
	}

	// SSL/Certificate placeholders (best effort)
	certSubject := ""
	certIssuer := ""
	certSerial := ""
	certFlags := "0"
	if r.TLS != nil && len(r.TLS.PeerCertificates) > 0 {
		cert := r.TLS.PeerCertificates[0]
		certSubject = cert.Subject.String()
		certIssuer = cert.Issuer.String()
		certSerial = cert.SerialNumber.String()
		certFlags = "1"
	}

	// All HTTP headers (HTTP_* form)
	var allHTTPBuilder strings.Builder
	for name, values := range r.Header {
		if len(values) == 0 {
			continue
		}
		headerName := "HTTP_" + strings.ToUpper(strings.ReplaceAll(name, "-", "_"))
		allHTTPBuilder.WriteString(headerName)
		allHTTPBuilder.WriteString(": ")
		allHTTPBuilder.WriteString(values[0])
		allHTTPBuilder.WriteString("\r\n")
		req.ServerVariables.Add(headerName, values[0])
	}

	// ALL_RAW (original header names)
	var allRawBuilder strings.Builder
	for name, values := range r.Header {
		if len(values) == 0 {
			continue
		}
		allRawBuilder.WriteString(name)
		allRawBuilder.WriteString(": ")
		allRawBuilder.WriteString(values[0])
		allRawBuilder.WriteString("\r\n")
	}

	// Core server variables
	req.ServerVariables.Add("ALL_HTTP", allHTTPBuilder.String())
	req.ServerVariables.Add("ALL_RAW", allRawBuilder.String())
	req.ServerVariables.Add("APPL_MD_PATH", "/LM/W3SVC/1/ROOT")
	req.ServerVariables.Add("APPL_PHYSICAL_PATH", absRoot)
	req.ServerVariables.Add("AUTH_PASSWORD", authPass)
	req.ServerVariables.Add("AUTH_TYPE", authType)
	req.ServerVariables.Add("AUTH_USER", authUser)
	req.ServerVariables.Add("CERT_COOKIE", "")
	req.ServerVariables.Add("CERT_FLAGS", certFlags)
	req.ServerVariables.Add("CERT_ISSUER", certIssuer)
	req.ServerVariables.Add("CERT_KEYSIZE", "0")
	req.ServerVariables.Add("CERT_SECRETKEYSIZE", "0")
	req.ServerVariables.Add("CERT_SERIALNUMBER", certSerial)
	req.ServerVariables.Add("CERT_SERVER_ISSUER", "")
	req.ServerVariables.Add("CERT_SERVER_SUBJECT", "")
	req.ServerVariables.Add("CERT_SUBJECT", certSubject)
	req.ServerVariables.Add("CONTENT_LENGTH", r.Header.Get("Content-Length"))
	req.ServerVariables.Add("CONTENT_TYPE", r.Header.Get("Content-Type"))
	req.ServerVariables.Add("GATEWAY_INTERFACE", "CGI/1.1")
	req.ServerVariables.Add("HTTPS", httpsVal)
	req.ServerVariables.Add("HTTPS_KEYSIZE", "0")
	req.ServerVariables.Add("HTTPS_SECRETKEYSIZE", "0")
	req.ServerVariables.Add("HTTPS_SERVER_ISSUER", "")
	req.ServerVariables.Add("HTTPS_SERVER_SUBJECT", "")
	req.ServerVariables.Add("INSTANCE_ID", "1")
	req.ServerVariables.Add("INSTANCE_META_PATH", "/LM/W3SVC/1")
	req.ServerVariables.Add("LOCAL_ADDR", localAddr)
	req.ServerVariables.Add("LOGON_USER", authUser)
	req.ServerVariables.Add("PATH_INFO", pathInfo)
	req.ServerVariables.Add("PATH_TRANSLATED", pathTranslated)
	req.ServerVariables.Add("QUERY_STRING", r.URL.RawQuery)
	req.ServerVariables.Add("REMOTE_ADDR", remoteAddr)
	req.ServerVariables.Add("REMOTE_HOST", remoteAddr)
	req.ServerVariables.Add("REMOTE_USER", authUser)
	req.ServerVariables.Add("REQUEST_METHOD", r.Method)
	req.ServerVariables.Add("SCRIPT_NAME", pathInfo)
	req.ServerVariables.Add("SERVER_NAME", serverName)
	req.ServerVariables.Add("SERVER_PORT", serverPort)
	req.ServerVariables.Add("SERVER_PORT_SECURE", serverPortSecure)
	req.ServerVariables.Add("SERVER_PROTOCOL", r.Proto)
	req.ServerVariables.Add("SERVER_SOFTWARE", "G3pix-AxonASP/Go")
	req.ServerVariables.Add("URL", pathInfo)
	req.ServerVariables.Add("REQUEST_PATH", r.URL.Path)
	req.ServerVariables.Add("HTTP_USER_AGENT", r.UserAgent())
	req.ServerVariables.Add("HTTP_REFERER", r.Referer())
	req.ServerVariables.Add("HTTP_COOKIE", r.Header.Get("Cookie"))
	req.ServerVariables.Add("HTTP_ACCEPT", r.Header.Get("Accept"))
	req.ServerVariables.Add("HTTP_ACCEPT_LANGUAGE", r.Header.Get("Accept-Language"))
	req.ServerVariables.Add("HTTP_ACCEPT_ENCODING", r.Header.Get("Accept-Encoding"))
	req.ServerVariables.Add("HTTP_HOST", r.Host)
	req.ServerVariables.Add("HTTP_CONNECTION", r.Header.Get("Connection"))
	req.ServerVariables.Add("HTTP_CACHE_CONTROL", r.Header.Get("Cache-Control"))
	req.ServerVariables.Add("HTTP_PRAGMA", r.Header.Get("Pragma"))
	req.ServerVariables.Add("HTTP_UPGRADE_INSECURE_REQUESTS", r.Header.Get("Upgrade-Insecure-Requests"))
	if xfwd := r.Header.Get("X-Forwarded-For"); xfwd != "" {
		req.ServerVariables.Add("HTTP_X_FORWARDED_FOR", xfwd)
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
		// fmt.Printf("DEBUG: visitMemberExpression ClassInstance prop=%s\n", propName)
		val, err := classInst.GetProperty(propName)
		if err != nil {
			return nil, err
		}
		if val != nil {
			return val, nil
		}
		// fmt.Printf("DEBUG: GetProperty returned nil, trying CallMethod for %s\n", propName)
		// If property not found, try method call (parameterless)
		// This emulates VBScript behavior where obj.Method is equivalent to obj.Method()
		return classInst.CallMethod(propName)
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
			// Check for built-in COM objects (New RegExp syntax)
			if strings.ToLower(className) == "regexp" {
				return NewRegExpLibrary(v.context), nil
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
		} else {
			if err := v.VisitStatement(fn.Body); err != nil {
				return nil, err
			}
		}
	}

	// Get return value
	val, _ := v.context.GetVariable(funcName)
	return val, nil
}

// executeFunctionWithRefs executes a user defined function with ByRef parameter support
func (v *ASPVisitor) executeFunctionWithRefs(fn *ast.FunctionDeclaration, arguments []ast.Expression, visitor *ASPVisitor) (interface{}, error) {
	v.context.PushScope()
	defer v.context.PopScope()

	// Map to track ByRef parameters and their original variable names
	byRefMap := make(map[string]string) // param name -> original var name

	// Bind parameters
	for i, param := range fn.Parameters {
		var val interface{}

		if i < len(arguments) {
			// Check if parameter is ByRef
			if param.Modifier == ast.ParameterModifierByRef {
				// For ByRef, we need the original variable name
				if ident, ok := arguments[i].(*ast.Identifier); ok {
					byRefMap[strings.ToLower(param.Identifier.Name)] = strings.ToLower(ident.Name)
					// Get value from caller's scope (could be parent scope or global)
					if parentVal, exists := v.context.GetVariableFromParentScope(ident.Name); exists {
						val = parentVal
					} else if globalVal, exists := v.context.GetVariable(ident.Name); exists {
						// Try global scope if not found in parent
						val = globalVal
					}
				} else {
					// If not an identifier, evaluate and use value (can't ByRef a literal)
					var err error
					val, err = visitor.visitExpression(arguments[i])
					if err != nil {
						return nil, err
					}
				}
			} else {
				// ByVal or default: evaluate the argument
				var err error
				val, err = visitor.visitExpression(arguments[i])
				if err != nil {
					return nil, err
				}
			}
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
		} else {
			if err := v.VisitStatement(fn.Body); err != nil {
				return nil, err
			}
		}
	}

	// Apply ByRef updates back to caller's scope
	for paramName, origVarName := range byRefMap {
		if newVal, exists := v.context.GetVariable(paramName); exists {
			// Try to set in parent scope first
			err := v.context.SetVariableInParentScope(origVarName, newVal)
			if err != nil {
				// If setting in parent scope fails, set globally
				v.context.SetVariable(origVarName, newVal)
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
		} else {
			if err := v.VisitStatement(sub.Body); err != nil {
				return nil, err
			}
		}
	}

	return nil, nil
}

// executeSubWithRefs executes a user defined subroutine with ByRef parameter support
func (v *ASPVisitor) executeSubWithRefs(sub *ast.SubDeclaration, arguments []ast.Expression, visitor *ASPVisitor) (interface{}, error) {
	v.context.PushScope()
	defer v.context.PopScope()

	// Map to track ByRef parameters and their original variable names
	byRefMap := make(map[string]string) // param name -> original var name

	// Bind parameters
	for i, param := range sub.Parameters {
		var val interface{}

		if i < len(arguments) {
			// Check if parameter is ByRef
			if param.Modifier == ast.ParameterModifierByRef {
				// For ByRef, we need the original variable name
				if ident, ok := arguments[i].(*ast.Identifier); ok {
					byRefMap[strings.ToLower(param.Identifier.Name)] = strings.ToLower(ident.Name)
					// Get value from caller's scope (could be parent scope or global)
					if parentVal, exists := v.context.GetVariableFromParentScope(ident.Name); exists {
						val = parentVal
					} else if globalVal, exists := v.context.GetVariable(ident.Name); exists {
						// Try global scope if not found in parent
						val = globalVal
					}
				} else {
					// If not an identifier, evaluate and use value (can't ByRef a literal)
					var err error
					val, err = visitor.visitExpression(arguments[i])
					if err != nil {
						return nil, err
					}
				}
			} else {
				// ByVal or default: evaluate the argument
				var err error
				val, err = visitor.visitExpression(arguments[i])
				if err != nil {
					return nil, err
				}
			}
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
		} else {
			if err := v.VisitStatement(sub.Body); err != nil {
				return nil, err
			}
		}
	}

	// Apply ByRef updates back to caller's scope
	for paramName, origVarName := range byRefMap {
		if newVal, exists := v.context.GetVariable(paramName); exists {
			// Try to set in parent scope first
			err := v.context.SetVariableInParentScope(origVarName, newVal)
			if err != nil {
				// If setting in parent scope fails, set globally
				v.context.SetVariable(origVarName, newVal)
			}
		}
	}

	return nil, nil
}

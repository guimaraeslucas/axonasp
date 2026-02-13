/*
 * AxonASP Server
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
	"bytes"
	"fmt"
	"io"
	"log"
	"math"
	"math/rand"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"runtime"
	"runtime/debug"

	"strconv"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/asp"
	"g3pix.com.br/axonasp/experimental"
	"g3pix.com.br/axonasp/vbscript/ast"
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
	scopeStack       []map[string]interface{}
	parentScopeStack []map[string]interface{} // For "Scope Isolation" (Dynamic Scope Read-Only fallback)
	scopeConstStack  []map[string]interface{}
	contextObject    interface{}  // For Class Instance (Me)
	currentExecutor  *ASPExecutor // Current executor for ExecuteGlobal/Execute

	// Mutex for thread safety
	mu sync.RWMutex

	// Session tracking
	isNewSession     bool
	sessionAbandoned bool

	// Cancellation support for timeout handling
	cancelChan chan struct{}
	cancelled  bool

	// Resource cleanup tracking for COM objects and other resources
	managedResources []interface{}
	resourceMutex    sync.Mutex

	// Include tracking for AxIncludeOnce
	includedOnce map[string]bool
	includeMutex sync.Mutex

	resumeOnError bool
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
		Request:          NewRequestObject(),
		Response:         NewResponseObject(w, r),
		Server:           nil, // Will be initialized below
		Session:          NewSessionObject(sessionID, sessionData.Data),
		Application:      GetGlobalApplication(),
		Err:              NewErrObject(),
		variables:        make(map[string]interface{}),
		constants:        make(map[string]interface{}),
		libraries:        make(map[string]interface{}),
		scopeStack:       make([]map[string]interface{}, 0),
		httpWriter:       w,
		httpRequest:      r,
		startTime:        time.Now(),
		timeout:          timeout,
		rng:              rand.New(rand.NewSource(time.Now().UnixNano())),
		sessionID:        sessionID,
		sessionManager:   sessionManager,
		isNewSession:     isNew,
		sessionAbandoned: false,
		compareMode:      ast.OptionCompareBinary,
		optionBase:       0,
		cancelChan:       make(chan struct{}),
		cancelled:        false,
		managedResources: make([]interface{}, 0),
		includedOnce:     make(map[string]bool),
		resumeOnError:    false,
	}

	// Initialize Server object with context reference
	ctx.Server = NewServerObjectWithContext(ctx)
	ctx.Server.SetHttpRequest(r)

	// Link Response to Request for emergency termination on infinite BinaryRead loops
	ctx.Request.SetResponse(ctx.Response)

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

	// Comparison mode constants
	ctx.constants["vbbinarycompare"] = 0
	ctx.constants["vbtextcompare"] = 1
	ctx.constants["vbdatabasecompare"] = 2

	// DateTime format constants (FormatDateTime)
	ctx.constants["vbgeneraldate"] = 0
	ctx.constants["vblongdate"] = 1
	ctx.constants["vbshortdate"] = 2
	ctx.constants["vblongtime"] = 3
	ctx.constants["vbshorttime"] = 4

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

// MarkSessionAbandoned marks the current session for deletion at request end.
func (ec *ExecutionContext) MarkSessionAbandoned() {
	ec.mu.Lock()
	defer ec.mu.Unlock()
	ec.sessionAbandoned = true
}

// IsSessionAbandoned returns true when Session.Abandon has been called.
func (ec *ExecutionContext) IsSessionAbandoned() bool {
	ec.mu.RLock()
	defer ec.mu.RUnlock()
	return ec.sessionAbandoned
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
	ec.scopeConstStack = append(ec.scopeConstStack, make(map[string]interface{}))
}

// PopScope pops the last local scope
func (ec *ExecutionContext) PopScope() {
	ec.mu.Lock()
	defer ec.mu.Unlock()
	if len(ec.scopeStack) > 0 {
		ec.scopeStack = ec.scopeStack[:len(ec.scopeStack)-1]
	}
	if len(ec.scopeConstStack) > 0 {
		ec.scopeConstStack = ec.scopeConstStack[:len(ec.scopeConstStack)-1]
	}
}

func (ec *ExecutionContext) hasResourceReferenceLocked(value interface{}) bool {
	for _, existing := range ec.variables {
		if sameRuntimeResource(existing, value) {
			return true
		}
		if classInst, ok := existing.(*ClassInstance); ok {
			for _, classVal := range classInst.Variables {
				if sameRuntimeResource(classVal, value) {
					return true
				}
			}
		}
	}

	for i := len(ec.scopeStack) - 1; i >= 0; i-- {
		for _, existing := range ec.scopeStack[i] {
			if sameRuntimeResource(existing, value) {
				return true
			}
			if classInst, ok := existing.(*ClassInstance); ok {
				for _, classVal := range classInst.Variables {
					if sameRuntimeResource(classVal, value) {
						return true
					}
				}
			}
		}
	}

	for i := len(ec.parentScopeStack) - 1; i >= 0; i-- {
		for _, existing := range ec.parentScopeStack[i] {
			if sameRuntimeResource(existing, value) {
				return true
			}
			if classInst, ok := existing.(*ClassInstance); ok {
				for _, classVal := range classInst.Variables {
					if sameRuntimeResource(classVal, value) {
						return true
					}
				}
			}
		}
	}

	if sameRuntimeResource(ec.contextObject, value) {
		return true
	}
	if classInst, ok := ec.contextObject.(*ClassInstance); ok {
		for _, classVal := range classInst.Variables {
			if sameRuntimeResource(classVal, value) {
				return true
			}
		}
	}

	return false
}

func (ec *ExecutionContext) ReleaseResourceIfUnreferenced(value interface{}) {
	if value == nil || isNothingValue(value) || isEmptyValue(value) {
		return
	}

	ec.mu.RLock()
	referenced := ec.hasResourceReferenceLocked(value)
	ec.mu.RUnlock()
	if referenced {
		return
	}

	releaseRuntimeResource(value)
}

func sameRuntimeResource(left interface{}, right interface{}) bool {
	switch l := left.(type) {
	case *ADODBConnection:
		switch r := right.(type) {
		case *ADODBConnection:
			return l == r
		case *ADOConnection:
			return r != nil && l == r.lib
		}
	case *ADOConnection:
		switch r := right.(type) {
		case *ADOConnection:
			return l == r
		case *ADODBConnection:
			return l != nil && l.lib == r
		}
	case *ADODBRecordset:
		switch r := right.(type) {
		case *ADODBRecordset:
			return l == r
		case *ADORecordset:
			return r != nil && l == r.lib
		}
	case *ADORecordset:
		switch r := right.(type) {
		case *ADORecordset:
			return l == r
		case *ADODBRecordset:
			return l != nil && l.lib == r
		}
	case *ADODBOLERecordset:
		switch r := right.(type) {
		case *ADODBOLERecordset:
			return l == r
		case *ADOOLERecordset:
			return r != nil && l == r.lib
		}
	case *ADOOLERecordset:
		switch r := right.(type) {
		case *ADOOLERecordset:
			return l == r
		case *ADODBOLERecordset:
			return l != nil && l.lib == r
		}
	case *COMObject:
		if r, ok := right.(*COMObject); ok {
			return l == r
		}
	}

	return false
}

func releaseRuntimeResource(value interface{}) {
	switch resource := value.(type) {
	case *ADODBRecordset:
		if resource != nil {
			resource.CallMethod("close")
		}
	case *ADORecordset:
		if resource != nil {
			_, _ = resource.CallMethod("close")
		}
	case *ADODBOLERecordset:
		if resource != nil {
			resource.CallMethod("close")
		}
	case *ADOOLERecordset:
		if resource != nil {
			_, _ = resource.CallMethod("close")
		}
	case *COMObject:
		if resource != nil {
			resource.release()
		}
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
// Returns error if attempting to overwrite a constant.
// Resolution order follows VBScript semantics: local constants → local scopes →
// class members (context object) → global constants → global variables.
// Class members take priority over global constants so that a class private
// field with the same name as a global constant can still be assigned.
func (ec *ExecutionContext) SetVariable(name string, value interface{}) error {
	nameLower := lowerKey(name)

	ec.mu.Lock()
	defer ec.mu.Unlock()

	// 1. Check local scope constants (current function / block constants)
	for i := len(ec.scopeConstStack) - 1; i >= 0; i-- {
		if _, exists := ec.scopeConstStack[i][nameLower]; exists {
			return fmt.Errorf("cannot reassign constant '%s'", name)
		}
	}

	// 2. Check current local scope only
	if len(ec.scopeStack) > 0 {
		currentScope := ec.scopeStack[len(ec.scopeStack)-1]
		if _, exists := currentScope[nameLower]; exists {
			currentScope[nameLower] = value
			return nil
		}
	}

	// 3. Check Context Object (Class Member) — takes priority over global constants
	if ec.contextObject != nil {
		if internal, ok := ec.contextObject.(interface {
			SetMember(string, interface{}) (bool, error)
		}); ok {
			if handled, err := internal.SetMember(nameLower, value); handled {
				return err
			}
		}
	}

	// 4. Check global constants — only after class member check
	if _, exists := ec.constants[nameLower]; exists {
		return fmt.Errorf("cannot reassign constant '%s'", name)
	}

	// 5. Existing Global Variable
	if _, exists := ec.variables[nameLower]; exists {
		ec.variables[nameLower] = value
		return nil
	}

	// 6. Implicit Local Variable (inside procedure/class scope)
	if len(ec.scopeStack) > 0 {
		ec.scopeStack[len(ec.scopeStack)-1][nameLower] = value
		return nil
	}

	// 7. Global Variable (module-level assignment)
	ec.variables[nameLower] = value
	return nil
}

// SetVariableGlobal sets a variable directly in global scope, bypassing local scopes
// Used by ExecuteGlobal to ensure variables are set globally even when called from class methods
func (ec *ExecutionContext) SetVariableGlobal(name string, value interface{}) error {
	nameLower := lowerKey(name)

	ec.mu.Lock()
	defer ec.mu.Unlock()

	// Check if this is a constant
	if _, exists := ec.constants[nameLower]; exists {
		return fmt.Errorf("cannot reassign constant '%s'", name)
	}

	// Set directly in global scope
	ec.variables[nameLower] = value
	return nil
}

// DefineVariable defines a variable in the current scope (Dim)
func (ec *ExecutionContext) DefineVariable(name string, value interface{}) error {
	nameLower := lowerKey(name)
	ec.mu.Lock()
	defer ec.mu.Unlock()

	// Prevent defining a variable with the same name as a constant in the current scope
	if len(ec.scopeConstStack) > 0 {
		if _, exists := ec.scopeConstStack[len(ec.scopeConstStack)-1][nameLower]; exists {
			return fmt.Errorf("cannot define variable with name of existing constant '%s'", name)
		}
	} else {
		if _, exists := ec.constants[nameLower]; exists {
			return fmt.Errorf("cannot define variable with name of existing constant '%s'", name)
		}
	}

	if len(ec.scopeStack) > 0 {
		// Define in top scope
		ec.scopeStack[len(ec.scopeStack)-1][nameLower] = value
	} else {
		// Define global
		ec.variables[nameLower] = value
	}
	return nil
}

// DefineVariableGlobal defines a variable directly in the global scope (for ExecuteGlobal)
// This bypasses local scopes, ensuring variables are always defined globally
func (ec *ExecutionContext) DefineVariableGlobal(name string, value interface{}) error {
	nameLower := lowerKey(name)
	ec.mu.Lock()
	defer ec.mu.Unlock()

	// Prevent defining a variable with the same name as a global constant
	if _, exists := ec.constants[nameLower]; exists {
		return fmt.Errorf("cannot define variable with name of existing constant '%s'", name)
	}

	// Always define in global scope
	ec.variables[nameLower] = value
	return nil
}

// GetVariable gets a variable from the execution context (case-insensitive)
func (ec *ExecutionContext) GetVariable(name string) (interface{}, bool) {
	nameLower := lowerKey(name)

	// Debug for rs
	if nameLower == "rs" {
		//fmt.Printf("[DEBUG] GetVariable(rs) called, scopeStack depth=%d\n", len(ec.scopeStack))
	}

	ec.mu.RLock()

	// 1. Check local constants (top-down)
	for i := len(ec.scopeConstStack) - 1; i >= 0; i-- {
		if val, exists := ec.scopeConstStack[i][nameLower]; exists {
			ec.mu.RUnlock()
			return val, true
		}
	}

	// 2. Check global constants
	if val, exists := ec.constants[nameLower]; exists {
		ec.mu.RUnlock()
		return val, true
	}

	// 3. Check all local scopes (top-down)
	for i := len(ec.scopeStack) - 1; i >= 0; i-- {
		if val, exists := ec.scopeStack[i][nameLower]; exists {
			ec.mu.RUnlock()
			return val, true
		}
	}

	// 4. Check Context Object (Class Member)
	// Local/procedure scopes are resolved first; class members are fallback
	// when no scoped symbol exists.
	// NOTE: We must release the lock BEFORE calling GetMember/GetProperty because
	// those can trigger property execution which needs to acquire the lock for PushScope
	ctxObj := ec.contextObject

	// Release lock before checking context object (may call property getters)
	ec.mu.RUnlock()

	// Check context object without holding the lock - this is step 4
	// It must come BEFORE parent scopes and global variables to maintain proper precedence
	if ctxObj != nil {
		// Try internal access first (allows Private members)
		if internal, ok := ctxObj.(interface {
			GetMember(string) (interface{}, bool, error)
		}); ok {
			if val, found, _ := internal.GetMember(nameLower); found {
				return val, true
			}
		} else if getter, ok := ctxObj.(interface{ GetProperty(string) interface{} }); ok {
			val := getter.GetProperty(nameLower)
			if val != nil {
				return val, true
			}
		}
	}

	// Re-acquire lock to check remaining lookups
	ec.mu.RLock()
	defer ec.mu.RUnlock()

	// 5. Check Parent Scope Stack (Isolated Caller Scopes) - Read Only Fallback
	// This restores visibility of caller's variables for Class Methods (pseudo-dynamic scoping)
	// without allowing modification via SetVariable (which only checks scopeStack).
	for i := len(ec.parentScopeStack) - 1; i >= 0; i-- {
		if val, exists := ec.parentScopeStack[i][nameLower]; exists {
			return val, true
		}
	}

	// 6. Check Global Variables
	if val, exists := ec.variables[nameLower]; exists {
		if nameLower == "rs" {
			//fmt.Printf("[DEBUG] GetVariable(rs): found in globals, val=%T\n", val)
		}
		return val, true
	}

	if nameLower == "rs" {
		//fmt.Printf("[DEBUG] GetVariable(rs): NOT FOUND!\n")
	}
	return nil, false
}

// HasVariableInCurrentScope reports whether a variable exists in the current local scope.
func (ec *ExecutionContext) HasVariableInCurrentScope(name string) bool {
	nameLower := lowerKey(name)
	ec.mu.RLock()
	defer ec.mu.RUnlock()
	if len(ec.scopeStack) == 0 {
		return false
	}
	_, exists := ec.scopeStack[len(ec.scopeStack)-1][nameLower]
	return exists
}

// HasVariableInAccessibleScopes reports whether a symbol exists in local or parent fallback scopes.
// This excludes global scope and context-object members; it is used to avoid implicit class-method
// dispatch when a scoped symbol already exists.
func (ec *ExecutionContext) HasVariableInAccessibleScopes(name string) bool {
	nameLower := lowerKey(name)
	ec.mu.RLock()
	defer ec.mu.RUnlock()

	for i := len(ec.scopeConstStack) - 1; i >= 0; i-- {
		if _, exists := ec.scopeConstStack[i][nameLower]; exists {
			return true
		}
	}

	for i := len(ec.scopeStack) - 1; i >= 0; i-- {
		if _, exists := ec.scopeStack[i][nameLower]; exists {
			return true
		}
	}

	for i := len(ec.parentScopeStack) - 1; i >= 0; i-- {
		if _, exists := ec.parentScopeStack[i][nameLower]; exists {
			return true
		}
	}

	return false
}

// buildClassIsolatedParentScopes prepares read-only fallback scopes for class method execution.
// It keeps previously inherited parents and, when entering a class boundary from non-class
// code, exposes the caller scope chain as read-only fallback (writes remain isolated).
func buildClassIsolatedParentScopes(prevParentScopeStack []map[string]interface{}, prevScopeStack []map[string]interface{}, oldContextObj interface{}) []map[string]interface{} {
	combinedParents := make([]map[string]interface{}, 0, len(prevParentScopeStack)+len(prevScopeStack))
	combinedParents = append(combinedParents, prevParentScopeStack...)

	_, calledFromClass := oldContextObj.(*ClassInstance)
	if !calledFromClass && len(prevScopeStack) > 0 {
		combinedParents = append(combinedParents, prevScopeStack...)
	}

	return combinedParents
}

// GetGlobalVariable gets a variable directly from global scope (case-insensitive).
// This bypasses local scopes and context objects.
func (ec *ExecutionContext) GetGlobalVariable(name string) (interface{}, bool) {
	nameLower := lowerKey(name)
	ec.mu.RLock()
	defer ec.mu.RUnlock()
	val, exists := ec.variables[nameLower]
	return val, exists
}

// SetResumeOnError updates the current On Error Resume Next state.
func (ec *ExecutionContext) SetResumeOnError(enabled bool) {
	ec.mu.Lock()
	defer ec.mu.Unlock()
	ec.resumeOnError = enabled
}

// IsResumeOnError reports whether On Error Resume Next is currently active.
func (ec *ExecutionContext) IsResumeOnError() bool {
	ec.mu.RLock()
	defer ec.mu.RUnlock()
	return ec.resumeOnError
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

		// Constants in parent scope
		if val, exists := ec.scopeConstStack[parentIndex][nameLower]; exists {
			return val, true
		}

		if val, exists := ec.scopeStack[parentIndex][nameLower]; exists {
			return val, true
		}
	}

	// Check globals as fallback (variable might be global)
	if val, exists := ec.constants[nameLower]; exists {
		return val, true
	}
	if val, exists := ec.variables[nameLower]; exists {
		return val, true
	}

	return nil, false
}

// SetVariableInParentScope sets a variable in the parent scope (for ByRef parameters)
// This mirrors SetVariable logic but skips the current (topmost) scope, ensuring
// that ByRef writebacks go to the correct location: existing parent scope entries,
// context object (class members), or global variables — never creating new entries
// in a scope where the variable didn't already exist.
func (ec *ExecutionContext) SetVariableInParentScope(name string, value interface{}) error {
	nameLower := lowerKey(name)

	ec.mu.Lock()

	// Check constants in all parent scopes — if the variable is a constant,
	// silently skip the writeback (VBScript allows passing constants ByRef
	// but the writeback is simply a no-op)
	for i := len(ec.scopeConstStack) - 2; i >= 0; i-- {
		if _, exists := ec.scopeConstStack[i][nameLower]; exists {
			ec.mu.Unlock()
			return nil // Silently skip — can't write back to a constant
		}
	}
	if _, exists := ec.constants[nameLower]; exists {
		ec.mu.Unlock()
		return nil // Silently skip — can't write back to a constant
	}

	// 1. Check parent local scopes (excluding topmost which is the current function)
	for i := len(ec.scopeStack) - 2; i >= 0; i-- {
		if _, exists := ec.scopeStack[i][nameLower]; exists {
			ec.scopeStack[i][nameLower] = value
			ec.mu.Unlock()
			return nil
		}
	}

	// 2. Check Context Object (Class Member) before falling through to globals
	ctxObj := ec.contextObject
	ec.mu.Unlock()

	if ctxObj != nil {
		if internal, ok := ctxObj.(interface {
			SetMember(string, interface{}) (bool, error)
		}); ok {
			if handled, err := internal.SetMember(nameLower, value); handled {
				return err
			}
		}
	}

	// 3. Global Variables (Default)
	ec.mu.Lock()
	ec.variables[nameLower] = value
	ec.mu.Unlock()
	return nil
}

// SetConstant sets a constant in the execution context (case-insensitive)
// Constants cannot be changed after initialization
func (ec *ExecutionContext) SetConstant(name string, value interface{}) error {
	nameLower := lowerKey(name)

	ec.mu.Lock()
	defer ec.mu.Unlock()

	// If we are inside a local scope, bind the constant to that scope
	if len(ec.scopeStack) > 0 {
		topIndex := len(ec.scopeConstStack) - 1
		consts := ec.scopeConstStack[topIndex]

		if _, exists := consts[nameLower]; exists {
			return fmt.Errorf("constant '%s' already defined", name)
		}

		if _, exists := ec.scopeStack[topIndex][nameLower]; exists {
			return fmt.Errorf("cannot define constant with name of existing variable '%s'", name)
		}

		consts[nameLower] = value
		return nil
	}

	// Global constant
	if _, exists := ec.constants[nameLower]; exists {
		return fmt.Errorf("constant '%s' already defined", name)
	}

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

// IsCancelled checks if execution has been cancelled (e.g., due to timeout)
func (ec *ExecutionContext) IsCancelled() bool {
	ec.mu.RLock()
	defer ec.mu.RUnlock()
	return ec.cancelled
}

// Cancel signals all goroutines to stop execution
func (ec *ExecutionContext) Cancel() {
	ec.mu.Lock()
	defer ec.mu.Unlock()
	if !ec.cancelled {
		ec.cancelled = true
		close(ec.cancelChan)
	}
}

// ShouldStop checks if execution should stop (cancelled, timed out, or response ended)
func (ec *ExecutionContext) ShouldStop() bool {
	// Check cancellation first (non-blocking)
	select {
	case <-ec.cancelChan:
		return true
	default:
	}
	// Check if Response.End() was called
	if ec.Response != nil && ec.Response.IsEnded() {
		return true
	}
	// Check timeout
	if time.Since(ec.startTime) > ec.timeout {
		return true
	}
	return false
}

// RegisterManagedResource registers a resource for automatic cleanup
func (ec *ExecutionContext) RegisterManagedResource(resource interface{}) {
	if resource == nil {
		return
	}
	ec.resourceMutex.Lock()
	defer ec.resourceMutex.Unlock()
	ec.managedResources = append(ec.managedResources, resource)
}

// Cleanup releases all managed resources
func (ec *ExecutionContext) Cleanup() {
	ec.resourceMutex.Lock()
	defer ec.resourceMutex.Unlock()

	// Release all resources safely, catching any panics
	for _, resource := range ec.managedResources {
		func() {
			defer func() {
				if r := recover(); r != nil {
					// Silently recover from cleanup errors
					fmt.Printf("[DEBUG] Cleanup: recovered from panic during resource release: %v\n", r)
				}
			}()

			switch res := resource.(type) {
			case *ADODBConnection:
				if res != nil {
					res.CallMethod("close")
				}
			case *ADODBRecordset:
				if res != nil {
					res.CallMethod("close")
				}
			case *ADODBOLERecordset:
				if res != nil {
					res.CallMethod("close")
				}
			case *COMObject:
				if res != nil {
					res.release()
				}
			}
		}()
	}
	ec.managedResources = ec.managedResources[:0]
}

// Server_MapPath converts a virtual path to an absolute file system path
func (ec *ExecutionContext) Server_MapPath(path string) string {
	rootDir := ec.RootDir
	if rootDir == "" {
		rootDir = "./www"
	}

	// Get absolute rootDir
	absRootDir, err := filepath.Abs(rootDir)
	if err != nil {
		absRootDir = rootDir
	}

	// The application root is the configured WEB_ROOT (absRootDir).
	// Virtual paths (starting with /) always resolve relative to this root,
	// matching classic ASP/IIS behavior where Server.MapPath("/path")
	// resolves from the web application root.
	appRoot := absRootDir

	// Handle different path formats
	if path == "/" || path == "" {
		return appRoot
	}

	cleanPath := strings.ReplaceAll(path, "\\", "/")

	// Absolute virtual path (starts with /) -> app root-based
	if strings.HasPrefix(cleanPath, "/") {
		cleanPath = strings.TrimPrefix(cleanPath, "/")
		fullPath := filepath.Join(appRoot, cleanPath)
		return fullPath
	}

	// Relative path -> prefer current script directory when available
	baseDir := absRootDir
	if ec.CurrentDir != "" {
		absCurrentDir, err := filepath.Abs(ec.CurrentDir)
		if err == nil {
			baseDir = absCurrentDir
		} else {
			baseDir = ec.CurrentDir
		}
	}
	fullPath := filepath.Join(baseDir, cleanPath)

	// Compatibility: if a relative path does not exist from current script dir,
	// fall back to app root for root-relative style paths used by frameworks.
	if baseDir != absRootDir && !strings.HasPrefix(cleanPath, "./") && !strings.HasPrefix(cleanPath, "../") {
		if _, err := os.Stat(fullPath); err != nil {
			rootPath := filepath.Join(appRoot, cleanPath)
			if _, rootErr := os.Stat(rootPath); rootErr == nil {
				return rootPath
			}
		}
	}

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

	resolvedContent := fileContent
	// Pre-process Includes if not already resolved
	if !alreadyResolved {
		var err error
		resolvedContent, err = asp.ResolveIncludes(fileContent, filePath, ae.config.RootDir, nil)
		if err != nil {
			return fmt.Errorf("include error: %w", err)
		}
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
		parsed, err := asp.ParseResolvedWithCache(resolvedContent, parsingOptions)
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
		if ae.config != nil && ae.config.UseVM {
			return ae.executeVM(resolvedContent, result)
		}
		// Always execute per-block to avoid re-emitting large HTML through VB conversion
		return ae.executeBlocks(result)
	}

	go func() {
		defer func() {
			if rec := recover(); rec != nil {
				if ae.config != nil && ae.config.DebugASP {
					done <- fmt.Errorf("runtime panic: %v\n%s", rec, debug.Stack())
					return
				}
				done <- fmt.Errorf("runtime panic: %v", rec)
			}
		}()

		err := execFunc()
		done <- err
	}()

	// Wait for execution or timeout
	stopByControlFlow := false
	select {
	case err := <-done:
		if err != nil {
			if err.Error() == "RESPONSE_END" || err == ErrServerTransfer || strings.Contains(err.Error(), "runtime panic: RESPONSE_END") {
				stopByControlFlow = true
			} else {
				return err
			}
		}
	case <-time.After(timeout):
		// Cancel the execution context to signal the goroutine to stop
		ae.context.Cancel()
		return fmt.Errorf("script execution timeout (>%d seconds)", ae.config.ScriptTimeout)
	}

	// Write response to HTTP ResponseWriter
	// Check if Response.End() was already called (which flushes headers/content)
	if !stopByControlFlow && !ae.context.Response.IsEnded() {
		// Flush the Response object (writes headers and buffer)
		if err := ae.context.Response.Flush(); err != nil {
			return fmt.Errorf("failed to flush response: %w", err)
		}
	}

	// Save session data to file after request completes
	if err := ae.saveSession(); err != nil {
		fmt.Printf("Warning: Failed to save session: %v\n", err)
	}

	// Cleanup all managed resources (COM objects, database connections, etc.)
	// Must be called at the very end, after everything else
	if ae.context != nil {
		ae.context.Cleanup()
	}

	if asp.ShouldForceFreeMemory(len(resolvedContent)) {
		runtime.GC()
		debug.FreeOSMemory()
	}

	return nil
}

// saveSession persists the current session data to file
func (ae *ASPExecutor) saveSession() error {
	if ae.context == nil || ae.context.sessionManager == nil {
		return fmt.Errorf("no session context available")
	}

	for key, value := range ae.context.Session.Data {
		if strings.Contains(strings.ToLower(key), "isauthenticatedintranet") {
			traceAuthFlow(ae.context, "session before save %s=%v", key, value)
		}
	}

	if ae.context.IsSessionAbandoned() {
		traceAuthFlow(ae.context, "session abandoned=true; deleting session")
		return ae.context.sessionManager.DeleteSession(ae.context.sessionID)
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

	// It shares Request, Response, Session, Application
	// But has its own scope for variables and constants
	childCtx := &ExecutionContext{
		Request:          ae.context.Request,
		Response:         ae.context.Response,
		Server:           nil, // Set below
		Session:          ae.context.Session,
		Application:      ae.context.Application,
		Err:              ae.context.Err,
		variables:        make(map[string]interface{}),
		constants:        make(map[string]interface{}),
		libraries:        make(map[string]interface{}),
		scopeStack:       make([]map[string]interface{}, 0),
		scopeConstStack:  make([]map[string]interface{}, 0),
		httpWriter:       ae.context.httpWriter,
		httpRequest:      ae.context.httpRequest,
		startTime:        ae.context.startTime,
		timeout:          ae.context.timeout,
		rng:              ae.context.rng,
		lastRnd:          ae.context.lastRnd,
		hasLastRnd:       ae.context.hasLastRnd,
		sessionID:        ae.context.sessionID,
		sessionManager:   ae.context.sessionManager,
		isNewSession:     false, // Session already exists
		RootDir:          ae.context.RootDir,
		CurrentFile:      physicalPath,
		CurrentDir:       filepath.Dir(physicalPath),
		compareMode:      ae.context.compareMode,
		optionBase:       ae.context.optionBase,
		cancelChan:       ae.context.cancelChan,
		cancelled:        ae.context.cancelled,
		managedResources: make([]interface{}, 0),
		includedOnce:     ae.context.includedOnce,
	}
	for name, value := range ae.context.constants {
		childCtx.constants[name] = value
	}

	// Initialize Server object for child context
	childCtx.Server = NewServerObjectWithContext(childCtx)
	childCtx.Server.SetHttpRequest(ae.context.httpRequest)
	childCtx.Server.SetRootDir(ae.context.RootDir)
	childCtx.Server.SetScriptTimeout(ae.context.Server.GetScriptTimeout())
	_ = childCtx.Server.SetProperty("_scriptDir", childCtx.CurrentDir)

	// Create a new executor for the child context
	childExecutor := &ASPExecutor{
		config:  ae.config,
		context: childCtx,
	}

	childCtx.Server.SetExecutor(childExecutor)

	// Add Document object
	childCtx.variables["document"] = NewDocumentObject(childCtx)
	childCtx.variables["err"] = childCtx.Err

	// Parse ASP code
	parsingOptions := &asp.ASPParsingOptions{
		SaveComments:      false,
		StrictMode:        false,
		AllowImplicitVars: true,
		DebugMode:         ae.config.DebugASP,
	}
	_, result, err := asp.ParseWithCache(content, physicalPath, ae.config.RootDir, parsingOptions)
	if err != nil {
		return fmt.Errorf("failed to parse included ASP file: %w", err)
	}

	// Check for parse errors
	if len(result.Errors) > 0 {
		return fmt.Errorf("ASP parse error in included file: %v", result.Errors[0])
	}

	// Execute blocks (combined program preserves HTML/control-flow semantics)
	return childExecutor.executeBlocks(result)
}

// executeBlocks executes all blocks in order (HTML and ASP)
func (ae *ASPExecutor) executeBlocks(result *asp.ASPParserResult) error {
	if result == nil {
		return nil
	}

	if result.CombinedProgram != nil {
		visitor := NewASPVisitor(ae.context, ae)
		ae.hoistDeclarations(visitor, result.CombinedProgram)
		if err := ae.executeVBProgram(result.CombinedProgram); err != nil {
			if err.Error() == "RESPONSE_END" {
				return nil
			}
			if err == ErrServerTransfer {
				return nil
			}
			return err
		}
		return nil
	}

	// Hoist declarations across all parsed programs up-front so forward references
	// across ASP blocks (including class members calling later global functions)
	// resolve before any execution begins.
	ae.hoistAllPrograms(result)

	for i, block := range result.Blocks {
		// Check if execution should stop (cancelled, timed out, or Response.End)
		if ae.context.ShouldStop() {
			return nil
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

func (ae *ASPExecutor) executeVM(resolvedContent string, result *asp.ASPParserResult) error {
	if result == nil || result.CombinedProgram == nil {
		return nil
	}

	// Apply options and hoist declarations so script-defined classes are available for CreateObject.
	ae.context.SetOptionCompare(result.CombinedProgram.OptionCompare)
	ae.context.SetOptionBase(result.CombinedProgram.OptionBase)
	visitor := NewASPVisitor(ae.context, ae)
	ae.hoistDeclarations(visitor, result.CombinedProgram)

	compileFn := func() (*experimental.Function, error) {
		compiler := experimental.NewCompiler()
		if err := compiler.Compile(result.CombinedProgram); err != nil {
			return nil, err
		}
		return compiler.MainFunction(), nil
	}

	fn, err := experimental.CompileWithCache(resolvedContent, "default", compileFn)
	if err != nil {
		return fmt.Errorf("VM compile error: %w", err)
	}

	vm := experimental.NewVM(fn, NewVMHostAdapter(ae.context, ae))
	if err := vm.Run(); err != nil {
		return fmt.Errorf("VM runtime error: %w", err)
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

			// Check if execution should stop (cancelled, timed out, or Response.End)
			if ae.context.ShouldStop() {
				return fmt.Errorf("RESPONSE_END")
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
		//log.Printf("[DEBUG] hoistDeclarations: stmt %d is type %T\n", i, stmt)
		ae.hoistStatement(v, stmt)
	}
}

// hoistStatement walks a statement and registers any declarations it finds
func (ae *ASPExecutor) hoistStatement(v *ASPVisitor, stmt ast.Statement) {
	switch s := stmt.(type) {
	case *ast.SubDeclaration:
		if s != nil && s.Identifier != nil {
			//log.Printf("[DEBUG] Hoisting Sub: %s\n", s.Identifier.Name)
			_ = v.context.SetVariable(s.Identifier.Name, s)
		}
	case *ast.FunctionDeclaration:
		if s != nil && s.Identifier != nil {
			//log.Printf("[DEBUG] Hoisting Function: %s\n", s.Identifier.Name)
			_ = v.context.SetVariable(s.Identifier.Name, s)
		}
	case *ast.ClassDeclaration:
		if s != nil && s.Identifier != nil {
			//log.Printf("[DEBUG] Hoisting Class: %s\n", s.Identifier.Name)
		}
		_ = v.visitClassDeclaration(s)
	// Hoist Dim declarations: In VBScript, module-level Dim is processed before
	// any code runs (like var hoisting in JavaScript). We pre-declare variables
	// so that assignments before the textual Dim statement work correctly.
	case *ast.VariableDeclaration:
		if s != nil && s.Identifier != nil {
			name := s.Identifier.Name
			if _, exists := v.context.GetVariable(name); !exists {
				if len(s.ArrayDims) > 0 {
					// Fixed-size array: pre-create the array
					dims := make([]int, len(s.ArrayDims))
					for i, dimExpr := range s.ArrayDims {
						dimVal, err := v.visitExpression(dimExpr)
						if err == nil {
							dims[i] = toInt(dimVal)
						}
					}
					arr := v.makeNestedArray(dims)
					_ = v.context.DefineVariable(name, arr)
				} else if s.IsDynamicArray {
					_ = v.context.DefineVariable(name, NewVBArray(v.context.OptionBase(), 0))
				} else {
					_ = v.context.DefineVariable(name, nil)
				}
			}
		}
	case *ast.VariablesDeclaration:
		if s != nil {
			for _, decl := range s.Variables {
				ae.hoistStatement(v, decl)
			}
		}
	case *ast.StatementList:
		for _, inner := range s.Statements {
			ae.hoistStatement(v, inner)
		}
	case *ast.IfStatement:
		if s.Consequent != nil {
			ae.hoistStatement(v, s.Consequent)
		}
		if s.Alternate != nil {
			ae.hoistStatement(v, s.Alternate)
		}
	case *ast.ElseIfStatement:
		if s.Consequent != nil {
			ae.hoistStatement(v, s.Consequent)
		}
		if s.Alternate != nil {
			ae.hoistStatement(v, s.Alternate)
		}
	case *ast.ForStatement:
		for _, inner := range s.Body {
			ae.hoistStatement(v, inner)
		}
	case *ast.ForEachStatement:
		for _, inner := range s.Body {
			ae.hoistStatement(v, inner)
		}
	case *ast.DoStatement:
		for _, inner := range s.Body {
			ae.hoistStatement(v, inner)
		}
	case *ast.WhileStatement:
		for _, inner := range s.Body {
			ae.hoistStatement(v, inner)
		}
	case *ast.SelectStatement:
		for _, c := range s.Cases {
			ae.hoistStatement(v, c)
		}
	case *ast.CaseStatement:
		for _, inner := range s.Body {
			ae.hoistStatement(v, inner)
		}
	case *ast.WithStatement:
		for _, inner := range s.Body {
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
	objType = strings.TrimSpace(objType)
	originalObjType := objType
	objType = strings.ToUpper(objType)
	//fmt.Printf("[DEBUG] CreateObject called: %s\n", objType)

	if ae.context != nil {
		if classVal, ok := ae.context.GetVariable(originalObjType); ok {
			if classDef, ok := classVal.(*ClassDef); ok {
				return NewClassInstance(classDef, ae.context)
			}
		}
	}

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
	case "G3FILEUPLOADER", "FILEUPLOADER":
		return NewFileUploaderLibrary(ae.context), nil
	case "G3DB", "DB":
		return NewG3DBLibrary(ae.context), nil
	case "G3ZIP", "ZIP":
		return NewZIPLibrary(ae.context), nil
	case "G3FC", "FC":
		return NewG3FCLibrary(ae.context), nil

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

	// ADOX Objects
	case "ADOX.CATALOG", "ADOX.CATALOG.6.0":
		return NewADOXCatalog(ae.context), nil

	// ADODB Objects
	case "ADODB.CONNECTION", "ADODB", "CONNECTION":
		return NewADOConnection(ae.context), nil
	case "ADODB.RECORDSET", "RECORDSET":
		return NewADORecordset(ae.context), nil
	case "ADODB.STREAM", "STREAM":
		return NewADOStream(ae.context), nil

	// WScript Objects
	case "WSCRIPT.SHELL", "SHELL", "WSCRIPT":
		return NewWScriptShell(ae.context), nil

	default:
		comObj, err := NewCOMObject(originalObjType)
		if err == nil {
			return comObj, nil
		}
		return nil, fmt.Errorf("unsupported object type: %s", objType)
	}
}

// ASPVisitor traverses and executes the VBScript AST
type ASPVisitor struct {
	context             *ExecutionContext
	executor            *ASPExecutor
	depth               int
	withStack           []interface{}
	resumeOnError       bool
	forceGlobal         bool   // When true, Dim statements define variables in global scope (for ExecuteGlobal)
	currentFunctionName string // Tracks the current function being executed (for return value assignment)
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
		resumeOnError: ctx.IsResumeOnError(),
		forceGlobal:   false,
	}
}

// NewASPVisitorGlobal creates a new ASP visitor that forces global scope for variable definitions
func NewASPVisitorGlobal(ctx *ExecutionContext, executor *ASPExecutor) *ASPVisitor {
	v := NewASPVisitor(ctx, executor)
	v.forceGlobal = true
	return v
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
	v.context.SetResumeOnError(v.resumeOnError)

	if node == nil {

		return nil

	}

	// Check if execution should stop (cancelled, timed out, or Response.End)
	if v.context.ShouldStop() {
		return fmt.Errorf("RESPONSE_END")
	}

	// DEBUG: Trace statement execution

	// log.Printf("DEBUG: VisitStatement type=%T\n", node)

	v.depth++

	if v.depth > 1000 {
		return fmt.Errorf("maximum call depth exceeded")
	}
	defer func() { v.depth-- }()

	switch stmt := node.(type) {
	case *ast.AssignmentStatement:
		return v.handleStatementError(v.visitAssignment(stmt))
	case *ast.CallStatement:
		// Handle implicit sub calls (e.g., obj.MoveNext without parentheses)
		// When the callee already carries arguments (IndexOrCall/WithMemberAccess),
		// evaluating the expression triggers the call; otherwise force a zero-arg call.
		switch stmt.Callee.(type) {
		case *ast.IndexOrCallExpression, *ast.WithMemberAccessExpression:
			_, err := v.visitExpression(stmt.Callee)
			return v.handleStatementError(err)
		default:
			_, err := v.resolveCall(stmt.Callee, []ast.Expression{})
			return v.handleStatementError(err)
		}
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
		v.context.SetResumeOnError(true)
		v.context.Err.Clear()
		return nil
	case *ast.OnErrorGoTo0Statement:
		v.resumeOnError = false
		v.context.SetResumeOnError(false)
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

	// Debug: log specific variable assignments
	//if ident, ok := stmt.Left.(*ast.Identifier); ok {
	//	nameLower := strings.ToLower(ident.Name)
	//	if nameLower == "plugins" {
	//		fmt.Printf("[DEBUG] visitAssignment: setting plugins = %T (%v)\n", value, value)
	//	}
	//}

	// Handle different left-hand side patterns

	// Case 1: Simple variable assignment (Dim x = 5 or x = 5)
	if ident, ok := stmt.Left.(*ast.Identifier); ok {
		oldValue, hadOldValue := v.context.GetVariable(ident.Name)

		// Check if this is a function return value assignment
		// If we're inside a function and assigning to the function name, use SetVariable (local scope)
		// even if forceGlobal is true, because the function return variable was defined in local scope
		if v.currentFunctionName != "" && strings.EqualFold(ident.Name, v.currentFunctionName) {
			//fmt.Printf("[DEBUG] Function return assignment: %s = %v (type: %T)\n", ident.Name, value, value)
			// Function return value - always set in current scope
			if err := v.context.SetVariable(ident.Name, value); err != nil {
				return err
			}
		} else if v.forceGlobal {
			// Use SetVariableGlobal when forceGlobal is active (ExecuteGlobal context)
			if err := v.context.SetVariableGlobal(ident.Name, value); err != nil {
				return err
			}
		} else {
			if err := v.context.SetVariable(ident.Name, value); err != nil {
				return err
			}
		}

		if hadOldValue && isNothingValue(value) {
			v.context.ReleaseResourceIfUnreferenced(oldValue)
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

			// Handle array index assignment (arr(0) = value or arr(0, 1) = value)
			if arrObj, ok := toVBArray(obj); ok {
				if len(indexCall.Indexes) > 0 {
					if len(indexCall.Indexes) == 1 {
						// Single-dimensional: arr(0) = value
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
					} else {
						// Multi-dimensional: arr(0, 1) = value
						// Navigate through nested VBArrays using all but the last index,
						// then Set on the innermost VBArray with the last index.
						current := arrObj
						for i := 0; i < len(indexCall.Indexes)-1; i++ {
							idxVal, err := v.visitExpression(indexCall.Indexes[i])
							if err != nil {
								return err
							}
							index := toInt(idxVal)
							inner, exists := current.Get(index)
							if !exists {
								return fmt.Errorf("subscript out of range")
							}
							innerArr, ok := toVBArray(inner)
							if !ok {
								return fmt.Errorf("subscript out of range")
							}
							current = innerArr
						}
						// Set the value on the innermost array using the last index
						lastIdxVal, err := v.visitExpression(indexCall.Indexes[len(indexCall.Indexes)-1])
						if err != nil {
							return err
						}
						lastIndex := toInt(lastIdxVal)
						if current.Set(lastIndex, value) {
							return nil
						}
						return fmt.Errorf("subscript out of range")
					}
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

			// Handle objects with non-error SetProperty signature
			if lib, ok := obj.(interface {
				SetProperty(string, interface{})
			}); ok && len(indexCall.Indexes) > 0 {
				key, err := v.visitExpression(indexCall.Indexes[0])
				if err != nil {
					return err
				}
				lib.SetProperty(fmt.Sprintf("%v", key), value)
				return nil
			}

			// Handle Response.Cookies collection assignment
			if cookiesCollection, ok := obj.(*ResponseCookiesCollection); ok {
				if len(indexCall.Indexes) > 0 {
					key, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					cookiesCollection.SetItem(fmt.Sprintf("%v", key), fmt.Sprintf("%v", value))
					return nil
				}
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

			// If it's an object wrapper with non-error SetProperty, set indexed property
			if lib, ok := obj.(interface {
				SetProperty(string, interface{})
			}); ok && len(indexCall.Indexes) > 0 {
				key, err := v.visitExpression(indexCall.Indexes[0])
				if err != nil {
					return err
				}
				lib.SetProperty(fmt.Sprintf("%v", key), value)
				return nil
			}

			// If it's a Response.Cookies collection, set the indexed cookie value
			if cookiesCollection, ok := obj.(*ResponseCookiesCollection); ok {
				if len(indexCall.Indexes) > 0 {
					key, err := v.visitExpression(indexCall.Indexes[0])
					if err != nil {
						return err
					}
					cookiesCollection.SetItem(fmt.Sprintf("%v", key), fmt.Sprintf("%v", value))
					return nil
				}
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

		// Handle ClassInstance first (VBScript user-defined classes)
		if classInst, ok := obj.(*ClassInstance); ok {
			return classInst.SetProperty(propName, value)
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

		// If it's an object with non-error SetProperty signature, set the property
		if lib, ok := obj.(interface {
			SetProperty(string, interface{})
		}); ok {
			lib.SetProperty(propName, value)
			return nil
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

	// DEBUG: Log visitFor entry
	//fmt.Printf("[DEBUG] visitFor: varName=%s\n", varName)

	// Evaluate From and To
	from, err := v.visitExpression(stmt.From)
	if err != nil {
		return err
	}

	to, err := v.visitExpression(stmt.To)
	if err != nil {
		return err
	}

	// DEBUG: Log from/to values
	//fmt.Printf("[DEBUG] visitFor: from=%v, to=%v\n", from, to)

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
			// Check if execution should stop (cancelled, timed out, or Response.End)
			if v.context.ShouldStop() {
				return nil
			}
			_ = v.context.SetVariable(varName, current)

			// Execute body
			if stmt.Body != nil {
				for _, s := range stmt.Body {
					if err := v.VisitStatement(s); err != nil {
						if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
							return nil
						}
						return err
					}
				}
			}

			current += step
		}
	} else if step < 0 {
		for current >= end {
			// Check if execution should stop (cancelled, timed out, or Response.End)
			if v.context.ShouldStop() {
				return nil
			}
			_ = v.context.SetVariable(varName, current)

			// Execute body
			if stmt.Body != nil {
				for _, s := range stmt.Body {
					if err := v.VisitStatement(s); err != nil {
						if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
							return nil
						}
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
			// Check if execution should stop (cancelled, timed out, or Response.End)
			if v.context.ShouldStop() {
				return nil
			}
			// Set loop variable
			_ = v.context.SetVariable(stmt.Identifier.Name, item)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
						return nil
					}
					return err
				}
			}
		}
	case []interface{}:
		// Iterate over array
		for _, item := range col {
			// Check if execution should stop
			if v.context.ShouldStop() {
				return nil
			}
			// Set loop variable
			_ = v.context.SetVariable(stmt.Identifier.Name, item)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
						return nil
					}
					return err
				}
			}
		}
	case [][]interface{}:
		// Iterate over 2D array (e.g., from GetRows)
		for _, row := range col {
			// Check if execution should stop
			if v.context.ShouldStop() {
				return nil
			}
			// Convert row to []interface{} for consistency
			rowInterface := make([]interface{}, len(row))
			copy(rowInterface, row)

			// Set loop variable
			_ = v.context.SetVariable(stmt.Identifier.Name, rowInterface)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
						return nil
					}
					return err
				}
			}
		}
	case map[string]interface{}:
		// Iterate over map (VBScript dictionary)
		for key := range col {
			// Check if execution should stop
			if v.context.ShouldStop() {
				return nil
			}
			// Set loop variable to key
			_ = v.context.SetVariable(stmt.Identifier.Name, key)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
						return nil
					}
					return err
				}
			}
		}
	case *Collection:
		// Iterate over Collection (Request.QueryString, Request.Form, etc.)
		keys := col.GetKeys()
		for _, key := range keys {
			// Check if execution should stop
			if v.context.ShouldStop() {
				return nil
			}
			// Set loop variable to key
			_ = v.context.SetVariable(stmt.Identifier.Name, key)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
						return nil
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
				// Check if execution should stop
				if v.context.ShouldStop() {
					return nil
				}
				// Set loop variable
				_ = v.context.SetVariable(stmt.Identifier.Name, item)

				// Execute loop body
				for _, body := range stmt.Body {
					if err := v.VisitStatement(body); err != nil {
						// Handle Exit For
						if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
							return nil
						}
						return err
					}
				}
			}
		}
	case interface{ Enumerate() ([]interface{}, error) }:
		items, enumErr := col.Enumerate()
		if enumErr != nil {
			return enumErr
		}
		for _, item := range items {
			if v.context.ShouldStop() {
				return nil
			}
			_ = v.context.SetVariable(stmt.Identifier.Name, item)

			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					if exitErr, ok := err.(*LoopExitError); ok && exitErr.LoopType == "for" {
						return nil
					}
					return err
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

	loopIter := 0
	for {
		loopIter++
		// Check if execution should stop (cancelled, timed out, or Response.End)
		if v.context.ShouldStop() {
			return nil
		}

		// Check pre-test condition if needed
		if stmt.TestType == ast.ConditionTestTypePreTest {
			condition, err := v.visitExpression(stmt.Condition)
			if err != nil {
				return err
			}

			// DEBUG: Log condition evaluation
			if loopIter <= 5 || loopIter%1000 == 0 {
				//fmt.Printf("[DEBUG] visitDo PreTest iter=%d: condition=%v (%T), LoopType=%v\n", loopIter, condition, condition, stmt.LoopType)
			}

			// Handle loop type (While vs Until)
			shouldContinue := isTruthy(condition)
			if stmt.LoopType == ast.LoopTypeUntil {
				shouldContinue = !shouldContinue
			}

			if loopIter <= 5 || loopIter%1000 == 0 {
				//fmt.Printf("[DEBUG] visitDo PreTest iter=%d: isTruthy=%v, shouldContinue=%v\n", loopIter, isTruthy(condition), shouldContinue)
			}

			if !shouldContinue {
				break
			}
		}

		// Execute loop body
		for _, body := range stmt.Body {
			// Check if execution should stop
			if v.context.ShouldStop() {
				return nil
			}
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
		// Check if execution should stop (cancelled, timed out, or Response.End)
		if v.context.ShouldStop() {
			return nil
		}

		condition, err := v.visitExpression(stmt.Condition)
		if err != nil {
			return err
		}

		if !isTruthy(condition) {
			break
		}

		// Execute loop body
		for _, body := range stmt.Body {
			// Check if execution should stop
			if v.context.ShouldStop() {
				return nil
			}
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
				dims := []int{}
				if len(field.ArrayDims) > 0 {
					dims = make([]int, len(field.ArrayDims))
					for i, dimExpr := range field.ArrayDims {
						dimVal, err := v.visitExpression(dimExpr)
						if err != nil {
							return err
						}
						dims[i] = toInt(dimVal)
					}
				}

				classDef.Variables[strings.ToLower(field.Identifier.Name)] = ClassMemberVar{
					Name:       field.Identifier.Name,
					Visibility: vis,
					Dims:       dims,
					IsDynamic:  field.IsDynamicArray,
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
		case *ast.InitializeSubDeclaration:
			nameLower := strings.ToLower(m.Identifier.Name)
			// Treat as regular sub
			if m.AccessModifier == ast.MethodAccessModifierPrivate {
				classDef.PrivateMethods[nameLower] = &m.SubDeclaration
			} else {
				classDef.Methods[nameLower] = &m.SubDeclaration
			}
		case *ast.TerminateSubDeclaration:
			nameLower := strings.ToLower(m.Identifier.Name)
			// Treat as regular sub
			if m.AccessModifier == ast.MethodAccessModifierPrivate {
				classDef.PrivateMethods[nameLower] = &m.SubDeclaration
			} else {
				classDef.Methods[nameLower] = &m.SubDeclaration
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
// In VBScript, Dim statements are hoisted (processed before code execution) at both
// module level AND inside functions/subs. At any scope, if a variable already exists
// (from hoisting or prior assignment), Dim should NOT reset it.
// The hoistDimDeclarations helper pre-scans function bodies to achieve this.
func (v *ASPVisitor) visitVariableDeclaration(stmt *ast.VariableDeclaration) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	varName := stmt.Identifier.Name

	// Helper to define variable in the appropriate scope
	defineVar := func(name string, value interface{}) error {
		if v.forceGlobal {
			return v.context.DefineVariableGlobal(name, value)
		}
		return v.context.DefineVariable(name, value)
	}

	// At global scope, Dim should not overwrite existing variables.
	// VBScript hoists all module-level Dim declarations before execution, so by the
	// time execution reaches the textual Dim statement, the variable is already
	// allocated. Re-running DefineVariable would reset it to nil and clobber any
	// value that was assigned earlier in the script.
	// VBScript hoists Dim declarations at ALL scopes (module + function/sub),
	// so by the time execution reaches the textual Dim, the variable was already
	// pre-defined by hoistDimDeclarations. Skip re-definition to avoid resetting.
	isGlobalScope := len(v.context.scopeStack) == 0

	// Check whether this variable already exists in the current scope.
	// If it does, the Dim is a no-op (either hoisting defined it or a prior
	// assignment already created it in this scope).
	alreadyInCurrentScope := false
	if isGlobalScope {
		v.context.mu.RLock()
		_, alreadyInCurrentScope = v.context.variables[lowerKey(varName)]
		v.context.mu.RUnlock()
	} else if len(v.context.scopeStack) > 0 {
		v.context.mu.RLock()
		_, alreadyInCurrentScope = v.context.scopeStack[len(v.context.scopeStack)-1][lowerKey(varName)]
		v.context.mu.RUnlock()
	}

	// Check if it's an array declaration
	if len(stmt.ArrayDims) > 0 {
		// Fixed-size array: Dim arr(5) or Dim arr(2,3)
		// Skip if already defined in current scope (hoisting handled it).
		if alreadyInCurrentScope {
			return nil
		}
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
		if err := defineVar(varName, arr); err != nil {
			return err
		}
	} else if stmt.IsDynamicArray {
		// Dynamic array: Dim arr() - initialize as empty array
		if alreadyInCurrentScope {
			return nil
		}
		if err := defineVar(varName, NewVBArray(v.context.OptionBase(), 0)); err != nil {
			return err
		}
	} else {
		// Regular variable
		// Skip if already defined in current scope (hoisted or assigned earlier)
		if alreadyInCurrentScope {
			return nil
		}
		// Initialize to nil (VBScript Empty)
		if err := defineVar(varName, nil); err != nil {
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

// hoistDimDeclarations pre-scans a function/sub body for all Dim statements
// and defines the variables in the current scope. This implements VBScript's
// variable hoisting: Dim declarations take effect from the start of the
// function regardless of where they appear in the source code.
// Only simple variables (no array dimensions) are hoisted; arrays with
// computed dimensions are left to be handled inline.
func (v *ASPVisitor) hoistDimDeclarations(body ast.Node) {
	if body == nil {
		return
	}
	var stmts []ast.Node
	switch b := body.(type) {
	case *ast.StatementList:
		for _, s := range b.Statements {
			stmts = append(stmts, s)
		}
	default:
		stmts = []ast.Node{body}
	}
	v.collectAndDefineDimVars(stmts)
}

// collectAndDefineDimVars walks a flat list of statements and defines any Dim
// variables it finds. It recurses into If/ElseIf/Else, For, While, Do, With,
// Select blocks to find Dim statements at any nesting level.
func (v *ASPVisitor) collectAndDefineDimVars(stmts []ast.Node) {
	for _, stmt := range stmts {
		if stmt == nil {
			continue
		}
		switch s := stmt.(type) {
		case *ast.VariableDeclaration:
			// Simple Dim: define as nil (Empty) in current scope
			if s.Identifier != nil && len(s.ArrayDims) == 0 && !s.IsDynamicArray {
				_ = v.context.DefineVariable(s.Identifier.Name, nil)
			}
		case *ast.VariablesDeclaration:
			for _, decl := range s.Variables {
				if decl != nil && decl.Identifier != nil && len(decl.ArrayDims) == 0 && !decl.IsDynamicArray {
					_ = v.context.DefineVariable(decl.Identifier.Name, nil)
				}
			}

		// Recurse into compound statements to find nested Dim declarations
		case *ast.IfStatement:
			if s.Consequent != nil {
				v.hoistDimDeclarations(s.Consequent)
			}
			if s.Alternate != nil {
				v.hoistDimDeclarations(s.Alternate)
			}
		case *ast.ElseIfStatement:
			if s.Consequent != nil {
				v.hoistDimDeclarations(s.Consequent)
			}
			if s.Alternate != nil {
				v.hoistDimDeclarations(s.Alternate)
			}
		case *ast.ForStatement:
			v.hoistStatementsSlice(s.Body)
		case *ast.ForEachStatement:
			v.hoistStatementsSlice(s.Body)
		case *ast.WhileStatement:
			v.hoistStatementsSlice(s.Body)
		case *ast.DoStatement:
			v.hoistStatementsSlice(s.Body)
		case *ast.WithStatement:
			v.hoistStatementsSlice(s.Body)
		case *ast.SelectStatement:
			for _, c := range s.Cases {
				if c != nil {
					v.hoistStatementsSlice(c.Body)
				}
			}
		}
	}
}

// hoistStatementsSlice converts []ast.Statement to []ast.Node and calls collectAndDefineDimVars
func (v *ASPVisitor) hoistStatementsSlice(stmts []ast.Statement) {
	nodes := make([]ast.Node, len(stmts))
	for i, s := range stmts {
		nodes[i] = s
	}
	v.collectAndDefineDimVars(nodes)
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

		// Intrinsic ASP objects are available when no scoped variable shadows them.
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

		// Check if we're inside a class and this identifier is a class method/function
		// This allows calling class methods without "Me." prefix (e.g., calling "dict" from within cls_asplite)
		if ctxObj := v.context.GetContextObject(); ctxObj != nil {
			if classInst, ok := ctxObj.(*ClassInstance); ok {
				nameLower := strings.ToLower(varName)
				// Check if this is a class function (returns value)
				if _, exists := classInst.ClassDef.Functions[nameLower]; exists {
					// Call the class function with no arguments
					return v.callClassMethodWithRefs(classInst, varName, []ast.Expression{})
				}
				// Check if this is a class method (sub)
				if _, exists := classInst.ClassDef.Methods[nameLower]; exists {
					// Call the class method with no arguments
					return v.callClassMethodWithRefs(classInst, varName, []ast.Expression{})
				}
				// Check if this is a private method/function
				if _, exists := classInst.ClassDef.PrivateMethods[nameLower]; exists {
					// Call the private method with no arguments
					return v.callClassMethodWithRefs(classInst, varName, []ast.Expression{})
				}
			}
		}

		// Check for built-in functions (parameterless call)
		if val, handled := EvalBuiltInFunction(varName, []interface{}{}, v.context); handled {
			return val, nil
		}

		// Undefined variable returns Empty in VBScript
		return EmptyValue{}, nil

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
		// Handle nil/Nothing case
		// In VBScript: Not Empty = True (Empty is 0, Not 0 = -1 = True)
		// Not Null = Null
		// Not Nothing should error with 91 (but if we got here, the error was already captured)
		//
		// If we reach here with nil, it means either:
		// 1. The value is truly Empty (should be treated as 0)
		// 2. An error occurred and was captured by On Error Resume Next
		//
		// The correct behavior for Empty/nil in VBScript is to treat as 0,
		// so Not 0 = -1 (True in VBScript)
		if operand == nil {
			// Treat nil as Empty (0), so Not 0 = -1 = True
			// But this can cause infinite loops in "Do While Not rs.eof" when rs is Nothing
			// The real fix is to error on property access of Nothing (done in visitMemberExpression)
			return int(^int32(0)), nil // This is -1 in VBScript, which is True
		}
		// Check for NothingValue type
		if _, isNothing := operand.(NothingValue); isNothing {
			// Not Nothing should error, but if we got here, return -1 (True) as per Empty behavior
			return int(^int32(0)), nil
		}
		// Check for boolean type to preserve it
		if b, ok := operand.(bool); ok {
			return !b, nil
		}
		// In VBScript, Not works as bitwise operator (invert all bits)
		// Truncate to int32 before complementing to match VBScript 32-bit semantics
		operandInt := int32(toFloat(operand))
		return int(^operandInt), nil
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

	// Debug: log what we're resolving when escape is involved
	// if member, ok := objectExpr.(*ast.MemberExpression); ok {
	// 	if member.Property != nil && strings.ToLower(member.Property.Name) == "escape" {
	// 		//fmt.Printf("[DEBUG] resolveCall FOR ESCAPE: member.Object=%T, argCount=%d\n", member.Object, len(arguments))
	// 		for i, arg := range arguments {
	// 			_ = i
	// 			_ = arg
	// 			//fmt.Printf("[DEBUG] resolveCall FOR ESCAPE: arguments[%d] type=%T\n", i, arg)
	// 		}
	// 	}
	// }

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
			if result, handled := EvalBuiltInFunction(ident.Name, args, v.context); handled {
				// Check if Response.End() was called during execution (e.g., inside ExecuteGlobal)
				if v.context.Response != nil && v.context.Response.IsEnded() {
					return result, fmt.Errorf("RESPONSE_END")
				}
				return result, nil
			}
		}
	}

	// CRITICAL: Handle member-method dispatch BEFORE evaluating base for MemberExpression
	// This prevents visitMemberExpression from calling the method with 0 args when
	// we actually want to call it with the provided arguments.
	if member, ok := objectExpr.(*ast.MemberExpression); ok {
		propName := ""
		if member.Property != nil {
			propName = member.Property.Name
		}
		//fmt.Printf("[DEBUG] resolveCall MemberExpression BEFORE eval: member.Object=%T, propName=%s\n", member.Object, propName)
		// Check if member.Object is itself a MemberExpression (nested case like aspl.json.escape)
		// if innerMember, isInner := member.Object.(*ast.MemberExpression); isInner {
		// 	innerProp := ""
		// 	if innerMember.Property != nil {
		// 		innerProp = innerMember.Property.Name
		// 	}
		// 	_ = innerProp
		// 	//fmt.Printf("[DEBUG] resolveCall NESTED MemberExpression: inner propName=%s, outer propName=%s\n", innerProp, propName)
		// }
		// Evaluate just the parent object, not the full member expression
		parentObj, err := v.visitExpression(member.Object)
		if propName == "json" || propName == "toJson" || propName == "escape" {
			//fmt.Printf("[DEBUG] resolveCall: member.Object evaluated for %s, parentObj=%T\n", propName, parentObj)
		}
		if err != nil {
			return nil, err
		}

		if parentObj != nil {
			propName := ""
			if member.Property != nil {
				propName = member.Property.Name
			}
			// Debug: log all member expressions on built-in objects
			//if _, isReq := parentObj.(*RequestObject); isReq {
			//	log.Printf("[DEBUG] MemberExpression on RequestObject: property=%q, args=%d\n", propName, len(arguments))
			//}
			// Log only for debugging specific methods
			// log.Printf("[DEBUG] resolveCall: MemberExpression parentObj=%T propName=%q\n", parentObj, propName)

			// Class instance methods - PASS EXPRESSIONS for ByRef support
			if classInst, ok := parentObj.(*ClassInstance); ok {
				//fmt.Printf("[DEBUG] resolveCall: calling callClassMethodWithRefs for %s.%s\n", classInst.ClassDef.Name, propName)
				return v.callClassMethodWithRefs(classInst, propName, arguments)
			}

			// Special handling for Request.BinaryRead with ByRef parameter
			// In Classic ASP, BinaryRead(count) modifies count to reflect bytes actually read
			if reqObj, ok := parentObj.(*RequestObject); ok {
				propNameLower := strings.ToLower(propName)
				if propNameLower == "binaryread" && len(arguments) > 0 {
					// Get the count value
					count := int64(0)
					if len(args) > 0 {
						switch c := args[0].(type) {
						case int:
							count = int64(c)
						case int32:
							count = int64(c)
						case int64:
							count = c
						case float64:
							count = int64(c)
						}
					}
					////fmt.Printf("[DEBUG] BinaryRead ByRef: count requested=%d\n", count)

					// Call BinaryRead
					data, err := reqObj.BinaryRead(count)
					if err != nil {
						return nil, err
					}

					// Update the ByRef parameter with actual bytes read
					actualBytesRead := int64(len(data))
					////fmt.Printf("[DEBUG] BinaryRead ByRef: actual bytes read=%d\n", actualBytesRead)
					if ident, ok := arguments[0].(*ast.Identifier); ok {
						// Update the variable with actual bytes read
						////fmt.Printf("[DEBUG] BinaryRead ByRef: updating variable '%s' to %d\n", ident.Name, actualBytesRead)
						v.context.SetVariable(ident.Name, actualBytesRead)
					} else {
						////fmt.Printf("[DEBUG] BinaryRead ByRef: argument is NOT an Identifier, type=%T\n", arguments[0])
					}

					return data, nil
				}
			}

			// ASPObject / server-native CallMethod dispatch (no ByRef support needed)
			if aspObj, ok := parentObj.(asp.ASPObject); ok {
				return aspObj.CallMethod(propName, args...)
			}
			if caller, ok := parentObj.(interface {
				CallMethod(string, ...interface{}) (interface{}, error)
			}); ok {
				return caller.CallMethod(propName, args...)
			}
			if caller, ok := parentObj.(interface {
				CallMethod(string, ...interface{}) interface{}
			}); ok {
				return caller.CallMethod(propName, args...), nil
			}
		}
	}

	// Handle WithMemberAccessExpression (e.g. .Open "..." inside With block)
	if withMember, ok := objectExpr.(*ast.WithMemberAccessExpression); ok {
		if len(v.withStack) > 0 {
			parentObj := v.withStack[len(v.withStack)-1]
			propName := withMember.Property.Name

			// Class instance methods
			if classInst, ok := parentObj.(*ClassInstance); ok {
				return v.callClassMethodWithRefs(classInst, propName, arguments)
			}

			// ASPObject dispatch
			if aspObj, ok := parentObj.(asp.ASPObject); ok {
				return aspObj.CallMethod(propName, args...)
			}
			if caller, ok := parentObj.(interface {
				CallMethod(string, ...interface{}) (interface{}, error)
			}); ok {
				return caller.CallMethod(propName, args...)
			}
			if caller, ok := parentObj.(interface {
				CallMethod(string, ...interface{}) interface{}
			}); ok {
				return caller.CallMethod(propName, args...), nil
			}
		}
	}

	// 1. Check if objectExpr is an Identifier referring to a User Function/Sub (Manual Lookup)
	if ident, ok := objectExpr.(*ast.Identifier); ok {
		localVal, hasVariableWithSameName := v.context.GetVariable(ident.Name)
		localValueIsCallable := false
		if hasVariableWithSameName {
			if _, ok := localVal.(asp.ASPObject); ok {
				localValueIsCallable = true
			} else if _, ok := localVal.(interface {
				CallMethod(string, ...interface{}) (interface{}, error)
			}); ok {
				localValueIsCallable = true
			} else if _, ok := localVal.(interface {
				CallMethod(string, ...interface{}) interface{}
			}); ok {
				localValueIsCallable = true
			}
		}

		// First check if we're inside a class and this is a class method call.
		// In VBScript, recursive calls from inside a Function must still resolve to the
		// method itself even though the function return variable has the same name.
		if ctxObj := v.context.GetContextObject(); ctxObj != nil && (!hasVariableWithSameName || !localValueIsCallable) {
			if classInst, ok := ctxObj.(*ClassInstance); ok {
				nameLower := strings.ToLower(ident.Name)
				// Check if this identifier refers to a class method
				if _, exists := classInst.ClassDef.Methods[nameLower]; exists {
					// Call the class method with ByRef support
					return v.callClassMethodWithRefs(classInst, ident.Name, arguments)
				}
				if _, exists := classInst.ClassDef.Functions[nameLower]; exists {
					// Call the class function with ByRef support
					return v.callClassMethodWithRefs(classInst, ident.Name, arguments)
				}
				if _, exists := classInst.ClassDef.PrivateMethods[nameLower]; exists {
					// Call the private method with ByRef support
					return v.callClassMethodWithRefs(classInst, ident.Name, arguments)
				}
				// Check Property Get with arguments (e.g., getClickLink(false))
				// callClassMethodWithRefs will fall through to CallMethod which handles Properties
				if props, ok := classInst.ClassDef.Properties[nameLower]; ok {
					for _, p := range props {
						if p.Type == PropGet {
							return v.callClassMethodWithRefs(classInst, ident.Name, arguments)
						}
					}
				}
			}
		}

		// Regular function/sub lookup
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

		// Fallback: if local/parent symbol shadows a function name with a non-callable value,
		// check global scope for function/sub declarations.
		if globalVal, exists := v.context.GetGlobalVariable(ident.Name); exists {
			if fn, ok := globalVal.(*ast.FunctionDeclaration); ok {
				return v.executeFunctionWithRefs(fn, arguments, v)
			}
			if sub, ok := globalVal.(*ast.SubDeclaration); ok {
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
			// 2. Check Variable first
			if val, exists := v.context.GetVariable(ident.Name); exists {
				base = val
			}

			// 3. Fallback to intrinsic built-in objects
			if base == nil {
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
				if result, handled := EvalBuiltInFunction(funcName, args, v.context); handled {
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

		// Finally, try CallMethod without error return
		if obj, ok := base.(interface {
			CallMethod(string, ...interface{}) interface{}
		}); ok {
			return obj.CallMethod(methodName, args...), nil
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
				if caller, ok := parentObj.(interface {
					CallMethod(string, ...interface{}) interface{}
				}); ok {
					return caller.CallMethod(propName, args...), nil
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
						log.Printf("[DEBUG] BinaryRead called with %d args\n", len(args))
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

				// Try SessionObject methods
				if sessionObj, ok := parentObj.(*SessionObject); ok {
					propNameLower := strings.ToLower(propName)
					switch propNameLower {
					case "abandon":
						sessionObj.Abandon()
						v.context.MarkSessionAbandoned()
						return nil, nil
					case "removeall":
						sessionObj.RemoveAll()
						return nil, nil
					case "lock":
						sessionObj.Lock()
						return nil, nil
					case "unlock":
						sessionObj.Unlock()
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

	// Handle array access (including multi-dimensional arrays)
	if arr, ok := toVBArray(base); ok && len(args) > 0 {
		//fmt.Printf("[DEBUG] resolveCall: returning from array access\n")
		// For single-dimensional access: arr(0)
		if len(args) == 1 {
			idx := toInt(args[0])
			if val, ok := arr.Get(idx); ok {
				return val, nil
			}
			return nil, nil
		}
		// For multi-dimensional access: arr(0, 1) or arr(0, 1, 2), etc.
		// Navigate through dimensions
		current := interface{}(arr)
		for i := 0; i < len(args); i++ {
			idx := toInt(args[i])
			if currArr, ok := toVBArray(current); ok {
				if val, exists := currArr.Get(idx); exists {
					current = val
				} else {
					return nil, nil // Out of bounds
				}
			} else {
				// current is not an array, can't index further
				return nil, nil
			}
		}
		return current, nil
	}

	// Handle objects with CallMethod interface - subscript access via default property (typically "Item")
	// This handles OLE objects like ADODBOLEFields when accessed as flds("iId")
	if obj, ok := base.(interface {
		CallMethod(string, ...interface{}) (interface{}, error)
	}); ok && len(args) > 0 {
		// Call the default property method "Item" with the subscript argument
		result, _ := obj.CallMethod("item", args...)
		return result, nil
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

	// Debug: trace final nil return
	// if ident, ok := objectExpr.(*ast.Identifier); ok {
	// 	_ = ident
	// 	//fmt.Printf("[DEBUG] resolveCall: returning nil, nil for Identifier %s (base was %T)\n", ident.Name, base)
	// } else {
	// 	//fmt.Printf("[DEBUG] resolveCall: returning nil, nil (objectExpr type=%T, base type=%T)\n", objectExpr, base)
	// }

	return nil, nil
}

// visitMemberExpression evaluates member access (obj.property)
func (v *ASPVisitor) visitMemberExpression(expr *ast.MemberExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	// Get property name first for debugging
	propName := ""
	if expr.Property != nil {
		propName = expr.Property.Name
	}

	// Debug logging for json property (commented out for performance)
	// if propName == "json" {
	// 	//fmt.Printf("[DEBUG] visitMemberExpression ENTER for prop=%s, expr.Object=%T\n", propName, expr.Object)
	// 	if ident, ok := expr.Object.(*ast.Identifier); ok {
	// 		_ = ident
	// 		//fmt.Printf("[DEBUG] vME: expr.Object is Identifier: %s\n", ident.Name)
	// 	} else if member, ok := expr.Object.(*ast.MemberExpression); ok {
	// 		innerProp := ""
	// 		if member.Property != nil {
	// 			innerProp = member.Property.Name
	// 		}
	// 		_ = innerProp
	// 		//fmt.Printf("[DEBUG] vME: expr.Object is MemberExpression with prop=%s\n", innerProp)
	// 	}
	// 	// Print mini stack trace
	// 	for i := 0; i < 8; i++ {
	// 		_, file, line, ok := runtime.Caller(i)
	// 		if !ok {
	// 			break
	// 		}
	// 		_ = line
	// 		if strings.Contains(file, "axonasp") {
	// 			//fmt.Printf("[DEBUG] vME entry caller[%d]: %s:%d\n", i, file, line)
	// 		}
	// 	}
	// }

	// Evaluate object
	obj, err := v.visitExpression(expr.Object)
	if err != nil {
		return nil, err
	}

	// Check if we're accessing a method/property on Nothing
	// In VBScript, this should raise Error 91: "Object variable or With block variable not set"
	// With On Error Resume Next, this error is captured and the value becomes Empty
	if obj == nil || isNothingValue(obj) || isEmptyValue(obj) {
		// Generate error 91 - this will be captured by On Error Resume Next if active
		// Include object identifier for debugging
		objIdentifier := "<unknown>"
		if idExpr, ok := expr.Object.(*ast.Identifier); ok {
			objIdentifier = idExpr.Name
		}
		errMsg := fmt.Sprintf("object variable not set when accessing '%s' on '%s' (Error 91)", propName, objIdentifier)

		// VBScript semantics: with On Error Resume Next enabled, convert runtime member access
		// failures into Empty and continue execution while updating Err object state.
		if v.context.IsResumeOnError() {
			v.context.Err.SetError(fmt.Errorf("%s", errMsg))
			return EmptyValue{}, nil
		}

		// Return the error so it can be handled properly
		return nil, fmt.Errorf("%s", errMsg)
	}

	if propName == "json" {
		//fmt.Printf("[DEBUG] visitMemberExpression for prop=%s, obj type=%T\n", propName, obj)
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

	// Handle RequestObject properties and methods
	// Must be before generic asp.ASPObject to handle methods like BinaryRead correctly
	if reqObj, ok := obj.(*RequestObject); ok {
		propNameLower := strings.ToLower(propName)
		//log.Printf("[DEBUG] visitMemberAccessExpression: RequestObject property=%q\n", propName)
		switch propNameLower {
		case "querystring":
			return reqObj.QueryString, nil
		case "form":
			return reqObj.Form, nil
		case "cookies":
			return reqObj.Cookies, nil
		case "servervariables":
			return reqObj.ServerVariables, nil
		case "clientcertificate":
			return reqObj.ClientCertificate, nil
		case "totalbytes":
			return reqObj.GetTotalBytes(), nil
		case "binaryread":
			// BinaryRead is a method, not a property - return a callable wrapper
			// This allows Request.BinaryRead(count) to work when parsed as member+call
			//log.Println("[DEBUG] visitMemberAccessExpression: Returning RequestBinaryReadWrapper")
			return &RequestBinaryReadWrapper{reqObj: reqObj}, nil
		}
		// For other properties, use generic GetProperty
		return reqObj.GetProperty(propName), nil
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
		propNameLower := strings.ToLower(propName)
		switch propNameLower {
		case "contents":
			return NewSessionContentsCollection(sessionObj), nil
		}
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
			return NewApplicationContentsCollection(appObj), nil
		}
	}

	// Handle ClassInstance
	if classInst, ok := obj.(*ClassInstance); ok {
		if propName == "json" {
			ctxObj := v.context.GetContextObject()
			_ = ctxObj // silence unused variable
			//fmt.Printf("[DEBUG] visitMemberExpression(3518) for json on %s (ptr=%p), ctxObj=%T (ptr=%p)\n",
			//	classInst.ClassDef.Name, classInst, ctxObj, ctxObj)
		}
		// Check for internal access (Me) to allow Private members/variables
		if ctxObj := v.context.GetContextObject(); ctxObj != nil {
			if currentInst, ok := ctxObj.(*ClassInstance); ok && currentInst == classInst {
				if propName == "json" {
					//fmt.Printf("[DEBUG] visitMemberExpression(3518) for json - INTERNAL ACCESS, using GetMember\n")
				}
				// Internal access: use GetMember to allow private members
				val, found, err := classInst.GetMember(propName)
				if err != nil {
					return nil, err
				}
				if found {
					// If it's an AST node (method definition), we treat it as not a value property
					// unless it's a Property Get (which GetMember executes and returns value for)
					if _, isNode := val.(ast.Node); !isNode {
						return val, nil
					}
				}
			}
		}

		val, err := classInst.GetProperty(propName)
		if err != nil {
			return nil, err
		}
		if val != nil {
			return val, nil
		}
		// Fallback to method call (parameterless)
		if propName == "json" {
			//fmt.Printf("[DEBUG] visitMemberExpression(3518) falling through to CallMethod for json\n")
		}
		result, err := classInst.CallMethod(propName)
		if propName == "json" {
			//fmt.Printf("[DEBUG] visitMemberExpression(3518) CallMethod returned for json, result=%T, err=%v\n", result, err)
		}
		return result, err
	}

	// Handle generic property access
	if aspObj, ok := obj.(asp.ASPObject); ok {
		return aspObj.GetProperty(propName), nil
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
		val := getter.GetProperty(propName)
		if val != nil {
			return val, nil
		}
	}

	// Fallback: if property not found, try zero-arg method call for objects that expose CallMethod
	if caller, ok := obj.(interface {
		CallMethod(string, ...interface{}) (interface{}, error)
	}); ok {
		return caller.CallMethod(propName)
	}
	if caller, ok := obj.(interface {
		CallMethod(string, ...interface{}) interface{}
	}); ok {
		return caller.CallMethod(propName), nil
	}

	return nil, nil
}

// Helper functions

// isTruthy checks if a value is truthy in VBScript
func isTruthy(val interface{}) bool {
	if val == nil {
		return false
	}
	if _, ok := val.(EmptyValue); ok {
		return false
	}
	if _, ok := val.(NullValue); ok {
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

func isEmptyValue(val interface{}) bool {
	if val == nil {
		return false
	}
	_, ok := val.(EmptyValue)
	return ok
}

func isNullValue(val interface{}) bool {
	if val == nil {
		return false
	}
	_, ok := val.(NullValue)
	return ok
}

// toString converts a value to string
func toString(val interface{}) string {
	if val == nil {
		return ""
	}
	val = normalizeOLEValue(val)
	if aspObj, ok := val.(asp.ASPObject); ok {
		if inner := aspObj.GetProperty("value"); inner != nil && inner != val {
			return toString(inner)
		}
	}
	if objWithErrProperty, ok := val.(interface {
		GetProperty(string) (interface{}, error)
	}); ok {
		if inner, err := objWithErrProperty.GetProperty("value"); err == nil && inner != nil && inner != val {
			return toString(inner)
		}
	}
	switch v := val.(type) {
	case EmptyValue:
		return ""
	case NothingValue:
		return ""
	case string:
		return v
	case []byte:
		if decoded, ok := decodeLikelyUTF16LE(v); ok {
			return normalizeOLEString(decoded)
		}
		// Binary data from BinaryRead - convert to string preserving each byte as a rune
		// This is similar to Latin-1 encoding where each byte maps directly to a Unicode code point
		runes := make([]rune, len(v))
		for i, b := range v {
			runes[i] = rune(b)
		}
		return string(runes)
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
	val = normalizeOLEValue(val)
	if aspObj, ok := val.(asp.ASPObject); ok {
		if inner := aspObj.GetProperty("value"); inner != nil && inner != val {
			return toInt(inner)
		}
	}
	if objWithErrProperty, ok := val.(interface {
		GetProperty(string) (interface{}, error)
	}); ok {
		if inner, err := objWithErrProperty.GetProperty("value"); err == nil && inner != nil && inner != val {
			return toInt(inner)
		}
	}
	switch v := val.(type) {
	case EmptyValue, NothingValue:
		return 0
	case int:
		return v
	case int8:
		return int(v)
	case int16:
		return int(v)
	case int32:
		return int(v)
	case int64:
		return int(v)
	case uint:
		return int(v)
	case uint8:
		return int(v)
	case uint16:
		return int(v)
	case uint32:
		return int(v)
	case uint64:
		return int(v)
	case float32:
		return int(v)
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
	val = normalizeOLEValue(val)
	if aspObj, ok := val.(asp.ASPObject); ok {
		if inner := aspObj.GetProperty("value"); inner != nil && inner != val {
			return toFloat(inner)
		}
	}
	if objWithErrProperty, ok := val.(interface {
		GetProperty(string) (interface{}, error)
	}); ok {
		if inner, err := objWithErrProperty.GetProperty("value"); err == nil && inner != nil && inner != val {
			return toFloat(inner)
		}
	}
	switch v := val.(type) {
	case EmptyValue, NothingValue:
		return 0
	case int:
		return float64(v)
	case int8:
		return float64(v)
	case int16:
		return float64(v)
	case int32:
		return float64(v)
	case int64:
		return float64(v)
	case uint:
		return float64(v)
	case uint8:
		return float64(v)
	case uint16:
		return float64(v)
	case uint32:
		return float64(v)
	case uint64:
		return float64(v)
	case float32:
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
	// VBScript: if either operand is Null, most operations yield Null (no runtime error)
	if op != ast.BinaryOperationIs {
		if isNullValue(left) || isNullValue(right) {
			return NullValue{}, nil
		}
	}

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
		// Enforce 32-bit signed integer semantics
		leftInt := int32(toFloat(left))
		rightInt := int32(toFloat(right))
		return int(leftInt & rightInt), nil
	case ast.BinaryOperationOr:
		// Check for boolean preservation
		if b1, ok1 := left.(bool); ok1 {
			if b2, ok2 := right.(bool); ok2 {
				return b1 || b2, nil
			}
		}
		// In VBScript, Or works as bitwise operator
		// Enforce 32-bit signed integer semantics
		leftInt := int32(toFloat(left))
		rightInt := int32(toFloat(right))
		return int(leftInt | rightInt), nil
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
		// VBScript IntDiv: Round both operands to Long first (banker's rounding), then divide
		leftNum := int32(math.Round(toFloat(left)))
		rightNum := int32(math.Round(toFloat(right)))
		if rightNum == 0 {
			return nil, fmt.Errorf("division by zero")
		}
		return int(leftNum / rightNum), nil
	case ast.BinaryOperationMod:
		// VBScript Mod: Round both operands to Long first (banker's rounding), then mod
		leftNum := int32(math.Round(toFloat(left)))
		rightNum := int32(math.Round(toFloat(right)))
		if rightNum == 0 {
			return nil, fmt.Errorf("division by zero")
		}
		return int(leftNum % rightNum), nil
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
		// For binary data ([]byte), concatenate bytes directly
		leftBytes, leftIsBinary := left.([]byte)
		rightBytes, rightIsBinary := right.([]byte)
		if leftIsBinary || rightIsBinary {
			// At least one operand is binary, so treat both as binary
			if !leftIsBinary {
				leftBytes = []byte(toString(left))
			}
			if !rightIsBinary {
				rightBytes = []byte(toString(right))
			}
			result := make([]byte, len(leftBytes)+len(rightBytes))
			copy(result, leftBytes)
			copy(result[len(leftBytes):], rightBytes)
			return result, nil
		}
		// Normal string concatenation
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
	case ast.BinaryOperationXor:
		// Boolean XOR
		if b1, ok1 := left.(bool); ok1 {
			if b2, ok2 := right.(bool); ok2 {
				return b1 != b2, nil
			}
		}
		// Enforce 32-bit semantics
		l := int32(toFloat(left))
		r := int32(toFloat(right))
		return int(l ^ r), nil
	case ast.BinaryOperationEqv:
		// Boolean equivalence (Not XOR)
		if b1, ok1 := left.(bool); ok1 {
			if b2, ok2 := right.(bool); ok2 {
				return b1 == b2, nil
			}
		}
		// Enforce 32-bit semantics
		l := int32(toFloat(left))
		r := int32(toFloat(right))
		return int(^(l ^ r)), nil
	case ast.BinaryOperationImp:
		// Boolean implication: (Not A) Or B
		if b1, ok1 := left.(bool); ok1 {
			if b2, ok2 := right.(bool); ok2 {
				return (!b1) || b2, nil
			}
		}
		// Enforce 32-bit semantics
		l := int32(toFloat(left))
		r := int32(toFloat(right))
		return int((^l) | r), nil
	default:
		return nil, fmt.Errorf("unknown binary operator: %d", op)
	}
}

// compareEqual compares two values for equality with VBScript rules
func compareEqual(left, right interface{}, mode ast.OptionCompareMode) bool {
	// Normalize nil to EmptyValue for consistent comparison
	// In VBScript, Empty = "" → True, Empty = 0 → True, Empty = False → True
	if left == nil {
		left = EmptyValue{}
	}
	if right == nil {
		right = EmptyValue{}
	}

	// Both Empty
	_, leftIsEmpty := left.(EmptyValue)
	_, rightIsEmpty := right.(EmptyValue)
	if leftIsEmpty && rightIsEmpty {
		return true
	}

	// Empty compared to another value — coerce Empty to the other operand's type
	if leftIsEmpty {
		switch r := right.(type) {
		case string:
			return r == ""
		case bool:
			return r == false
		default:
			if rn, rok := toNumeric(right); rok {
				return rn == 0
			}
			return toString(right) == ""
		}
	}
	if rightIsEmpty {
		switch l := left.(type) {
		case string:
			return l == ""
		case bool:
			return l == false
		default:
			if ln, lok := toNumeric(left); lok {
				return ln == 0
			}
			return toString(left) == ""
		}
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
	case EmptyValue:
		return 0, true
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

func shouldTraceAuthFlow() bool {
	trace := strings.TrimSpace(strings.ToLower(os.Getenv("TRACE_LOGIN_FLOW")))
	if trace == "1" || trace == "true" || trace == "yes" || trace == "on" {
		return true
	}
	trace = strings.TrimSpace(strings.ToLower(os.Getenv("TRACE_AUTH_FLOW")))
	return trace == "1" || trace == "true" || trace == "yes" || trace == "on"
}

func traceAuthFlow(ctx *ExecutionContext, format string, args ...interface{}) {
	if !shouldTraceAuthFlow() {
		return
	}
	prefix := "[TRACE_AUTH]"
	if ctx != nil {
		requestPath := ""
		if ctx.httpRequest != nil && ctx.httpRequest.URL != nil {
			requestPath = ctx.httpRequest.URL.Path
		}
		if requestPath != "" {
			prefix = fmt.Sprintf("[TRACE_AUTH sid=%s path=%s]", ctx.sessionID, requestPath)
		} else {
			prefix = fmt.Sprintf("[TRACE_AUTH sid=%s]", ctx.sessionID)
		}
	}
	fmt.Printf(prefix+" "+format+"\n", args...)
}

// populateRequestData fills a RequestObject with data from HTTP request
func populateRequestData(req *RequestObject, r *http.Request, ctx *ExecutionContext) {
	// IMPORTANT: Read and preserve the body BEFORE ParseForm consumes it
	// This is needed because Classic ASP scripts may use Request.BinaryRead to read raw body
	var bodyBytes []byte
	if r.Body != nil && r.ContentLength > 0 {
		bodyBytes, _ = io.ReadAll(r.Body)
		r.Body.Close()
		// Restore body for ParseForm AND for potential BinaryRead later
		r.Body = io.NopCloser(bytes.NewReader(bodyBytes))
	}

	// Set the HTTP request for BinaryRead support
	// Also pre-populate the body buffer so BinaryRead has data to read
	req.SetHTTPRequest(r)
	if len(bodyBytes) > 0 {
		req.PreloadBody(bodyBytes)
	}

	// Parse form data
	// For multipart, we parse to extract text fields but preserve the raw body for BinaryRead
	contentType := r.Header.Get("Content-Type")
	isMultipart := strings.HasPrefix(contentType, "multipart/form-data")

	if isMultipart {
		// For multipart, we need to parse to get form fields but body is already preserved
		// Restore body for parsing
		r.Body = io.NopCloser(bytes.NewReader(bodyBytes))
		err := r.ParseMultipartForm(32 << 20) // 32MB max memory
		if err != nil {
			// Log error but continue - parsing may fail but body is preserved
		}
	} else {
		// For non-multipart, restore body and parse
		r.Body = io.NopCloser(bytes.NewReader(bodyBytes))
		r.ParseForm()
	}

	// Set query string parameters (from URL only)
	for key, values := range r.URL.Query() {
		if len(values) > 0 {
			value := values[0]
			if len(values) > 1 {
				value = strings.Join(values, ", ")
			}
			req.QueryString.Add(key, value)
		}
	}

	// Set form parameters (POST data only)
	// For multipart and URL-encoded forms, r.PostForm contains only body fields.
	// Do not use r.Form here because it also includes query-string values.
	if r.Method == "POST" || r.Method == "PUT" {
		for key, values := range r.PostForm {
			if len(values) > 0 {
				value := values[0]
				if len(values) > 1 {
					value = strings.Join(values, ", ")
				}
				req.Form.Add(key, value)
			}
		}
	}
	// NOTE: The raw body is preserved via PreloadBody() for BinaryRead to work

	traceAuthFlow(ctx,
		"request prepared method=%s pageAction(q=%v,f=%v) iId(q=%v,f=%v) iPostID(q=%v,f=%v)",
		r.Method,
		req.QueryString.Get("pageAction"), req.Form.Get("pageAction"),
		req.QueryString.Get("iId"), req.Form.Get("iId"),
		req.QueryString.Get("iPostID"), req.Form.Get("iPostID"),
	)

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
		propNameLower := strings.ToLower(propName)
		switch propNameLower {
		case "contents":
			return NewSessionContentsCollection(sessionObj), nil
		}
		return sessionObj.GetProperty(propName), nil
	}

	// Handle ClassInstance
	if classInst, ok := obj.(*ClassInstance); ok {
		// Check for internal access (Me) to allow Private members/variables
		if ctxObj := v.context.GetContextObject(); ctxObj != nil {
			if currentInst, ok := ctxObj.(*ClassInstance); ok && currentInst == classInst {
				// Internal access: use GetMember to allow private members
				val, found, err := classInst.GetMember(propName)
				if err != nil {
					return nil, err
				}
				if found {
					// If it's an AST node (method definition), we treat it as not a value property
					// unless it's a Property Get (which GetMember executes and returns value for)
					if _, isNode := val.(ast.Node); !isNode {
						return val, nil
					}
				}
			}
		}

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
		if propName == "json" {
			//fmt.Printf("[DEBUG] Falling through to CallMethod for json - this creates recursion\n")
		}
		result, err := classInst.CallMethod(propName)
		if propName == "json" || propName == "toJson" {
			//fmt.Printf("[DEBUG] visitMemberExpression CallMethod(%s) returned type=%T, err=%v\n", propName, result, err)
		}
		return result, err
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
		//log.Printf("[DEBUG] visitNewExpression: looking for class %q\n", className)

		// Lookup ClassDef
		classDefVal, exists := v.context.GetVariable(className)
		if !exists {
			// Check for built-in COM objects (New RegExp syntax)
			if strings.ToLower(className) == "regexp" {
				return NewRegExpLibrary(v.context), nil
			}
			//log.Printf("[DEBUG] visitNewExpression: class %q not defined\n", className)
			return nil, fmt.Errorf("class not defined: %s", className)
		}

		if classDef, ok := classDefVal.(*ClassDef); ok {
			//log.Printf("[DEBUG] visitNewExpression: creating instance of %q\n", className)
			return NewClassInstance(classDef, v.context)
		}

		// Also support *ast.ClassDeclaration (which is how ExecuteGlobal stores classes)
		if classDecl, ok := classDefVal.(*ast.ClassDeclaration); ok {
			// Create a ClassDef from the declaration
			classDef := v.NewClassDefFromDecl(classDecl)
			return NewClassInstance(classDef, v.context)
		}

		return nil, fmt.Errorf("%s is not a class", className)
	}

	return nil, fmt.Errorf("invalid New expression")
}

// NewClassDefFromDecl converts a ClassDeclaration AST node to a ClassDef runtime object
func (v *ASPVisitor) NewClassDefFromDecl(stmt *ast.ClassDeclaration) *ClassDef {
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
			vis := VisPublic
			if m.Modifier == ast.FieldAccessModifierPrivate {
				vis = VisPrivate
			}
			for _, field := range m.Fields {
				dims := []int{}
				if len(field.ArrayDims) > 0 {
					dims = make([]int, len(field.ArrayDims))
					for i, dimExpr := range field.ArrayDims {
						dimVal, _ := v.visitExpression(dimExpr)
						dims[i] = toInt(dimVal)
					}
				}
				classDef.Variables[strings.ToLower(field.Identifier.Name)] = ClassMemberVar{
					Name:       field.Identifier.Name,
					Visibility: vis,
					Dims:       dims,
					IsDynamic:  field.IsDynamicArray,
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
	return classDef
}

// executeFunction executes a user defined function
func (v *ASPVisitor) executeFunction(fn *ast.FunctionDeclaration, args []interface{}) (interface{}, error) {
	v.context.PushScope()
	defer v.context.PopScope()

	// Track current function name for return value assignment
	prevFunctionName := v.currentFunctionName
	v.currentFunctionName = fn.Identifier.Name
	defer func() { v.currentFunctionName = prevFunctionName }()

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

	// Hoist Dim declarations so variables defined later in the function body
	// are available from the start (VBScript hoisting semantics)
	if fn.Body != nil {
		v.hoistDimDeclarations(fn.Body)
	}

	// Execute body
	if fn.Body != nil {
		if list, ok := fn.Body.(*ast.StatementList); ok {
			for _, stmt := range list.Statements {
				if err := v.VisitStatement(stmt); err != nil {
					if pe, ok := err.(*ProcedureExitError); ok {
						if pe.Kind == "function" {
							break
						}
						return nil, err
					}
					return nil, err
				}
			}
		} else {
			if err := v.VisitStatement(fn.Body); err != nil {
				if pe, ok := err.(*ProcedureExitError); ok {
					if pe.Kind == "function" {
						goto functionDone
					}
					return nil, err
				}
				return nil, err
			}
		}
	}

functionDone:

	// Get return value
	val, _ := v.context.GetVariable(funcName)
	return val, nil
}

// executeFunctionWithRefs executes a user defined function with ByRef parameter support
func (v *ASPVisitor) executeFunctionWithRefs(fn *ast.FunctionDeclaration, arguments []ast.Expression, visitor *ASPVisitor) (interface{}, error) {
	v.context.PushScope()
	defer v.context.PopScope()

	// Track current function name for return value assignment
	prevFunctionName := v.currentFunctionName
	v.currentFunctionName = fn.Identifier.Name
	defer func() { v.currentFunctionName = prevFunctionName }()

	// Map to track ByRef parameters and their original variable names
	byRefMap := make(map[string]string) // param name -> original var name

	// Bind parameters
	for i, param := range fn.Parameters {
		var val interface{}

		if i < len(arguments) {
			// VBScript default: parameters are ByRef unless explicitly ByVal
			// Check if parameter is ByVal (explicitly)
			if param.Modifier == ast.ParameterModifierByVal {
				// ByVal: evaluate the argument
				var err error
				val, err = visitor.visitExpression(arguments[i])
				if err != nil {
					return nil, err
				}
			} else {
				// ByRef (explicit or default): we need the original variable name
				if ident, ok := arguments[i].(*ast.Identifier); ok {
					byRefMap[strings.ToLower(param.Identifier.Name)] = strings.ToLower(ident.Name)
					// Handle "Me" keyword - not stored as regular variable
					if strings.EqualFold(ident.Name, "me") {
						if ctxObj := v.context.GetContextObject(); ctxObj != nil {
							val = ctxObj
						}
					} else if parentVal, exists := v.context.GetVariableFromParentScope(ident.Name); exists {
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
			}
		}

		_ = v.context.DefineVariable(param.Identifier.Name, val)
	}

	// Define return variable
	funcName := fn.Identifier.Name
	_ = v.context.DefineVariable(funcName, nil)

	// Hoist Dim declarations (VBScript hoisting semantics)
	if fn.Body != nil {
		v.hoistDimDeclarations(fn.Body)
	}

	// Execute body
	if fn.Body != nil {
		if list, ok := fn.Body.(*ast.StatementList); ok {
			for _, stmt := range list.Statements {
				if err := v.VisitStatement(stmt); err != nil {
					if pe, ok := err.(*ProcedureExitError); ok {
						if pe.Kind == "function" {
							break
						}
						return nil, err
					}
					return nil, err
				}
			}
		} else {
			if err := v.VisitStatement(fn.Body); err != nil {
				if pe, ok := err.(*ProcedureExitError); ok {
					if pe.Kind == "function" {
						goto functionWithRefsDone
					}
					return nil, err
				}
				return nil, err
			}
		}
	}

functionWithRefsDone:

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

	// Hoist Dim declarations (VBScript hoisting semantics)
	if sub.Body != nil {
		v.hoistDimDeclarations(sub.Body)
	}

	// Execute body
	if sub.Body != nil {
		if list, ok := sub.Body.(*ast.StatementList); ok {
			for _, stmt := range list.Statements {
				if err := v.VisitStatement(stmt); err != nil {
					if pe, ok := err.(*ProcedureExitError); ok {
						if pe.Kind == "sub" {
							break
						}
						return nil, err
					}
					return nil, err
				}
			}
		} else {
			if err := v.VisitStatement(sub.Body); err != nil {
				if pe, ok := err.(*ProcedureExitError); ok {
					if pe.Kind == "sub" {
						goto subDone
					}
					return nil, err
				}
				return nil, err
			}
		}
	}

subDone:

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
			// VBScript default: parameters are ByRef unless explicitly ByVal
			// Check if parameter is ByVal (explicitly)
			if param.Modifier == ast.ParameterModifierByVal {
				// ByVal: evaluate the argument
				var err error
				val, err = visitor.visitExpression(arguments[i])
				if err != nil {
					return nil, err
				}
			} else {
				// ByRef (explicit or default): we need the original variable name
				if ident, ok := arguments[i].(*ast.Identifier); ok {
					byRefMap[strings.ToLower(param.Identifier.Name)] = strings.ToLower(ident.Name)
					// Handle "Me" keyword - not stored as regular variable
					if strings.EqualFold(ident.Name, "me") {
						if ctxObj := v.context.GetContextObject(); ctxObj != nil {
							val = ctxObj
						}
					} else if parentVal, exists := v.context.GetVariableFromParentScope(ident.Name); exists {
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
			}
		}

		_ = v.context.DefineVariable(param.Identifier.Name, val)
	}

	// Hoist Dim declarations (VBScript hoisting semantics)
	if sub.Body != nil {
		v.hoistDimDeclarations(sub.Body)
	}

	// Execute body
	if sub.Body != nil {
		if list, ok := sub.Body.(*ast.StatementList); ok {
			for _, stmt := range list.Statements {
				if err := v.VisitStatement(stmt); err != nil {
					if pe, ok := err.(*ProcedureExitError); ok {
						if pe.Kind == "sub" {
							break
						}
						return nil, err
					}
					return nil, err
				}
			}
		} else {
			if err := v.VisitStatement(sub.Body); err != nil {
				if pe, ok := err.(*ProcedureExitError); ok {
					if pe.Kind == "sub" {
						goto subWithRefsDone
					}
					return nil, err
				}
				return nil, err
			}
		}
	}

subWithRefsDone:

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

// callClassMethodWithRefs calls a class instance method with ByRef parameter support
func (v *ASPVisitor) callClassMethodWithRefs(classInst *ClassInstance, methodName string, arguments []ast.Expression) (interface{}, error) {
	// Debug logging commented out for performance
	// if ctxObj := v.context.GetContextObject(); ctxObj != nil {
	// 	if ci, ok := ctxObj.(*ClassInstance); ok {
	// 		_ = ci
	// 		//fmt.Printf("[DEBUG] callClassMethodWithRefs ENTRY: currentContextObj=%s (calling %s.%s)\n", ci.ClassDef.Name, classInst.ClassDef.Name, methodName)
	// 	} else {
	// 		//fmt.Printf("[DEBUG] callClassMethodWithRefs ENTRY: currentContextObj=%T (calling %s.%s)\n", ctxObj, classInst.ClassDef.Name, methodName)
	// 	}
	// } else {
	// 	//fmt.Printf("[DEBUG] callClassMethodWithRefs ENTRY: currentContextObj=nil (calling %s.%s)\n", classInst.ClassDef.Name, methodName)
	// }
	//fmt.Printf("[DEBUG] callClassMethodWithRefs: class=%s method=%s args=%d\n", classInst.ClassDef.Name, methodName, len(arguments))
	// Debug: print argument types (commented out for performance)
	// for i, arg := range arguments {
	// 	_ = i
	// 	//fmt.Printf("[DEBUG] callClassMethodWithRefs: arg[%d] type=%T\n", i, arg)
	// 	if ident, ok := arg.(*ast.Identifier); ok {
	// 		_ = ident
	// 		//fmt.Printf("[DEBUG] callClassMethodWithRefs: arg[%d] is Identifier: %s\n", i, ident.Name)
	// 	} else if member, ok := arg.(*ast.MemberExpression); ok {
	// 		prop := ""
	// 		if member.Property != nil {
	// 			prop = member.Property.Name
	// 		}
	// 		_ = prop
	// 		//fmt.Printf("[DEBUG] callClassMethodWithRefs: arg[%d] is MemberExpression with prop=%s\n", i, prop)
	// 	} else if binExpr, ok := arg.(*ast.BinaryExpression); ok {
	// 		_ = binExpr
	// 		//fmt.Printf("[DEBUG] callClassMethodWithRefs: arg[%d] is BinaryExpression with Op=%v\n", i, binExpr.Operation)
	// 	} else if indexCall, ok := arg.(*ast.IndexOrCallExpression); ok {
	// 		_ = indexCall
	// 		//fmt.Printf("[DEBUG] callClassMethodWithRefs: arg[%d] is IndexOrCallExpression with Object=%T\n", i, indexCall.Object)
	// 	}
	// }
	nameLower := strings.ToLower(methodName)
	var methodNode ast.Node
	var params []*ast.Parameter
	var funcName string

	//fmt.Printf("[DEBUG] callClassMethodWithRefs: looking up method %s in class %s\n", nameLower, classInst.ClassDef.Name)
	//fmt.Printf("[DEBUG] callClassMethodWithRefs: class has %d Methods, %d Functions, %d PrivateMethods\n",
	//	len(classInst.ClassDef.Methods), len(classInst.ClassDef.Functions), len(classInst.ClassDef.PrivateMethods))
	if sub, ok := classInst.ClassDef.Methods[nameLower]; ok {
		//fmt.Printf("[DEBUG] callClassMethodWithRefs: found in Methods\n")
		methodNode = sub
		params = sub.Parameters
	}
	if fn, ok := classInst.ClassDef.Functions[nameLower]; ok {
		//fmt.Printf("[DEBUG] callClassMethodWithRefs: found in Functions\n")
		methodNode = fn
		params = fn.Parameters
		funcName = fn.Identifier.Name
	}
	if node, ok := classInst.ClassDef.PrivateMethods[nameLower]; ok {
		//fmt.Printf("[DEBUG] callClassMethodWithRefs: found in PrivateMethods\n")
		methodNode = node
		switch n := node.(type) {
		case *ast.SubDeclaration:
			params = n.Parameters
		case *ast.FunctionDeclaration:
			params = n.Parameters
			funcName = n.Identifier.Name
		}
	}
	// Check Properties (PropGet) - Handles Private Properties too since we are inside the context or explicitly asking for it
	// NOTE: We only care about PropGet for function calls.
	if methodNode == nil {
		if props, ok := classInst.ClassDef.Properties[nameLower]; ok {
			for _, p := range props {
				if p.Type == PropGet { // Assuming valid visibility check is implicit or we allow it via callClassMethodWithRefs?
					// callClassMethodWithRefs is called by resolveCall which verifies current context object.
					// If we are calling Me.Prop(), we have access.
					methodNode = p.Node
					if pg, ok := p.Node.(*ast.PropertyGetDeclaration); ok {
						params = pg.Parameters
						funcName = pg.Identifier.Name
					}
					break
				}
			}
		}
	}

	if methodNode == nil {
		//fmt.Printf("[DEBUG] callClassMethodWithRefs: methodNode is NIL, falling back to CallMethod\n")
		args := make([]interface{}, len(arguments))
		for i, arg := range arguments {
			val, err := v.visitExpression(arg)
			if err != nil {
				return nil, err
			}
			args[i] = val
		}
		return classInst.CallMethod(methodName, args...)
	}

	// NOTE: Scope push delayed until after argument evaluation

	// Evaluate all arguments in the CALLER'S context BEFORE switching context
	evaluatedArgs := make([]interface{}, len(arguments))
	byRefMap := make(map[string]string)

	//fmt.Printf("[DEBUG] callClassMethodWithRefs: ABOUT TO EVALUATE ARGS, numArgs=%d\n", len(arguments))
	for i, arg := range arguments {
		if i < len(params) {
			param := params[i]
			if param.Modifier == ast.ParameterModifierByVal {
				// ByVal: evaluate in caller's context
				//fmt.Printf("[DEBUG] callClassMethodWithRefs: evaluating arg[%d] (ByVal)\n", i)
				val, err := v.visitExpression(arg)
				if err != nil {
					return nil, err
				}
				evaluatedArgs[i] = val
			} else {
				// ByRef: evaluate in caller's context and track variable name
				if ident, ok := arg.(*ast.Identifier); ok {
					paramNameLower := strings.ToLower(param.Identifier.Name)
					identNameLower := strings.ToLower(ident.Name)
					byRefMap[paramNameLower] = identNameLower
					// Handle "Me" keyword - it's not a regular variable
					if strings.EqualFold(ident.Name, "me") {
						if ctxObj := v.context.GetContextObject(); ctxObj != nil {
							evaluatedArgs[i] = ctxObj
						} else {
							evaluatedArgs[i] = nil
						}
					} else if existingVal, exists := v.context.GetVariable(ident.Name); exists {
						evaluatedArgs[i] = existingVal
					} else {
						evaluatedArgs[i] = nil
					}
				} else {
					// Non-identifier ByRef argument, evaluate in caller's context
					//fmt.Printf("[DEBUG] callClassMethodWithRefs: evaluating arg[%d] (ByRef non-ident)\n", i)
					val, err := v.visitExpression(arg)
					if err != nil {
						return nil, err
					}
					evaluatedArgs[i] = val
				}
			}

		}
	}

	// NOW switch the context to the class instance (Isolate Scope)
	// 1. Save Scope Stack
	prevScopeStack := classInst.Context.scopeStack
	prevScopeConstStack := classInst.Context.scopeConstStack
	prevParentScopeStack := classInst.Context.parentScopeStack
	oldCtxObj := classInst.Context.GetContextObject()

	// 2. Clear Scope Stack (Isolate Class Method)
	classInst.Context.scopeStack = make([]map[string]interface{}, 0)
	classInst.Context.scopeConstStack = make([]map[string]interface{}, 0)
	classInst.Context.parentScopeStack = buildClassIsolatedParentScopes(prevParentScopeStack, prevScopeStack, oldCtxObj)

	// 3. Push Scope for Method
	classInst.Context.PushScope()

	// 4. Set Context Object
	classInst.Context.SetContextObject(classInst)

	// Defer Cleanup & Restoration
	defer func() {
		// Capture ByRef values from Method Scope (before popping)
		byRefResults := make(map[string]interface{})
		for paramName := range byRefMap {
			if val, exists := classInst.Context.GetVariable(paramName); exists {
				byRefResults[paramName] = val
			}
		}

		classInst.Context.PopScope()
		classInst.Context.SetContextObject(oldCtxObj)

		// Restore Caller Scope Stack
		classInst.Context.scopeStack = prevScopeStack
		classInst.Context.scopeConstStack = prevScopeConstStack
		classInst.Context.parentScopeStack = prevParentScopeStack

		// Apply ByRef updates to Caller Scope
		for paramName, origVarName := range byRefMap {
			if newVal, exists := byRefResults[paramName]; exists {
				// Use SetVariable (Caller Context is active)
				v.context.SetVariable(origVarName, newVal)
			}
		}
	}()

	// Define parameters in the class instance context with the evaluated values
	for i, param := range params {
		if i < len(evaluatedArgs) {
			_ = classInst.Context.DefineVariable(param.Identifier.Name, evaluatedArgs[i])
		}
	}

	if funcName != "" {
		_ = classInst.Context.DefineVariable(funcName, nil)
	}

	classVisitor := NewASPVisitor(classInst.Context, v.executor)
	// Set currentFunctionName so the function can read/write its return value without recursion
	classVisitor.currentFunctionName = funcName

	// Hoist Dim declarations in the method body (VBScript hoisting semantics)
	switch n := methodNode.(type) {
	case *ast.SubDeclaration:
		if n.Body != nil {
			classVisitor.hoistDimDeclarations(n.Body)
		}
	case *ast.FunctionDeclaration:
		if n.Body != nil {
			classVisitor.hoistDimDeclarations(n.Body)
		}
	case *ast.PropertyGetDeclaration:
		// Property Body is []ast.Statement, not ast.Statement
		for _, stmt := range n.Body {
			classVisitor.hoistDimDeclarations(stmt)
		}
	}

	switch n := methodNode.(type) {
	case *ast.SubDeclaration:
		if n.Body != nil {
			if list, ok := n.Body.(*ast.StatementList); ok {
				for _, stmt := range list.Statements {
					if err := classVisitor.VisitStatement(stmt); err != nil {
						if pe, ok := err.(*ProcedureExitError); ok && pe.Kind == "sub" {
							break
						}
						return nil, err
					}
				}
			} else {
				if err := classVisitor.VisitStatement(n.Body); err != nil {
					return nil, err
				}
			}
		}
	case *ast.FunctionDeclaration:
		if n.Body != nil {
			if list, ok := n.Body.(*ast.StatementList); ok {
				for _, stmt := range list.Statements {
					if err := classVisitor.VisitStatement(stmt); err != nil {
						if pe, ok := err.(*ProcedureExitError); ok && pe.Kind == "function" {
							break
						}
						return nil, err
					}
				}
			} else {
				if err := classVisitor.VisitStatement(n.Body); err != nil {
					return nil, err
				}
			}
		}
	case *ast.PropertyGetDeclaration:
		for _, stmt := range n.Body {
			if err := classVisitor.VisitStatement(stmt); err != nil {
				if pe, ok := err.(*ProcedureExitError); ok && pe.Kind == "property" {
					break
				}
				return nil, err
			}
		}
	}

	if funcName != "" {
		if val, exists := classInst.Context.GetVariable(funcName); exists {
			return val, nil
		}
	}
	return nil, nil
}

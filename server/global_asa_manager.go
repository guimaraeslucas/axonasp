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
	"g3pix.com.br/axonasp/asp"
	"g3pix.com.br/axonasp/vbscript/ast"
	"os"
	"path/filepath"
	"strings"
	"sync"
)

// GlobalASAManager handles Global.asa loading and event callbacks
type GlobalASAManager struct {
	mu sync.RWMutex

	// Parsed procedures from Global.asa
	applicationOnStart *ast.SubDeclaration
	applicationOnEnd   *ast.SubDeclaration
	sessionOnStart     *ast.SubDeclaration
	sessionOnEnd       *ast.SubDeclaration

	// Variables declared with <OBJECT> tags in Global.asa
	staticObjects map[string]interface{}
	// Global.asa file content
	globalASAContent string

	// Track if Global.asa has been loaded
	loaded bool

	// Track which events are defined
	hasApplicationOnStart bool
	hasApplicationOnEnd   bool
	hasSessionOnStart     bool
	hasSessionOnEnd       bool
}

// globalASAManagerSingleton holds the singleton instance
var globalASAManagerInstance *GlobalASAManager
var globalASAManagerOnce sync.Once

// GetGlobalASAManager returns the singleton GlobalASAManager instance
func GetGlobalASAManager() *GlobalASAManager {
	globalASAManagerOnce.Do(func() {
		globalASAManagerInstance = &GlobalASAManager{
			staticObjects: make(map[string]interface{}),
		}
	})
	return globalASAManagerInstance
}

// LoadGlobalASA loads and parses the Global.asa file from disk
// This should be called once during server startup
func (gam *GlobalASAManager) LoadGlobalASA(webRoot string) error {
	gam.mu.Lock()
	defer gam.mu.Unlock()

	// Construct path to Global.asa
	globalASAPath := filepath.Join(webRoot, "global.asa")

	// Check if file exists
	if _, err := os.Stat(globalASAPath); os.IsNotExist(err) {
		// Global.asa is optional, so just return without error
		gam.loaded = true
		return nil
	}

	// Read Global.asa content with UTF-8 decoding
	content, err := asp.ReadFileText(globalASAPath)
	if err != nil {
		return fmt.Errorf("failed to read Global.asa: %w", err)
	}

	gam.globalASAContent = content
	// Parse Global.asa
	parsingOptions := &asp.ASPParsingOptions{
		SaveComments: false,
		StrictMode:   false,
		DebugMode:    false,
	}

	parser := asp.NewASPParserWithOptions(gam.globalASAContent, parsingOptions)
	result, err := parser.Parse()
	if err != nil {
		return fmt.Errorf("failed to parse Global.asa: %w", err)
	}

	// Check for parse errors
	if len(result.Errors) > 0 {
		return fmt.Errorf("Global.asa parsing errors: %v", result.Errors)
	}

	// Extract event handlers and OBJECT declarations
	err = gam.extractFromVBPrograms(result.VBPrograms)
	if err != nil {
		return fmt.Errorf("failed to extract Global.asa content: %w", err)
	}

	gam.loaded = true
	return nil
}

// extractFromVBPrograms extracts procedures and static objects from parsed VB programs
func (gam *GlobalASAManager) extractFromVBPrograms(programs map[int]*ast.Program) error {
	if len(programs) == 0 {
		return nil
	}

	// Iterate through all programs and extract Sub procedures
	for _, program := range programs {
		if program == nil {
			continue
		}

		for _, stmt := range program.Body {
			gam.checkForEventHandler(stmt)
		}
	}

	return nil
}

// checkForEventHandler checks if a statement is an event handler Sub
func (gam *GlobalASAManager) checkForEventHandler(stmt interface{}) {
	subStmt, ok := stmt.(*ast.SubDeclaration)
	if !ok {
		return
	}

	// Get the sub name from the Identifier
	if subStmt.Identifier == nil {
		return
	}

	subName := strings.ToLower(strings.TrimSpace(subStmt.Identifier.Name))
	switch subName {
	case "application_onstart":
		gam.applicationOnStart = subStmt
		gam.hasApplicationOnStart = true
	case "application_onend":
		gam.applicationOnEnd = subStmt
		gam.hasApplicationOnEnd = true
	case "session_onstart":
		gam.sessionOnStart = subStmt
		gam.hasSessionOnStart = true
	case "session_onend":
		gam.sessionOnEnd = subStmt
		gam.hasSessionOnEnd = true
	}
}

// HasApplicationOnStart returns true if Application_OnStart is defined
func (gam *GlobalASAManager) HasApplicationOnStart() bool {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.hasApplicationOnStart
}

// HasApplicationOnEnd returns true if Application_OnEnd is defined
func (gam *GlobalASAManager) HasApplicationOnEnd() bool {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.hasApplicationOnEnd
}

// HasSessionOnStart returns true if Session_OnStart is defined
func (gam *GlobalASAManager) HasSessionOnStart() bool {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.hasSessionOnStart
}

// HasSessionOnEnd returns true if Session_OnEnd is defined
func (gam *GlobalASAManager) HasSessionOnEnd() bool {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.hasSessionOnEnd
}

// GetApplicationOnStart returns the Application_OnStart Sub procedure
func (gam *GlobalASAManager) GetApplicationOnStart() *ast.SubDeclaration {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.applicationOnStart
}

// GetApplicationOnEnd returns the Application_OnEnd Sub procedure
func (gam *GlobalASAManager) GetApplicationOnEnd() *ast.SubDeclaration {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.applicationOnEnd
}

// GetSessionOnStart returns the Session_OnStart Sub procedure
func (gam *GlobalASAManager) GetSessionOnStart() *ast.SubDeclaration {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.sessionOnStart
}

// GetSessionOnEnd returns the Session_OnEnd Sub procedure
func (gam *GlobalASAManager) GetSessionOnEnd() *ast.SubDeclaration {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.sessionOnEnd
}

// ExecuteApplicationOnStart executes the Application_OnStart event
// This should be called once when the server starts
func (gam *GlobalASAManager) ExecuteApplicationOnStart(executor *ASPExecutor, ctx *ExecutionContext) error {
	if !gam.HasApplicationOnStart() {
		return nil
	}

	sub := gam.GetApplicationOnStart()
	if sub == nil {
		return nil
	}

	// Create a visitor to execute the Sub body
	visitor := NewASPVisitor(ctx, executor)

	ctx.PushScope()
	defer ctx.PopScope()

	if sub.Body != nil {
		err := visitor.VisitStatement(sub.Body)
		if err != nil {
			return fmt.Errorf("error in Application_OnStart: %w", err)
		}
	}

	return nil
}

// ExecuteSessionOnStart executes the Session_OnStart event for a new session
func (gam *GlobalASAManager) ExecuteSessionOnStart(executor *ASPExecutor, ctx *ExecutionContext) error {
	if !gam.HasSessionOnStart() {
		return nil
	}

	sub := gam.GetSessionOnStart()
	if sub == nil {
		return nil
	}

	// Create a visitor to execute the Sub body
	visitor := NewASPVisitor(ctx, executor)

	ctx.PushScope()
	defer ctx.PopScope()

	if sub.Body != nil {
		err := visitor.VisitStatement(sub.Body)
		if err != nil {
			return fmt.Errorf("error in Session_OnStart: %w", err)
		}
	}

	return nil
}

// ExecuteSessionOnEnd executes the Session_OnEnd event for a session that is ending
func (gam *GlobalASAManager) ExecuteSessionOnEnd(executor *ASPExecutor, ctx *ExecutionContext) error {
	if !gam.HasSessionOnEnd() {
		return nil
	}

	sub := gam.GetSessionOnEnd()
	if sub == nil {
		return nil
	}

	// Create a visitor to execute the Sub body
	visitor := NewASPVisitor(ctx, executor)

	ctx.PushScope()
	defer ctx.PopScope()

	if sub.Body != nil {
		err := visitor.VisitStatement(sub.Body)
		if err != nil {
			return fmt.Errorf("error in Session_OnEnd: %w", err)
		}
	}

	return nil
}

// ExecuteApplicationOnEnd executes the Application_OnEnd event when server is shutting down
func (gam *GlobalASAManager) ExecuteApplicationOnEnd(executor *ASPExecutor, ctx *ExecutionContext) error {
	if !gam.HasApplicationOnEnd() {
		return nil
	}

	sub := gam.GetApplicationOnEnd()
	if sub == nil {
		return nil
	}

	// Create a visitor to execute the Sub body
	visitor := NewASPVisitor(ctx, executor)

	ctx.PushScope()
	defer ctx.PopScope()

	if sub.Body != nil {
		err := visitor.VisitStatement(sub.Body)
		if err != nil {
			return fmt.Errorf("error in Application_OnEnd: %w", err)
		}
	}

	return nil
}

// IsLoaded returns whether Global.asa has been loaded
func (gam *GlobalASAManager) IsLoaded() bool {
	gam.mu.RLock()
	defer gam.mu.RUnlock()
	return gam.loaded
}

// Reset clears all loaded Global.asa data (useful for testing)
func (gam *GlobalASAManager) Reset() {
	gam.mu.Lock()
	defer gam.mu.Unlock()

	gam.applicationOnStart = nil
	gam.applicationOnEnd = nil
	gam.sessionOnStart = nil
	gam.sessionOnEnd = nil
	gam.staticObjects = make(map[string]interface{})
	gam.globalASAContent = ""
	gam.loaded = false
}

// AddStaticObject adds a static object to the Global.asa manager
// These objects are typically created from <OBJECT> tags
func (gam *GlobalASAManager) AddStaticObject(name string, obj interface{}) {
	gam.mu.Lock()
	defer gam.mu.Unlock()

	nameLower := strings.ToLower(name)
	gam.staticObjects[nameLower] = obj
}

// GetStaticObject retrieves a static object by name
func (gam *GlobalASAManager) GetStaticObject(name string) (interface{}, bool) {
	gam.mu.RLock()
	defer gam.mu.RUnlock()

	nameLower := strings.ToLower(name)
	obj, exists := gam.staticObjects[nameLower]
	return obj, exists
}

// GetAllStaticObjects returns all static objects
func (gam *GlobalASAManager) GetAllStaticObjects() map[string]interface{} {
	gam.mu.RLock()
	defer gam.mu.RUnlock()

	// Return a copy
	copy := make(map[string]interface{})
	for k, v := range gam.staticObjects {
		copy[k] = v
	}
	return copy
}

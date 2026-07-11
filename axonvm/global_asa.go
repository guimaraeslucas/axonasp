/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Updated by @mlaadd 06/05/2026
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
package axonvm

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"

	"g3pix.com.br/axonasp/axonvm/asp"
	"g3pix.com.br/axonasp/vbscript"
)

// GlobalASA manages the parsed and compiled state of the Application's global.asa file.
// It executes top-level object declarations and handles Application and Session events.
type GlobalASA struct {
	mu           sync.RWMutex
	compiler     *Compiler
	bytecode     []byte
	constants    []Value
	globalsCount int
	isLoaded     bool

	hasAppOnStart  bool
	hasAppOnEnd    bool
	hasSessOnStart bool
	hasSessOnEnd   bool

	appOnStartIdx  int
	appOnEndIdx    int
	sessOnStartIdx int
	sessOnEndIdx   int

	sessionStaticObjects []*vbscript.ASPObjectToken
}

var (
	globalASAInstance *GlobalASA
	globalASAOnce     sync.Once
)

// GetGlobalASA returns the singleton GlobalASA instance.
func GetGlobalASA() *GlobalASA {
	globalASAOnce.Do(func() {
		globalASAInstance = &GlobalASA{}
	})
	return globalASAInstance
}

// ResetGlobalASA discards the singleton and allows a fresh LoadAndCompile
// with a different global.asa. Intended for test use only.
func ResetGlobalASA() {
	globalASAInstance = nil
	globalASAOnce = sync.Once{}
}

// LoadAndCompile reads, compiles, and registers the global.asa file from the specified path.
func (g *GlobalASA) LoadAndCompile(webRoot string, app *asp.Application) error {
	g.mu.Lock()
	defer g.mu.Unlock()

	globalASAPath := filepath.Join(webRoot, "global.asa")

	if _, err := os.Stat(globalASAPath); os.IsNotExist(err) {
		g.isLoaded = true
		return nil
	}

	content, err := os.ReadFile(globalASAPath)
	if err != nil {
		return fmt.Errorf("failed to read global.asa: %w", err)
	}

	// Strip UTF-8 BOM if present to prevent parsing errors
	content = stripUTF8BOM(content)

	compiler := NewASPCompiler(string(content))
	compiler.SetSourceName(globalASAPath)
	compiler.SetIncludeSiteRoot(webRoot)

	if err := compiler.Compile(); err != nil {
		return fmt.Errorf("failed to compile global.asa: %w", err)
	}

	g.bytecode = compiler.Bytecode()
	g.constants = compiler.Constants()
	g.globalsCount = compiler.GlobalsCount()
	g.compiler = compiler

	g.appOnStartIdx, g.hasAppOnStart = compiler.Globals.Get("Application_OnStart")
	g.appOnEndIdx, g.hasAppOnEnd = compiler.Globals.Get("Application_OnEnd")
	g.sessOnStartIdx, g.hasSessOnStart = compiler.Globals.Get("Session_OnStart")
	g.sessOnEndIdx, g.hasSessOnEnd = compiler.Globals.Get("Session_OnEnd")

	for _, objToken := range compiler.ObjectDeclarations {
		scope := strings.ToLower(strings.TrimSpace(objToken.Scope))
		progID := strings.TrimSpace(objToken.ProgID)
		if progID == "" {
			progID = strings.TrimSpace(objToken.ClassID)
		}
		var val asp.ApplicationValue
		if progID == "" {
			val = asp.NewApplicationEmpty()
		} else {
			val = asp.NewApplicationString(staticObjectProgIDPrefix + progID)
		}

		if scope == "application" && app != nil {
			app.AddStaticObject(objToken.ID, val)
		} else if scope == "session" {
			g.sessionStaticObjects = append(g.sessionStaticObjects, objToken)
		}
	}

	g.isLoaded = true
	return nil
}

// PopulateSessionStaticObjects adds the globally defined Session scope static objects to a new Session.
func (g *GlobalASA) PopulateSessionStaticObjects(session *asp.Session) {
	g.mu.RLock()
	defer g.mu.RUnlock()

	for _, objToken := range g.sessionStaticObjects {
		progID := strings.TrimSpace(objToken.ProgID)
		if progID == "" {
			progID = strings.TrimSpace(objToken.ClassID)
		}
		var val asp.ApplicationValue
		if progID == "" {
			val = asp.NewApplicationEmpty()
		} else {
			val = asp.NewApplicationString(staticObjectProgIDPrefix + progID)
		}
		session.AddStaticObject(objToken.ID, val)
	}
}

func (g *GlobalASA) executeHandler(host ASPHostEnvironment, handlerIdx int, handlerName string) error {
	vm := AcquireVMFromCompiler(g.compiler)
	vm.SetHost(host)
	defer vm.Release()

	// Suppress standard response output for Global.asa handlers to match IIS behavior.
	response := host.Response()
	originalOutput := response.Output
	response.Output = nil
	defer func() { response.Output = originalOutput }()

	// Run the top-level code to populate Sub/Function declarations into global variables
	// and the JScript environment.
	if err := vm.Run(); err != nil {
		return err
	}

	// 1) Check VBScript Globals for a VTUserSub handler (VBScript-defined handlers).
	if handlerIdx >= 0 && handlerIdx < len(vm.Globals) {
		handler := vm.Globals[handlerIdx]
		if handler.Type == VTUserSub {
			if vm.beginUserSubCall(handler, nil, true, 0) {
				return vm.Run()
			}
		}
	}

	// 2) Fallback: look up the handler as a JScript function in the JS environment.
	//    JScript function declarations inside <script language="JScript" runat="server">
	//    blocks are stored in the JS env frames (via OpJSDeclareName/OpJSSetName),
	//    NOT as VTUserSub entries in VBScript Globals.  Without this fallback,
	//    Application_OnStart / Session_OnStart / etc. defined in JScript are silently
	//    skipped, breaking IIS compatibility.
	if handlerName != "" {
		jsHandler := vm.jsGetName(handlerName)
		if jsHandler.Type == VTJSFunction {
			if vm.jsBeginDirectCall(jsHandler, Value{Type: VTJSUndefined}, nil) {
				return vm.Run()
			}
			// If jsBeginDirectCall returned false (e.g. generator/async/bound),
			// fall through to jsCall.
			vm.jsCall(jsHandler, Value{Type: VTJSUndefined}, nil)
		}
	}

	return nil
}

// ExecuteApplicationOnStart executes the Application_OnStart event.
// Supports both VBScript Sub (via VTUserSub in Globals) and JScript function
// (via the JS environment fallback in executeHandler).
func (g *GlobalASA) ExecuteApplicationOnStart(host ASPHostEnvironment) error {
	g.mu.RLock()
	idx := g.appOnStartIdx
	g.mu.RUnlock()
	return g.executeHandler(host, idx, "Application_OnStart")
}

// ExecuteApplicationOnEnd executes the Application_OnEnd event.
func (g *GlobalASA) ExecuteApplicationOnEnd(host ASPHostEnvironment) error {
	g.mu.RLock()
	idx := g.appOnEndIdx
	g.mu.RUnlock()
	return g.executeHandler(host, idx, "Application_OnEnd")
}

// ExecuteSessionOnStart executes the Session_OnStart event.
func (g *GlobalASA) ExecuteSessionOnStart(host ASPHostEnvironment) error {
	g.mu.RLock()
	idx := g.sessOnStartIdx
	g.mu.RUnlock()
	return g.executeHandler(host, idx, "Session_OnStart")
}

// ExecuteSessionOnEnd executes the Session_OnEnd event.
func (g *GlobalASA) ExecuteSessionOnEnd(host ASPHostEnvironment) error {
	g.mu.RLock()
	idx := g.sessOnEndIdx
	g.mu.RUnlock()
	return g.executeHandler(host, idx, "Session_OnEnd")
}

// IsLoaded returns whether Global.asa has been loaded.
func (g *GlobalASA) IsLoaded() bool {
	g.mu.RLock()
	defer g.mu.RUnlock()
	return g.isLoaded
}

// Reset clears the GlobalASA state (useful for testing).
func (g *GlobalASA) Reset() {
	g.mu.Lock()
	defer g.mu.Unlock()

	g.compiler = nil
	g.bytecode = nil
	g.constants = nil
	g.isLoaded = false
	g.hasAppOnStart = false
	g.hasAppOnEnd = false
	g.hasSessOnStart = false
	g.hasSessOnEnd = false
	g.sessionStaticObjects = nil
}

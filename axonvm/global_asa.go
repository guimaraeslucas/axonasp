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

func (g *GlobalASA) executeHandler(host ASPHostEnvironment, handlerIdx int) error {
	vm := AcquireVMFromCompiler(g.compiler)
	vm.SetHost(host)
	defer vm.Release()

	// Suppress standard response output for Global.asa handlers to match IIS behavior.
	host.Response().Output = nil

	// Run the top-level code to populate Sub declarations into global variables.
	if err := vm.Run(); err != nil {
		return err
	}

	if handlerIdx < len(vm.Globals) {
		handler := vm.Globals[handlerIdx]
		if handler.Type == VTUserSub {
			if vm.beginUserSubCall(handler, nil, true, 0) {
				return vm.Run()
			}
		}
	}
	return nil
}

// ExecuteApplicationOnStart executes the Application_OnStart event.
func (g *GlobalASA) ExecuteApplicationOnStart(host ASPHostEnvironment) error {
	g.mu.RLock()
	has := g.hasAppOnStart
	idx := g.appOnStartIdx
	g.mu.RUnlock()

	if !has {
		return nil
	}
	return g.executeHandler(host, idx)
}

// ExecuteApplicationOnEnd executes the Application_OnEnd event.
func (g *GlobalASA) ExecuteApplicationOnEnd(host ASPHostEnvironment) error {
	g.mu.RLock()
	has := g.hasAppOnEnd
	idx := g.appOnEndIdx
	g.mu.RUnlock()

	if !has {
		return nil
	}
	return g.executeHandler(host, idx)
}

// ExecuteSessionOnStart executes the Session_OnStart event.
func (g *GlobalASA) ExecuteSessionOnStart(host ASPHostEnvironment) error {
	g.mu.RLock()
	has := g.hasSessOnStart
	idx := g.sessOnStartIdx
	g.mu.RUnlock()

	if !has {
		return nil
	}
	return g.executeHandler(host, idx)
}

// ExecuteSessionOnEnd executes the Session_OnEnd event.
func (g *GlobalASA) ExecuteSessionOnEnd(host ASPHostEnvironment) error {
	g.mu.RLock()
	has := g.hasSessOnEnd
	idx := g.sessOnEndIdx
	g.mu.RUnlock()

	if !has {
		return nil
	}
	return g.executeHandler(host, idx)
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

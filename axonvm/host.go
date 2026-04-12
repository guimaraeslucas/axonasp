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
	"io"
	"os"

	"g3pix.com.br/axonasp/axonvm/asp"
)

// ASPHostEnvironment abstracts all interactions with the outside world.
type ASPHostEnvironment interface {
	Response() *asp.Response
	Request() *asp.Request
	Server() *asp.Server
	Session() *asp.Session
	Application() *asp.Application
	SetSessionEnabled(enabled bool)
	SessionEnabled() bool

	// ExecuteASPFile compiles and executes another ASP file sharing the current host context.
	// The caller is responsible for providing an absolute file system path.
	ExecuteASPFile(absPath string) error

	// Compatibility methods for VM/CLI (optional, could be removed later)
	Write(p []byte) (n int, err error)
	WriteString(s string)
}

// MockHost is a purely in-memory implementation for testing and CLI.
type MockHost struct {
	response         *asp.Response
	request          *asp.Request
	server           *asp.Server
	session          *asp.Session
	application      *asp.Application
	sessionEnabled   bool
	executeCacheHits int
	executeCompiles  int
}

// NewMockHost creates an in-memory ASP host with all intrinsic objects initialized.
func NewMockHost() *MockHost {
	m := &MockHost{
		response:       asp.NewResponse(nil), // Will be set later or use a buffer
		request:        asp.NewRequest(),
		server:         asp.NewServer(),
		session:        asp.NewSession(),
		application:    asp.NewApplication(),
		sessionEnabled: true,
	}
	return m
}

func (m *MockHost) Response() *asp.Response       { return m.response }
func (m *MockHost) Request() *asp.Request         { return m.request }
func (m *MockHost) Server() *asp.Server           { return m.server }
func (m *MockHost) Session() *asp.Session         { return m.session }
func (m *MockHost) Application() *asp.Application { return m.application }

// SetApplication injects a custom Application instance.
func (m *MockHost) SetApplication(app *asp.Application) { m.application = app }

// SetSession injects a custom Session instance.
func (m *MockHost) SetSession(session *asp.Session) { m.session = session }

// SetSessionEnabled toggles session availability for the current page execution.
func (m *MockHost) SetSessionEnabled(enabled bool) { m.sessionEnabled = enabled }

// SessionEnabled reports whether session state is enabled for the current page.
func (m *MockHost) SessionEnabled() bool { return m.sessionEnabled }

// ExecuteCacheHits reports how many child ASP executions reused cached compilation.
func (m *MockHost) ExecuteCacheHits() int { return m.executeCacheHits }

// ExecuteCompiles reports how many child ASP executions required compilation.
func (m *MockHost) ExecuteCompiles() int { return m.executeCompiles }

// Compatibility: Delegate to Response
func (m *MockHost) Write(p []byte) (int, error) {
	if m.response.Output != nil {
		return m.response.Output.Write(p)
	}
	return 0, nil
}

func (m *MockHost) WriteString(s string) {
	m.response.Write(s)
}

// ExecuteASPFile compiles and executes another ASP file within the current MockHost context.
// The child script shares the same Response, Session, and Application as the parent.
func (m *MockHost) ExecuteASPFile(absPath string) error {
	previousRequestPath := m.server.GetRequestPath()
	m.server.SetRequestPath(m.server.VirtualPathFromAbsolutePath(absPath))
	defer m.server.SetRequestPath(previousRequestPath)

	cache := getExecuteScriptCache()
	cacheHit := false
	program := CachedProgram{}
	if cache != nil {
		if cached, found := cache.Get(absPath); found {
			program = cached
			cacheHit = true
		}
	}
	if !cacheHit {
		var err error
		if cache != nil {
			program, err = cache.LoadOrCompile(absPath)
		} else {
			content, readErr := os.ReadFile(absPath)
			if readErr != nil {
				return readErr
			}
			// Strip UTF-8 BOM if present to prevent parsing errors
			content = stripUTF8BOM(content)
			compiler := NewASPCompiler(string(content))
			compiler.SetSourceName(absPath)
			if compileErr := compiler.Compile(); compileErr != nil {
				return compileErr
			}
			program = buildCachedProgramFromCompiler(compiler)
		}
		if err != nil {
			return err
		}
		m.executeCompiles++
	} else {
		m.executeCacheHits++
	}

	childVM := AcquireVMFromCachedProgram(program)
	childVM.SetHost(m)
	defer childVM.Release()
	return childVM.Run()
}

func (m *MockHost) SetOutput(w io.Writer) {
	m.response.Output = w
}

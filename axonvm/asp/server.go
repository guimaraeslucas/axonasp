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
package asp

import (
	"fmt"
	"html"
	"net/url"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/vbscript"
)

// Server provides server utility methods.
type Server struct {
	mu            sync.RWMutex
	scriptTimeout int
	rootDir       string
	requestPath   string
	lastError     *ASPError
	execStart     time.Time
	execDepth     int
}

// NewServer creates a new Server object with ASP-compatible defaults.
func NewServer() *Server {
	return &Server{
		scriptTimeout: 90,
		rootDir:       "./www",
		requestPath:   "/",
	}
}

// SetRootDir defines the web root used by MapPath for absolute virtual paths.
func (s *Server) SetRootDir(rootDir string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	if strings.TrimSpace(rootDir) == "" {
		s.rootDir = "./www"
		return
	}
	s.rootDir = rootDir
}

// SetRequestPath stores the current request URL path for relative MapPath resolution.
func (s *Server) SetRequestPath(path string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	if strings.TrimSpace(path) == "" {
		s.requestPath = "/"
		return
	}
	s.requestPath = strings.ReplaceAll(path, "\\", "/")
}

// GetRequestPath returns the current request path used for relative MapPath resolution.
func (s *Server) GetRequestPath() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	if s.requestPath == "" {
		return "/"
	}
	return s.requestPath
}

// VirtualPathFromAbsolutePath converts one absolute file path under the root directory back into a virtual path.
func (s *Server) VirtualPathFromAbsolutePath(absPath string) string {
	s.mu.RLock()
	rootDir := s.rootDir
	s.mu.RUnlock()
	if strings.TrimSpace(absPath) == "" {
		return "/"
	}
	absRoot, err := filepath.Abs(rootDir)
	if err != nil {
		return "/" + filepath.ToSlash(filepath.Base(absPath))
	}
	absTarget, err := filepath.Abs(absPath)
	if err != nil {
		return "/" + filepath.ToSlash(filepath.Base(absPath))
	}
	relPath, err := filepath.Rel(absRoot, absTarget)
	if err != nil {
		return "/" + filepath.ToSlash(filepath.Base(absPath))
	}
	cleaned := filepath.ToSlash(relPath)
	cleaned = strings.TrimPrefix(cleaned, "./")
	cleaned = strings.TrimPrefix(cleaned, "/")
	if strings.HasPrefix(cleaned, "../") || cleaned == ".." {
		return "/" + filepath.ToSlash(filepath.Base(absTarget))
	}
	if cleaned == "." || cleaned == "" {
		return "/"
	}
	return "/" + cleaned
}

// GetScriptTimeout returns the script timeout in seconds.
func (s *Server) GetScriptTimeout() int {
	s.mu.RLock()
	defer s.mu.RUnlock()
	if s.scriptTimeout < 1 {
		return 90
	}
	return s.scriptTimeout
}

// SetScriptTimeout updates timeout value and validates ASP-compatible range.
func (s *Server) SetScriptTimeout(timeout int) error {
	if timeout < 1 {
		return fmt.Errorf("script timeout must be at least 1 second")
	}
	if timeout > 2147483647 {
		return fmt.Errorf("script timeout exceeds maximum value")
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	s.scriptTimeout = timeout
	return nil
}

// BeginExecution starts or nests script execution timing for one request.
func (s *Server) BeginExecution() {
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.execDepth == 0 {
		s.execStart = time.Now()
	}
	s.execDepth++
}

// EndExecution ends one nested script execution scope.
func (s *Server) EndExecution() {
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.execDepth > 0 {
		s.execDepth--
	}
	if s.execDepth == 0 {
		s.execStart = time.Time{}
	}
}

// HasTimedOut reports whether the current request execution exceeded the current ScriptTimeout value.
func (s *Server) HasTimedOut() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	if s.execDepth == 0 || s.execStart.IsZero() {
		return false
	}
	timeout := s.scriptTimeout
	if timeout < 1 {
		timeout = 90
	}
	return time.Since(s.execStart) >= time.Duration(timeout)*time.Second
}

// SetLastError stores the latest ASP error for GetLastError access.
func (s *Server) SetLastError(err *ASPError) {
	s.mu.Lock()
	defer s.mu.Unlock()
	if err == nil {
		s.lastError = nil
		return
	}

	s.lastError = err.Clone()
}

// GetLastError returns the current ASP error object, or a default empty error.
func (s *Server) GetLastError() *ASPError {
	s.mu.RLock()
	defer s.mu.RUnlock()
	if s.lastError != nil {
		return s.lastError
	}
	return NewASPError()
}

// ClearLastError resets the current ASP error state.
func (s *Server) ClearLastError() {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.lastError = nil
}

// HTMLEncode escapes special HTML characters using ASP-compatible behavior.
func (s *Server) HTMLEncode(str string) string {
	return html.EscapeString(str)
}

// URLEncode escapes text for query usage using RFC3986-compatible escaping.
func (s *Server) URLEncode(str string) string {
	return url.QueryEscape(str)
}

// URLPathEncode escapes URL path segments while preserving slash separators.
func (s *Server) URLPathEncode(str string) string {
	parts := strings.Split(str, "/")
	for index, part := range parts {
		parts[index] = url.PathEscape(part)
	}
	return strings.Join(parts, "/")
}

// MapPath resolves absolute and relative ASP paths into local file system paths.
func (s *Server) MapPath(path string) string {
	s.mu.RLock()
	rootDir := s.rootDir
	requestPath := s.requestPath
	s.mu.RUnlock()
	if path == "" || path == "/" || path == "\\" {
		absRoot, err := filepath.Abs(rootDir)
		if err != nil {
			return rootDir
		}
		return absRoot
	}

	normalized := strings.ReplaceAll(path, "\\", "/")
	if strings.HasPrefix(normalized, "/") {
		fullPath := filepath.Join(rootDir, strings.TrimPrefix(normalized, "/"))
		absPath, err := filepath.Abs(fullPath)
		if err != nil {
			return fullPath
		}
		return absPath
	}

	scriptDir := filepath.Dir(strings.ReplaceAll(requestPath, "\\", "/"))
	if scriptDir == "." {
		scriptDir = "/"
	}

	fullPath := filepath.Join(rootDir, scriptDir, normalized)
	absPath, err := filepath.Abs(fullPath)
	if err != nil {
		return fullPath
	}
	return absPath
}

// CreateObject provides a VM-compatible API surface and captures unsupported errors.
func (s *Server) CreateObject(progID string) (ApplicationValue, error) {
	trimmed := strings.TrimSpace(progID)
	if trimmed == "" {
		err := fmt.Errorf("CreateObject requires a ProgID")
		s.SetLastError(NewVBScriptASPError(vbscript.ActiveXCannotCreateObject, "Server.CreateObject", "ASP", err.Error(), "", 0, 0))
		return NewApplicationEmpty(), err
	}

	err := fmt.Errorf("ActiveX cannot create object: %s", trimmed)
	s.SetLastError(NewVBScriptASPError(vbscript.ActiveXCannotCreateObject, "Server.CreateObject", "ASP", err.Error(), "", 0, 0))
	return NewApplicationEmpty(), err
}

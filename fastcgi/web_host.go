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
package main

import (
	"io"
	"net"
	"net/http"
	"os"
	"strconv"
	"sync"

	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
)

// sharedFastCGIApplication stores process-wide application state in FastCGI mode.
var sharedFastCGIApplication = asp.NewApplication()

var (
	fastCGIExecuteScriptCacheOnce sync.Once
	fastCGIExecuteScriptCache     *axonvm.ScriptCache
)

func getFastCGIExecuteScriptCache() *axonvm.ScriptCache {
	fastCGIExecuteScriptCacheOnce.Do(func() {
		fastCGIExecuteScriptCache = axonvm.NewScriptCache(axonvm.BytecodeCacheMemoryOnly, "", 64)
	})
	return fastCGIExecuteScriptCache
}

// GetSharedApplication returns the singleton application object for global use.
func GetSharedApplication() *asp.Application {
	return sharedFastCGIApplication
}

const fastCGISessionCookieName = "ASPSESSIONID"

// FastCGIHost implements axonvm.ASPHostEnvironment for FastCGI requests.
type FastCGIHost struct {
	response       *asp.Response
	request        *asp.Request
	server         *asp.Server
	session        *asp.Session
	application    *asp.Application
	sessionEnabled bool
}

// NewFastCGIHost creates a host object bound to one FastCGI request.
func NewFastCGIHost(w http.ResponseWriter, r *http.Request) *FastCGIHost {
	session, isNew := loadOrCreateFastCGISession(r)

	host := &FastCGIHost{
		response:       asp.NewResponse(w),
		request:        asp.NewRequest(),
		server:         asp.NewServer(),
		session:        session,
		application:    sharedFastCGIApplication,
		sessionEnabled: true,
	}
	host.response.SetRequest(r)
	host.response.SetMaxBufferBytes(ResponseBufferLimitBytes)
	host.request.SetHTTPRequest(r)
	host.server.SetRootDir(RootDir)
	host.server.SetRequestPath(r.URL.Path)
	_ = host.server.SetScriptTimeout(ScriptTimeout)

	for key, values := range r.URL.Query() {
		host.request.QueryString.AddValues(key, values)
	}

	host.request.SetBodyLoader(func() ([]byte, error) {
		if r.Body == nil {
			return []byte{}, nil
		}
		loadedBody, err := io.ReadAll(r.Body)
		if err != nil {
			return nil, err
		}
		return loadedBody, nil
	})
	host.request.SetFormLoader(func() (map[string][]string, error) {
		if r.Method != http.MethodPost {
			return map[string][]string{}, nil
		}
		if err := r.ParseForm(); err != nil {
			return nil, err
		}
		return r.PostForm, nil
	})

	for _, cookie := range r.Cookies() {
		host.request.Cookies.AddCookie(cookie.Name, cookie.Value)
	}

	host.request.ServerVars.Add("REMOTE_ADDR", fastCGIRequestRemoteAddr(r.RemoteAddr))
	host.request.ServerVars.Add("REQUEST_METHOD", r.Method)
	host.request.ServerVars.Add("SERVER_NAME", fastCGIRequestServerName(r))
	host.request.ServerVars.Add("SERVER_PORT", fastCGIRequestServerPort(r))
	host.request.ServerVars.Add("SCRIPT_NAME", r.URL.Path)
	host.request.ServerVars.Add("URL", r.URL.Path)
	host.request.ServerVars.Add("HTTP_USER_AGENT", r.UserAgent())
	host.request.ServerVars.Add("HTTP_ACCEPT_LANGUAGE", r.Header.Get("Accept-Language"))
	host.request.ServerVars.Add("CONTENT_LENGTH", strconv.FormatInt(host.request.TotalBytes(), 10))
	host.request.ServerVars.Add("CONTENT_TYPE", r.Header.Get("Content-Type"))

	host.setSessionCookie()

	if isNew && axonvm.GetGlobalASA().IsLoaded() {
		axonvm.GetGlobalASA().PopulateSessionStaticObjects(session)
		_ = axonvm.GetGlobalASA().ExecuteSessionOnStart(host)
	}

	return host
}

// Response returns the ASP Response intrinsic object.
func (h *FastCGIHost) Response() *asp.Response { return h.response }

// Request returns the ASP Request intrinsic object.
func (h *FastCGIHost) Request() *asp.Request { return h.request }

// Server returns the ASP Server intrinsic object.
func (h *FastCGIHost) Server() *asp.Server { return h.server }

// Session returns the ASP Session intrinsic object.
func (h *FastCGIHost) Session() *asp.Session { return h.session }

// Application returns the ASP Application intrinsic object.
func (h *FastCGIHost) Application() *asp.Application { return h.application }

// SetSessionEnabled toggles session state for the current FastCGI page execution.
func (h *FastCGIHost) SetSessionEnabled(enabled bool) { h.sessionEnabled = enabled }

// SessionEnabled reports whether session state is enabled for the current FastCGI page.
func (h *FastCGIHost) SessionEnabled() bool { return h.sessionEnabled }

// Write forwards raw bytes into the ASP Response buffer.
func (h *FastCGIHost) Write(p []byte) (int, error) {
	h.response.Write(string(p))
	return len(p), nil
}

// WriteString forwards text output into the ASP Response buffer.
func (h *FastCGIHost) WriteString(s string) {
	h.response.Write(s)
}

// ExecuteASPFile compiles and executes another ASP file within the current FastCGIHost context.
// The child script shares the same Response, Session, and Application as the parent.
func (h *FastCGIHost) ExecuteASPFile(absPath string) error {
	previousRequestPath := h.server.GetRequestPath()
	h.server.SetRequestPath(h.server.VirtualPathFromAbsolutePath(absPath))
	defer h.server.SetRequestPath(previousRequestPath)

	cache := getFastCGIExecuteScriptCache()
	program := axonvm.CachedProgram{}
	if cache != nil {
		if cached, found := cache.Get(absPath); found {
			program = cached
		} else {
			compiled, compileErr := cache.LoadOrCompile(absPath)
			if compileErr != nil {
				return compileErr
			}
			program = compiled
		}
	} else {
		content, err := os.ReadFile(absPath)
		if err != nil {
			return err
		}
		// Strip UTF-8 BOM if present to prevent parsing errors
		if len(content) >= 3 && content[0] == 0xEF && content[1] == 0xBB && content[2] == 0xBF {
			content = content[3:]
		}
		compiler := axonvm.NewASPCompiler(string(content))
		compiler.SetSourceName(absPath)
		if err := compiler.Compile(); err != nil {
			return err
		}
		childVM := axonvm.AcquireVMFromCompiler(compiler)
		childVM.SetHost(h)
		defer childVM.Release()
		return childVM.Run()
	}

	childVM := axonvm.AcquireVMFromCachedProgram(program)
	childVM.SetHost(h)
	defer childVM.Release()
	return childVM.Run()
}

// PersistSession saves or rotates session state after request execution.
func (h *FastCGIHost) PersistSession() {
	if h.session == nil || !h.sessionEnabled {
		return
	}

	if h.session.IsAbandoned() {
		_ = h.session.Delete()
		newSession, err := asp.CreateSession()
		if err == nil {
			h.session = newSession
			h.setSessionCookie()
		}
		return
	}

	_ = h.session.SaveIfDirty()
	h.setSessionCookie()
}

// loadOrCreateFastCGISession resolves session from cookie or creates one.
func loadOrCreateFastCGISession(r *http.Request) (*asp.Session, bool) {
	var sessionID string
	if cookie, err := r.Cookie(fastCGISessionCookieName); err == nil && cookie != nil {
		sessionID = cookie.Value
	}

	session, isNew, err := asp.GetOrCreateSession(sessionID)
	if err != nil {
		return asp.NewSession(), true
	}

	return session, isNew
}

// setSessionCookie updates ASPSESSIONID to match the current session ID.
func (h *FastCGIHost) setSessionCookie() {
	if h.response == nil || h.response.Output == nil || h.session == nil {
		return
	}

	writer, ok := h.response.Output.(http.ResponseWriter)
	if !ok {
		return
	}

	http.SetCookie(writer, &http.Cookie{
		Name:     fastCGISessionCookieName,
		Value:    h.session.ID,
		Path:     "/",
		HttpOnly: true,
		SameSite: http.SameSiteLaxMode,
	})
}

// fastCGIRequestRemoteAddr normalizes RemoteAddr into the client host without the port suffix.
func fastCGIRequestRemoteAddr(remoteAddr string) string {
	host, _, err := net.SplitHostPort(remoteAddr)
	if err == nil && host != "" {
		return host
	}
	return remoteAddr
}

// fastCGIRequestServerName resolves the ASP SERVER_NAME variable from the request host.
func fastCGIRequestServerName(r *http.Request) string {
	if r == nil {
		return ""
	}
	if host := r.URL.Hostname(); host != "" {
		return host
	}
	host, _, err := net.SplitHostPort(r.Host)
	if err == nil && host != "" {
		return host
	}
	return r.Host
}

// fastCGIRequestServerPort resolves the ASP SERVER_PORT variable using explicit or default scheme ports.
func fastCGIRequestServerPort(r *http.Request) string {
	if r == nil {
		return ""
	}
	if port := r.URL.Port(); port != "" {
		return port
	}
	_, port, err := net.SplitHostPort(r.Host)
	if err == nil && port != "" {
		return port
	}
	if r.TLS != nil {
		return "443"
	}
	return "80"
}

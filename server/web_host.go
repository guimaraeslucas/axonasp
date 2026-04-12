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

// sharedApplication stores application-wide state across all web requests.
var sharedApplication = asp.NewApplication()

var (
	webHostExecuteScriptCacheOnce sync.Once
	webHostExecuteScriptCache     *axonvm.ScriptCache
)

func getWebHostExecuteScriptCache() *axonvm.ScriptCache {
	webHostExecuteScriptCacheOnce.Do(func() {
		webHostExecuteScriptCache = axonvm.NewScriptCache(axonvm.BytecodeCacheMemoryOnly, "", 64)
	})
	return webHostExecuteScriptCache
}

// GetSharedApplication returns the singleton application object for global use.
func GetSharedApplication() *asp.Application {
	return sharedApplication
}

const sessionCookieName = "ASPSESSIONID"

// WebHost implements axonvm.ASPHostEnvironment for a real HTTP request.
type WebHost struct {
	response       *asp.Response
	request        *asp.Request
	server         *asp.Server
	session        *asp.Session
	application    *asp.Application
	sessionEnabled bool
}

// NewWebHost creates a new WebHost instance from a real HTTP request/response.
func NewWebHost(w http.ResponseWriter, r *http.Request) *WebHost {
	session, isNew := loadOrCreateSession(r)

	host := &WebHost{
		response:       asp.NewResponse(w),
		request:        asp.NewRequest(),
		server:         asp.NewServer(),
		session:        session,
		application:    sharedApplication,
		sessionEnabled: true,
	}
	host.response.SetRequest(r)
	host.response.SetMaxBufferBytes(ResponseBufferLimitBytes)
	host.request.SetHTTPRequest(r)
	host.server.SetRootDir(RootDir)
	host.server.SetRequestPath(r.URL.Path)
	_ = host.server.SetScriptTimeout(ScriptTimeout)

	// Populate Request object
	// 1. QueryString
	for k, v := range r.URL.Query() {
		host.request.QueryString.AddValues(k, v)
	}

	// 2. Lazy body + form loaders to keep GET/HEAD hot paths allocation-free.
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

	// 3. Cookies
	for _, cookie := range r.Cookies() {
		host.request.Cookies.AddCookie(cookie.Name, cookie.Value)
	}

	// 4. ServerVariables (subset)
	hostName := requestServerName(r)
	port := requestServerPort(r)
	queryString := ""
	requestURI := r.URL.Path
	if r.URL != nil {
		queryString = r.URL.RawQuery
		requestURI = r.URL.RequestURI()
	}
	httpsValue := "off"
	if r.TLS != nil {
		httpsValue = "on"
	}
	host.request.ServerVars.Add("QUERY_STRING", queryString)
	host.request.ServerVars.Add("HTTP_HOST", r.Host)
	host.request.ServerVars.Add("HTTP_CONTENT_TYPE", r.Header.Get("Content-Type"))
	host.request.ServerVars.Add("HTTPS", httpsValue)
	host.request.ServerVars.Add("SERVER_PROTOCOL", r.Proto)
	host.request.ServerVars.Add("REQUEST_URI", requestURI)
	host.request.ServerVars.Add("PATH_INFO", r.URL.Path)
	host.request.ServerVars.Add("REMOTE_ADDR", requestRemoteAddr(r.RemoteAddr))
	host.request.ServerVars.Add("REQUEST_METHOD", r.Method)
	host.request.ServerVars.Add("SERVER_NAME", hostName)
	host.request.ServerVars.Add("SERVER_PORT", port)
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

func (h *WebHost) Response() *asp.Response       { return h.response }
func (h *WebHost) Request() *asp.Request         { return h.request }
func (h *WebHost) Server() *asp.Server           { return h.server }
func (h *WebHost) Session() *asp.Session         { return h.session }
func (h *WebHost) Application() *asp.Application { return h.application }

// SetSessionEnabled toggles session state for the current ASP page execution.
func (h *WebHost) SetSessionEnabled(enabled bool) { h.sessionEnabled = enabled }

// SessionEnabled reports whether session state is enabled for the current ASP page.
func (h *WebHost) SessionEnabled() bool { return h.sessionEnabled }

// PersistSession commits or removes session data after request execution.
func (h *WebHost) PersistSession() {
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

// loadOrCreateSession resolves session from ASPSESSIONID cookie or creates a new one.
func loadOrCreateSession(r *http.Request) (*asp.Session, bool) {
	var sessionID string
	if cookie, err := r.Cookie(sessionCookieName); err == nil && cookie != nil {
		sessionID = cookie.Value
	}

	session, isNew, err := asp.GetOrCreateSession(sessionID)
	if err != nil {
		fallback := asp.NewSession()
		return fallback, true
	}

	return session, isNew
}

// setSessionCookie updates ASPSESSIONID cookie to match current host session.
func (h *WebHost) setSessionCookie() {
	if h.response == nil || h.response.Output == nil || h.session == nil {
		return
	}

	writer, ok := h.response.Output.(http.ResponseWriter)
	if !ok {
		return
	}

	http.SetCookie(writer, &http.Cookie{
		Name:     sessionCookieName,
		Value:    h.session.ID,
		Path:     "/",
		HttpOnly: true,
		SameSite: http.SameSiteLaxMode,
	})
}

// Write forwards raw bytes into the ASP Response buffer.
func (h *WebHost) Write(p []byte) (int, error) {
	h.response.Write(string(p))
	return len(p), nil
}

// WriteString forwards text into the ASP Response buffer.
func (h *WebHost) WriteString(s string) {
	h.response.Write(s)
}

// ExecuteASPFile compiles and executes another ASP file within the current WebHost context.
// The child script shares the same Response, Session, and Application as the parent.
func (h *WebHost) ExecuteASPFile(absPath string) error {
	previousRequestPath := h.server.GetRequestPath()
	h.server.SetRequestPath(h.server.VirtualPathFromAbsolutePath(absPath))
	defer h.server.SetRequestPath(previousRequestPath)

	cache := getWebHostExecuteScriptCache()
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

// requestRemoteAddr normalizes RemoteAddr into the client host without the port suffix.
func requestRemoteAddr(remoteAddr string) string {
	host, _, err := net.SplitHostPort(remoteAddr)
	if err == nil && host != "" {
		return host
	}
	return remoteAddr
}

// requestServerName resolves the ASP SERVER_NAME variable from the request host.
func requestServerName(r *http.Request) string {
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

// requestServerPort resolves the ASP SERVER_PORT variable using explicit or default scheme ports.
func requestServerPort(r *http.Request) string {
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

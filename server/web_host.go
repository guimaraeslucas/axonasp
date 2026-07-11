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
	"path/filepath"
	"slices"
	"sort"
	"strconv"
	"strings"
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
	engineMode     axonvm.EngineMode
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
		engineMode:     ServerEngineMode,
	}
	host.response.SetRequest(r)
	host.response.SetMaxBufferBytes(ResponseBufferLimitBytes)
	host.request.SetHTTPRequest(r)
	host.server.SetRootDir(RootDir)
	host.server.SetRequestPath(r.URL.Path)
	_ = host.server.SetScriptTimeout(ScriptTimeout)

	// Populate Request object
	// 1. QueryString
	if len(r.URL.RawQuery) > 0 {
		host.request.QueryString.SetLazyPayload([]byte(r.URL.RawQuery))
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
	host.request.ServerVars.Add("HTTP_X_G3AXONLIVE", r.Header.Get("X-G3AxonLive"))
	host.request.ServerVars.Add("HTTP_X_G3AXONLIVE_SESSIONID", r.Header.Get("X-G3AxonLive-SessionId"))
	host.request.ServerVars.Add("HTTP_X_G3AXONLIVE_COMPONENTID", r.Header.Get("X-G3AxonLive-ComponentId"))
	host.request.ServerVars.Add("HTTP_X_G3AXONLIVE_EVENTNAME", r.Header.Get("X-G3AxonLive-EventName"))
	host.request.ServerVars.Add("HTTP_X_G3AXONLIVE_EVENTARGS", r.Header.Get("X-G3AxonLive-EventArgs"))
	host.request.ServerVars.Add("HTTPS", httpsValue)
	host.request.ServerVars.Add("AUTH_TYPE", requestAuthType(r))
	host.request.ServerVars.Add("SERVER_ADDR", requestServerAddr(r))
	host.request.ServerVars.Add("GATEWAY_INTERFACE", "CGI/1.1")
	host.request.ServerVars.Add("SERVER_SOFTWARE", "AxonASP")
	host.request.ServerVars.Add("SERVER_PROTOCOL", r.Proto)
	host.request.ServerVars.Add("REQUEST_URI", requestURI)
	host.request.ServerVars.Add("PATH_INFO", r.URL.Path)
	host.request.ServerVars.Add("PATH_TRANSLATED", host.server.MapPath(r.URL.Path))
	host.request.ServerVars.Add("APPL_PHYSICAL_PATH", host.server.MapPath("/"))
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
	allHTTP, allRaw := buildAggregateHeaderServerVariables(r.Header)
	host.request.ServerVars.Add("ALL_HTTP", allHTTP)
	host.request.ServerVars.Add("ALL_RAW", allRaw)

	// Expose all request headers for ASP access (HTTP_<HEADER_NAME>).
	for headerName, values := range r.Header {
		if len(values) == 0 {
			continue
		}
		host.request.ServerVars.Add(serverVariableFromHeader(headerName), strings.Join(values, ","))
	}

	host.setSessionCookie()

	if isNew && axonvm.GetGlobalASA().IsLoaded() {
		axonvm.GetGlobalASA().PopulateSessionStaticObjects(session)
		_ = axonvm.GetGlobalASA().ExecuteSessionOnStart(host)
		// Session_OnStart may call Response.End/Redirect which sets ended=true.
		// The handler suppresses output (Output=nil) so flushInternal is a no-op,
		// but ended stays true and would silently discard all page output.
		// Reset it so the page executes with a clean response state.
		if host.response.IsEnded() {
			host.response.ResetEnded()
		}
		// Clear any residual output written during Session_OnStart so it
		// does not leak into the page response body.
		host.response.Clear()
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

// EngineMode returns the current language mode for the host.
func (h *WebHost) EngineMode() axonvm.EngineMode { return h.engineMode }

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

	// Optimization: Use asynchronous write-behind for session persistence
	// to avoid blocking the HTTP handler and release the VM back to pool faster.
	h.session.QueueSaveIfDirty()
	h.setSessionCookie()
}

// buildAggregateHeaderServerVariables builds Classic ASP-style ALL_HTTP and ALL_RAW values.
func buildAggregateHeaderServerVariables(header http.Header) (string, string) {
	if len(header) == 0 {
		return "", ""
	}

	names := make([]string, 0, len(header))
	for name := range header {
		names = append(names, name)
	}
	sort.Strings(names)

	var allHTTP strings.Builder
	var allRaw strings.Builder
	for _, name := range names {
		values := header.Values(name)
		if len(values) == 0 {
			continue
		}
		joined := strings.Join(values, ",")

		if allHTTP.Len() > 0 {
			allHTTP.WriteString("\r\n")
		}
		allHTTP.WriteString(serverVariableFromHeader(name))
		allHTTP.WriteString(":")
		allHTTP.WriteString(joined)

		if allRaw.Len() > 0 {
			allRaw.WriteString("\r\n")
		}
		allRaw.WriteString(name)
		allRaw.WriteString(":")
		allRaw.WriteString(joined)
	}

	return allHTTP.String(), allRaw.String()
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

	replaceResponseCookie(writer, sessionCookieName)
	http.SetCookie(writer, &http.Cookie{
		Name:     sessionCookieName,
		Value:    h.session.ID,
		Path:     "/",
		HttpOnly: true,
		SameSite: http.SameSiteLaxMode,
	})
}

// replaceResponseCookie removes any pending Set-Cookie header for one cookie name before rewriting it.
func replaceResponseCookie(writer http.ResponseWriter, cookieName string) {
	if writer == nil {
		return
	}

	headers := writer.Header()
	if headers == nil {
		return
	}

	existing := headers.Values("Set-Cookie")
	if len(existing) == 0 {
		return
	}

	prefix := cookieName + "="
	filtered := make([]string, 0, len(existing))
	for _, value := range existing {
		if strings.HasPrefix(value, prefix) {
			continue
		}
		filtered = append(filtered, value)
	}

	headers.Del("Set-Cookie")
	for _, value := range filtered {
		headers.Add("Set-Cookie", value)
	}
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
			compiled, compileErr := cache.LoadOrCompileWithOptions(absPath, axonvm.ScriptCompileOptions{IncludeSiteRoot: h.server.MapPath("/")})
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

		var compiler *axonvm.Compiler
		ext := strings.ToLower(filepath.Ext(absPath))
		isVBS := slices.Contains(ExecuteAsVBScriptExtensions, ext)
		isJS := slices.Contains(ExecuteAsJavaScriptExtensions, ext)

		if isJS {
			compiler = axonvm.NewJavaScriptCompiler(string(content))
		} else if isVBS {
			compiler = axonvm.NewCompiler(string(content))
		} else {
			compiler = axonvm.NewASPCompiler(string(content))
		}

		compiler.SetSourceName(absPath)
		compiler.SetIncludeSiteRoot(h.server.MapPath("/"))
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

// requestServerAddr resolves the ASP SERVER_ADDR variable using the best available host or remote address.
func requestServerAddr(r *http.Request) string {
	if r == nil {
		return ""
	}
	if host := requestServerName(r); host != "" {
		return host
	}
	return requestRemoteAddr(r.RemoteAddr)
}

// requestAuthType resolves the ASP AUTH_TYPE variable from the Authorization header when available.
func requestAuthType(r *http.Request) string {
	if r == nil {
		return ""
	}
	authorization := strings.TrimSpace(r.Header.Get("Authorization"))
	if authorization == "" {
		return ""
	}
	if space := strings.IndexByte(authorization, ' '); space > 0 {
		return authorization[:space]
	}
	return authorization
}

// serverVariableFromHeader converts an HTTP header name to classic ASP
// ServerVariables key format: HTTP_<UPPERCASE_WITH_UNDERSCORES>.
func serverVariableFromHeader(headerName string) string {
	normalized := strings.ToUpper(strings.ReplaceAll(headerName, "-", "_"))
	return "HTTP_" + normalized
}

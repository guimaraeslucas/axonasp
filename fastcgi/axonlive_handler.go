//go:build !lib_g3axonlive_disabled

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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

// Package main provides the G3AxonLive HTTP communication endpoint for the AxonASP FastCGI server.
// This file registers the /g3al/ route for synchronous fetch POST requests only.
//
// NOTE: WebSocket upgrades are NOT supported in FastCGI mode because the FastCGI protocol
// transports requests via fcgi.Serve which does not expose a raw TCP connection for hijacking.
package main

import (
	"bytes"
	"encoding/json"
	"io"
	"net/http"
	"net/url"
	"path/filepath"
	"strings"

	"g3pix.com.br/axonasp/axonvm"
)

// fcgiG3alEndpoint is the URL prefix for G3AxonLive communication in FastCGI mode.
const fcgiG3alEndpoint = "/g3al/"

// fcgiG3alEndpointNoSlash avoids implicit redirect behavior for /g3al requests.
const fcgiG3alEndpointNoSlash = "/g3al"

// fcgiG3alMaxBodyBytes limits the fetch POST body to 256 KiB in FastCGI mode.
const fcgiG3alMaxBodyBytes = 256 * 1024

// FCGIAxonLiveEvent is the JSON payload sent by the browser client.
type FCGIAxonLiveEvent struct {
	SessionID   string            `json:"sessionId"`
	ComponentID string            `json:"componentId"`
	EventName   string            `json:"eventName"`
	EventArgs   map[string]string `json:"eventArgs,omitempty"`
}

// FCGIAxonLiveResponse and FCGIAxonPatch are type aliases for the shared axonvm types.
// Using the axonvm package types ensures the JSON serialization is always consistent
// between the VM library and the FastCGI handler layer.
type FCGIAxonLiveResponse = axonvm.G3ALResponse
type FCGIAxonPatch = axonvm.G3ALPatch

// RegisterG3AxonLiveEndpoint checks the G3AxonLiveActive flag and, when active, starts the
// background session cleanup goroutine and registers the /g3al/ handler on the provided mux.
// WebSocket is not registered in FastCGI mode.
func RegisterG3AxonLiveEndpoint(mux *http.ServeMux) {
	if !G3AxonLiveActive {
		return
	}
	axonvm.G3ALStartCleanup(30)
	mux.HandleFunc(fcgiG3alEndpoint, fastCGIMiddleware(handleFCGIG3AxonLive))
	mux.HandleFunc(fcgiG3alEndpointNoSlash, fastCGIMiddleware(handleFCGIG3AxonLive))
}

// ShutdownG3AxonLiveEndpoint stops the background session cleanup goroutine.
func ShutdownG3AxonLiveEndpoint() {
	axonvm.G3ALStopCleanup()
}

// handleFCGIG3AxonLive dispatches /g3al/ requests in FastCGI mode.
func handleFCGIG3AxonLive(w http.ResponseWriter, r *http.Request) {
	if strings.EqualFold(r.Header.Get("Upgrade"), "websocket") {
		http.Error(w, "WebSocket is not supported in FastCGI mode. Use the AxonASP HTTP server for real-time connections.", http.StatusNotImplemented)
		return
	}
	if r.Method == http.MethodPost {
		handleFCGIG3AxonLiveFetch(w, r)
		return
	}
	http.Error(w, "Method Not Allowed", http.StatusMethodNotAllowed)
}

// handleFCGIG3AxonLiveFetch processes a synchronous fetch POST from the G3AxonLive client engine
// in FastCGI mode. It validates the request, looks up the ASP page registered for the session,
// then re-executes that page with the original event body so it can return JSON component patches.
func handleFCGIG3AxonLiveFetch(w http.ResponseWriter, r *http.Request) {
	if !strings.EqualFold(r.Header.Get("X-G3AxonLive"), "true") {
		writeFCGIG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALMissingXHeader])
		return
	}

	body, err := io.ReadAll(io.LimitReader(r.Body, fcgiG3alMaxBodyBytes))
	if err != nil {
		writeFCGIG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALBodyReadFailed])
		return
	}
	_ = r.Body.Close()

	var event FCGIAxonLiveEvent
	if err := json.Unmarshal(body, &event); err != nil {
		writeFCGIG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidJSONPayload])
		return
	}

	if strings.TrimSpace(event.SessionID) == "" {
		writeFCGIG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidSessionID])
		return
	}
	if strings.TrimSpace(event.ComponentID) == "" {
		writeFCGIG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidComponentID])
		return
	}
	if strings.TrimSpace(event.EventName) == "" {
		writeFCGIG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidEventName])
		return
	}

	// Look up which ASP page is registered for this session.
	scriptURL := resolveFCGIG3ALScriptURL(r, event.SessionID)
	if scriptURL == "" {
		writeFCGIG3AlJSONError(w, http.StatusNotFound, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALSessionNotRegistered])
		return
	}

	// Resolve to a filesystem path within RootDir.
	relativePath := strings.TrimPrefix(scriptURL, "/")
	fullPath := filepath.Join(RootDir, filepath.FromSlash(relativePath))

	// Security: ensure the resolved path stays within the configured web root.
	absRoot, rootErr := filepath.Abs(RootDir)
	absPath, pathErr := filepath.Abs(fullPath)
	if rootErr != nil || pathErr != nil || !strings.HasPrefix(absPath+string(filepath.Separator), absRoot+string(filepath.Separator)) {
		writeFCGIG3AlJSONError(w, http.StatusForbidden, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALPagePathOutsideRoot])
		return
	}

	// Clone the request, re-inject the original body, and fix the URL path so
	// SCRIPT_NAME reflects the real page path, not /g3al/.
	r2 := r.Clone(r.Context())
	r2.Method = http.MethodPost
	r2.URL.Path = scriptURL
	if r2.Header == nil {
		r2.Header = make(http.Header)
	}
	r2.Header.Set("Content-Type", "application/json")
	eventArgsJSON, _ := json.Marshal(event.EventArgs)
	r2.Header.Set("X-G3AxonLive-SessionId", event.SessionID)
	r2.Header.Set("X-G3AxonLive-ComponentId", event.ComponentID)
	r2.Header.Set("X-G3AxonLive-EventName", event.EventName)
	r2.Header.Set("X-G3AxonLive-EventArgs", string(eventArgsJSON))
	r2.Body = io.NopCloser(bytes.NewReader(body))
	r2.ContentLength = int64(len(body))

	// Execute the ASP page — it detects the X-G3AxonLive header, processes the event,
	// and writes the JSON patch response directly.
	executeASP(w, r2, absPath)
}

// resolveFCGIG3ALScriptURL resolves the ASP page path for an incoming event.
// It first uses the payload sessionId, then falls back to ASPSESSIONID cookie,
// and finally to same-origin Referer path to tolerate registry edge cases.
func resolveFCGIG3ALScriptURL(r *http.Request, sessionID string) string {
	scriptURL := axonvm.G3ALGetPageForSession(strings.TrimSpace(sessionID))
	if scriptURL != "" {
		return scriptURL
	}

	if c, err := r.Cookie("ASPSESSIONID"); err == nil && c != nil {
		scriptURL = axonvm.G3ALGetPageForSession(strings.TrimSpace(c.Value))
		if scriptURL != "" {
			return scriptURL
		}
	}

	ref := strings.TrimSpace(r.Referer())
	if ref == "" {
		return ""
	}
	u, err := url.Parse(ref)
	if err != nil {
		return ""
	}
	if u.Path == "" || u.Path == fcgiG3alEndpoint || u.Path == fcgiG3alEndpointNoSlash {
		return ""
	}
	if strings.HasPrefix(u.Path, "/") {
		return u.Path
	}
	return ""
}

// writeFCGIG3AlJSONError writes a JSON-encoded FCGIAxonLiveResponse with Success=false.
func writeFCGIG3AlJSONError(w http.ResponseWriter, status int, message string) {
	resp := FCGIAxonLiveResponse{Success: false, Error: message}
	data, _ := json.Marshal(resp)
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.Header().Set("Cache-Control", "no-store")
	w.WriteHeader(status)
	_, _ = w.Write(data)
}

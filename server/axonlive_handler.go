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

// Package main provides the G3AxonLive HTTP communication endpoint for the AxonASP HTTP server.
// This file implements the /g3al/ route that accepts both synchronous fetch POST requests and
// long-lived WebSocket connections for real-time reactive component updates.
package main

import (
	"bytes"
	"crypto/sha1"
	"encoding/base64"
	"encoding/binary"
	"encoding/json"
	"io"
	"net/http"
	"net/url"
	"path/filepath"
	"strings"

	"g3pix.com.br/axonasp/axonvm"
)

// g3alEndpoint is the URL prefix for all G3AxonLive communication. Every fetch
// POST and WebSocket upgrade for the reactive component framework is handled here.
const g3alEndpoint = "/g3al/"

// g3alEndpointNoSlash avoids implicit redirect behavior for /g3al requests.
const g3alEndpointNoSlash = "/g3al"

// wsGUID is the RFC 6455 magic GUID concatenated with the client Sec-WebSocket-Key
// before SHA-1 hashing to produce Sec-WebSocket-Accept.
const wsGUID = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"

// wsMaxPayloadBytes is the maximum accepted WebSocket frame payload length (1 MiB).
const wsMaxPayloadBytes = 1 * 1024 * 1024

// g3alMaxBodyBytes is the maximum allowed POST body size for fetch requests (256 KiB).
const g3alMaxBodyBytes = 256 * 1024

// WebSocket opcodes as defined in RFC 6455 §5.2.
const (
	wsOpcodeContinuation = 0x0
	wsOpcodeText         = 0x1
	wsOpcodeBinary       = 0x2
	wsOpcodeClose        = 0x8
	wsOpcodePing         = 0x9
	wsOpcodePong         = 0xA
)

// G3AxonLiveEvent is the JSON payload sent by the browser client.
type G3AxonLiveEvent struct {
	SessionID   string            `json:"sessionId"`
	ComponentID string            `json:"componentId"`
	EventName   string            `json:"eventName"`
	EventArgs   map[string]string `json:"eventArgs,omitempty"`
}

// G3AxonLiveResponse and G3AxonPatch are type aliases for the shared axonvm types.
// Using the axonvm package types ensures the JSON serialization is always consistent
// between the VM library and the HTTP handler layer.
type G3AxonLiveResponse = axonvm.G3ALResponse
type G3AxonPatch = axonvm.G3ALPatch

// RegisterG3AxonLiveEndpoint checks the G3AxonLiveActive flag and, when active, starts the
// background session cleanup goroutine and registers the /g3al/ handler on the provided mux.
func RegisterG3AxonLiveEndpoint(mux *http.ServeMux) {
	if !G3AxonLiveActive {
		return
	}
	axonvm.G3ALStartCleanup(30)
	mux.HandleFunc(g3alEndpoint, handleG3AxonLive)
	mux.HandleFunc(g3alEndpointNoSlash, handleG3AxonLive)
}

// ShutdownG3AxonLiveEndpoint stops the background session cleanup goroutine.
// Call this during server shutdown when G3AxonLive is active.
func ShutdownG3AxonLiveEndpoint() {
	axonvm.G3ALStopCleanup()
}

// handleG3AxonLive is the central dispatcher for all /g3al/ requests.
func handleG3AxonLive(w http.ResponseWriter, r *http.Request) {
	if isWebSocketUpgrade(r) {
		handleG3AxonLiveWebSocket(w, r)
		return
	}
	if r.Method == http.MethodPost {
		handleG3AxonLiveFetch(w, r)
		return
	}
	http.Error(w, "Method Not Allowed", http.StatusMethodNotAllowed)
}

// handleG3AxonLiveFetch processes a synchronous fetch POST from the G3AxonLive client engine.
// It validates the request, looks up the ASP page registered for the session, then re-executes
// that page with the original event body so the page can return JSON component patches.
func handleG3AxonLiveFetch(w http.ResponseWriter, r *http.Request) {
	if !strings.EqualFold(r.Header.Get("X-G3AxonLive"), "true") {
		writeG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALMissingXHeader])
		return
	}

	body, err := io.ReadAll(io.LimitReader(r.Body, g3alMaxBodyBytes))
	if err != nil {
		writeG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALBodyReadFailed])
		return
	}
	_ = r.Body.Close()

	var event G3AxonLiveEvent
	if err := json.Unmarshal(body, &event); err != nil {
		writeG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidJSONPayload])
		return
	}

	if strings.TrimSpace(event.SessionID) == "" {
		writeG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidSessionID])
		return
	}
	if strings.TrimSpace(event.ComponentID) == "" {
		writeG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidComponentID])
		return
	}
	if strings.TrimSpace(event.EventName) == "" {
		writeG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidEventName])
		return
	}

	// Look up which ASP page is registered for this session.
	scriptURL := resolveG3ALScriptURL(r, event.SessionID)
	if scriptURL == "" {
		writeG3AlJSONError(w, http.StatusNotFound, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALSessionNotRegistered])
		return
	}

	// Resolve the registered URL to a filesystem path within RootDir.
	relativePath := strings.TrimPrefix(scriptURL, "/")
	fullPath := filepath.Join(RootDir, filepath.FromSlash(relativePath))

	// Security: ensure the resolved path stays within the configured web root.
	absRoot, rootErr := filepath.Abs(RootDir)
	absPath, pathErr := filepath.Abs(fullPath)
	if rootErr != nil || pathErr != nil || !strings.HasPrefix(absPath+string(filepath.Separator), absRoot+string(filepath.Separator)) {
		writeG3AlJSONError(w, http.StatusForbidden, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALPagePathOutsideRoot])
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
	// updates component state, and writes the JSON patch response directly.
	executeASP(w, r2, absPath)
}

// resolveG3ALScriptURL resolves the ASP page path for an incoming event.
// It first uses the payload sessionId, then falls back to ASPSESSIONID cookie,
// and finally to same-origin Referer path to tolerate registry edge cases.
func resolveG3ALScriptURL(r *http.Request, sessionID string) string {
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
	if u.Path == "" || u.Path == g3alEndpoint || u.Path == g3alEndpointNoSlash {
		return ""
	}
	if strings.HasPrefix(u.Path, "/") {
		return u.Path
	}
	return ""
}

// handleG3AxonLiveWebSocket upgrades an HTTP connection to a WebSocket per RFC 6455.
func handleG3AxonLiveWebSocket(w http.ResponseWriter, r *http.Request) {
	clientKey := strings.TrimSpace(r.Header.Get("Sec-WebSocket-Key"))
	if clientKey == "" {
		http.Error(w, "Bad Request: missing Sec-WebSocket-Key", http.StatusBadRequest)
		return
	}

	acceptKey := computeWebSocketAccept(clientKey)

	hijacker, ok := w.(http.Hijacker)
	if !ok {
		http.Error(w, "Internal Server Error: WebSocket not supported by this transport", http.StatusInternalServerError)
		return
	}

	conn, buf, err := hijacker.Hijack()
	if err != nil {
		http.Error(w, "Internal Server Error: connection hijack failed", http.StatusInternalServerError)
		return
	}
	defer conn.Close()

	_, err = buf.WriteString(
		"HTTP/1.1 101 Switching Protocols\r\n" +
			"Upgrade: websocket\r\n" +
			"Connection: Upgrade\r\n" +
			"Sec-WebSocket-Accept: " + acceptKey + "\r\n" +
			"\r\n",
	)
	if err != nil {
		return
	}
	if err = buf.Flush(); err != nil {
		return
	}

	serveG3AxonLiveWebSocketLoop(conn, buf.Reader)
}

// serveG3AxonLiveWebSocketLoop is the main I/O loop for an established G3AxonLive WebSocket connection.
func serveG3AxonLiveWebSocketLoop(w io.Writer, r io.Reader) {
	for {
		opcode, payload, err := readWebSocketFrame(r)
		if err != nil {
			return
		}

		switch opcode {
		case wsOpcodeClose:
			_ = writeWebSocketFrame(w, wsOpcodeClose, nil)
			return

		case wsOpcodePing:
			_ = writeWebSocketFrame(w, wsOpcodePong, payload)

		case wsOpcodeText, wsOpcodeBinary:
			var event G3AxonLiveEvent
			if jsonErr := json.Unmarshal(payload, &event); jsonErr != nil {
				errData, _ := json.Marshal(G3AxonLiveResponse{Success: false, Error: axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidJSONPayload]})
				_ = writeWebSocketFrame(w, wsOpcodeText, errData)
				continue
			}

			if strings.TrimSpace(event.SessionID) == "" ||
				strings.TrimSpace(event.ComponentID) == "" ||
				strings.TrimSpace(event.EventName) == "" {
				errData, _ := json.Marshal(G3AxonLiveResponse{Success: false, Error: axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidSessionID]})
				_ = writeWebSocketFrame(w, wsOpcodeText, errData)
				continue
			}

			// TODO (Phase 3): WebSocket path — execute ASP page and push patches.
			resp := G3AxonLiveResponse{
				Success:    true,
				Components: []G3AxonPatch{},
			}
			data, _ := json.Marshal(resp)
			_ = writeWebSocketFrame(w, wsOpcodeText, data)

		case wsOpcodeContinuation:
			// Fragmented messages not yet supported.
		}
	}
}

// isWebSocketUpgrade reports whether the HTTP request is a WebSocket upgrade.
func isWebSocketUpgrade(r *http.Request) bool {
	return strings.EqualFold(r.Header.Get("Upgrade"), "websocket") &&
		strings.Contains(strings.ToLower(r.Header.Get("Connection")), "upgrade")
}

// computeWebSocketAccept derives the Sec-WebSocket-Accept response header value.
func computeWebSocketAccept(clientKey string) string {
	h := sha1.New()
	_, _ = io.WriteString(h, clientKey+wsGUID)
	return base64.StdEncoding.EncodeToString(h.Sum(nil))
}

// readWebSocketFrame reads a single RFC 6455 WebSocket frame from the reader.
func readWebSocketFrame(r io.Reader) (opcode byte, payload []byte, err error) {
	var header [2]byte
	if _, err = io.ReadFull(r, header[:]); err != nil {
		return 0, nil, err
	}

	opcode = header[0] & 0x0F
	masked := (header[1] & 0x80) != 0
	rawLen := int64(header[1] & 0x7F)

	switch rawLen {
	case 126:
		var ext [2]byte
		if _, err = io.ReadFull(r, ext[:]); err != nil {
			return 0, nil, err
		}
		rawLen = int64(binary.BigEndian.Uint16(ext[:]))
	case 127:
		var ext [8]byte
		if _, err = io.ReadFull(r, ext[:]); err != nil {
			return 0, nil, err
		}
		rawLen = int64(binary.BigEndian.Uint64(ext[:]))
	}

	if rawLen > wsMaxPayloadBytes {
		return 0, nil, io.ErrUnexpectedEOF
	}

	var maskKey [4]byte
	if masked {
		if _, err = io.ReadFull(r, maskKey[:]); err != nil {
			return 0, nil, err
		}
	}

	payload = make([]byte, int(rawLen))
	if _, err = io.ReadFull(r, payload); err != nil {
		return 0, nil, err
	}

	if masked {
		for i := range payload {
			payload[i] ^= maskKey[i%4]
		}
	}

	return opcode, payload, nil
}

// writeWebSocketFrame encodes and writes a single unmasked RFC 6455 WebSocket frame.
func writeWebSocketFrame(w io.Writer, opcode byte, payload []byte) error {
	length := len(payload)
	firstByte := byte(0x80) | (opcode & 0x0F)

	var header []byte
	switch {
	case length <= 125:
		header = []byte{firstByte, byte(length)}
	case length <= 65535:
		header = []byte{firstByte, 126, byte(length >> 8), byte(length)}
	default:
		n := uint64(length)
		header = []byte{
			firstByte, 127,
			byte(n >> 56), byte(n >> 48), byte(n >> 40), byte(n >> 32),
			byte(n >> 24), byte(n >> 16), byte(n >> 8), byte(n),
		}
	}

	frame := make([]byte, len(header)+length)
	copy(frame, header)
	copy(frame[len(header):], payload)

	_, err := w.Write(frame)
	return err
}

// writeG3AlJSONError writes a JSON-encoded G3AxonLiveResponse with Success=false.
func writeG3AlJSONError(w http.ResponseWriter, status int, message string) {
	resp := G3AxonLiveResponse{Success: false, Error: message}
	data, _ := json.Marshal(resp)
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.Header().Set("Cache-Control", "no-store")
	w.WriteHeader(status)
	_, _ = w.Write(data)
}

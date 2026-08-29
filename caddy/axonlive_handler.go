package caddy

import (
	"bytes"
	"crypto/sha1"
	"encoding/base64"
	"encoding/binary"
	"encoding/json"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"

	"g3pix.com.br/axonasp/v2/axonvm"
	"github.com/caddyserver/caddy/v2/modules/caddyhttp"
)

// g3alEndpoint is the URL prefix for all G3AxonLive communication.
const g3alEndpoint = "/g3al/"

// g3alEndpointNoSlash avoids implicit redirect behavior for /g3al requests.
const g3alEndpointNoSlash = "/g3al"

// wsGUID is the RFC 6455 magic GUID concatenated with the client Sec-WebSocket-Key.
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
type G3AxonLiveResponse = axonvm.G3ALResponse
type G3AxonPatch = axonvm.G3ALPatch

func isAxonLiveEndpoint(path string) bool {
	lower := strings.ToLower(path)
	return lower == g3alEndpointNoSlash || lower == g3alEndpoint || strings.HasPrefix(lower, g3alEndpoint)
}

func (a *AxonASP) isG3AxonLiveActive() bool {
	if a.config == nil {
		return true
	}
	if a.config.IsSet("g3axonlive.g3axonlive_active") {
		return a.config.GetBool("g3axonlive.g3axonlive_active")
	}
	return true
}

func (a *AxonASP) resolveWebRoot(r *http.Request) string {
	if a.GlobalAsaPath != "" {
		return filepath.Dir(a.GlobalAsaPath)
	}
	if docRoot, ok := caddyhttp.GetVar(r.Context(), "root").(string); ok && docRoot != "" {
		return docRoot
	}
	return "."
}

func (a *AxonASP) handleAxonLive(w http.ResponseWriter, r *http.Request) error {
	if isWebSocketUpgrade(r) {
		a.handleAxonLiveWebSocket(w, r)
		return nil
	}
	if r.Method == http.MethodPost {
		a.handleAxonLiveFetch(w, r)
		return nil
	}
	http.Error(w, "Method Not Allowed", http.StatusMethodNotAllowed)
	return nil
}

func (a *AxonASP) handleAxonLiveFetch(w http.ResponseWriter, r *http.Request) {
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

	if strings.TrimSpace(event.ComponentID) == "" {
		writeG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidComponentID])
		return
	}
	if strings.TrimSpace(event.EventName) == "" {
		writeG3AlJSONError(w, http.StatusBadRequest, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidEventName])
		return
	}

	authenticatedSessionID := authenticatedSessionIDFromRequest(r)
	normalizedSessionID, authorized := normalizeAndAuthorizeG3ALSessionID(event.SessionID, authenticatedSessionID)
	if !authorized {
		writeG3AlJSONError(w, http.StatusForbidden, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALInvalidSessionID])
		return
	}
	event.SessionID = normalizedSessionID

	// Look up which ASP page is registered for the authenticated session.
	scriptURL := resolveG3ALScriptURL(event.SessionID)
	if scriptURL == "" {
		writeG3AlJSONError(w, http.StatusNotFound, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALSessionNotRegistered])
		return
	}

	webRoot := a.resolveWebRoot(r)

	// Resolve the registered URL to a filesystem path within webRoot.
	relativePath := strings.TrimPrefix(scriptURL, "/")
	fullPath := filepath.Join(webRoot, filepath.FromSlash(relativePath))

	// Security: ensure the resolved path stays within the configured web root.
	absRoot, rootErr := filepath.Abs(webRoot)
	absPath, pathErr := filepath.Abs(fullPath)
	if rootErr != nil || pathErr != nil || !strings.HasPrefix(absPath+string(filepath.Separator), absRoot+string(filepath.Separator)) {
		writeG3AlJSONError(w, http.StatusForbidden, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALPagePathOutsideRoot])
		return
	}

	if _, err := os.Stat(absPath); err != nil {
		writeG3AlJSONError(w, http.StatusNotFound, axonvm.AxonASPErrorMessages[axonvm.ErrG3ALSessionNotRegistered])
		return
	}

	// Clone the request, re-inject the original body, and fix the URL path so
	// SCRIPT_NAME reflects the real page path, not /g3al/.
	r2 := r.Clone(r.Context())
	r2.Method = http.MethodPost
	r2.URL.Path = scriptURL
	normalizedBody, _ := json.Marshal(event)
	if r2.Header == nil {
		r2.Header = make(http.Header)
	}
	r2.Header.Set("Content-Type", "application/json")
	eventArgsJSON, _ := json.Marshal(event.EventArgs)
	r2.Header.Set("X-G3AxonLive-SessionId", event.SessionID)
	r2.Header.Set("X-G3AxonLive-ComponentId", event.ComponentID)
	r2.Header.Set("X-G3AxonLive-EventName", event.EventName)
	r2.Header.Set("X-G3AxonLive-EventArgs", string(eventArgsJSON))
	r2.Body = io.NopCloser(bytes.NewReader(normalizedBody))
	r2.ContentLength = int64(len(normalizedBody))

	// Execute the ASP page — it detects the X-G3AxonLive header, processes the event,
	// updates component state, and writes the JSON patch response directly.
	_ = a.executeASPFile(w, r2, webRoot, absPath)
}

func resolveG3ALScriptURL(sessionID string) string {
	return axonvm.G3ALGetPageForSession(strings.TrimSpace(sessionID))
}

func authenticatedSessionIDFromRequest(r *http.Request) string {
	if c, err := r.Cookie("ASPSESSIONID"); err == nil && c != nil {
		return strings.TrimSpace(c.Value)
	}
	return ""
}

func normalizeAndAuthorizeG3ALSessionID(payloadSessionID, authenticatedSessionID string) (string, bool) {
	authID := strings.TrimSpace(authenticatedSessionID)
	if authID == "" {
		return "", false
	}
	eventID := strings.TrimSpace(payloadSessionID)
	if eventID != "" && !strings.EqualFold(eventID, authID) {
		return "", false
	}
	return authID, true
}

func (a *AxonASP) handleAxonLiveWebSocket(w http.ResponseWriter, r *http.Request) {
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

			resp := G3AxonLiveResponse{
				Success:    true,
				Components: []G3AxonPatch{},
			}
			data, _ := json.Marshal(resp)
			_ = writeWebSocketFrame(w, wsOpcodeText, data)

		case wsOpcodeContinuation:
		}
	}
}

func isWebSocketUpgrade(r *http.Request) bool {
	return strings.EqualFold(r.Header.Get("Upgrade"), "websocket") &&
		strings.Contains(strings.ToLower(r.Header.Get("Connection")), "upgrade")
}

func computeWebSocketAccept(clientKey string) string {
	h := sha1.New()
	_, _ = io.WriteString(h, clientKey+wsGUID)
	return base64.StdEncoding.EncodeToString(h.Sum(nil))
}

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

func writeG3AlJSONError(w http.ResponseWriter, status int, message string) {
	resp := G3AxonLiveResponse{Success: false, Error: message}
	data, _ := json.Marshal(resp)
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.Header().Set("Cache-Control", "no-store")
	w.WriteHeader(status)
	_, _ = w.Write(data)
}

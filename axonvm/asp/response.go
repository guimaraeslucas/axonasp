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
	"io"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"time"
)

// ResponseEndSignal is used to terminate script execution after Response.End/Redirect.
const ResponseEndSignal = "RESPONSE_END"

// ResponseCookie stores one response cookie and its optional properties.
type ResponseCookie struct {
	Name       string
	Value      string
	Domain     string
	Path       string
	ExpiresRaw string
	Secure     bool
	HTTPOnly   bool
}

// Response controls the output sent to the client.
type Response struct {
	Output io.Writer
	w      http.ResponseWriter
	req    *http.Request

	buffer         []byte
	mu             sync.RWMutex
	ended          bool
	flushed        bool
	maxBufferBytes int

	bufferEnabled bool
	cacheControl  string
	charset       string
	codePage      int
	contentType   string
	expires       int
	expiresAbsRaw string
	pics          string
	status        string

	headers    map[string]string
	cookies    map[string]*ResponseCookie
	logEntries []string
}

// DefaultResponseBufferLimitBytes defines the default buffered response safety limit.
const DefaultResponseBufferLimitBytes = 4 * 1024 * 1024

// ResponseBufferLimitError reports buffered output exceeding the configured limit.
type ResponseBufferLimitError struct {
	LimitBytes   int
	CurrentBytes int
}

// Error returns one human-readable buffered output limit failure message.
func (e *ResponseBufferLimitError) Error() string {
	if e == nil {
		return ""
	}
	return fmt.Sprintf("Response buffering exceeded the configured limit of %d bytes", e.LimitBytes)
}

// NewResponse creates a new response object with ASP-compatible defaults.
func NewResponse(output io.Writer) *Response {
	r := &Response{
		Output:         output,
		buffer:         make([]byte, 0),
		bufferEnabled:  true,
		maxBufferBytes: DefaultResponseBufferLimitBytes,
		cacheControl:   "Private",
		charset:        "utf-8",
		contentType:    "text/html",
		status:         "200 OK",
		headers:        make(map[string]string),
		cookies:        make(map[string]*ResponseCookie),
		logEntries:     make([]string, 0),
		codePage:       65001,
	}
	if w, ok := output.(http.ResponseWriter); ok {
		r.w = w
	}
	return r
}

// SetMaxBufferBytes updates the buffered response safety limit.
func (r *Response) SetMaxBufferBytes(limit int) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if limit <= 0 {
		limit = DefaultResponseBufferLimitBytes
	}
	r.maxBufferBytes = limit
}

// GetMaxBufferBytes returns the buffered response safety limit in bytes.
func (r *Response) GetMaxBufferBytes() int {
	r.mu.RLock()
	defer r.mu.RUnlock()
	if r.maxBufferBytes <= 0 {
		return DefaultResponseBufferLimitBytes
	}
	return r.maxBufferBytes
}

// ensureBufferCapacityLocked validates the pending buffer size while the response mutex is held.
func (r *Response) ensureBufferCapacityLocked(nextSize int) {
	limit := r.maxBufferBytes
	if limit <= 0 {
		limit = DefaultResponseBufferLimitBytes
	}
	if r.bufferEnabled && nextSize > limit {
		panic(&ResponseBufferLimitError{LimitBytes: limit, CurrentBytes: nextSize})
	}
}

// SetCodePage sets the response code page using ASP-compatible fallback rules.
func (r *Response) SetCodePage(codePage int) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	if codePage <= 0 {
		r.codePage = 65001
	} else {
		r.codePage = codePage
	}
	if r.codePage == 65001 && r.charset == "" {
		r.charset = "utf-8"
	}
}

// GetCodePage returns the current response code page.
func (r *Response) GetCodePage() int {
	r.mu.RLock()
	defer r.mu.RUnlock()
	if r.codePage <= 0 {
		return 65001
	}
	return r.codePage
}

// SetRequest stores the HTTP request for IsClientConnected checks.
func (r *Response) SetRequest(req *http.Request) {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.req = req
}

// contentTypeSupportsCharset reports whether a content type semantically accepts a charset parameter.
func contentTypeSupportsCharset(contentType string) bool {
	mediaType := strings.ToLower(strings.TrimSpace(contentType))
	if mediaType == "" {
		return true
	}
	if semi := strings.IndexByte(mediaType, ';'); semi >= 0 {
		mediaType = strings.TrimSpace(mediaType[:semi])
	}
	if strings.HasPrefix(mediaType, "text/") {
		return true
	}
	if strings.HasSuffix(mediaType, "+json") || strings.HasSuffix(mediaType, "+xml") {
		return true
	}
	switch mediaType {
	case "application/json", "application/javascript", "application/xml", "application/xhtml+xml", "application/rss+xml", "application/atom+xml", "application/x-www-form-urlencoded", "image/svg+xml":
		return true
	default:
		return false
	}
}

// Write appends a string to the HTTP output buffer.
func (r *Response) Write(s string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.ended || r.Output == nil {
		return
	}
	r.ensureBufferCapacityLocked(len(r.buffer) + len(s))
	r.buffer = append(r.buffer, s...)
	if !r.bufferEnabled {
		r.flushInternal()
	}
}

// BinaryWrite appends binary bytes to the response buffer.
func (r *Response) BinaryWrite(data []byte) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.ended || r.Output == nil {
		return
	}
	r.ensureBufferCapacityLocked(len(r.buffer) + len(data))
	r.buffer = append(r.buffer, data...)
	if !r.bufferEnabled {
		r.flushInternal()
	}
}

// AddHeader sets an HTTP header if output was not flushed yet.
func (r *Response) AddHeader(name string, value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.flushed || r.ended {
		return
	}
	r.headers[name] = value
}

// AppendToLog appends one message into response internal log entries and runtime log output.
func (r *Response) AppendToLog(message string) {
	r.mu.Lock()
	r.logEntries = append(r.logEntries, message)
	r.mu.Unlock()

	fmt.Println(message)
	appendRuntimeLogLine(message)
}

// appendRuntimeLogLine writes one log line to temp/<runtime>.log using best-effort semantics.
func appendRuntimeLogLine(message string) {
	cwd, err := os.Getwd()
	if err != nil {
		return
	}

	tempDir := filepath.Join(cwd, "temp")
	if mkErr := os.MkdirAll(tempDir, 0o755); mkErr != nil {
		return
	}

	execName := strings.ToLower(filepath.Base(os.Args[0]))
	logName := "server.log"
	switch {
	case strings.Contains(execName, "fastcgi"):
		logName = "fastcgi.log"
	case strings.Contains(execName, "cli"):
		logName = "cli.log"
	}

	line := message + "\n"
	file, openErr := os.OpenFile(filepath.Join(tempDir, logName), os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0o644)
	if openErr != nil {
		return
	}
	defer file.Close()
	_, _ = file.WriteString(line)
}

// Clear clears the current response buffer when response is not ended.
func (r *Response) Clear() {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.ended {
		return
	}
	r.buffer = r.buffer[:0]
}

// Flush sends buffered output to the client.
func (r *Response) Flush() {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.flushInternal()
}

// End stops the script execution.
func (r *Response) End() {
	r.mu.Lock()
	r.flushInternal()
	r.ended = true
	r.mu.Unlock()
	panic(ResponseEndSignal)
}

// Redirect sends a 302 redirect and terminates the current response.
func (r *Response) Redirect(location string) {
	r.mu.Lock()
	if r.ended {
		r.mu.Unlock()
		return
	}
	r.buffer = r.buffer[:0]
	r.headers["Location"] = location
	r.status = "302 Found"
	r.flushInternal()
	r.ended = true
	r.mu.Unlock()
	panic(ResponseEndSignal)
}

// SetContentType sets the Content-Type header.
func (r *Response) SetContentType(contentType string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	r.contentType = contentType
}

// GetContentType returns the configured content type.
func (r *Response) GetContentType() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.contentType
}

// SetBuffer enables or disables response buffering.
func (r *Response) SetBuffer(enabled bool) {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.bufferEnabled = enabled
}

// GetBuffer returns whether buffering is currently enabled.
func (r *Response) GetBuffer() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.bufferEnabled
}

// SetCacheControl sets cache control header value.
func (r *Response) SetCacheControl(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	r.cacheControl = value
}

// GetCacheControl returns cache control header value.
func (r *Response) GetCacheControl() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.cacheControl
}

// SetCharset sets response charset.
func (r *Response) SetCharset(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	r.charset = value
}

// GetCharset returns response charset.
func (r *Response) GetCharset() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.charset
}

// SetExpires sets expires value in minutes.
func (r *Response) SetExpires(minutes int) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	r.expires = minutes
}

// GetExpires returns configured expires minutes.
func (r *Response) GetExpires() int {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.expires
}

// SetExpiresAbsoluteRaw sets absolute expires value as raw string.
func (r *Response) SetExpiresAbsoluteRaw(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	r.expiresAbsRaw = value
}

// GetExpiresAbsoluteRaw returns raw absolute expires value.
func (r *Response) GetExpiresAbsoluteRaw() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.expiresAbsRaw
}

// SetPICS sets PICS label header value.
func (r *Response) SetPICS(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	r.pics = value
}

// GetPICS returns PICS label header value.
func (r *Response) GetPICS() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.pics
}

// SetStatus sets raw status line value.
func (r *Response) SetStatus(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	r.status = value
}

// GetStatus returns raw status line value.
func (r *Response) GetStatus() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.status
}

// IsClientConnected checks whether the HTTP client connection is still alive.
func (r *Response) IsClientConnected() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	if r.req == nil {
		return true
	}
	select {
	case <-r.req.Context().Done():
		return false
	default:
		return true
	}
}

// SetCookieValue sets response cookie value.
func (r *Response) SetCookieValue(name string, value string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}
	key := strings.ToLower(name)
	if cookie, exists := r.cookies[key]; exists {
		cookie.Value = value
		return
	}
	r.cookies[key] = &ResponseCookie{Name: name, Value: value, Path: "/"}
}

// GetCookieValue returns response cookie value.
func (r *Response) GetCookieValue(name string) string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	if cookie, exists := r.cookies[strings.ToLower(name)]; exists {
		return cookie.Value
	}
	return ""
}

// SetCookieProperty sets one cookie property by name.
func (r *Response) SetCookieProperty(cookieName string, propertyName string, propertyValue string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.flushed {
		return
	}

	key := strings.ToLower(cookieName)
	cookie, exists := r.cookies[key]
	if !exists {
		cookie = &ResponseCookie{Name: cookieName, Path: "/"}
		r.cookies[key] = cookie
	}

	switch strings.ToLower(propertyName) {
	case "value":
		cookie.Value = propertyValue
	case "domain":
		cookie.Domain = propertyValue
	case "path":
		cookie.Path = propertyValue
	case "expires":
		cookie.ExpiresRaw = propertyValue
	case "secure":
		cookie.Secure = strings.EqualFold(propertyValue, "true") || propertyValue == "1"
	case "httponly":
		cookie.HTTPOnly = strings.EqualFold(propertyValue, "true") || propertyValue == "1"
	}
}

// GetCookieProperty returns one cookie property by name.
func (r *Response) GetCookieProperty(cookieName string, propertyName string) string {
	r.mu.RLock()
	defer r.mu.RUnlock()

	cookie, exists := r.cookies[strings.ToLower(cookieName)]
	if !exists {
		return ""
	}

	switch strings.ToLower(propertyName) {
	case "name":
		return cookie.Name
	case "value":
		return cookie.Value
	case "domain":
		return cookie.Domain
	case "path":
		return cookie.Path
	case "expires":
		return cookie.ExpiresRaw
	case "secure":
		if cookie.Secure {
			return "True"
		}
		return "False"
	case "httponly":
		if cookie.HTTPOnly {
			return "True"
		}
		return "False"
	default:
		return ""
	}
}

// GetCookieKeys returns a snapshot of all response cookie names.
func (r *Response) GetCookieKeys() []string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	keys := make([]string, 0, len(r.cookies))
	for k := range r.cookies {
		keys = append(keys, k)
	}
	return keys
}

// IsEnded reports whether response was ended.
func (r *Response) IsEnded() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.ended
}

// flushInternal writes headers, cookies, and buffered content to the HTTP output.
func (r *Response) flushInternal() {
	if r.Output == nil {
		return
	}

	if r.w != nil && !r.flushed {
		r.w.Header().Set("Cache-Control", r.cacheControl)
		if r.pics != "" {
			r.w.Header().Set("PICS-Label", r.pics)
		}
		for name, value := range r.headers {
			r.w.Header().Set(name, value)
		}

		contentType := r.contentType
		if r.charset != "" && !strings.Contains(strings.ToLower(contentType), "charset=") && contentTypeSupportsCharset(contentType) {
			contentType = contentType + "; charset=" + r.charset
		}
		r.w.Header().Set("Content-Type", contentType)

		if r.expiresAbsRaw != "" {
			r.w.Header().Set("Expires", r.expiresAbsRaw)
		} else if r.expires != 0 {
			expiresTime := time.Now().Add(time.Duration(r.expires) * time.Minute)
			r.w.Header().Set("Expires", expiresTime.Format(http.TimeFormat))
		}

		for _, cookie := range r.cookies {
			httpCookie := &http.Cookie{
				Name:     cookie.Name,
				Value:    cookie.Value,
				Domain:   cookie.Domain,
				Path:     cookie.Path,
				Secure:   cookie.Secure,
				HttpOnly: cookie.HTTPOnly,
			}
			if httpCookie.Path == "" {
				httpCookie.Path = "/"
			}
			if cookie.ExpiresRaw != "" {
				if unix, err := strconv.ParseInt(cookie.ExpiresRaw, 10, 64); err == nil {
					httpCookie.Expires = time.Unix(unix, 0).UTC()
				} else if parsed, err := time.Parse(time.RFC3339, cookie.ExpiresRaw); err == nil {
					httpCookie.Expires = parsed
				}
			}
			http.SetCookie(r.w, httpCookie)
		}

		if r.status != "" && !strings.EqualFold(r.status, "200 OK") {
			statusCode := 0
			_, _ = fmt.Sscanf(r.status, "%d", &statusCode)
			if statusCode > 0 {
				r.w.WriteHeader(statusCode)
			}
		}

		r.flushed = true
	}

	if len(r.buffer) > 0 {
		_, _ = r.Output.Write(r.buffer)
		r.buffer = r.buffer[:0]
	}

	if flusher, ok := r.Output.(http.Flusher); ok {
		flusher.Flush()
	}
}

// IsClientAbortError checks whether an error indicates client disconnect.
func IsClientAbortError(err error) bool {
	if err == nil {
		return false
	}
	if netErr, ok := err.(*net.OpError); ok {
		message := strings.ToLower(netErr.Error())
		return strings.Contains(message, "broken pipe") || strings.Contains(message, "connection reset")
	}
	message := strings.ToLower(err.Error())
	return strings.Contains(message, "broken pipe") || strings.Contains(message, "connection reset")
}

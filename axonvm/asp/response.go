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
	"bytes"
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
	"unicode/utf16"

	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/encoding/traditionalchinese"
	"golang.org/x/text/transform"
)

// ResponseEndSignal is used to terminate script execution after Response.End/Redirect.
const ResponseEndSignal = "RESPONSE_END"

// responseBufferPool pools pre-allocated *bytes.Buffer instances that back the per-request
// Response body buffer. Reusing these buffers across requests eliminates the
// per-request heap allocation and reduces GC pressure under concurrent load.
const maxPooledBufferCap = 256 * 1024 // 256 KB upper limit for pooled buffers

var responseBufferPool = sync.Pool{
	New: func() any {
		// Start with 24 KB — comfortably covers most classic ASP pages.
		return bytes.NewBuffer(make([]byte, 0, 24*1024))
	},
}

// ResponseCookie stores one response cookie and its optional properties.
type ResponseCookie struct {
	Name             string
	Value            string
	SubKeys          map[string]string
	subKeyOrder      []string
	valueInitialized bool
	Domain           string
	Path             string
	ExpiresRaw       string
	Secure           bool
	HTTPOnly         bool
}

// rebuildSubKeyValue serializes cookie sub-keys as key=value&key2=value2
// in deterministic insertion order.
func (c *ResponseCookie) rebuildSubKeyValue() {
	if len(c.subKeyOrder) == 0 {
		c.Value = ""
		return
	}
	var sb strings.Builder
	for i, k := range c.subKeyOrder {
		if i > 0 {
			sb.WriteByte('&')
		}
		sb.WriteString(k)
		sb.WriteByte('=')
		sb.WriteString(c.SubKeys[k])
	}
	c.Value = sb.String()
}

type responseSegment struct {
	start    int
	end      int
	codePage int
	binary   bool
}

// Response controls the output sent to the client.
type Response struct {
	Output io.Writer
	w      http.ResponseWriter
	req    *http.Request

	buffer         *bytes.Buffer
	segments       []responseSegment
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

	headers     http.Header
	cookies     map[string]*ResponseCookie
	cookieOrder []string
	logEntries  []string
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
	buf := responseBufferPool.Get().(*bytes.Buffer)
	buf.Reset()
	r := &Response{
		Output:         output,
		buffer:         buf,
		segments:       make([]responseSegment, 0, 8),
		bufferEnabled:  true,
		maxBufferBytes: DefaultResponseBufferLimitBytes,
		cacheControl:   "Private",
		charset:        "utf-8",
		contentType:    "text/html",
		status:         "200 OK",
		headers:        make(http.Header),
		cookies:        make(map[string]*ResponseCookie),
		cookieOrder:    make([]string, 0),
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

func responseEncoding(codePage int) encoding.Encoding {
	switch codePage {
	case 37:
		return charmap.CodePage037
	case 437:
		return charmap.CodePage437
	case 850:
		return charmap.CodePage850
	case 852:
		return charmap.CodePage852
	case 855:
		return charmap.CodePage855
	case 858:
		return charmap.CodePage858
	case 860:
		return charmap.CodePage860
	case 862:
		return charmap.CodePage862
	case 863:
		return charmap.CodePage863
	case 865:
		return charmap.CodePage865
	case 866:
		return charmap.CodePage866
	case 874:
		return charmap.Windows874
	case 932:
		return japanese.ShiftJIS
	case 936:
		return simplifiedchinese.GBK
	case 949:
		return korean.EUCKR
	case 950:
		return traditionalchinese.Big5
	case 1250:
		return charmap.Windows1250
	case 1251:
		return charmap.Windows1251
	case 1252:
		return charmap.Windows1252
	case 1253:
		return charmap.Windows1253
	case 1254:
		return charmap.Windows1254
	case 1255:
		return charmap.Windows1255
	case 1256:
		return charmap.Windows1256
	case 1257:
		return charmap.Windows1257
	case 1258:
		return charmap.Windows1258
	default:
		return nil
	}
}

func encodeResponseBody(data []byte, codePage int) []byte {
	enc := responseEncoding(codePage)
	if enc == nil || len(data) == 0 {
		return data
	}
	var encoded bytes.Buffer
	encoded.Grow(len(data))
	for _, codeUnit := range utf16.Encode([]rune(string(data))) {
		part, _, err := transform.String(enc.NewEncoder(), string(rune(codeUnit)))
		if err != nil {
			encoded.WriteByte('?')
			continue
		}
		encoded.WriteString(part)
	}
	return encoded.Bytes()
}

func (r *Response) appendSegmentLocked(start, end int, binary bool) {
	if start >= end {
		return
	}
	if count := len(r.segments); count > 0 {
		last := &r.segments[count-1]
		if last.end == start && last.binary == binary && (binary || last.codePage == r.codePage) {
			last.end = end
			return
		}
	}
	r.segments = append(r.segments, responseSegment{start: start, end: end, codePage: r.codePage, binary: binary})
}

// Write appends a string to the HTTP output buffer.
func (r *Response) Write(s string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.ended || r.buffer == nil {
		return
	}
	start := r.buffer.Len()
	r.ensureBufferCapacityLocked(r.buffer.Len() + len(s))
	// Optimization: Use WriteString to avoid heap allocation from implicit []byte(s) conversion.
	_, _ = r.buffer.WriteString(s)
	r.appendSegmentLocked(start, r.buffer.Len(), false)
	if !r.bufferEnabled {
		r.flushInternal()
	}
}

// BinaryWrite appends binary bytes to the response buffer.
func (r *Response) BinaryWrite(data []byte) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.ended || r.buffer == nil {
		return
	}
	start := r.buffer.Len()
	r.ensureBufferCapacityLocked(r.buffer.Len() + len(data))
	_, _ = r.buffer.Write(data)
	r.appendSegmentLocked(start, r.buffer.Len(), true)
	if !r.bufferEnabled {
		r.flushInternal()
	}
}

// AddHeader adds an HTTP header if output was not flushed yet. Classic ASP
// permits repeated Set-Cookie values, but a later value for the same cookie
// name replaces the earlier pending value.
func (r *Response) AddHeader(name string, value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.flushed || r.ended {
		return
	}
	if strings.EqualFold(name, "Set-Cookie") {
		value = normalizeSetCookiePath(value)
		replaceSetCookieHeader(r.headers, value)
		return
	}
	r.headers.Add(name, value)
}

func replaceSetCookieHeader(headers http.Header, value string) {
	name := setCookieName(value)
	if name == "" {
		headers.Add("Set-Cookie", value)
		return
	}

	existing := headers.Values("Set-Cookie")
	headers.Del("Set-Cookie")
	for _, current := range existing {
		if !strings.EqualFold(setCookieName(current), name) {
			headers.Add("Set-Cookie", current)
		}
	}
	headers.Add("Set-Cookie", value)
}

func setCookieName(value string) string {
	pair := value
	if separator := strings.IndexByte(pair, ';'); separator >= 0 {
		pair = pair[:separator]
	}
	equals := strings.IndexByte(pair, '=')
	if equals <= 0 {
		return ""
	}
	return strings.TrimSpace(pair[:equals])
}

// normalizeSetCookiePath prevents legacy apps that concatenate a customer
// prefix and application path from issuing Path=//. RFC 6265 requires a
// single leading slash; normalizing here makes the cookie usable by browsers
// without changing application code.
func normalizeSetCookiePath(value string) string {
	parts := strings.Split(value, ";")
	for i := range parts {
		trimmed := strings.TrimSpace(parts[i])
		if len(trimmed) < len("path=") || !strings.EqualFold(trimmed[:len("path=")], "path=") {
			continue
		}
		path := strings.TrimSpace(trimmed[len("path="):])
		if strings.HasPrefix(path, "//") {
			path = "/" + strings.TrimLeft(path, "/")
			parts[i] = " path=" + path
		}
	}
	return strings.Join(parts, ";")
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
	tempDir := resolveConfiguredTempDir()
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

	if r.ended || r.buffer == nil {
		return
	}
	r.buffer.Reset()
	r.segments = r.segments[:0]
	r.flushed = false
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
	if r.buffer != nil {
		r.buffer.Reset()
		r.segments = r.segments[:0]
	}
	r.headers.Set("Location", location)
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
		cookie.SubKeys = nil
		cookie.subKeyOrder = nil
		cookie.valueInitialized = true
		return
	}
	r.cookies[key] = &ResponseCookie{
		Name:             name,
		Value:            value,
		Path:             "/",
		valueInitialized: true,
	}
	r.cookieOrder = append(r.cookieOrder, name)
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
		r.cookieOrder = append(r.cookieOrder, cookieName)
	}

	switch strings.ToLower(propertyName) {
	case "value":
		cookie.Value = propertyValue
		cookie.SubKeys = nil
		cookie.subKeyOrder = nil
		cookie.valueInitialized = true
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

// incomingRequestCookieValueLocked reads a raw incoming HTTP cookie value matching
// cookieName case-insensitively while r.mu is held.
func (r *Response) incomingRequestCookieValueLocked(cookieName string) (string, bool) {
	if r.req == nil {
		return "", false
	}
	if incomingCookie, err := r.req.Cookie(cookieName); err == nil && incomingCookie != nil {
		return incomingCookie.Value, true
	}
	for _, incomingCookie := range r.req.Cookies() {
		if incomingCookie != nil && strings.EqualFold(incomingCookie.Name, cookieName) {
			return incomingCookie.Value, true
		}
	}
	return "", false
}

// HasInitializedCookieValue reports whether a response cookie value or sub-key
// has already been initialized in the current response.
func (r *Response) HasInitializedCookieValue(cookieName string) bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	cookie, exists := r.cookies[strings.ToLower(cookieName)]
	return exists && cookie.valueInitialized
}

// HasCookieKeys reports whether a response cookie stores dictionary sub-keys.
func (r *Response) HasCookieKeys(cookieName string) bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	cookie, exists := r.cookies[strings.ToLower(cookieName)]
	if !exists {
		return false
	}
	if cookie.SubKeys != nil {
		return len(cookie.SubKeys) > 0
	}
	order, _ := parseCookieSubKeysOrdered(cookie.Value)
	return len(order) > 0
}

// GetCookieSubKeys returns the ordered sub-key names of a response dictionary cookie.
func (r *Response) GetCookieSubKeys(cookieName string) []string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	cookie, exists := r.cookies[strings.ToLower(cookieName)]
	if !exists {
		return nil
	}
	if cookie.SubKeys != nil {
		keys := make([]string, len(cookie.subKeyOrder))
		copy(keys, cookie.subKeyOrder)
		return keys
	}
	order, _ := parseCookieSubKeysOrdered(cookie.Value)
	return order
}

// SetCookieSubKey appends or updates a sub-key value within the cookie value using
// key=value&key=value encoding compatible with Classic ASP cookie sub-keys.
// If the cookie value has not been initialized in this response, existing sub-keys
// are loaded from the incoming HTTP request before merging.
func (r *Response) SetCookieSubKey(cookieName string, subKey string, value string) {
	r.SetCookieSubKeyWithFallback(cookieName, subKey, value, "", false)
}

// SetCookieSubKeyWithFallback appends or updates a sub-key value within the cookie value.
// When the cookie value has not been initialized in this response, it loads existing
// sub-keys from the incoming HTTP request, or from fallbackRaw when hasFallback is true
// and no matching HTTP request cookie is present.
func (r *Response) SetCookieSubKeyWithFallback(cookieName string, subKey string, value string, fallbackRaw string, hasFallback bool) {
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
		r.cookieOrder = append(r.cookieOrder, cookieName)
	}

	if !cookie.valueInitialized {
		if incomingVal, found := r.incomingRequestCookieValueLocked(cookieName); found {
			cookie.Value = incomingVal
		} else if hasFallback {
			cookie.Value = fallbackRaw
		}
		cookie.SubKeys = nil
		cookie.subKeyOrder = nil
	}

	if cookie.SubKeys == nil {
		cookie.subKeyOrder, cookie.SubKeys = parseCookieSubKeysOrdered(cookie.Value)
	}

	normalizedSubKey := strings.ToLower(subKey)
	if _, subExists := cookie.SubKeys[normalizedSubKey]; !subExists {
		cookie.subKeyOrder = append(cookie.subKeyOrder, normalizedSubKey)
	}
	cookie.SubKeys[normalizedSubKey] = value
	cookie.rebuildSubKeyValue()
	cookie.valueInitialized = true
}

// GetCookieSubKey returns the sub-key value from a cookie value encoded as
// key=value&key=value, or empty string if the sub-key does not exist.
func (r *Response) GetCookieSubKey(cookieName string, subKey string) string {
	r.mu.RLock()
	defer r.mu.RUnlock()

	cookie, exists := r.cookies[strings.ToLower(cookieName)]
	if !exists {
		return ""
	}

	if cookie.SubKeys != nil {
		return cookie.SubKeys[strings.ToLower(subKey)]
	}
	subKeys := parseCookieSubKeys(cookie.Value)
	if subKeys == nil {
		return ""
	}
	return subKeys[strings.ToLower(subKey)]
}

// GetCookieProperty returns one cookie property by name.
func (r *Response) GetCookieProperty(cookieName string, propertyName string) string {
	r.mu.RLock()
	defer r.mu.RUnlock()

	cookie, exists := r.cookies[strings.ToLower(cookieName)]
	if !exists {
		if strings.EqualFold(propertyName, "haskeys") {
			return "False"
		}
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
	case "haskeys":
		if cookie.SubKeys != nil {
			if len(cookie.SubKeys) > 0 {
				return "True"
			}
			return "False"
		}
		order, _ := parseCookieSubKeysOrdered(cookie.Value)
		if len(order) > 0 {
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
	keys := make([]string, len(r.cookieOrder))
	copy(keys, r.cookieOrder)
	return keys
}

// GetCookieCount returns the number of response cookies in insertion order.
func (r *Response) GetCookieCount() int {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return len(r.cookieOrder)
}

// GetCookieKey returns one response cookie name by ASP-compatible 1-based index.
func (r *Response) GetCookieKey(index int) string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	if index < 1 || index > len(r.cookieOrder) {
		return ""
	}
	return r.cookieOrder[index-1]
}

// IsEnded reports whether response was ended.
func (r *Response) IsEnded() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.ended
}

// ResetEnded clears the ended and flushed flags so the response can continue
// accepting output. This is used after Session_OnStart execution, where
// Response.End/Redirect may have been called inside the suppressed-output scope
// and would otherwise silently discard all subsequent page output.
func (r *Response) ResetEnded() {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.ended = false
	r.flushed = false
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
		for name, values := range r.headers {
			for _, value := range values {
				r.w.Header().Add(name, value)
			}
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

		for _, cookieName := range r.cookieOrder {
			cookie, exists := r.cookies[strings.ToLower(cookieName)]
			if !exists || cookie == nil {
				continue
			}
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
				} else if parsed, err := time.Parse(time.RFC1123, cookie.ExpiresRaw); err == nil {
					httpCookie.Expires = parsed
				} else if parsed, err := time.Parse(time.RFC1123Z, cookie.ExpiresRaw); err == nil {
					httpCookie.Expires = parsed
				} else if parsed, err := time.Parse(time.RFC850, cookie.ExpiresRaw); err == nil {
					httpCookie.Expires = parsed
				} else if parsed, err := time.Parse(time.ANSIC, cookie.ExpiresRaw); err == nil {
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

	if r.buffer != nil && r.buffer.Len() > 0 {
		body := r.buffer.Bytes()
		if len(r.segments) > 0 {
			for _, segment := range r.segments {
				if segment.start < 0 || segment.end > len(body) || segment.start >= segment.end {
					continue
				}
				part := body[segment.start:segment.end]
				if !segment.binary {
					part = encodeResponseBody(part, segment.codePage)
				}
				_, _ = r.Output.Write(part)
			}
		} else {
			_, _ = r.Output.Write(body)
		}
		r.buffer.Reset()
		r.segments = r.segments[:0]
	}

	if flusher, ok := r.Output.(http.Flusher); ok {
		flusher.Flush()
	}
}

// ReleaseBuffer returns the internal body buffer to the global pool.
func (r *Response) ReleaseBuffer() {
	if r == nil {
		return
	}
	r.mu.Lock()
	buf := r.buffer
	r.buffer = nil
	r.segments = nil
	r.mu.Unlock()
	if buf != nil && buf.Cap() > 0 && buf.Cap() <= maxPooledBufferCap {
		buf.Reset()
		responseBufferPool.Put(buf)
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

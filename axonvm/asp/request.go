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
	"io"
	"net/http"
	"strings"
	"sync"
)

// RequestCollectionValue stores values and optional cookie-style subkeys.
type RequestCollectionValue struct {
	Values     []string
	Attributes map[string]string
}

// NewRequestCollectionValue creates a collection value from one or many values.
func NewRequestCollectionValue(values []string) RequestCollectionValue {
	normalized := make([]string, 0, len(values))
	for _, value := range values {
		normalized = append(normalized, value)
	}
	return RequestCollectionValue{Values: normalized}
}

// NewRequestCookieValue creates a collection value and parses cookie subkeys.
func NewRequestCookieValue(rawValue string) RequestCollectionValue {
	value := RequestCollectionValue{Values: []string{rawValue}}
	attributes := parseCookieSubKeys(rawValue)
	if len(attributes) > 0 {
		value.Attributes = attributes
	}
	return value
}

// Joined returns ASP-compatible string output for multi-value collection entries.
func (v RequestCollectionValue) Joined() string {
	if len(v.Values) == 0 {
		return ""
	}
	if len(v.Values) == 1 {
		return v.Values[0]
	}
	return strings.Join(v.Values, ", ")
}

// Item resolves one item by index (1-based) or attribute key.
func (v RequestCollectionValue) Item(selector string) string {
	if selector == "" {
		return v.Joined()
	}

	if index, ok := parsePositiveInt(selector); ok {
		if index >= 1 && index <= len(v.Values) {
			return v.Values[index-1]
		}
		return ""
	}

	if len(v.Attributes) > 0 {
		if attributeValue, exists := v.Attributes[strings.ToLower(selector)]; exists {
			return attributeValue
		}
	}

	return ""
}

// Count returns the number of values in the collection item.
func (v RequestCollectionValue) Count() int {
	return len(v.Values)
}

// HasKeys reports whether the value has cookie-style key/value attributes.
func (v RequestCollectionValue) HasKeys() bool {
	return len(v.Attributes) > 0
}

// RequestCollection stores one ASP Request collection using case-insensitive keys.
type RequestCollection struct {
	data     map[string]RequestCollectionValue
	keys     []string
	mu       sync.RWMutex
	onAccess func()
}

// NewRequestCollection creates an empty request collection.
func NewRequestCollection() *RequestCollection {
	return &RequestCollection{
		data: make(map[string]RequestCollectionValue),
		keys: make([]string, 0),
	}
}

// Add stores one key with one value.
func (c *RequestCollection) Add(key string, value string) {
	c.AddValues(key, []string{value})
}

// AddValues stores one key with one or many values.
func (c *RequestCollection) AddValues(key string, values []string) {
	c.mu.Lock()
	defer c.mu.Unlock()

	normalizedKey := strings.ToLower(key)
	if _, exists := c.data[normalizedKey]; !exists {
		c.keys = append(c.keys, normalizedKey)
	}
	c.data[normalizedKey] = NewRequestCollectionValue(values)
}

// AddCookie stores one cookie value and parses subkeys when present.
func (c *RequestCollection) AddCookie(key string, rawValue string) {
	c.mu.Lock()
	defer c.mu.Unlock()

	normalizedKey := strings.ToLower(key)
	if _, exists := c.data[normalizedKey]; !exists {
		c.keys = append(c.keys, normalizedKey)
	}
	c.data[normalizedKey] = NewRequestCookieValue(rawValue)
}

// SetOnAccess configures a callback invoked before read operations.
func (c *RequestCollection) SetOnAccess(callback func()) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.onAccess = callback
}

func (c *RequestCollection) notifyAccess() {
	c.mu.RLock()
	callback := c.onAccess
	c.mu.RUnlock()
	if callback != nil {
		callback()
	}
}

// Get returns the joined string value for one key.
func (c *RequestCollection) Get(key string) string {
	c.notifyAccess()

	c.mu.RLock()
	defer c.mu.RUnlock()

	value, exists := c.data[strings.ToLower(key)]
	if !exists {
		return ""
	}
	return value.Joined()
}

// GetValue returns the full structured value for one key.
func (c *RequestCollection) GetValue(key string) (RequestCollectionValue, bool) {
	c.notifyAccess()

	c.mu.RLock()
	defer c.mu.RUnlock()

	value, exists := c.data[strings.ToLower(key)]
	return value, exists
}

// Exists reports whether a key exists in the collection.
func (c *RequestCollection) Exists(key string) bool {
	c.notifyAccess()

	c.mu.RLock()
	defer c.mu.RUnlock()

	_, exists := c.data[strings.ToLower(key)]
	return exists
}

// Count returns the total number of keys in the collection.
func (c *RequestCollection) Count() int {
	c.notifyAccess()

	c.mu.RLock()
	defer c.mu.RUnlock()

	return len(c.keys)
}

// Key returns one key by ASP-compatible 1-based index.
func (c *RequestCollection) Key(index int) string {
	c.notifyAccess()

	c.mu.RLock()
	defer c.mu.RUnlock()

	if index < 1 || index > len(c.keys) {
		return ""
	}
	return c.keys[index-1]
}

// GetByIndex returns one joined value by ASP-compatible 1-based index.
func (c *RequestCollection) GetByIndex(index int) string {
	c.notifyAccess()

	c.mu.RLock()
	defer c.mu.RUnlock()

	if index < 1 || index > len(c.keys) {
		return ""
	}
	key := c.keys[index-1]
	return c.data[key].Joined()
}

// GetKeys returns a snapshot of all keys in insertion order.
func (c *RequestCollection) GetKeys() []string {
	c.notifyAccess()

	c.mu.RLock()
	defer c.mu.RUnlock()
	result := make([]string, len(c.keys))
	copy(result, c.keys)
	return result
}

// Request accesses data sent by the client (QueryString, Form, etc.).
type Request struct {
	QueryString       *RequestCollection
	Form              *RequestCollection
	Cookies           *RequestCollection
	ServerVars        *RequestCollection
	ClientCertificate *RequestCollection

	bodyBytes      []byte
	bodyPos        int64
	totalBytes     int64
	httpRequest    *http.Request
	bodyLoaded     bool
	formLoaded     bool
	bodyLoader     func() ([]byte, error)
	formLoader     func() (map[string][]string, error)
	bodyLoadOnce   sync.Once
	formLoadOnce   sync.Once
	binaryReadUsed bool // set once BinaryRead has been called; blocks Form access
	formUsed       bool // set once Form collection is read; blocks BinaryRead
	mu             sync.RWMutex
}

// NewRequest creates a new Request object.
func NewRequest() *Request {
	r := &Request{
		QueryString:       NewRequestCollection(),
		Form:              NewRequestCollection(),
		Cookies:           NewRequestCollection(),
		ServerVars:        NewRequestCollection(),
		ClientCertificate: NewRequestCollection(),
	}
	r.Form.SetOnAccess(r.handleFormReadAccess)
	return r
}

// SetHTTPRequest stores the native Go HTTP request.
func (r *Request) SetHTTPRequest(req *http.Request) {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.httpRequest = req
	if req != nil && req.ContentLength > 0 {
		r.totalBytes = req.ContentLength
	}
}

// SetBodyLoader sets a lazy body loader used on BinaryRead and POST TotalBytes access.
func (r *Request) SetBodyLoader(loader func() ([]byte, error)) {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.bodyLoader = loader
}

// SetFormLoader sets a lazy form loader used when Request.Form is accessed.
func (r *Request) SetFormLoader(loader func() (map[string][]string, error)) {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.formLoader = loader
}

// HTTPRequest returns the native Go HTTP request.
func (r *Request) HTTPRequest() *http.Request {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.httpRequest
}

// GetValue retrieves a value from Request collections using ASP default lookup order.
func (r *Request) GetValue(key string) string {
	if value := r.QueryString.Get(key); value != "" {
		return value
	}
	if value := r.Form.Get(key); value != "" {
		return value
	}
	if value := r.Cookies.Get(key); value != "" {
		return value
	}
	if value := r.ClientCertificate.Get(key); value != "" {
		return value
	}
	if value := r.ServerVars.Get(key); value != "" {
		return value
	}
	return ""
}

// GetCollectionValue retrieves one value from a specific collection.
func (r *Request) GetCollectionValue(collectionName string, key string) string {
	collection := r.GetCollection(collectionName)
	if collection == nil {
		return ""
	}
	if index, ok := parsePositiveInt(key); ok {
		return collection.GetByIndex(index)
	}
	return collection.Get(key)
}

// GetCollectionEntry retrieves one value and reports whether the key was present.
// This implements IIS-compatible semantics: a missing key returns ("", false) so
// callers can return VBScript Empty rather than an empty string, matching the
// Classic ASP behaviour that IsEmpty(Request.Form("missingKey")) = True.
func (r *Request) GetCollectionEntry(collectionName string, key string) (string, bool) {
	collection := r.GetCollection(collectionName)
	if collection == nil {
		return "", false
	}
	if index, ok := parsePositiveInt(key); ok {
		val := collection.GetByIndex(index)
		return val, val != "" || index <= collection.Count()
	}
	v, ok := collection.GetValue(key)
	if !ok {
		return "", false
	}
	return v.Joined(), true
}

// GetCollectionProperty returns common collection properties such as Count and Key(index).
func (r *Request) GetCollectionProperty(collectionName string, propertyName string, keyArg string) string {
	collection := r.GetCollection(collectionName)
	if collection == nil {
		return ""
	}

	propertyLower := strings.ToLower(propertyName)
	switch propertyLower {
	case "count":
		return intToString(collection.Count())
	case "key":
		index, ok := parsePositiveInt(keyArg)
		if !ok {
			return ""
		}
		return collection.Key(index)
	default:
		return ""
	}
}

// GetCookieAttribute retrieves a subkey from a cookie value.
func (r *Request) GetCookieAttribute(cookieName string, subKey string) string {
	value, exists := r.Cookies.GetValue(cookieName)
	if !exists {
		return ""
	}
	return value.Item(subKey)
}

// TotalBytes returns the request body size in bytes.
func (r *Request) TotalBytes() int64 {
	r.mu.RLock()
	total := r.totalBytes
	bodyLoaded := r.bodyLoaded
	req := r.httpRequest
	r.mu.RUnlock()

	if total > 0 || bodyLoaded {
		return total
	}
	if req != nil && req.ContentLength >= 0 {
		return req.ContentLength
	}
	return total
}

// SetBody preloads request body bytes for BinaryRead support.
func (r *Request) SetBody(body []byte) {
	r.mu.Lock()
	defer r.mu.Unlock()

	r.bodyBytes = make([]byte, len(body))
	copy(r.bodyBytes, body)
	r.bodyPos = 0
	r.totalBytes = int64(len(body))
	r.bodyLoaded = true
}

// IsBinaryReadUsed reports whether BinaryRead has been called on this request.
func (r *Request) IsBinaryReadUsed() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.binaryReadUsed
}

// IsFormUsed reports whether the Form collection has been accessed on this request.
func (r *Request) IsFormUsed() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.formUsed
}

// MarkFormUsed records that the Form collection has been accessed.
// This must be called from the VM before allowing Form reads.
func (r *Request) MarkFormUsed() {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.formUsed = true
}

// BinaryRead returns up to count bytes from the current request body position.
func (r *Request) BinaryRead(count int64) []byte {
	r.ensureBodyLoaded()

	r.mu.Lock()
	defer r.mu.Unlock()

	if r.formUsed {
		return []byte{}
	}

	if count <= 0 || len(r.bodyBytes) == 0 {
		return []byte{}
	}

	remaining := int64(len(r.bodyBytes)) - r.bodyPos
	if remaining <= 0 {
		return []byte{}
	}
	if count > remaining {
		count = remaining
	}

	start := r.bodyPos
	end := start + count
	result := make([]byte, count)
	copy(result, r.bodyBytes[start:end])
	r.bodyPos = end
	r.binaryReadUsed = true
	return result
}

func (r *Request) handleFormReadAccess() {
	r.MarkFormUsed()

	r.mu.RLock()
	blocked := r.binaryReadUsed
	r.mu.RUnlock()
	if blocked {
		return
	}

	r.ensureFormLoaded()
}

func (r *Request) ensureBodyLoaded() {
	r.bodyLoadOnce.Do(func() {
		r.mu.RLock()
		if r.bodyLoaded {
			r.mu.RUnlock()
			return
		}
		loader := r.bodyLoader
		req := r.httpRequest
		r.mu.RUnlock()

		var body []byte
		if loader != nil {
			loaded, err := loader()
			if err == nil {
				body = loaded
			}
		} else if req != nil && req.Body != nil {
			loaded, err := io.ReadAll(req.Body)
			if err == nil {
				body = loaded
				req.Body = io.NopCloser(bytes.NewReader(loaded))
			}
		}

		r.mu.Lock()
		r.bodyBytes = make([]byte, len(body))
		copy(r.bodyBytes, body)
		r.bodyPos = 0
		r.totalBytes = int64(len(body))
		r.bodyLoaded = true
		r.mu.Unlock()
	})
}

func (r *Request) ensureFormLoaded() {
	r.formLoadOnce.Do(func() {
		r.mu.RLock()
		loader := r.formLoader
		req := r.httpRequest
		r.mu.RUnlock()

		var values map[string][]string
		if loader != nil {
			loaded, err := loader()
			if err == nil {
				values = loaded
			}
		} else if req != nil {
			if err := req.ParseForm(); err == nil {
				values = req.PostForm
			}
		}

		for key, vals := range values {
			r.Form.AddValues(key, vals)
		}

		r.mu.Lock()
		r.formLoaded = true
		r.mu.Unlock()
	})
}

// GetCollection returns one request collection by case-insensitive name.
func (r *Request) GetCollection(name string) *RequestCollection {
	switch strings.ToLower(name) {
	case "querystring":
		return r.QueryString
	case "form":
		return r.Form
	case "cookies":
		return r.Cookies
	case "servervariables":
		return r.ServerVars
	case "clientcertificate":
		return r.ClientCertificate
	default:
		return nil
	}
}

// parseCookieSubKeys parses cookie subkey pairs in key=value&key2=value2 style.
func parseCookieSubKeys(rawValue string) map[string]string {
	parts := strings.Split(rawValue, "&")
	parsed := make(map[string]string)
	for _, part := range parts {
		trimmedPart := strings.TrimSpace(part)
		if trimmedPart == "" {
			continue
		}
		pair := strings.SplitN(trimmedPart, "=", 2)
		if len(pair) == 0 {
			continue
		}
		key := strings.ToLower(strings.TrimSpace(pair[0]))
		if key == "" {
			continue
		}
		value := ""
		if len(pair) > 1 {
			value = strings.TrimSpace(pair[1])
		}
		parsed[key] = value
	}
	return parsed
}

// parsePositiveInt parses positive integers from text values.
func parsePositiveInt(text string) (int, bool) {
	if text == "" {
		return 0, false
	}
	value := 0
	for _, character := range text {
		if character < '0' || character > '9' {
			return 0, false
		}
		value = value*10 + int(character-'0')
	}
	if value < 1 {
		return 0, false
	}
	return value, true
}

// intToString converts a non-negative integer to string without fmt allocations.
func intToString(value int) string {
	if value == 0 {
		return "0"
	}
	if value < 0 {
		return ""
	}

	digits := [20]byte{}
	index := len(digits)
	for value > 0 {
		index--
		digits[index] = byte('0' + (value % 10))
		value /= 10
	}
	return string(digits[index:])
}

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
package server

import (
	"bytes"
	"fmt"
	"io"
	"net/http"
	"strings"
	"sync"
)

// Collection represents a collection object in ASP (QueryString, Form, etc.)
// Supports indexed access and iteration like the real ASP collections
type Collection struct {
	data map[string]interface{}
	keys []string // Maintain insertion order for iteration
	mu   sync.RWMutex
}

// NewCollection creates a new empty collection
func NewCollection() *Collection {
	return &Collection{
		data: make(map[string]interface{}),
		keys: make([]string, 0),
	}
}

// Add adds a key-value pair to the collection
// If key already exists, it's updated
func (c *Collection) Add(key string, value interface{}) {
	c.mu.Lock()
	defer c.mu.Unlock()

	keyLower := strings.ToLower(key)
	if _, exists := c.data[keyLower]; !exists {
		c.keys = append(c.keys, keyLower)
	}
	c.data[keyLower] = value
}

// Get retrieves a value by key (case-insensitive)
func (c *Collection) Get(key string) interface{} {
	c.mu.RLock()
	defer c.mu.RUnlock()

	keyLower := strings.ToLower(key)
	if val, exists := c.data[keyLower]; exists {
		return val
	}
	return ""
}

// Exists checks if a key exists in the collection (case-insensitive)
func (c *Collection) Exists(key string) bool {
	c.mu.RLock()
	defer c.mu.RUnlock()

	keyLower := strings.ToLower(key)
	_, exists := c.data[keyLower]
	return exists
}

// Count returns the number of items in the collection
func (c *Collection) Count() int {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return len(c.keys)
}

// GetKeys returns all keys in the collection (for iteration)
func (c *Collection) GetKeys() []string {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.keys
}

// GetData returns the underlying map (for JSON serialization, etc.)
func (c *Collection) GetData() map[string]interface{} {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.data
}

// GetProperty retrieves a property by name (ASPObject interface)
func (c *Collection) GetProperty(name string) interface{} {
	nameLower := strings.ToLower(name)
	switch nameLower {
	case "count":
		return c.Count()
	case "item":
		// Item property is usually indexed, but can return the whole collection if no index
		// For property access without index, return self? Or nil?
		// In VBScript, Collection.Item("key") is handled by CallMethod or Index access.
		return nil
	}
	// Try to get item by key if property name matches a key?
	// Standard ASP Collections don't expose keys as properties directly, only via Item(key).
	return nil
}

// SetProperty sets a property value (ASPObject interface)
func (c *Collection) SetProperty(name string, value interface{}) error {
	return nil // Read-only mostly
}

// GetName returns the object name
func (c *Collection) GetName() string {
	return "Collection"
}

// CallMethod calls a method on the collection (ASPObject interface)
func (c *Collection) CallMethod(name string, args ...interface{}) (interface{}, error) {
	nameLower := strings.ToLower(name)

	// Default method (Item)
	// Also handle collection names because VBScript-Go sometimes passes the property name
	// (e.g. "QueryString") as the method name when accessing Request.QueryString("key")
	if nameLower == "" || nameLower == "item" ||
		nameLower == "querystring" || nameLower == "form" ||
		nameLower == "cookies" || nameLower == "servervariables" ||
		nameLower == "clientcertificate" {
		if len(args) > 0 {
			key := fmt.Sprintf("%v", args[0])
			// If index is integer, it might be numeric index (1-based in ASP?)
			// ASP Request collections allow string keys or numeric index.
			// Our Collection only supports string keys effectively.
			// But lets support string keys.
			return c.Get(key), nil
		}
	}

	switch nameLower {
	case "key":
		// Key(index)
		if len(args) > 0 {
			if idx, ok := args[0].(int); ok {
				keys := c.GetKeys()
				// ASP collections are 1-based usually
				if idx >= 1 && idx <= len(keys) {
					return keys[idx-1], nil
				}
			}
		}
	}

	return nil, nil
}

// RequestObject represents the ASP Request object
// Provides access to QueryString, Form, Cookies, ServerVariables, and ClientCertificate collections
// Implements all properties and methods from Classic ASP Request Object
type RequestObject struct {
	QueryString       *Collection
	Form              *Collection
	Cookies           *Collection
	ServerVariables   *Collection
	ClientCertificate *Collection
	properties        map[string]interface{}

	// Internal state for BinaryRead
	httpRequest    *http.Request
	bodyBytes      []byte
	bodyBuffer     *bytes.Reader
	totalBytes     int64
	bytesRead      int64
	emptyReadCount int // Track consecutive empty reads to prevent infinite loops
	mu             sync.RWMutex

	// Reference to Response for emergency termination on infinite loops
	response *ResponseObject
}

// NewRequestObject creates a new Request object
func NewRequestObject() *RequestObject {
	return &RequestObject{
		QueryString:       NewCollection(),
		Form:              NewCollection(),
		Cookies:           NewCollection(),
		ServerVariables:   NewCollection(),
		ClientCertificate: NewCollection(),
		properties:        make(map[string]interface{}),
		totalBytes:        0,
		bytesRead:         0,
	}
}

// GetName returns the object name
func (r *RequestObject) GetName() string {
	return "Request"
}

// GetProperty retrieves a property by name
func (r *RequestObject) GetProperty(name string) interface{} {
	nameLower := strings.ToLower(name)

	// Return collections directly as properties
	switch nameLower {
	case "querystring":
		return r.QueryString
	case "form":
		return r.Form
	case "cookies":
		return r.Cookies
	case "servervariables":
		return r.ServerVariables
	case "clientcertificate":
		return r.ClientCertificate
	case "totalbytes":
		// TotalBytes property - returns total bytes in request body
		r.mu.RLock()
		defer r.mu.RUnlock()
		return r.totalBytes
	}

	// Check custom properties
	if val, exists := r.properties[name]; exists {
		return val
	}
	return nil
}

// SetProperty sets a property value
func (r *RequestObject) SetProperty(name string, value interface{}) error {
	r.properties[name] = value
	return nil
}

// CallMethod calls a method on the Request object
// Supports: Request.QueryString(key), Request.Form(key), Request.Cookies(key), Request.ServerVariables(key), Request.ClientCertificate(key), Request.BinaryRead(count)
func (r *RequestObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	nameLower := strings.ToLower(name)

	switch nameLower {
	case "querystring":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.QueryString, nil
		}
		// Return specific value
		key := fmt.Sprintf("%v", args[0])
		return r.QueryString.Get(key), nil

	case "form":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.Form, nil
		}
		// Return specific value
		key := fmt.Sprintf("%v", args[0])
		return r.Form.Get(key), nil

	case "cookies":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.Cookies, nil
		}
		// Return specific value
		key := fmt.Sprintf("%v", args[0])
		return r.Cookies.Get(key), nil

	case "servervariables":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.ServerVariables, nil
		}
		// Return specific value
		key := fmt.Sprintf("%v", args[0])
		return r.ServerVariables.Get(key), nil

	case "clientcertificate":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.ClientCertificate, nil
		}
		// Return specific value
		key := fmt.Sprintf("%v", args[0])
		return r.ClientCertificate.Get(key), nil

	case "binaryread":
		// BinaryRead(count) - reads binary data from request body
		if len(args) < 1 {
			return []byte{}, nil
		}

		count := int64(0)
		switch v := args[0].(type) {
		case int:
			count = int64(v)
		case int32:
			count = int64(v)
		case int64:
			count = v
		case float64:
			count = int64(v)
		default:
			return []byte{}, nil
		}

		return r.BinaryRead(count)

	default:
		// Default lookup order: QueryString, Form, Cookies, ClientCertificate, ServerVariables
		if nameLower == "" && len(args) > 0 {
			key := fmt.Sprintf("%v", args[0])

			// 1. QueryString
			if r.QueryString.Exists(key) {
				return r.QueryString.Get(key), nil
			}

			// 2. Form
			if r.Form.Exists(key) {
				return r.Form.Get(key), nil
			}

			// 3. Cookies
			if r.Cookies.Exists(key) {
				return r.Cookies.Get(key), nil
			}

			// 4. ClientCertificate
			if r.ClientCertificate.Exists(key) {
				return r.ClientCertificate.Get(key), nil
			}

			// 5. ServerVariables
			if r.ServerVariables.Exists(key) {
				return r.ServerVariables.Get(key), nil
			}

			// Not found returns nil (which prints as empty string or Empty)
			return nil, nil
		}
		return nil, nil
	}
}

// ==================== METHODS ====================

// BinaryRead reads binary data from the request body
// Usage: byteArray = Request.BinaryRead(count)
// Classic ASP: Once BinaryRead is called, Form collection cannot be accessed
// Note: To prevent infinite loops in legacy upload code, this tracks consecutive
// empty reads and terminates script after detecting the pattern
func (r *RequestObject) BinaryRead(count int64) ([]byte, error) {
	r.mu.Lock()
	defer r.mu.Unlock()

	// fmt.Printf("[DEBUG] Request.BinaryRead: Requested=%d, TotalBytes=%d, BytesRead=%d\n", count, r.totalBytes, r.bytesRead)

	// Check if we've already signaled script termination
	if r.response != nil && r.response.IsEnded() {
		return []byte{}, nil
	}

	// Prevent infinite loop: after multiple consecutive empty reads, terminate script
	if r.emptyReadCount >= 3 {
		if r.response != nil {
			//fmt.Println("[DEBUG] Request.BinaryRead: Infinite loop detected (3 empty reads), ending response.")
			r.response.End()
		}
		return []byte{}, nil
	}

	if count <= 0 {
		return []byte{}, nil
	}

	// If we haven't read the body yet, read it now
	if r.bodyBuffer == nil {
		if r.httpRequest == nil {
			return []byte{}, nil
		}
		if r.httpRequest.Body == nil {
			return []byte{}, nil
		}

		// Read entire body into memory
		bodyBytes, err := io.ReadAll(r.httpRequest.Body)
		if err != nil {
			return []byte{}, err
		}
		r.httpRequest.Body.Close()

		r.bodyBytes = bodyBytes
		r.bodyBuffer = bytes.NewReader(bodyBytes)
		r.totalBytes = int64(len(bodyBytes))
		r.emptyReadCount = 0
		//fmt.Printf("[DEBUG] Request.BinaryRead: Initialized body buffer. Size=%d bytes\n", r.totalBytes)
	}

	// Calculate how many bytes to read
	remaining := r.totalBytes - r.bytesRead
	if count > remaining {
		count = remaining
	}

	if count <= 0 {
		// Track consecutive empty reads to detect infinite loop pattern
		r.emptyReadCount++
		// After reaching threshold, terminate script on next iteration
		//fmt.Printf("[DEBUG] Request.BinaryRead: EOF reached (count=%d, remaining=%d). EmptyReadCount=%d\n", count, remaining, r.emptyReadCount)
		return []byte{}, nil
	}

	// Read from buffer
	buffer := make([]byte, count)
	n, err := r.bodyBuffer.Read(buffer)
	if err != nil && err != io.EOF {
		return []byte{}, err
	}

	r.bytesRead += int64(n)
	
	// Reset empty read counter when we successfully read data
	if n > 0 {
		r.emptyReadCount = 0
	} else {
		// Track empty reads
		r.emptyReadCount++
	}
	
	//fmt.Printf("[DEBUG] Request.BinaryRead: Returning %d bytes. TotalRead=%d/%d\n", n, r.bytesRead, r.totalBytes)
	return buffer[:n], nil
}

// SetResponse links the Response object for emergency termination
func (r *RequestObject) SetResponse(resp *ResponseObject) {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.response = resp
}

// SetHTTPRequest sets the underlying HTTP request for BinaryRead support
// This should be called when creating the RequestObject
func (r *RequestObject) SetHTTPRequest(req *http.Request) {
	r.mu.Lock()
	defer r.mu.Unlock()

	r.httpRequest = req

	// Set TotalBytes from Content-Length header
	if req != nil {
		r.totalBytes = req.ContentLength
		if r.totalBytes < 0 {
			r.totalBytes = 0
		}
	}
}

// PreloadBody pre-populates the body buffer with data that was already read
// This is used when the body needs to be read before SetHTTPRequest is called
// (e.g., to preserve it before ParseForm consumes it)
func (r *RequestObject) PreloadBody(data []byte) {
	r.mu.Lock()
	defer r.mu.Unlock()

	r.bodyBytes = data
	r.bodyBuffer = bytes.NewReader(data)
	r.totalBytes = int64(len(data))
	r.bytesRead = 0
	r.emptyReadCount = 0
}

// GetTotalBytes returns the total bytes in the request body
func (r *RequestObject) GetTotalBytes() int64 {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.totalBytes
}

// RequestBinaryReadWrapper wraps Request.BinaryRead as a callable object
// This allows Request.BinaryRead(count) to work when the AST is parsed as
// a MemberExpression followed by a Call expression
type RequestBinaryReadWrapper struct {
	reqObj *RequestObject
}

// GetName returns the name for ASPObject interface
func (w *RequestBinaryReadWrapper) GetName() string {
	return "Request.BinaryRead"
}

// GetProperty returns nil - BinaryRead is a method
func (w *RequestBinaryReadWrapper) GetProperty(name string) interface{} {
	return nil
}

// SetProperty returns error - BinaryRead is read-only
func (w *RequestBinaryReadWrapper) SetProperty(name string, value interface{}) error {
	return fmt.Errorf("Request.BinaryRead is a method, not a settable property")
}

// CallMethod implements the actual BinaryRead call
func (w *RequestBinaryReadWrapper) CallMethod(name string, args ...interface{}) (interface{}, error) {
	// When called with empty method name (default call), execute BinaryRead
	if name == "" && len(args) > 0 {
		return w.reqObj.CallMethod("binaryread", args...)
	}
	// When called with explicit method name
	if strings.ToLower(name) == "binaryread" || name == "" {
		return w.reqObj.CallMethod("binaryread", args...)
	}
	return nil, fmt.Errorf("unknown method: %s", name)
}

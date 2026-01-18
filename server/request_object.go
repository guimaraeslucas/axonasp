package server

import (
	"bytes"
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
	httpRequest *http.Request
	bodyBytes   []byte
	bodyBuffer  *bytes.Reader
	totalBytes  int64
	bytesRead   int64
	mu          sync.RWMutex
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
		key := args[0].(string)
		return r.QueryString.Get(key), nil

	case "form":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.Form, nil
		}
		// Return specific value
		key := args[0].(string)
		return r.Form.Get(key), nil

	case "cookies":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.Cookies, nil
		}
		// Return specific value
		key := args[0].(string)
		return r.Cookies.Get(key), nil

	case "servervariables":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.ServerVariables, nil
		}
		// Return specific value
		key := args[0].(string)
		return r.ServerVariables.Get(key), nil

	case "clientcertificate":
		if len(args) == 0 {
			// Return the collection itself for iteration
			return r.ClientCertificate, nil
		}
		// Return specific value
		key := args[0].(string)
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
		return nil, nil
	}
}

// ==================== METHODS ====================

// BinaryRead reads binary data from the request body
// Usage: byteArray = Request.BinaryRead(count)
// Classic ASP: Once BinaryRead is called, Form collection cannot be accessed
func (r *RequestObject) BinaryRead(count int64) ([]byte, error) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if count <= 0 {
		return []byte{}, nil
	}

	// If we haven't read the body yet, read it now
	if r.bodyBuffer == nil {
		if r.httpRequest == nil || r.httpRequest.Body == nil {
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
	}

	// Calculate how many bytes to read
	remaining := r.totalBytes - r.bytesRead
	if count > remaining {
		count = remaining
	}

	if count <= 0 {
		return []byte{}, nil
	}

	// Read from buffer
	buffer := make([]byte, count)
	n, err := r.bodyBuffer.Read(buffer)
	if err != nil && err != io.EOF {
		return []byte{}, err
	}

	r.bytesRead += int64(n)
	return buffer[:n], nil
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

// GetTotalBytes returns the total bytes in the request body
func (r *RequestObject) GetTotalBytes() int64 {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.totalBytes
}

package server

import (
	"strings"
)

// Collection represents a collection object in ASP (QueryString, Form, etc.)
// Supports indexed access and iteration like the real ASP collections
type Collection struct {
	data map[string]interface{}
	keys []string // Maintain insertion order for iteration
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
	keyLower := strings.ToLower(key)
	if _, exists := c.data[keyLower]; !exists {
		c.keys = append(c.keys, keyLower)
	}
	c.data[keyLower] = value
}

// Get retrieves a value by key (case-insensitive)
func (c *Collection) Get(key string) interface{} {
	keyLower := strings.ToLower(key)
	if val, exists := c.data[keyLower]; exists {
		return val
	}
	return ""
}

// Exists checks if a key exists in the collection (case-insensitive)
func (c *Collection) Exists(key string) bool {
	keyLower := strings.ToLower(key)
	_, exists := c.data[keyLower]
	return exists
}

// Count returns the number of items in the collection
func (c *Collection) Count() int {
	return len(c.keys)
}

// GetKeys returns all keys in the collection (for iteration)
func (c *Collection) GetKeys() []string {
	return c.keys
}

// GetData returns the underlying map (for JSON serialization, etc.)
func (c *Collection) GetData() map[string]interface{} {
	return c.data
}

// RequestObject represents the ASP Request object
// Provides access to QueryString, Form, Cookies, and ServerVariables collections
type RequestObject struct {
	QueryString     *Collection
	Form            *Collection
	Cookies         *Collection
	ServerVariables *Collection
	properties      map[string]interface{}
}

// NewRequestObject creates a new Request object
func NewRequestObject() *RequestObject {
	return &RequestObject{
		QueryString:     NewCollection(),
		Form:            NewCollection(),
		Cookies:         NewCollection(),
		ServerVariables: NewCollection(),
		properties:      make(map[string]interface{}),
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
// Supports: Request.QueryString(key), Request.Form(key), Request.Cookies(key), Request.ServerVariables(key)
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

	default:
		return nil, nil
	}
}

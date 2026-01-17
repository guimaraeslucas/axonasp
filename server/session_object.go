package server

import (
	"strings"
)

// SessionObject wraps session data and provides ASP-style access
type SessionObject struct {
	ID   string
	Data map[string]interface{}
}

// NewSessionObject creates a new session object wrapper
func NewSessionObject(sessionID string, data map[string]interface{}) *SessionObject {
	return &SessionObject{
		ID:   sessionID,
		Data: data,
	}
}

// GetProperty gets a session property (case-insensitive)
func (s *SessionObject) GetProperty(name string) interface{} {
	nameLower := strings.ToLower(name)
	
	// Handle built-in properties
	if nameLower == "sessionid" {
		return s.ID
	}
	
	// Get from session data
	if val, exists := s.Data[nameLower]; exists {
		return val
	}
	
	return nil
}

// SetProperty sets a session property (case-insensitive)
func (s *SessionObject) SetProperty(name string, value interface{}) error {
	nameLower := strings.ToLower(name)
	
	// Don't allow setting SessionID
	if nameLower == "sessionid" {
		return nil // Ignore attempts to set SessionID
	}
	
	s.Data[nameLower] = value
	return nil
}

// GetIndex gets a session value by index (for Session("key") syntax)
func (s *SessionObject) GetIndex(index interface{}) interface{} {
	// Convert index to string
	key := ""
	switch v := index.(type) {
	case string:
		key = v
	default:
		return nil
	}
	
	return s.GetProperty(key)
}

// SetIndex sets a session value by index (for Session("key") = value syntax)
func (s *SessionObject) SetIndex(index interface{}, value interface{}) error {
	// Convert index to string
	key := ""
	switch v := index.(type) {
	case string:
		key = v
	default:
		return nil
	}
	
	return s.SetProperty(key, value)
}

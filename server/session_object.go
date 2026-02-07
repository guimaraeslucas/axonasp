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
	"hash/fnv"
	"strings"
	"sync"
	"time"
)

// SessionObject wraps session data and provides ASP-style access
type SessionObject struct {
	ID        string
	NumericID int64
	Data      map[string]interface{}
	TimeOut   int // Timeout in minutes
	CreatedAt time.Time
	mu        sync.RWMutex
	locked    bool
	lockCount int
}

// NewSessionObject creates a new session object wrapper
func NewSessionObject(sessionID string, data map[string]interface{}) *SessionObject {
	return &SessionObject{
		ID:        sessionID,
		NumericID: numericSessionID(sessionID),
		Data:      data,
		TimeOut:   20, // Default ASP session timeout in minutes
		CreatedAt: time.Now(),
		locked:    false,
		lockCount: 0,
	}
}

// GetProperty gets a session property (case-insensitive)
func (s *SessionObject) GetProperty(name string) interface{} {
	nameLower := strings.ToLower(name)

	// Handle built-in properties
	if nameLower == "sessionid" {
		return s.NumericID
	}

	// Get from session data
	if val, exists := s.Data[nameLower]; exists {
		return val
	}

	return nil
}

func numericSessionID(sessionID string) int64 {
	if sessionID == "" {
		return 0
	}
	h := fnv.New32a()
	_, _ = h.Write([]byte(sessionID))
	value := int64(h.Sum32() & 0x7fffffff)
	if value == 0 {
		value = 1
	}
	return value
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

// Lock prevents other clients from modifying Session variables
// Supports nested locks - each Lock must have a corresponding Unlock
func (s *SessionObject) Lock() {
	s.mu.Lock()
	s.locked = true
	s.lockCount++
}

// Unlock allows other clients to modify Session variables
// Must be called the same number of times as Lock was called
func (s *SessionObject) Unlock() {
	if s.locked && s.lockCount > 0 {
		s.lockCount--
		if s.lockCount == 0 {
			s.locked = false
			s.mu.Unlock()
		}
	}
}

// IsLocked returns whether the Session object is currently locked
func (s *SessionObject) IsLocked() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.locked
}

// GetLockCount returns the current lock count
func (s *SessionObject) GetLockCount() int {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.lockCount
}

// Abandon terminates the session (removes it from storage)
// This is called by the Session.Abandon() method in ASP
func (s *SessionObject) Abandon() {
	s.mu.Lock()
	defer s.mu.Unlock()
	// Clear all data
	s.Data = make(map[string]interface{})
}

// RemoveAll clears all variables from the Session
func (s *SessionObject) RemoveAll() {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.Data = make(map[string]interface{})
}

// GetContents returns a copy of all session variables for enumeration
func (s *SessionObject) GetContents() map[string]interface{} {
	s.mu.RLock()
	defer s.mu.RUnlock()

	copy := make(map[string]interface{})
	for k, v := range s.Data {
		copy[k] = v
	}
	return copy
}

// GetAllKeys returns all variable names in the session
func (s *SessionObject) GetAllKeys() []string {
	s.mu.RLock()
	defer s.mu.RUnlock()

	keys := make([]string, 0, len(s.Data))
	for k := range s.Data {
		keys = append(keys, k)
	}
	return keys
}

// Count returns the number of variables in the session
func (s *SessionObject) Count() int {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return len(s.Data)
}

// SetTimeout sets the session timeout in minutes
func (s *SessionObject) SetTimeout(minutes int) {
	s.mu.Lock()
	defer s.mu.Unlock()
	if minutes > 0 {
		s.TimeOut = minutes
	}
}

// GetTimeout returns the session timeout in minutes
func (s *SessionObject) GetTimeout() int {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.TimeOut
}

// IsExpired checks if the session has expired based on creation time and timeout
func (s *SessionObject) IsExpired() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	timeout := time.Duration(s.TimeOut) * time.Minute
	return time.Since(s.CreatedAt) > timeout
}

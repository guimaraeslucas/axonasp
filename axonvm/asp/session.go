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
	"crypto/rand"
	"encoding/json"
	"hash/fnv"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
)

// Session maintains user state between requests.
type Session struct {
	ID string

	data          map[string]ApplicationValue
	staticObjects map[string]ApplicationValue

	LCID     int
	CodePage int
	Timeout  int

	CreatedAt    time.Time
	LastAccessed time.Time

	mu        sync.RWMutex
	stateMu   sync.Mutex
	locked    bool
	lockCount int
	abandoned bool
	dirty     bool
	version   uint64
}

// sessionDiskPayload stores session fields serialized to disk.
type sessionDiskPayload struct {
	ID            string                      `json:"id"`
	Data          map[string]ApplicationValue `json:"data"`
	StaticObjects map[string]ApplicationValue `json:"static_objects"`
	LCID          int                         `json:"lcid"`
	CodePage      int                         `json:"codepage"`
	Timeout       int                         `json:"timeout"`
	CreatedAt     time.Time                   `json:"created_at"`
	LastAccessed  time.Time                   `json:"last_accessed"`
}

const (
	defaultSessionTimeoutMinutes = 20
	defaultSessionLCID           = 1033
	defaultSessionCodePage       = 65001
)

var sessionStorageDir = filepath.Join("temp", "session")

var (
	defaultSessionLCIDOnce   sync.Once
	cachedDefaultSessionLCID = defaultSessionLCID

	sessionRegistryMu sync.RWMutex
	sessionRegistry   = make(map[string]*Session)

	sessionAutoFlushMu   sync.Mutex
	sessionAutoFlushStop chan struct{}
	sessionAutoFlushDone chan struct{}
)

// resolveDefaultSessionLCID reads the configured default LCID with safe fallbacks.
func resolveDefaultSessionLCID() int {
	defaultSessionLCIDOnce.Do(func() {
		if configuredLCID := axonconfig.NewViper().GetInt("global.default_mslcid"); configuredLCID > 0 {
			cachedDefaultSessionLCID = configuredLCID
		}
	})
	return cachedDefaultSessionLCID
}

// NewSession creates a new Session object with defaults.
func NewSession() *Session {
	return NewSessionWithID("")
}

// NewSessionWithID creates a new Session object with a specific ID.
func NewSessionWithID(sessionID string) *Session {
	now := time.Now()
	return &Session{
		ID:            sessionID,
		data:          make(map[string]ApplicationValue),
		staticObjects: make(map[string]ApplicationValue),
		LCID:          resolveDefaultSessionLCID(),
		CodePage:      defaultSessionCodePage,
		Timeout:       defaultSessionTimeoutMinutes,
		CreatedAt:     now,
		LastAccessed:  now,
		dirty:         true,
		version:       1,
	}
}

// SessionIDNumeric returns a numeric representation of SessionID for ASP compatibility.
func (s *Session) SessionIDNumeric() int64 {
	return numericSessionID(s.ID)
}

// Set stores one value in session contents using case-insensitive keys.
func (s *Session) Set(key string, value ApplicationValue) {
	s.mu.Lock()
	defer s.mu.Unlock()

	s.data[normalizeApplicationKey(key)] = value
	s.markDirtyLocked()
}

// Get retrieves one value from session contents using case-insensitive keys.
func (s *Session) Get(key string) (ApplicationValue, bool) {
	s.mu.RLock()
	defer s.mu.RUnlock()

	value, ok := s.data[normalizeApplicationKey(key)]
	return value, ok
}

// Contains reports whether a session key exists.
func (s *Session) Contains(key string) bool {
	s.mu.RLock()
	defer s.mu.RUnlock()

	_, ok := s.data[normalizeApplicationKey(key)]
	return ok
}

// Remove deletes one session value by key.
func (s *Session) Remove(key string) {
	s.mu.Lock()
	defer s.mu.Unlock()

	normalized := normalizeApplicationKey(key)
	if _, exists := s.data[normalized]; exists {
		delete(s.data, normalized)
		s.markDirtyLocked()
	}
}

// RemoveAll clears all session values.
func (s *Session) RemoveAll() {
	s.mu.Lock()
	defer s.mu.Unlock()

	if len(s.data) == 0 {
		return
	}
	s.data = make(map[string]ApplicationValue)
	s.markDirtyLocked()
}

// Count returns the number of session content values.
func (s *Session) Count() int {
	s.mu.RLock()
	defer s.mu.RUnlock()

	return len(s.data)
}

// AddStaticObject inserts a value into Session.StaticObjects.
func (s *Session) AddStaticObject(key string, value ApplicationValue) {
	s.mu.Lock()
	defer s.mu.Unlock()

	s.staticObjects[strings.ToLower(key)] = value
	s.markDirtyLocked()
}

// GetStaticObject retrieves a value from Session.StaticObjects.
func (s *Session) GetStaticObject(key string) (ApplicationValue, bool) {
	s.mu.RLock()
	defer s.mu.RUnlock()

	value, ok := s.staticObjects[strings.ToLower(key)]
	return value, ok
}

// ContainsStaticObject reports whether a Session.StaticObjects key exists.
func (s *Session) ContainsStaticObject(key string) bool {
	s.mu.RLock()
	defer s.mu.RUnlock()

	_, ok := s.staticObjects[strings.ToLower(key)]
	return ok
}

// GetStaticObjectsCopy returns a safe snapshot copy for enumeration.
func (s *Session) GetStaticObjectsCopy() map[string]ApplicationValue {
	s.mu.RLock()
	defer s.mu.RUnlock()

	copyMap := make(map[string]ApplicationValue, len(s.staticObjects))
	for k, v := range s.staticObjects {
		copyMap[k] = v
	}
	return copyMap
}

// GetContentsCopy returns a safe snapshot copy for enumeration.
func (s *Session) GetContentsCopy() map[string]ApplicationValue {
	s.mu.RLock()
	defer s.mu.RUnlock()

	copyMap := make(map[string]ApplicationValue, len(s.data))
	for key, value := range s.data {
		copyMap[key] = value
	}
	return copyMap
}

// GetAllKeys returns all keys stored in session contents.
func (s *Session) GetAllKeys() []string {
	s.mu.RLock()
	defer s.mu.RUnlock()

	keys := make([]string, 0, len(s.data))
	for key := range s.data {
		keys = append(keys, key)
	}
	return keys
}

// SetTimeout sets Session.Timeout using ASP default fallback rules.
func (s *Session) SetTimeout(minutes int) {
	s.mu.Lock()
	defer s.mu.Unlock()

	newTimeout := minutes
	if minutes <= 0 {
		newTimeout = defaultSessionTimeoutMinutes
	}
	if s.Timeout != newTimeout {
		s.Timeout = newTimeout
		s.markDirtyLocked()
	}
}

// GetTimeout returns Session.Timeout.
func (s *Session) GetTimeout() int {
	s.mu.RLock()
	defer s.mu.RUnlock()

	if s.Timeout <= 0 {
		return defaultSessionTimeoutMinutes
	}
	return s.Timeout
}

// SetLCID sets Session.LCID using ASP default fallback rules.
func (s *Session) SetLCID(lcid int) {
	s.mu.Lock()
	defer s.mu.Unlock()

	newLCID := lcid
	if lcid <= 0 {
		newLCID = resolveDefaultSessionLCID()
	}
	if s.LCID != newLCID {
		s.LCID = newLCID
		s.markDirtyLocked()
	}
}

// GetLCID returns Session.LCID.
func (s *Session) GetLCID() int {
	s.mu.RLock()
	defer s.mu.RUnlock()

	if s.LCID <= 0 {
		return resolveDefaultSessionLCID()
	}
	return s.LCID
}

// SetCodePage sets Session.CodePage using ASP default fallback rules.
func (s *Session) SetCodePage(codePage int) {
	s.mu.Lock()
	defer s.mu.Unlock()

	newCodePage := codePage
	if codePage <= 0 {
		newCodePage = defaultSessionCodePage
	}
	if s.CodePage != newCodePage {
		s.CodePage = newCodePage
		s.markDirtyLocked()
	}
}

// GetCodePage returns Session.CodePage.
func (s *Session) GetCodePage() int {
	s.mu.RLock()
	defer s.mu.RUnlock()

	if s.CodePage <= 0 {
		return defaultSessionCodePage
	}
	return s.CodePage
}

// Lock enters a session critical section and supports nested locks.
func (s *Session) Lock() {
	s.stateMu.Lock()
	defer s.stateMu.Unlock()

	s.lockCount++
	s.locked = s.lockCount > 0
}

// Unlock leaves one nested level from a session critical section.
func (s *Session) Unlock() {
	s.stateMu.Lock()
	defer s.stateMu.Unlock()

	if s.lockCount > 0 {
		s.lockCount--
	}
	s.locked = s.lockCount > 0
}

// IsLocked reports whether session lock is currently active.
func (s *Session) IsLocked() bool {
	s.stateMu.Lock()
	defer s.stateMu.Unlock()

	return s.locked
}

// GetLockCount returns the current nested lock count.
func (s *Session) GetLockCount() int {
	s.stateMu.Lock()
	defer s.stateMu.Unlock()

	return s.lockCount
}

// Abandon clears all contents and marks the session as abandoned.
func (s *Session) Abandon() {
	s.mu.Lock()
	defer s.mu.Unlock()

	if len(s.data) > 0 {
		s.data = make(map[string]ApplicationValue)
	}
	if !s.abandoned {
		s.abandoned = true
	}
	s.markDirtyLocked()
}

// IsAbandoned reports whether Abandon has been called.
func (s *Session) IsAbandoned() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()

	return s.abandoned
}

// IsExpired reports whether session timeout elapsed since last access.
func (s *Session) IsExpired() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()

	timeout := time.Duration(s.GetTimeout()) * time.Minute
	return time.Since(s.LastAccessed) > timeout
}

// Touch updates the LastAccessed timestamp.
func (s *Session) Touch() {
	s.mu.Lock()
	defer s.mu.Unlock()

	s.touchLocked()
}

// IsDirty reports whether session state has pending durable changes.
func (s *Session) IsDirty() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.dirty
}

// SaveIfDirty persists only when session state changed since the last successful save.
func (s *Session) SaveIfDirty() error {
	s.mu.RLock()
	dirty := s.dirty
	s.mu.RUnlock()
	if !dirty {
		return nil
	}
	return s.Save()
}

// Save persists session state as JSON in temp/session.
func (s *Session) Save() error {
	s.mu.RLock()
	version := s.version
	payload := sessionDiskPayload{
		ID:            s.ID,
		Data:          copySessionData(s.data),
		StaticObjects: copySessionData(s.staticObjects),
		LCID:          s.LCID,
		CodePage:      s.CodePage,
		Timeout:       s.Timeout,
		CreatedAt:     s.CreatedAt,
		LastAccessed:  s.LastAccessed,
	}
	if payload.LCID <= 0 {
		payload.LCID = resolveDefaultSessionLCID()
	}
	if payload.CodePage <= 0 {
		payload.CodePage = defaultSessionCodePage
	}
	if payload.Timeout <= 0 {
		payload.Timeout = defaultSessionTimeoutMinutes
	}
	s.mu.RUnlock()

	if err := os.MkdirAll(sessionStorageDir, 0o755); err != nil {
		return err
	}

	bytes, err := json.Marshal(payload)
	if err != nil {
		return err
	}

	if err := os.WriteFile(sessionFilePath(s.ID), bytes, 0o600); err != nil {
		return err
	}

	s.mu.Lock()
	if s.version == version {
		s.dirty = false
	}
	s.mu.Unlock()
	return nil
}

// Delete removes persisted session JSON from disk.
func (s *Session) Delete() error {
	err := os.Remove(sessionFilePath(s.ID))
	if err != nil && !os.IsNotExist(err) {
		return err
	}
	unregisterSession(s.ID)
	return nil
}

// LoadSession loads a persisted session by ID, or returns false if not found.
func LoadSession(sessionID string) (*Session, bool, error) {
	bytes, err := os.ReadFile(sessionFilePath(sessionID))
	if err != nil {
		if os.IsNotExist(err) {
			return nil, false, nil
		}
		return nil, false, err
	}

	var payload sessionDiskPayload
	if err := json.Unmarshal(bytes, &payload); err != nil {
		return nil, false, err
	}

	session := NewSessionWithID(payload.ID)
	session.data = payload.Data
	if session.data == nil {
		session.data = make(map[string]ApplicationValue)
	}
	session.staticObjects = payload.StaticObjects
	if session.staticObjects == nil {
		session.staticObjects = make(map[string]ApplicationValue)
	}
	session.LCID = payload.LCID
	session.CodePage = payload.CodePage
	session.Timeout = payload.Timeout
	session.CreatedAt = payload.CreatedAt
	session.LastAccessed = payload.LastAccessed
	session.dirty = false
	session.version = 0

	if session.IsExpired() {
		_ = session.Delete()
		return nil, false, nil
	}

	session.Touch()
	return session, true, nil
}

// CreateSession creates a brand-new session with generated ID and persists it.
func CreateSession() (*Session, error) {
	session := NewSessionWithID(newSessionID())
	if err := session.Save(); err != nil {
		return nil, err
	}
	return registerSession(session), nil
}

// GetOrCreateSession loads an existing session or creates a new one.
func GetOrCreateSession(sessionID string) (*Session, bool, error) {
	normalizedID := strings.TrimSpace(sessionID)
	if normalizedID != "" {
		if existing := getRegisteredSession(normalizedID); existing != nil {
			existing.Touch()
			return existing, false, nil
		}
	}

	if sessionID != "" {
		session, found, err := LoadSession(sessionID)
		if err != nil {
			return nil, false, err
		}
		if found {
			return registerSession(session), false, nil
		}
	}

	session, err := CreateSession()
	if err != nil {
		return nil, false, err
	}
	return session, true, nil
}

func getRegisteredSession(sessionID string) *Session {
	if strings.TrimSpace(sessionID) == "" {
		return nil
	}
	sessionRegistryMu.RLock()
	defer sessionRegistryMu.RUnlock()
	return sessionRegistry[sessionID]
}

func registerSession(session *Session) *Session {
	if session == nil || strings.TrimSpace(session.ID) == "" {
		return session
	}
	sessionRegistryMu.Lock()
	defer sessionRegistryMu.Unlock()
	if existing := sessionRegistry[session.ID]; existing != nil {
		return existing
	}
	sessionRegistry[session.ID] = session
	return session
}

func unregisterSession(sessionID string) {
	if strings.TrimSpace(sessionID) == "" {
		return
	}
	sessionRegistryMu.Lock()
	defer sessionRegistryMu.Unlock()
	delete(sessionRegistry, sessionID)
}

// StartSessionAutoFlush starts a background flusher for dirty sessions.
func StartSessionAutoFlush(interval time.Duration) {
	if interval <= 0 {
		return
	}

	sessionAutoFlushMu.Lock()
	defer sessionAutoFlushMu.Unlock()

	if sessionAutoFlushStop != nil {
		close(sessionAutoFlushStop)
		if sessionAutoFlushDone != nil {
			<-sessionAutoFlushDone
		}
	}

	stop := make(chan struct{})
	done := make(chan struct{})
	sessionAutoFlushStop = stop
	sessionAutoFlushDone = done

	go func(stopCh <-chan struct{}, doneCh chan<- struct{}) {
		defer close(doneCh)
		ticker := time.NewTicker(interval)
		defer ticker.Stop()

		for {
			select {
			case <-ticker.C:
				_ = FlushRegisteredSessions(false)
			case <-stopCh:
				_ = FlushRegisteredSessions(true)
				return
			}
		}
	}(stop, done)
}

// StopSessionAutoFlush stops the background flusher and flushes pending dirty sessions.
func StopSessionAutoFlush() {
	sessionAutoFlushMu.Lock()
	stop := sessionAutoFlushStop
	done := sessionAutoFlushDone
	sessionAutoFlushStop = nil
	sessionAutoFlushDone = nil
	sessionAutoFlushMu.Unlock()

	if stop != nil {
		close(stop)
	}
	if done != nil {
		<-done
	}
}

// FlushRegisteredSessions persists registered sessions.
// When force is false, only dirty sessions are persisted.
func FlushRegisteredSessions(force bool) error {
	sessionRegistryMu.RLock()
	sessions := make([]*Session, 0, len(sessionRegistry))
	for _, session := range sessionRegistry {
		sessions = append(sessions, session)
	}
	sessionRegistryMu.RUnlock()

	var firstErr error
	for _, session := range sessions {
		if session == nil {
			continue
		}
		var err error
		if force {
			err = session.Save()
		} else {
			err = session.SaveIfDirty()
		}
		if err != nil && firstErr == nil {
			firstErr = err
		}
	}

	return firstErr
}

// SetSessionStorageDir changes session persistence directory for tests and tooling.
func SetSessionStorageDir(path string) {
	if strings.TrimSpace(path) == "" {
		sessionStorageDir = filepath.Join("temp", "session")
		return
	}
	sessionStorageDir = path
}

// numericSessionID converts session ID string to positive numeric ID.
func numericSessionID(sessionID string) int64 {
	if sessionID == "" {
		return 0
	}
	hash := fnv.New32a()
	_, _ = hash.Write([]byte(sessionID))
	value := int64(hash.Sum32() & 0x7fffffff)
	if value == 0 {
		return 1
	}
	return value
}

// newSessionID generates a random session ID that matches the Classic ASP / IIS ASPSESSIONID
// cookie value format: 24 uppercase letters (A-Z only, no digits, no hyphens).
// Classic ASP uses this exact character set — lowercase hex would not match.
func newSessionID() string {
	b := make([]byte, 24)
	if _, err := rand.Read(b); err != nil {
		// Fallback: derive from timestamp digits mapped to uppercase letters.
		ts := strings.ReplaceAll(time.Now().UTC().Format("20060102150405.000000000"), ".", "")
		for len(ts) < 24 {
			ts += ts
		}
		ts = ts[:24]
		result := make([]byte, 24)
		for i := 0; i < 24; i++ {
			result[i] = 'A' + (ts[i]-'0')%26
		}
		return string(result)
	}
	result := make([]byte, 24)
	for i := 0; i < 24; i++ {
		result[i] = 'A' + (b[i] % 26)
	}
	return string(result)
}

// sessionFilePath builds the absolute session JSON file path for an ID.
func sessionFilePath(sessionID string) string {
	safeID := strings.ReplaceAll(sessionID, "/", "_")
	safeID = strings.ReplaceAll(safeID, "\\", "_")
	return filepath.Join(sessionStorageDir, safeID+".json")
}

// copySessionData copies typed session map data.
func copySessionData(source map[string]ApplicationValue) map[string]ApplicationValue {
	copyMap := make(map[string]ApplicationValue, len(source))
	for key, value := range source {
		copyMap[key] = value
	}
	return copyMap
}

// touchLocked updates LastAccessed and assumes write lock is already held.
func (s *Session) touchLocked() {
	s.LastAccessed = time.Now()
}

// markDirtyLocked updates LastAccessed and marks state as changed.
func (s *Session) markDirtyLocked() {
	s.touchLocked()
	s.dirty = true
	s.version++
}

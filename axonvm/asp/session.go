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
	"crypto/rand"
	"encoding/binary"
	"hash/fnv"
	"maps"
	"math"
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

	lastSavedVersion uint64
}

// sessionWriteQueue is a buffered channel for asynchronous session persistence.
var sessionWriteQueue = make(chan *Session, 10000)

func init() {
	// Start background session writers to offload disk I/O from request threads.
	// 4 workers provide controlled concurrency to avoid OS thread starvation.
	for range 4 {
		go sessionWriterWorker()
	}
}

// sessionWriterWorker listens for dirty sessions and persists them to disk.
func sessionWriterWorker() {
	for s := range sessionWriteQueue {
		if s == nil {
			continue
		}
		// Save performing actual disk I/O.
		_ = s.Save()
	}
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

// sessionBufferPool reduces allocations during session serialization.
var sessionBufferPool = sync.Pool{
	New: func() any {
		return new(bytes.Buffer)
	},
}

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
	maps.Copy(copyMap, s.staticObjects)
	return copyMap
}

// GetContentsCopy returns a safe snapshot copy for enumeration.
func (s *Session) GetContentsCopy() map[string]ApplicationValue {
	s.mu.RLock()
	defer s.mu.RUnlock()

	copyMap := make(map[string]ApplicationValue, len(s.data))
	maps.Copy(copyMap, s.data)
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
	version := s.version
	lastSaved := s.lastSavedVersion
	s.mu.RUnlock()

	if !dirty || version <= lastSaved {
		return nil
	}
	return s.Save()
}

// QueueSaveIfDirty pushes the session to the background write queue if it has pending changes.
func (s *Session) QueueSaveIfDirty() bool {
	s.mu.RLock()
	dirty := s.dirty
	version := s.version
	lastSaved := s.lastSavedVersion
	s.mu.RUnlock()

	if !dirty || version <= lastSaved {
		return false
	}

	select {
	case sessionWriteQueue <- s:
		return true
	default:
		println("Warning: Session write queue overflow. Persistence delayed for session:", s.ID)
		return false
	}
}

// Save persists session state using a high-performance binary format (.g3ses).
// Binary Format Specification:
// - Magic Bytes: [6]byte{'G','3','S','E','S', 0x00}
// - Version: uint8 (Value: 1)
// - Session ID: [24]byte
// - CreatedAt: int64 (Unix seconds)
// - LastAccessed: int64 (Unix seconds)
// - Timeout: int16
// - LCID: uint16
// - CodePage: uint32
// - Block Separator: 0x1E (RS) followed by uint32 (Payload Length)
// - Data Block 1: session.data map
// - Data Block 2: session.staticObjects map
func (s *Session) Save() error {
	s.mu.RLock()
	version := s.version
	id := s.ID
	data := s.data
	static := s.staticObjects
	lcid := uint16(s.LCID)
	codePage := uint32(s.CodePage)
	timeout := int16(s.Timeout)
	createdAt := s.CreatedAt.Unix()
	lastAccessed := s.LastAccessed.Unix()

	if lcid == 0 {
		lcid = uint16(resolveDefaultSessionLCID())
	}
	if codePage == 0 {
		codePage = defaultSessionCodePage
	}
	if timeout == 0 {
		timeout = defaultSessionTimeoutMinutes
	}
	s.mu.RUnlock()

	if err := os.MkdirAll(sessionStorageDir, 0o755); err != nil {
		return err
	}

	buf := sessionBufferPool.Get().(*bytes.Buffer)
	buf.Reset()
	defer sessionBufferPool.Put(buf)

	// Magic + Version + ID + Created + Accessed + Timeout + LCID + CodePage
	// 6 + 1 + 24 + 8 + 8 + 2 + 2 + 4 = 55 bytes
	var header [55]byte
	copy(header[0:6], "G3SES\x00")
	header[6] = 1
	copy(header[7:31], id)
	binary.LittleEndian.PutUint64(header[31:39], uint64(createdAt))
	binary.LittleEndian.PutUint64(header[39:47], uint64(lastAccessed))
	binary.LittleEndian.PutUint16(header[47:49], uint16(timeout))
	binary.LittleEndian.PutUint16(header[49:51], lcid)
	binary.LittleEndian.PutUint32(header[51:55], codePage)
	buf.Write(header[:])

	// Block 1: Data
	buf.WriteByte(0x1E)
	dataBuf := sessionBufferPool.Get().(*bytes.Buffer)
	dataBuf.Reset()
	serializeApplicationMap(dataBuf, data)
	writeUint32(buf, uint32(dataBuf.Len()))
	buf.Write(dataBuf.Bytes())
	sessionBufferPool.Put(dataBuf)

	// Block 2: Static Objects
	buf.WriteByte(0x1E)
	staticBuf := sessionBufferPool.Get().(*bytes.Buffer)
	staticBuf.Reset()
	serializeApplicationMap(staticBuf, static)
	writeUint32(buf, uint32(staticBuf.Len()))
	buf.Write(staticBuf.Bytes())
	sessionBufferPool.Put(staticBuf)

	if err := os.WriteFile(sessionFilePath(id), buf.Bytes(), 0o600); err != nil {
		return err
	}

	s.mu.Lock()
	if s.version == version {
		s.dirty = false
		s.lastSavedVersion = version
	}
	s.mu.Unlock()
	return nil
}

// Delete removes persisted session from disk.
func (s *Session) Delete() error {
	err := os.Remove(sessionFilePath(s.ID))
	if err != nil && !os.IsNotExist(err) {
		return err
	}
	unregisterSession(s.ID)
	return nil
}

// LoadSession loads a persisted binary session by ID.
func LoadSession(sessionID string) (*Session, bool, error) {
	data, err := os.ReadFile(sessionFilePath(sessionID))
	if err != nil {
		if os.IsNotExist(err) {
			return nil, false, nil
		}
		return nil, false, err
	}

	if len(data) < 55 || string(data[0:6]) != "G3SES\x00" {
		return nil, false, nil
	}

	ver := data[6]
	if ver != 1 {
		return nil, false, nil
	}

	createdAt := int64(binary.LittleEndian.Uint64(data[31:39]))
	lastAccessed := int64(binary.LittleEndian.Uint64(data[39:47]))
	timeout := int16(binary.LittleEndian.Uint16(data[47:49]))
	lcid := int(binary.LittleEndian.Uint16(data[49:51]))
	codePage := int(binary.LittleEndian.Uint32(data[51:55]))

	s := NewSessionWithID(sessionID)
	s.CreatedAt = time.Unix(createdAt, 0)
	s.LastAccessed = time.Unix(lastAccessed, 0)
	s.Timeout = int(timeout)
	s.LCID = lcid
	s.CodePage = codePage

	r := bytes.NewReader(data[55:])

	// Read Data Block
	if sep, _ := r.ReadByte(); sep == 0x1E {
		length := readUint32(r)
		_ = length
		s.data = deserializeApplicationMap(r)
	}

	// Read Static Objects Block
	if sep, _ := r.ReadByte(); sep == 0x1E {
		length := readUint32(r)
		_ = length
		s.staticObjects = deserializeApplicationMap(r)
	}

	s.dirty = false
	s.version = 0

	if s.IsExpired() {
		_ = s.Delete()
		return nil, false, nil
	}

	s.Touch()
	return s, true, nil
}

// CreateSession creates a brand-new session and persists it asynchronously.
func CreateSession() (*Session, error) {
	session := NewSessionWithID(newSessionID())
	session.mu.Lock()
	session.markDirtyLocked()
	session.mu.Unlock()
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
func FlushRegisteredSessions(force bool) error {
	sessionRegistryMu.RLock()
	sessions := make([]*Session, 0, len(sessionRegistry))
	registeredIDs := make(map[string]struct{}, len(sessionRegistry))
	for _, session := range sessionRegistry {
		sessions = append(sessions, session)
		if session != nil {
			registeredIDs[session.ID] = struct{}{}
		}
	}
	sessionRegistryMu.RUnlock()

	var firstErr error
	now := time.Now()
	for _, session := range sessions {
		if session == nil {
			continue
		}

		if session.IsAbandoned() || session.isExpiredAt(now) {
			err := session.Delete()
			if err != nil && firstErr == nil {
				firstErr = err
			}
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

	if err := cleanupExpiredSessionFilesOnDisk(now, registeredIDs); err != nil && firstErr == nil {
		firstErr = err
	}

	return firstErr
}

// cleanupExpiredSessionFilesOnDisk removes expired persisted sessions that are not registered in memory.
func cleanupExpiredSessionFilesOnDisk(now time.Time, registeredIDs map[string]struct{}) error {
	entries, err := os.ReadDir(sessionStorageDir)
	if err != nil {
		if os.IsNotExist(err) {
			return nil
		}
		return err
	}

	var firstErr error
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := entry.Name()
		if !strings.HasSuffix(name, ".g3ses") {
			continue
		}

		sessionID := strings.TrimSuffix(name, ".g3ses")
		if _, ok := registeredIDs[sessionID]; ok {
			continue
		}

		path := filepath.Join(sessionStorageDir, name)
		expired, checkErr := persistedSessionFileExpired(path, now)
		if checkErr != nil {
			if firstErr == nil {
				firstErr = checkErr
			}
			continue
		}
		if !expired {
			continue
		}

		if removeErr := os.Remove(path); removeErr != nil && !os.IsNotExist(removeErr) {
			if firstErr == nil {
				firstErr = removeErr
			}
		}
	}

	return firstErr
}

// persistedSessionFileExpired checks expiration using only timeout and last access fields from binary header.
func persistedSessionFileExpired(path string, now time.Time) (bool, error) {
	file, err := os.Open(path)
	if err != nil {
		if os.IsNotExist(err) {
			return false, nil
		}
		return false, err
	}
	defer file.Close()

	// Magic [6] + Version [1] + ID [24] + Created [8] + Accessed [8] + Timeout [2] = 49 bytes
	var header [49]byte
	n, err := file.Read(header[:])
	if err != nil || n < 49 {
		return true, nil // Treat corrupt as expired
	}

	if string(header[0:6]) != "G3SES\x00" {
		return true, nil
	}

	lastAccessedUnix := int64(binary.LittleEndian.Uint64(header[39:47]))
	timeoutMinutes := int16(binary.LittleEndian.Uint16(header[47:49]))

	if timeoutMinutes <= 0 {
		timeoutMinutes = defaultSessionTimeoutMinutes
	}

	lastAccessed := time.Unix(lastAccessedUnix, 0)
	return now.Sub(lastAccessed) > time.Duration(timeoutMinutes)*time.Minute, nil
}

// SetSessionStorageDir changes session persistence directory.
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

// newSessionID generates a random session ID (24 uppercase letters).
func newSessionID() string {
	b := make([]byte, 24)
	if _, err := rand.Read(b); err != nil {
		ts := strings.ReplaceAll(time.Now().UTC().Format("20060102150405.000000000"), ".", "")
		for len(ts) < 24 {
			ts += ts
		}
		ts = ts[:24]
		result := make([]byte, 24)
		for i := range 24 {
			result[i] = 'A' + (ts[i]-'0')%26
		}
		return string(result)
	}
	result := make([]byte, 24)
	for i := range 24 {
		result[i] = 'A' + (b[i] % 26)
	}
	return string(result)
}

// sessionFilePath builds the absolute session binary file path.
func sessionFilePath(sessionID string) string {
	safeID := strings.ReplaceAll(sessionID, "/", "_")
	safeID = strings.ReplaceAll(safeID, "\\", "_")
	return filepath.Join(sessionStorageDir, safeID+".g3ses")
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

// isExpiredAt reports whether timeout elapsed using a provided timestamp.
func (s *Session) isExpiredAt(now time.Time) bool {
	s.mu.RLock()
	defer s.mu.RUnlock()

	timeout := s.Timeout
	if timeout <= 0 {
		timeout = defaultSessionTimeoutMinutes
	}
	return now.Sub(s.LastAccessed) > time.Duration(timeout)*time.Minute
}

// Binary serialization helpers (Zero Reflection)

func writeUint32(buf *bytes.Buffer, v uint32) {
	var b [4]byte
	binary.LittleEndian.PutUint32(b[:], v)
	buf.Write(b[:])
}

func readUint32(r *bytes.Reader) uint32 {
	var b [4]byte
	_, _ = r.Read(b[:])
	return binary.LittleEndian.Uint32(b[:])
}

func readUint64(r *bytes.Reader) uint64 {
	var b [8]byte
	_, _ = r.Read(b[:])
	return binary.LittleEndian.Uint64(b[:])
}

func serializeApplicationMap(buf *bytes.Buffer, m map[string]ApplicationValue) {
	writeUint32(buf, uint32(len(m)))
	for k, v := range m {
		writeUint32(buf, uint32(len(k)))
		buf.WriteString(k)
		serializeApplicationValue(buf, v)
	}
}

func deserializeApplicationMap(r *bytes.Reader) map[string]ApplicationValue {
	count := readUint32(r)
	m := make(map[string]ApplicationValue, count)
	for range count {
		kLen := readUint32(r)
		kBuf := make([]byte, kLen)
		_, _ = r.Read(kBuf)
		k := string(kBuf)
		val, _ := deserializeApplicationValue(r)
		m[k] = val
	}
	return m
}

func serializeApplicationValue(buf *bytes.Buffer, v ApplicationValue) {
	buf.WriteByte(byte(v.Type))
	switch v.Type {
	case ApplicationValueString:
		writeUint32(buf, uint32(len(v.Str)))
		buf.WriteString(v.Str)
	case ApplicationValueInteger:
		var b [8]byte
		binary.LittleEndian.PutUint64(b[:], uint64(v.Num))
		buf.Write(b[:])
	case ApplicationValueDouble:
		var b [8]byte
		binary.LittleEndian.PutUint64(b[:], math.Float64bits(v.Flt))
		buf.Write(b[:])
	case ApplicationValueBool:
		if v.Num != 0 {
			buf.WriteByte(1)
		} else {
			buf.WriteByte(0)
		}
	case ApplicationValueArray:
		writeUint32(buf, uint32(v.ArrLower))
		writeUint32(buf, uint32(len(v.Arr)))
		for i := range v.Arr {
			serializeApplicationValue(buf, v.Arr[i])
		}
	}
}

func deserializeApplicationValue(r *bytes.Reader) (ApplicationValue, error) {
	t, err := r.ReadByte()
	if err != nil {
		return ApplicationValue{}, err
	}
	v := ApplicationValue{Type: ApplicationValueType(t)}
	switch v.Type {
	case ApplicationValueString:
		l := readUint32(r)
		str := make([]byte, l)
		_, _ = r.Read(str)
		v.Str = string(str)
	case ApplicationValueInteger:
		v.Num = int64(readUint64(r))
	case ApplicationValueDouble:
		v.Flt = math.Float64frombits(readUint64(r))
	case ApplicationValueBool:
		b, _ := r.ReadByte()
		if b != 0 {
			v.Num = 1
		} else {
			v.Num = 0
		}
	case ApplicationValueArray:
		v.ArrLower = int(readUint32(r))
		l := readUint32(r)
		v.Arr = make([]ApplicationValue, l)
		for i := range l {
			elem, err := deserializeApplicationValue(r)
			if err != nil {
				return v, err
			}
			v.Arr[i] = elem
		}
	}
	return v, nil
}

package server

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

// SessionManager handles file-based session storage
type SessionManager struct {
	sessionDir string
	mu         sync.RWMutex
}

// SessionData represents session data stored in file
type SessionData struct {
	ID           string                 `json:"id"`
	Data         map[string]interface{} `json:"data"`
	CreatedAt    time.Time              `json:"created_at"`
	LastAccessed time.Time              `json:"last_accessed"`
	Timeout      int                    `json:"timeout"` // Minutes, default 20
}

// globalSessionManager is the singleton session manager
var globalSessionManager *SessionManager
var sessionManagerOnce sync.Once

// GetSessionManager returns the global session manager instance
func GetSessionManager() *SessionManager {
	sessionManagerOnce.Do(func() {
		globalSessionManager = &SessionManager{
			sessionDir: "temp/session",
		}
		// Ensure session directory exists
		if err := os.MkdirAll(globalSessionManager.sessionDir, 0755); err != nil {
			fmt.Printf("Warning: Failed to create session directory: %v\n", err)
		}
	})
	return globalSessionManager
}

// getSessionFilePath returns the file path for a session ID
func (sm *SessionManager) getSessionFilePath(sessionID string) string {
	// Use sanitized session ID as filename
	safeID := strings.ReplaceAll(sessionID, "/", "_")
	safeID = strings.ReplaceAll(safeID, "\\", "_")
	return filepath.Join(sm.sessionDir, safeID+".json")
}

// LoadSession loads session data from file
func (sm *SessionManager) LoadSession(sessionID string) (*SessionData, error) {
	sm.mu.RLock()
	defer sm.mu.RUnlock()

	filePath := sm.getSessionFilePath(sessionID)

	// Check if file exists
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, fmt.Errorf("session not found")
	}

	// Read file
	data, err := os.ReadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to read session file: %w", err)
	}

	// Parse JSON
	var session SessionData
	if err := json.Unmarshal(data, &session); err != nil {
		return nil, fmt.Errorf("failed to parse session data: %w", err)
	}

	// Check if session has expired
	timeout := time.Duration(session.Timeout) * time.Minute
	if time.Since(session.LastAccessed) > timeout {
		// Session expired, delete file
		sm.mu.RUnlock()
		sm.DeleteSession(sessionID)
		sm.mu.RLock()
		return nil, fmt.Errorf("session expired")
	}

	// Update last accessed time
	session.LastAccessed = time.Now()

	return &session, nil
}

// SaveSession saves session data to file
func (sm *SessionManager) SaveSession(session *SessionData) error {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	filePath := sm.getSessionFilePath(session.ID)

	// Ensure directory exists
	if err := os.MkdirAll(sm.sessionDir, 0755); err != nil {
		return fmt.Errorf("failed to create session directory: %w", err)
	}

	// Convert to JSON
	data, err := json.Marshal(session)
	if err != nil {
		return fmt.Errorf("failed to marshal session data: %w", err)
	}

	// Write to file
	if err := os.WriteFile(filePath, data, 0600); err != nil {
		return fmt.Errorf("failed to write session file: %w", err)
	}

	return nil
}

// DeleteSession deletes session file
func (sm *SessionManager) DeleteSession(sessionID string) error {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	filePath := sm.getSessionFilePath(sessionID)

	// Remove file if exists
	if err := os.Remove(filePath); err != nil && !os.IsNotExist(err) {
		return fmt.Errorf("failed to delete session file: %w", err)
	}

	return nil
}

// CreateSession creates a new session
func (sm *SessionManager) CreateSession(sessionID string) (*SessionData, error) {
	now := time.Now()
	session := &SessionData{
		ID:           sessionID,
		Data:         make(map[string]interface{}),
		CreatedAt:    now,
		LastAccessed: now,
		Timeout:      20, // Default ASP timeout in minutes
	}

	// Save to file
	if err := sm.SaveSession(session); err != nil {
		return nil, err
	}

	return session, nil
}

// GetOrCreateSession retrieves an existing session or creates a new one
// Returns the session data and a boolean indicating if the session was newly created
func (sm *SessionManager) GetOrCreateSession(sessionID string) (*SessionData, bool, error) {
	// Try to load existing session
	session, err := sm.LoadSession(sessionID)
	if err == nil {
		return session, false, nil // Existing session
	}

	// Create new session if not found or expired
	session, err = sm.CreateSession(sessionID)
	if err != nil {
		return nil, false, err
	}
	return session, true, nil // New session
}

// CleanupExpiredSessions removes all expired session files
func (sm *SessionManager) CleanupExpiredSessions() error {
	sm.mu.Lock()
	defer sm.mu.Unlock()

	// Read all session files
	files, err := os.ReadDir(sm.sessionDir)
	if err != nil {
		if os.IsNotExist(err) {
			return nil // Directory doesn't exist yet
		}
		return fmt.Errorf("failed to read session directory: %w", err)
	}

	// Check each file
	for _, file := range files {
		if file.IsDir() || !strings.HasSuffix(file.Name(), ".json") {
			continue
		}

		filePath := filepath.Join(sm.sessionDir, file.Name())

		// Read and parse session data
		data, err := os.ReadFile(filePath)
		if err != nil {
			continue
		}

		var session SessionData
		if err := json.Unmarshal(data, &session); err != nil {
			continue
		}

		// Check if expired
		timeout := time.Duration(session.Timeout) * time.Minute
		if time.Since(session.LastAccessed) > timeout {
			// Remove expired session
			os.Remove(filePath)
		}
	}

	return nil
}

// StartCleanupRoutine starts a background goroutine to cleanup expired sessions
func (sm *SessionManager) StartCleanupRoutine(interval time.Duration) {
	go func() {
		ticker := time.NewTicker(interval)
		defer ticker.Stop()

		for range ticker.C {
			if err := sm.CleanupExpiredSessions(); err != nil {
				fmt.Printf("Session cleanup error: %v\n", err)
			}
		}
	}()
}

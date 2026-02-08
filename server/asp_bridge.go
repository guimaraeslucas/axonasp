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
	"crypto/rand"
	"fmt"
	"net/http"
	"strings"

	"g3pix.com.br/axonasp/asp"
)

// ASPProcessorConfig contains configuration for ASP processing
type ASPProcessorConfig struct {
	RootDir       string
	ScriptTimeout int  // in seconds
	DebugASP      bool // Enable debug output for ASP parsing and execution
}

// COM provider mode for Access connections: "auto" or "code"
var comProviderMode = "auto"

// SetCOMProviderMode configures how Access OLEDB providers are selected.
func SetCOMProviderMode(mode string) {
	mode = strings.ToLower(strings.TrimSpace(mode))
	if mode != "code" && mode != "auto" {
		mode = "auto"
	}
	comProviderMode = mode
}

// GetCOMProviderMode returns the configured COM provider mode.
func GetCOMProviderMode() string {
	return comProviderMode
}

// ASPProcessor handles ASP file execution
// Delegates to ASPExecutor for actual code execution
type ASPProcessor struct {
	config *ASPProcessorConfig
}

// NewASPProcessor creates a new ASP processor
func NewASPProcessor(config *ASPProcessorConfig) *ASPProcessor {
	if config == nil {
		config = &ASPProcessorConfig{
			RootDir:       "./www",
			ScriptTimeout: 30,
		}
	}
	return &ASPProcessor{
		config: config,
	}
}

// ExecuteASPFile processes and executes an ASP file
// Takes the file content as string and returns the rendered output
// Delegates to ASPExecutor in executor.go
func (ap *ASPProcessor) ExecuteASPFile(fileContent string, filePath string, w http.ResponseWriter, r *http.Request) error {
	resolvedContent, parsedResult, err := ap.getParsed(filePath, fileContent)
	if err != nil {
		return err
	}

	// Create the executor with configuration
	executor := NewASPExecutor(ap.config)

	// Generate session ID from request cookie or create new one
	sessionID := generateSessionID(r)

	// Execute using cached parse tree
	return executor.ExecuteWithParsed(resolvedContent, parsedResult, filePath, w, r, sessionID)
}

// GetConfig returns the configuration of this ASP processor
func (ap *ASPProcessor) GetConfig() *ASPProcessorConfig {
	return ap.config
}

// generateSessionID creates or retrieves a session ID from request cookies
func generateSessionID(r *http.Request) string {
	// Look for existing ASPSESSIONID cookie
	if cookie, err := r.Cookie("ASPSESSIONID"); err == nil {
		return cookie.Value
	}

	// Generate new session ID
	return generateUniqueID()
}

// generateUniqueID generates a unique identifier for sessions
func generateUniqueID() string {
	// Simple implementation - in production use crypto/rand with proper UUID
	b := make([]byte, 16)
	_, err := rand.Read(b)
	if err != nil {
		return "AXONINVALIDSESSION"
	}
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	result := fmt.Sprintf("%x-%x-%x-%x-%x", b[0:4], b[4:6], b[6:8], b[8:10], b[10:])
	return fmt.Sprintf("AXON%s", result)
}

func (ap *ASPProcessor) getParsed(filePath, rawContent string) (string, *asp.ASPParserResult, error) {
	parsingOptions := &asp.ASPParsingOptions{
		SaveComments:      false,
		StrictMode:        false,
		AllowImplicitVars: true,
		DebugMode:         ap.config.DebugASP,
	}
	resolvedContent, result, err := asp.ParseWithCache(rawContent, filePath, ap.config.RootDir, parsingOptions)
	if err != nil {
		return "", nil, fmt.Errorf("failed to parse ASP code: %w", err)
	}

	if len(result.Errors) > 0 {
		return "", nil, fmt.Errorf("ASP parse error: %v", result.Errors[0])
	}
	return resolvedContent, result, nil
}

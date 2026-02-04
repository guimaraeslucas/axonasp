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
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// WebConfigError represents a custom error handler configuration
type WebConfigError struct {
	StatusCode   string `xml:"statusCode,attr"`
	Path         string `xml:"path,attr"`
	ResponseMode string `xml:"responseMode,attr"`
}

// WebConfigHTTPErrors contains the httpErrors configuration
type WebConfigHTTPErrors struct {
	Errors []WebConfigError `xml:"error"`
}

// WebConfigSystemWebServer contains the system.webServer configuration
type WebConfigSystemWebServer struct {
	HTTPErrors WebConfigHTTPErrors `xml:"httpErrors"`
}

// WebConfig represents the root web.config structure
type WebConfig struct {
	XMLName         xml.Name                 `xml:"configuration"`
	SystemWebServer WebConfigSystemWebServer `xml:"system.webServer"`
}

// WebConfigParser handles parsing and retrieving web.config settings
type WebConfigParser struct {
	config     *WebConfig
	configPath string
	rootDir    string
}

// NewWebConfigParser creates a new web.config parser
func NewWebConfigParser(rootDir string) *WebConfigParser {
	return &WebConfigParser{
		rootDir:    rootDir,
		configPath: filepath.Join(rootDir, "web.config"),
	}
}

// Load parses the web.config file
func (p *WebConfigParser) Load() error {
	// Check if file exists
	if _, err := os.Stat(p.configPath); os.IsNotExist(err) {
		return fmt.Errorf("web.config not found at %s", p.configPath)
	}

	// Read file
	data, err := os.ReadFile(p.configPath)
	if err != nil {
		return fmt.Errorf("failed to read web.config: %w", err)
	}

	// Parse XML
	var config WebConfig
	err = xml.Unmarshal(data, &config)
	if err != nil {
		return fmt.Errorf("failed to parse web.config XML: %w", err)
	}

	p.config = &config
	return nil
}

// GetErrorHandlerPath returns the path for a specific status code
// Returns empty string if no handler is configured
func (p *WebConfigParser) GetErrorHandlerPath(statusCode int) (string, string) {
	if p.config == nil {
		return "", ""
	}

	statusCodeStr := fmt.Sprintf("%d", statusCode)
	for _, errorConfig := range p.config.SystemWebServer.HTTPErrors.Errors {
		if errorConfig.StatusCode == statusCodeStr {
			// Clean the path - remove leading slash if present
			path := strings.TrimPrefix(errorConfig.Path, "/")
			return path, errorConfig.ResponseMode
		}
	}

	return "", ""
}

// GetFullErrorHandlerPath returns the full file system path for an error handler
func (p *WebConfigParser) GetFullErrorHandlerPath(statusCode int) (string, string) {
	relativePath, responseMode := p.GetErrorHandlerPath(statusCode)
	if relativePath == "" {
		return "", ""
	}

	// Build full path relative to root directory
	fullPath := filepath.Join(p.rootDir, relativePath)
	return fullPath, responseMode
}

// IsLoaded returns true if web.config has been successfully loaded
func (p *WebConfigParser) IsLoaded() bool {
	return p.config != nil
}

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Jeffrey He (@jeffreyheping)
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
package main

import (
	"bufio"
	"fmt"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

// PathAlias maps a virtual URL prefix to a real filesystem directory.
// Example: VirtualPrefix="/music/" RealDir="D:\Music"
type PathAlias struct {
	VirtualPrefix string
	RealDir       string
}

var (
	pathAliases   []PathAlias
	pathAliasesMu sync.RWMutex
	lastAliasLoad time.Time
)

// aliasFlag implements flag.Value for repeated --alias /prefix=real/path flags.
type aliasFlag []PathAlias

func (a *aliasFlag) String() string { return "" }

func (a *aliasFlag) Set(value string) error {
	parts := strings.SplitN(value, "=", 2)
	if len(parts) != 2 {
		return fmt.Errorf("invalid alias format %q (expected /prefix=real/path)", value)
	}
	vp := strings.TrimSpace(parts[0])
	rd := strings.TrimSpace(parts[1])
	if vp == "" || rd == "" {
		return fmt.Errorf("alias prefix and path must not be empty")
	}
	if !strings.HasPrefix(vp, "/") {
		vp = "/" + vp
	}
	if !strings.HasSuffix(vp, "/") {
		vp += "/"
	}
	*a = append(*a, PathAlias{VirtualPrefix: vp, RealDir: rd})
	return nil
}

// cliAliases holds aliases set via --alias flags (loaded once, never re-read).
var cliAliases aliasFlag

// LoadPathAliases reads virtual path mappings from data/path_aliases.dat.
// File format: one mapping per line: /virtual_prefix|real_dir_path
// Lines starting with ; or # are comments. Blank lines are skipped.
func LoadPathAliases(appDir string) error {
	configPath := filepath.Join(appDir, "data", "path_aliases.dat")
	f, err := os.Open(configPath)
	if os.IsNotExist(err) {
		pathAliasesMu.Lock()
		pathAliases = mergeAliases(nil, cliAliases)
		pathAliasesMu.Unlock()
		return nil
	}
	if err != nil {
		return err
	}
	defer f.Close()

	var fileAliases []PathAlias
	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "" || line[0] == ';' || line[0] == '#' {
			continue
		}
		parts := strings.SplitN(line, "|", 2)
		if len(parts) != 2 {
			continue
		}
		vp := strings.TrimSpace(parts[0])
		rd := strings.TrimSpace(parts[1])
		if vp == "" || rd == "" {
			continue
		}
		if !strings.HasPrefix(vp, "/") {
			vp = "/" + vp
		}
		if !strings.HasSuffix(vp, "/") {
			vp += "/"
		}
		fileAliases = append(fileAliases, PathAlias{VirtualPrefix: vp, RealDir: rd})
	}

	pathAliasesMu.Lock()
	pathAliases = mergeAliases(fileAliases, cliAliases)
	pathAliasesMu.Unlock()

	return scanner.Err()
}

// mergeAliases combines file-based and CLI-based aliases.
// CLI aliases take priority over file aliases for the same prefix.
func mergeAliases(file, cli []PathAlias) []PathAlias {
	result := make([]PathAlias, 0, len(file)+len(cli))
	result = append(result, file...)
	for _, ca := range cli {
		found := false
		for i := range result {
			if result[i].VirtualPrefix == ca.VirtualPrefix {
				result[i].RealDir = ca.RealDir
				found = true
				break
			}
		}
		if !found {
			result = append(result, ca)
		}
	}
	return result
}

// resolveAlias checks if the request path matches any virtual prefix,
// reloading the config file if it has been >500ms since last load.
func resolveAlias(urlPath string) (PathAlias, string) {
	ensureAliasesLoaded()
	pathAliasesMu.RLock()
	defer pathAliasesMu.RUnlock()
	for _, a := range pathAliases {
		if strings.HasPrefix(urlPath, a.VirtualPrefix) {
			relPath := strings.TrimPrefix(urlPath, a.VirtualPrefix)
			return a, relPath
		}
	}
	return PathAlias{}, ""
}

// ensureAliasesLoaded reloads the alias config if >500ms have passed.
// This allows ASP pages to write path_aliases.dat and have the Go server
// pick up changes on the next request without restarting.
func ensureAliasesLoaded() {
	now := time.Now()
	if now.Sub(lastAliasLoad) < 500*time.Millisecond {
		return
	}
	lastAliasLoad = now
	_ = LoadPathAliases(appDir)
}

// ServeAliasFile serves a static file from an aliased directory.
// It rejects path traversal attempts and ensures the resolved path
// stays within the aliased directory.
func ServeAliasFile(w http.ResponseWriter, r *http.Request, alias PathAlias, relPath string) {
	if relPath == "" {
		// List files in the aliased directory (optional, for now just 404)
		http.Error(w, "Not Found", http.StatusNotFound)
		return
	}
	if strings.Contains(relPath, "..") {
		http.Error(w, "Bad Request", http.StatusBadRequest)
		return
	}

	filePath := filepath.Join(alias.RealDir, filepath.FromSlash(relPath))
	cleanPath := filepath.Clean(filePath)
	cleanAlias := filepath.Clean(alias.RealDir)

	if !strings.HasPrefix(cleanPath, cleanAlias+string(filepath.Separator)) &&
		cleanPath != cleanAlias {
		http.Error(w, "Forbidden", http.StatusForbidden)
		return
	}

	if _, err := os.Stat(cleanPath); err != nil {
		http.Error(w, "Not Found", http.StatusNotFound)
		return
	}

	serveStaticFile(w, r, cleanPath)
}

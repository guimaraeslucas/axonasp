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
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"log"
	"mime"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
)

// defaultPages defines the lookup order for directory index files.
var defaultPages = []string{"index.hta", "default.hta", "index.asp", "default.asp", "index.html", "default.html"}

// handleRequest is the main HTTP handler that routes requests to static file
// serving, HTA execution, or ASP execution based on the file extension.
// The response writer is wrapped so that appRuntimeJS is automatically
// injected into every HTML response (including ASP-generated pages).
func handleRequest(w http.ResponseWriter, r *http.Request) {
	// Top-level panic recovery: any panic in the handler, host, or ASP
	// engine is caught here so the process stays alive. Without this, a
	// panic in a code path not covered by inner recoveries would crash
	// axonhta.exe and the browser would show "ERR_CONNECTION_REFUSED".
	defer func() {
		if rec := recover(); rec != nil {
			log.Printf("Panic recovered in handleRequest: %v", rec)
		}
	}()

	// Heartbeat endpoint: page-side JS pings this every 5 seconds so the
	// server knows the browser window is still open.
	if r.URL.Path == "/__heartbeat__" {
		lastHeartbeat.Store(time.Now().Unix())
		w.WriteHeader(http.StatusOK)
		return
	}

	iw := newHTMLInjectWriter(w)
	defer iw.FinalFlush()

	path := r.URL.Path
	if path == "" {
		path = "/"
	}

	// Check virtual path aliases first (auto-reloads config every 500ms).
	if alias, relPath := resolveAlias(path); alias.VirtualPrefix != "" {
		ServeAliasFile(iw, r, alias, relPath)
		return
	}

	relativePath := strings.TrimPrefix(path, "/")
	fullPath := filepath.Join(appDir, filepath.FromSlash(relativePath))
	cleanPath := filepath.Clean(fullPath)

	if !strings.HasPrefix(cleanPath, appDir) && cleanPath != appDir {
		http.Error(w, "Forbidden", http.StatusForbidden)
		return
	}

	info, err := os.Stat(cleanPath)
	if os.IsNotExist(err) {
		http.Error(w, "Not Found", http.StatusNotFound)
		return
	}
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	if info.IsDir() {
		if !strings.HasSuffix(path, "/") {
			http.Redirect(w, r, path+"/", http.StatusMovedPermanently)
			return
		}

		for _, page := range defaultPages {
			candidate := filepath.Join(cleanPath, page)
			if candidateInfo, err := os.Stat(candidate); err == nil && !candidateInfo.IsDir() {
				fullPath = candidate
				break
			}
		}

		if info, err := os.Stat(fullPath); err != nil || info.IsDir() {
			http.Error(w, "Not Found", http.StatusNotFound)
			return
		}
	}

	ext := strings.ToLower(filepath.Ext(fullPath))
	if ext == ".hta" {
		// Treat .hta files as ASP with HTA tag stripping.
		executeHTA(iw, r, fullPath)
		return
	}
	if ext != ".asp" && ext != ".vbs" {
		serveStaticFile(iw, r, fullPath)
		return
	}

	executeASP(iw, r, fullPath)
}

// serveStaticFile serves a static file with the appropriate Content-Type
// header based on its extension.
func serveStaticFile(w http.ResponseWriter, r *http.Request, filePath string) {
	ext := strings.ToLower(filepath.Ext(filePath))
	if contentType := mime.TypeByExtension(ext); contentType != "" {
		w.Header().Set("Content-Type", contentType)
	}
	http.ServeFile(w, r, filePath)
}

// executeASP compiles and executes an ASP file within a DesktopHost,
// recovering from panics and rendering any compilation or runtime errors
// as an HTML error page.
func executeASP(w http.ResponseWriter, r *http.Request, filePath string) {
	host := NewDesktopHost(w, r, appDir, app)

	program, err := scriptCache.LoadOrCompileWithOptions(filePath, axonvm.ScriptCompileOptions{
		IncludeSiteRoot: host.Server().MapPath("/"),
	})
	if err != nil {
		aspErr := axonvm.CompilerErrorToASPError(err, filePath)
		renderError(w, "Compilation Error", aspErr)
		return
	}

	vm := axonvm.AcquireVMFromCachedProgram(program)
	vm.SetHost(host)

	type vmResult struct{ err error }
	done := make(chan vmResult, 1)
	go func() {
		// Outer recovery: protects vm.Release() and channel send, which
		// run outside the inner recovery that only covers vm.Run().
		// Without this, a panic in vm.Release() (e.g. during ADODB/COM
		// cleanup in CleanupRequestResources) kills the entire process.
		defer func() {
			if rec := recover(); rec != nil {
				done <- vmResult{err: fmt.Errorf("panic recovered in VM goroutine: %v", rec)}
			}
		}()
		defer vm.Release()
		runErr := func() (err error) {
			defer func() {
				if recovered := recover(); recovered != nil {
					err = fmt.Errorf("panic recovered in vm.Run: %v", recovered)
				}
			}()
			return vm.Run()
		}()
		done <- vmResult{err: runErr}
	}()

	res := <-done
	if res.err != nil {
		aspErr := axonvm.RuntimeErrorToASPError(res.err, filePath)
		renderError(w, "Runtime Error", aspErr)
		return
	}

	host.PersistSession()
	host.Response().Flush()
	host.Response().ReleaseBuffer()
}

// htaCacheDir stores cleaned .hta content for ScriptCache reuse.
// Each .hta file maps to a deterministic path based on content hash,
// so the ScriptCache recognizes the same file across requests.
var htaCacheDir string

// init creates a temporary directory for cleaned .hta file content so
// the ScriptCache can reuse compiled bytecode across requests.
func init() {
	d, err := os.MkdirTemp("", "axonhta-cache-")
	if err == nil {
		htaCacheDir = d
	}
}

// executeHTA reads the .hta file, strips the <hta:application> tag,
// then executes the remaining content as ASP. It derives a deterministic
// temp path from the file's mtime and content hash so ScriptCache can
// reuse the compiled bytecode across requests.
func executeHTA(w http.ResponseWriter, r *http.Request, filePath string) {
	data, err := os.ReadFile(filePath)
	if err != nil {
		http.Error(w, "Failed to read HTA file", http.StatusInternalServerError)
		return
	}

	cleaned := StripHTATag(string(data))

	// Derive a deterministic path from the file's mtime + content hash.
	// This ensures ScriptCache reuses the compiled bytecode on subsequent requests.
	info, _ := os.Stat(filePath)
	mtime := ""
	if info != nil {
		mtime = fmt.Sprintf("%d", info.ModTime().UnixNano())
	}
	hashInput := filePath + ":" + mtime + ":" + cleaned
	hash := sha256.Sum256([]byte(hashInput))
	hashHex := hex.EncodeToString(hash[:])[:16]
	baseName := strings.TrimSuffix(filepath.Base(filePath), filepath.Ext(filePath))
	tmpPath := filepath.Join(htaCacheDir, baseName+"_"+hashHex+".asp")

	// Only rewrite if the cached file is missing or stale.
	if _, err := os.Stat(tmpPath); err != nil {
		if err := os.WriteFile(tmpPath, []byte(cleaned), 0644); err != nil {
			http.Error(w, "Failed to process HTA file", http.StatusInternalServerError)
			return
		}
	}

	executeASP(w, r, tmpPath)
}

// renderError writes an HTML error page to the response with details about
// the ASP compilation or runtime error.
func renderError(w http.ResponseWriter, stage string, err *asp.ASPError) {
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	w.WriteHeader(http.StatusInternalServerError)
	fmt.Fprintf(w, "<html><head><meta charset=\"utf-8\"><title>Error - AxonHTA</title><style>body{font-family:sans-serif;padding:40px;background:#f5f5f5}.card{background:#fff;border:1px solid #ddd;padding:24px;max-width:700px;margin:0 auto;border-radius:8px}h1{color:#d32f2f;font-size:20px}td{padding:8px 12px;border:1px solid #eee}td.k{background:#f8f8f8;font-weight:600;width:100px}</style></head><body><div class=\"card\"><h1>%s</h1><table><tr><td class=\"k\">Source</td><td>%s</td></tr><tr><td class=\"k\">Description</td><td>%s</td></tr><tr><td class=\"k\">File</td><td>%s</td></tr><tr><td class=\"k\">Line</td><td>%d</td></tr><tr><td class=\"k\">Column</td><td>%d</td></tr></table></div></body></html>",
		stage, err.Source, err.Description, err.File, err.Line, err.Column)
}

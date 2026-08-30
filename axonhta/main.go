/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Jeffrey He (@jeffreyheping), Lucas Guimarães (@guimaraeslucas)
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
//Use go install github.com/josephspurrier/goversioninfo/cmd/goversioninfo@latest
//Then run "go generate" in the project root to embed version info into the executable
//You need to specify -64=false/-arm=true if you're trying to create an 32-bit or ARM windows binary, this is required by the new version of golang
//go:generate goversioninfo -icon=icon_hta.ico -64=true
package main

import (
	"context"
	"flag"
	"fmt"
	"io"
	"log"
	"mime"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"sync"
	"time"

	"g3pix.com.br/axonasp/v2/axonvm"
	"g3pix.com.br/axonasp/v2/axonvm/asp"
)

var (
	// Version is set at build time via -ldflags "-X main.Version=..."
	Version = "dev"

	appDir      string
	title       string
	width       int
	height      int
	port        string
	devMode     bool
	app         *asp.Application
	scriptCache *axonvm.ScriptCache

	// httpServer is the embedded HTTP server. Stored as a field so shutdown()
	// can gracefully close it.
	httpServer   *http.Server
	shutdownOnce sync.Once

	// logFile is the optional log file opened by setupLogFile().
	logFile *os.File
)

// init registers the SVG MIME type so static file serving works correctly.
func init() {
	_ = mime.AddExtensionType(".svg", "image/svg+xml; charset=utf-8;")
}

// htaConfig holds the parsed <hta:application> tag from the entry file.
var htaConfig *HtaConfig

// waitForServer polls the HTTP server's lightweight heartbeat endpoint until
// it responds or a timeout is reached. This ensures the server is ready
// before the browser window navigates to it, preventing "ERR_CONNECTION_REFUSED"
// on first launch.
//
// The heartbeat endpoint (/__heartbeat__) responds immediately with 200 OK
// without triggering any ASP compilation or execution, so the check is fast
// and reliable even on cold starts where the first page compilation takes
// several seconds.
func waitForServer(url string) {
	heartbeatURL := url + "__heartbeat__"
	client := &http.Client{Timeout: 300 * time.Millisecond}
	for range 100 {
		resp, err := client.Get(heartbeatURL)
		if err == nil {
			resp.Body.Close()
			return
		}
		time.Sleep(100 * time.Millisecond)
	}
	log.Printf("Warning: server not responding after 10 seconds, opening window anyway")
}

// main is the entry point for AxonHTA. It parses command-line flags,
// resolves the application directory, detects an HTA entry file for window
// configuration, starts an embedded HTTP server, and opens a Chromium app
// window (or system browser) pointing at the server.
func main() {
	// resolveHTAArg must run before flag.Parse so that drag-and-drop or
	// positional-argument .hta paths can seed appDir. The -app flag retains
	// priority: if the user explicitly passes -app=..., the default "./"
	// sentinel is replaced and appDir will be overwritten below by Abs().
	resolveHTAArg()

	flag.StringVar(&appDir, "app", appDir, "Application directory containing ASP/HTML/HTA files")
	flag.StringVar(&title, "title", "AxonHTA", "Window title")
	flag.IntVar(&width, "width", 1024, "Window width")
	flag.IntVar(&height, "height", 768, "Window height")
	flag.StringVar(&port, "port", "0", "HTTP server port (0 for random)")
	flag.Var(&cliAliases, "alias", "Virtual path alias, repeatable (e.g. --alias /music/=D:\\Music)")
	flag.BoolVar(&devMode, "dev", false, "Enable developer mode (F12 DevTools, no context menu suppression)")
	flag.Parse()

	absAppDir, err := filepath.Abs(appDir)
	if err != nil {
		log.Fatalf("Failed to resolve app directory: %v", err)
	}
	appDir = absAppDir

	// Re-anchor htaFileOverride to the (now-absolute) appDir so that if the
	// user also passed -app=..., we point at the right directory.
	if htaFileOverride != "" {
		htaFileOverride, _ = filepath.Abs(htaFileOverride)
	}

	// Set up log file (data/axonhta.log) alongside stdout.
	setupLogFile(filepath.Join(appDir, "data", "axonhta.log"))

	// Determine the HTA entry file for window configuration. Prefer the
	// explicitly supplied override (drag-and-drop / CLI argument) over the
	// default filesystem search so that title/size come from the right file.
	entryPath := htaFileOverride
	if entryPath == "" {
		entryPath = FindEntryFile(appDir)
	}
	if entryPath != "" {
		if cfg := ParseHTATag(entryPath); cfg != nil {
			htaConfig = cfg
			if cfg.ApplicationName != "" && title == "AxonHTA" {
				title = cfg.ApplicationName
			}
			log.Printf("Parsed <hta:application> from %s (applicationname=%q)",
				filepath.Base(entryPath), cfg.ApplicationName)
		}
	}

	app = asp.NewApplication()
	scriptCache = axonvm.NewScriptCache(axonvm.BytecodeCacheMemoryOnly, "", 64)

	// HTA apps are trusted desktop applications that need real-time filesystem
	// access (e.g. rescanning a music folder for new files). Disable the FSO
	// metadata cache so every GetFolder/FileExists/ReadDir reads from disk.
	axonvm.SetFSOCacheDisabled(true)

	// Load virtual path aliases from data/path_aliases.dat
	if err := LoadPathAliases(appDir); err != nil {
		log.Printf("Warning: failed to load path aliases: %v", err)
	}

	http.HandleFunc("/", handleRequest)

	listener, err := net.Listen("tcp", "127.0.0.1:"+port)
	if err != nil {
		log.Fatalf("Failed to start server: %v", err)
	}

	actualPort := listener.Addr().(*net.TCPAddr).Port
	baseURL := fmt.Sprintf("http://127.0.0.1:%d/", actualPort)

	// When the user supplied a specific .hta file (drag-and-drop or CLI),
	// open the browser directly at that file's URL path so it doesn't
	// depend on defaultPages discovery.
	url := baseURL
	if htaFileOverride != "" {
		rel, err := filepath.Rel(appDir, htaFileOverride)
		if err == nil {
			// Convert OS path separators to forward slashes for the URL.
			urlPath := filepath.ToSlash(rel)
			url = baseURL + urlPath
		}
	}

	log.Printf("AxonHTA %s starting...", Version)
	log.Printf("App directory: %s", appDir)
	log.Printf("Server URL: %s", url)
	if devMode {
		log.Println("Developer mode enabled")
	}

	httpServer = &http.Server{Handler: http.DefaultServeMux}

	go func() {
		if err := httpServer.Serve(listener); err != nil && err != http.ErrServerClosed {
			log.Fatalf("HTTP server error: %v", err)
		}
	}()

	waitForServer(baseURL)

	openWindow(url)
}

// shutdown performs a graceful shutdown of the HTTP server and cleans up
// temporary files. It is safe to call multiple times; only the first call
// has effect.
func shutdown() {
	shutdownOnce.Do(func() {
		log.Println("Shutting down AxonHTA...")

		// Gracefully stop the HTTP server (waits for in-flight requests).
		if httpServer != nil {
			ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
			defer cancel()
			if err := httpServer.Shutdown(ctx); err != nil {
				log.Printf("HTTP server shutdown error: %v", err)
			}
		}

		// Clean up the temporary HTA cache directory.
		if htaCacheDir != "" {
			_ = os.RemoveAll(htaCacheDir)
		}

		// Close the log file if one was opened.
		closeLogFile()

		os.Exit(0)
	})
}

// setupLogFile opens a log file and configures the standard log package to
// write to both stdout and the file. If the file cannot be opened (e.g. the
// data directory does not exist), logging falls back to stdout only.
func setupLogFile(path string) {
	dir := filepath.Dir(path)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return
	}
	f, err := os.OpenFile(path, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err != nil {
		return
	}
	logFile = f
	log.SetOutput(io.MultiWriter(os.Stdout, f))
}

// closeLogFile closes the log file if one was opened.
func closeLogFile() {
	if logFile != nil {
		_ = logFile.Close()
		logFile = nil
	}
}

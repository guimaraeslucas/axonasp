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
	"flag"
	"fmt"
	"log"
	"mime"
	"net"
	"net/http"
	"path/filepath"
	"time"

	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
)

var (
	appDir      string
	title       string
	width       int
	height      int
	port        string
	app         *asp.Application
	scriptCache *axonvm.ScriptCache
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
	for i := 0; i < 100; i++ {
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
	flag.StringVar(&appDir, "app", "./", "Application directory containing ASP/HTML/HTA files")
	flag.StringVar(&title, "title", "AxonHTA", "Window title")
	flag.IntVar(&width, "width", 1024, "Window width")
	flag.IntVar(&height, "height", 768, "Window height")
	flag.StringVar(&port, "port", "0", "HTTP server port (0 for random)")
	flag.Var(&cliAliases, "alias", "Virtual path alias, repeatable (e.g. --alias /music/=D:\\Music)")
	flag.Parse()

	absAppDir, err := filepath.Abs(appDir)
	if err != nil {
		log.Fatalf("Failed to resolve app directory: %v", err)
	}
	appDir = absAppDir

	// Try to find and parse an HTA entry file for window configuration.
	// Command-line flags take priority over HTA tag attributes.
	if entryPath := FindEntryFile(appDir); entryPath != "" {
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
	url := fmt.Sprintf("http://127.0.0.1:%d/", actualPort)

	log.Printf("AxonHTA starting...")
	log.Printf("App directory: %s", appDir)
	log.Printf("Server URL: %s", url)

	go func() {
		if err := http.Serve(listener, nil); err != nil && err != http.ErrServerClosed {
			log.Fatalf("HTTP server error: %v", err)
		}
	}()

	waitForServer(url)

	openWindow(url)
}

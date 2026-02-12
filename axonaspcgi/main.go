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
package main

//go:generate goversioninfo -icon=icon_cgi.ico

import (
	"flag"
	"fmt"
	"log"
	"net"
	"net/http"
	"net/http/fcgi"
	"os"
	"path/filepath"
	"strings"
	"time"

	"g3pix.com.br/axonasp/server"
	"github.com/joho/godotenv"
)

var (
	listenAddr    string
	webRoot       string
	timezone      string
	defaultPage   string
	scriptTimeout int
	debugASP      bool
)

func init() {
	// Load .env file if it exists
	_ = godotenv.Load()

	// Command line flags
	flag.StringVar(&listenAddr, "listen", getEnv("FCGI_LISTEN", "127.0.0.1:9000"), "FastCGI listen address (host:port or unix:/path/to/socket)")
	flag.StringVar(&webRoot, "root", getEnv("WEB_ROOT", "./www"), "Web root directory")
	flag.StringVar(&timezone, "timezone", getEnv("TIMEZONE", "America/Sao_Paulo"), "Server timezone")
	flag.StringVar(&defaultPage, "default", getEnv("DEFAULT_PAGE", "default.asp"), "Default page")
	flag.IntVar(&scriptTimeout, "timeout", getEnvInt("SCRIPT_TIMEOUT", 30), "Script execution timeout in seconds")
	flag.BoolVar(&debugASP, "debug", getEnvBool("DEBUG_ASP", false), "Enable ASP debug mode")
}

// dummyResponseWriter is a no-op http.ResponseWriter for Application_OnStart
type dummyResponseWriter struct {
	headers http.Header
	written bool
}

func (d *dummyResponseWriter) Header() http.Header {
	if d.headers == nil {
		d.headers = make(http.Header)
	}
	return d.headers
}

func (d *dummyResponseWriter) Write(b []byte) (int, error) {
	d.written = true
	return len(b), nil
}

func (d *dummyResponseWriter) WriteHeader(statusCode int) {
	d.written = true
}

func main() {
	flag.Parse()

	// Set timezone
	loc, err := time.LoadLocation(timezone)
	if err != nil {
		log.Printf("Warning: Could not load timezone %s, using UTC: %v", timezone, err)
		loc = time.UTC
	}
	time.Local = loc

	// Resolve absolute web root path
	absWebRoot, err := filepath.Abs(webRoot)
	if err != nil {
		log.Fatalf("Failed to resolve web root path: %v", err)
	}
	webRoot = absWebRoot

	// Verify web root exists
	if _, err := os.Stat(webRoot); os.IsNotExist(err) {
		log.Fatalf("Web root directory does not exist: %s", webRoot)
	}

	// Initialize session manager
	sessionManager := server.GetSessionManager()
	sessionManager.StartCleanupRoutine(20 * time.Minute)

	// Load global.asa if it exists
	globalASAManager := server.GetGlobalASAManager()
	err = globalASAManager.LoadGlobalASA(webRoot)
	if err == nil {
		log.Println("global.asa loaded successfully")

		// Execute Application_OnStart if defined
		if globalASAManager.HasApplicationOnStart() {
			log.Println("Executing Application_OnStart...")

			// Create temporary processor and executor
			processorConfig := &server.ASPProcessorConfig{
				RootDir:       webRoot,
				ScriptTimeout: scriptTimeout,
				DebugASP:      debugASP,
				UseVM:         false,
			}
			tempProcessor := server.NewASPProcessor(processorConfig)
			executor := server.NewASPExecutor(tempProcessor.GetConfig())

			// Create dummy writer and request for Application_OnStart
			dummyWriter := &dummyResponseWriter{}
			dummyRequest := &http.Request{Header: make(http.Header)}

			// Create execution context
			ctx := server.NewExecutionContext(dummyWriter, dummyRequest, "app_startup", time.Duration(scriptTimeout)*time.Second)

			// Execute Application_OnStart
			if err := globalASAManager.ExecuteApplicationOnStart(executor, ctx); err != nil {
				log.Printf("Warning: Error executing Application_OnStart: %v", err)
			} else {
				log.Println("Application_OnStart executed successfully")
			}
		}
	} else {
		log.Printf("No global.asa found or failed to load: %v", err)
	}

	// Create ASP processor for handling requests
	processorConfig := &server.ASPProcessorConfig{
		RootDir:       webRoot,
		ScriptTimeout: scriptTimeout,
		DebugASP:      debugASP,
		UseVM:         false, // VM mode disabled by default for FastCGI
	}
	processor := server.NewASPProcessor(processorConfig)

	// Create FastCGI handler
	handler := &FastCGIHandler{
		webRoot:     webRoot,
		defaultPage: defaultPage,
		processor:   processor,
	}

	// Determine listen type (TCP or Unix socket)
	var listener net.Listener
	if strings.HasPrefix(listenAddr, "unix:") {
		socketPath := strings.TrimPrefix(listenAddr, "unix:")

		// Remove existing socket file
		os.Remove(socketPath)

		listener, err = net.Listen("unix", socketPath)
		if err != nil {
			log.Fatalf("Failed to listen on Unix socket %s: %v", socketPath, err)
		}

		// Set socket permissions
		os.Chmod(socketPath, 0666)

		log.Printf("AxonASP FastCGI server listening on Unix socket: %s", socketPath)
	} else {
		listener, err = net.Listen("tcp", listenAddr)
		if err != nil {
			log.Fatalf("Failed to listen on %s: %v", listenAddr, err)
		}

		log.Printf("AxonASP FastCGI server listening on: %s", listenAddr)
	}

	log.Printf("Web root: %s", webRoot)
	log.Printf("Default page: %s", defaultPage)
	log.Printf("Script timeout: %d seconds", scriptTimeout)
	log.Printf("Debug mode: %v", debugASP)
	log.Println("Ready to accept FastCGI requests...")

	// Serve FastCGI requests
	if err := fcgi.Serve(listener, handler); err != nil {
		log.Fatalf("FastCGI server error: %v", err)
	}
}

// FastCGIHandler implements http.Handler for FastCGI requests
type FastCGIHandler struct {
	webRoot     string
	defaultPage string
	processor   *server.ASPProcessor
}

func (h *FastCGIHandler) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	// Get the script filename from FastCGI environment
	scriptFilename := r.Header.Get("X-Fcgi-Script-Filename")
	if scriptFilename == "" {
		// Try SCRIPT_FILENAME CGI variable
		scriptFilename = os.Getenv("SCRIPT_FILENAME")
	}

	// If still not found, construct from DOCUMENT_ROOT and SCRIPT_NAME
	if scriptFilename == "" {
		docRoot := r.Header.Get("X-Fcgi-Document-Root")
		if docRoot == "" {
			docRoot = h.webRoot
		}
		scriptName := r.URL.Path
		if scriptName == "" || scriptName == "/" {
			scriptName = "/" + h.defaultPage
		}
		scriptFilename = filepath.Join(docRoot, filepath.FromSlash(scriptName))
	}

	// Clean and validate the path
	scriptFilename = filepath.Clean(scriptFilename)

	// Check if file exists
	fileInfo, err := os.Stat(scriptFilename)
	if err != nil {
		if os.IsNotExist(err) {
			http.Error(w, "404 Not Found", http.StatusNotFound)
			return
		}
		http.Error(w, "500 Internal Server Error", http.StatusInternalServerError)
		log.Printf("Error accessing file %s: %v", scriptFilename, err)
		return
	}

	// If it's a directory, try default page
	if fileInfo.IsDir() {
		scriptFilename = filepath.Join(scriptFilename, h.defaultPage)
		if _, err := os.Stat(scriptFilename); err != nil {
			http.Error(w, "404 Not Found", http.StatusNotFound)
			return
		}
	}

	// Check if it's an ASP file
	ext := strings.ToLower(filepath.Ext(scriptFilename))
	if ext != ".asp" && ext != ".aspx" {
		// Not an ASP file, serve as static file
		http.ServeFile(w, r, scriptFilename)
		return
	}

	// Execute ASP file
	h.executeASP(w, r, scriptFilename)
}

func (h *FastCGIHandler) executeASP(w http.ResponseWriter, r *http.Request, filename string) {
	// Read the ASP file content
	content, err := os.ReadFile(filename)
	if err != nil {
		log.Printf("Error reading ASP file %s: %v", filename, err)
		http.Error(w, "500 Internal Server Error", http.StatusInternalServerError)
		return
	}

	// Parse form data for POST/PUT requests
	if r.Method == "POST" || r.Method == "PUT" {
		contentType := r.Header.Get("Content-Type")
		if strings.HasPrefix(contentType, "multipart/form-data") {
			// Parse multipart form (max 32MB in memory)
			r.ParseMultipartForm(32 << 20)
		} else {
			r.ParseForm()
		}
	} else {
		r.ParseForm()
	}

	// Execute the ASP file using the processor
	err = h.processor.ExecuteASPFile(string(content), filename, w, r)
	if err != nil {
		log.Printf("Error executing ASP file %s: %v", filename, err)

		// Check if headers were already sent
		if w.Header().Get("Content-Type") == "" {
			http.Error(w, "500 Internal Server Error", http.StatusInternalServerError)
		}
		return
	}
}

// Utility functions
func getEnv(key, defaultValue string) string {
	if value := os.Getenv(key); value != "" {
		return value
	}
	return defaultValue
}

func getEnvInt(key string, defaultValue int) int {
	if value := os.Getenv(key); value != "" {
		var intValue int
		if _, err := fmt.Sscanf(value, "%d", &intValue); err == nil {
			return intValue
		}
	}
	return defaultValue
}

func getEnvBool(key string, defaultValue bool) bool {
	if value := os.Getenv(key); value != "" {
		return strings.ToLower(value) == "true" || value == "1"
	}
	return defaultValue
}

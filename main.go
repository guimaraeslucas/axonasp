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

import (
	"fmt"
	"g3pix.com.br/axonasp/asp"
	"g3pix.com.br/axonasp/server"
	"net/http"
	_ "net/http/pprof"
	"os"
	"path/filepath"
	"runtime/debug"
	"strconv"
	"strings"
	"time"

	"github.com/joho/godotenv"
)

// Configuration variables
var (
	Port              = "4050"
	RootDir           = "./www"
	DefaultTimezone   = "America/Sao_Paulo"
	DefaultPage       = "default.asp"
	ScriptTimeout     = 30 // in seconds
	DebugASP          = false
	BlockedExtensions = ".asax,.ascx,.master,.skin,.browser,.sitemap,.config,.cs,.csproj,.vb,.vbproj,.webinfo,.licx,.resx,.resources,.mdb,.vjsproj,.java,.jsl,.ldb,.dsdgm,.ssdgm,.lsad,.ssmap,.cd,.dsprototype,.lsaprototype,.sdm,.sdmDocument,.mdf,.ldf,.ad,.dd,.ldd,.sd,.adprototype,.lddprototype,.exclude,.refresh,.compiled,.msgx,.vsdisco,.rules,.asa,.inc,.exe,.dll,.env,.config,.htaccess,.env.local,.json,.yaml,.yml"
)

func init() {
	// Load .env file
	err := godotenv.Load()
	if err != nil {
		fmt.Println("Info: No .env file found, using defaults or system environment. If you have an .env file, make sure it's in the currentworking directory.")
	}

	// Override configuration from Environment Variables
	if val := os.Getenv("SERVER_PORT"); val != "" {
		Port = val
	}
	if val := os.Getenv("WEB_ROOT"); val != "" {
		RootDir = val
	}
	if val := os.Getenv("TIMEZONE"); val != "" {
		DefaultTimezone = val
	}
	if val := os.Getenv("DEFAULT_PAGE"); val != "" {
		DefaultPage = val
	}
	if val := os.Getenv("SCRIPT_TIMEOUT"); val != "" {
		if i, err := strconv.Atoi(val); err == nil {
			ScriptTimeout = i
		}
	}
	if val := os.Getenv("DEBUG_ASP"); val == "TRUE" {
		DebugASP = true
	}
	if val := os.Getenv("BLOCKED_EXTENSIONS"); val != "" {
		BlockedExtensions = val
	}

	// Set timezone
	os.Setenv("TZ", DefaultTimezone)
}

// isBlockedExtension checks if a file extension is in the blocked list
func isBlockedExtension(ext string) bool {
	if ext == "" {
		return false
	}
	blockedList := strings.Split(BlockedExtensions, ",")
	for _, blocked := range blockedList {
		if strings.EqualFold(strings.TrimSpace(blocked), ext) {
			return true
		}
	}
	return false
}

func main() {
	fmt.Print("\033[2K")
	fmt.Printf("\033[48;5;240m\033[37mStarting G3pix AxonASP on http://localhost:%s ► \033[0m\n", Port)
	fmt.Printf("Serving files from %s\n", RootDir)

	http.HandleFunc("/", handleRequest)

	// Initialize session manager and start cleanup routine
	sessionManager := server.GetSessionManager()
	sessionManager.StartCleanupRoutine(20 * time.Minute) // Cleanup every 20 minutes

	// Load Global.asa file
	globalASAManager := server.GetGlobalASAManager()
	err := globalASAManager.LoadGlobalASA(RootDir)
	if err != nil {
		fmt.Printf("Warning: Failed to load Global.asa: %v\n", err)
	}

	fmt.Printf("Application_OnStart defined: %v\n", globalASAManager.HasApplicationOnStart())
	fmt.Printf("Session_OnStart defined: %v\n", globalASAManager.HasSessionOnStart())

	// Execute Application_OnStart if defined
	if globalASAManager.HasApplicationOnStart() {
		func() {
			// Wrap in a function with defer to catch any panics
			defer func() {
				if r := recover(); r != nil {
					fmt.Printf("[DEBUG] Error executing Application_OnStart: %v\n", r)
					debug.PrintStack()
				}
			}()

			// Create temporary executor for Application_OnStart
			processor := server.NewASPProcessor(&server.ASPProcessorConfig{
				RootDir:       RootDir,
				ScriptTimeout: ScriptTimeout,
				DebugASP:      DebugASP,
			})
			executor := server.NewASPExecutor(processor.GetConfig())

			// Create a dummy response writer for Application_OnStart
			// Since it's not tied to a specific request, we can use a no-op writer
			dummyWriter := NewDummyResponseWriter()

			// Create a dummy request
			dummyRequest := &http.Request{
				Header: make(http.Header),
			}

			// Create execution context for Application_OnStart
			ctx := server.NewExecutionContext(dummyWriter, dummyRequest, "app_startup", time.Duration(ScriptTimeout)*time.Second)
			ctx.RootDir = RootDir

			// Set the context in executor
			executor.SetContext(ctx)

			if err := globalASAManager.ExecuteApplicationOnStart(executor, ctx); err != nil {
				fmt.Printf("[DEBUG] Error in Application_OnStart: %v\n", err)
			} else {
				fmt.Println("[DEBUG] Application_OnStart executed successfully")
			}
		}()
	}

	if DebugASP {
		fmt.Println("[DEBUG] DEBUG_ASP mode is enabled")
		//Display build info for debugging purposes
		if info, ok := debug.ReadBuildInfo(); ok {
			fmt.Printf("[DEBUG] Go Version: %s\n", info.GoVersion)
			for _, setting := range info.Settings {
				switch setting.Key {
				case "vcs.revision":
					fmt.Printf("[DEBUG] VCS Revision: %s\n", setting.Value)
				case "vcs.time":
					fmt.Printf("[DEBUG] Build Date: %s\n", setting.Value)
				}
			}
		}
	}

	err = http.ListenAndServe(":"+Port, nil)
	if err != nil {
		fmt.Printf("Fatal error starting G3Pix AxonASP server:\n  %v\n", err)
	}
}

// DummyResponseWriter is a no-op response writer for Application_OnStart
type DummyResponseWriter struct{}

func (w *DummyResponseWriter) Header() http.Header {
	return make(http.Header)
}

func (w *DummyResponseWriter) Write([]byte) (int, error) {
	return 0, nil
}

func (w *DummyResponseWriter) WriteHeader(statusCode int) {
}

func NewDummyResponseWriter() *DummyResponseWriter {
	return &DummyResponseWriter{}
}

func handleRequest(w http.ResponseWriter, r *http.Request) {
	path := r.URL.Path
	if path == "/" {
		path = "/" + DefaultPage
	}

	fullPath := filepath.Join(RootDir, path)

	// Security check: prevent directory traversal
	cleanPath := filepath.Clean(fullPath)
	cleanRoot := filepath.Clean(RootDir)
	if !strings.HasPrefix(cleanPath, cleanRoot) {
		serveErrorPage(w, 403)
		return
	}

	// Security check: block direct access to restricted file extensions, we use not found for safety
	fileExt := strings.ToLower(filepath.Ext(fullPath))
	if isBlockedExtension(fileExt) {
		serveErrorPage(w, 404)
		return
	}

	// Check if file exists
	info, err := os.Stat(fullPath)
	if os.IsNotExist(err) {
		serveErrorPage(w, 404)
		return
	}

	// If it's a directory, try to serve the default page
	if info.IsDir() {
		fullPath = filepath.Join(fullPath, DefaultPage)
		if _, err := os.Stat(fullPath); os.IsNotExist(err) {
			serveErrorPage(w, 404)
			return
		}
	}

	// Serve static files if not ASP
	if !strings.HasSuffix(strings.ToLower(fullPath), ".asp") {
		http.ServeFile(w, r, fullPath)
		return
	}

	// Process ASP file
	content, err := asp.ReadFileText(fullPath)
	if err != nil {
		serveErrorPage(w, 500)
		return
	}

	// Recover from panics to avoid crashing server
	defer func() {
		if r := recover(); r != nil {
			fmt.Printf("G3Pix AxonASP Runtime panic in %s: %v\n", path, r)

			// Check if debug mode is enabled
			isDebug := os.Getenv("DEBUG_ASP") == "TRUE"

			if !isDebug {
				serveErrorPage(w, 500)
				return
			}

			w.Header().Set("Content-Type", "text/html; charset=utf-8")
			w.WriteHeader(http.StatusInternalServerError)

			fmt.Fprintf(w, "<br><hr style='border-top: 1px dashed red;'>\n")
			fmt.Fprintf(w, "<div style='color: red; font-family: monospace; background: #ffe6e6; padding: 10px; border: 1px solid red;'>\n")

			if isDebug {
				// Detailed error output with stack trace
				stack := string(debug.Stack())
				stack = strings.ReplaceAll(stack, "<", "&lt;")
				stack = strings.ReplaceAll(stack, ">", "&gt;")

				fmt.Fprintf(w, "<strong>G3pix AxonASP panic</strong><br>\n")
				fmt.Fprintf(w, "Error: %v<br>\n", r)
				fmt.Fprintf(w, "<pre>%s</pre>\n", stack)
			} else {
				// Simple error output
				fmt.Fprintf(w, "<strong>G3pix AxonASP error</strong><br>\n")
				fmt.Fprintf(w, "Description: %v<br>\n", r)
			}

			fmt.Fprintf(w, "</div>\n")
		}
	}()

	// Create ASP processor and execute
	processor := server.NewASPProcessor(&server.ASPProcessorConfig{
		RootDir:       RootDir,
		ScriptTimeout: ScriptTimeout,
		DebugASP:      DebugASP,
	})

	err = processor.ExecuteASPFile(content, fullPath, w, r)
	if err != nil {
		//stack := string(debug.Stack())
		fmt.Printf("[DEBUG] ASP processing error in %s: %v\n", path, err)
		//fmt.Printf("[DEBUG] STACK %s\n", stack)
	}
}

// serveErrorPage serves a custom HTML error page from the errorpages directory
func serveErrorPage(w http.ResponseWriter, statusCode int) {
	filename := fmt.Sprintf("%d.html", statusCode)
	filePath := filepath.Join("errorpages", filename)
	content, err := os.ReadFile(filePath)
	if err != nil {
		// Fallback to default text if custom page is missing
		fmt.Printf("[DEBUG] Could not read error page %s: %s\n", filename, err)
		http.Error(w, fmt.Sprintf("G3Pix AxonASP Error: %d", statusCode), statusCode)
		return
	}
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	w.WriteHeader(statusCode)
	w.Write(content)
}

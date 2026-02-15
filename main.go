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
//Use go install github.com/josephspurrier/goversioninfo/cmd/goversioninfo@latest
//Then run "go generate" in the project root to embed version info into the executable
//go:generate goversioninfo
package main

import (
	"fmt"
	"html"
	"net/http"
	_ "net/http/pprof"
	"os"
	"os/signal"
	"path/filepath"
	"runtime/debug"
	"strconv"
	"strings"
	"syscall"
	"time"
	_ "time/tzdata"

	"g3pix.com.br/axonasp/asp"
	"g3pix.com.br/axonasp/experimental"
	"g3pix.com.br/axonasp/server"

	"github.com/joho/godotenv"
)

// Configuration variables
var (
	Port              = "4050"
	RootDir           = "./www"
	DefaultTimezone   = "America/Sao_Paulo"
	DefaultPages      = []string{"index.asp", "default.asp", "index.html", "default.html", "default.htm", "index.htm", "index.txt"}
	ScriptTimeout     = 30 // in seconds
	DebugASP          = false
	CleanupSessions   = false
	ASTCacheType      = "disk" // "memory" or "disk"
	MemoryLimitMB     = 0      // 0 means no limit
	ASTCacheTTLMin    = 0      // 0 means keep forever
	UseVM             = false
	VMCacheType       = "disk" // "memory" or "disk"
	VMCacheTTLMin     = 0
	BlockedExtensions = ".asax,.ascx,.master,.skin,.browser,.sitemap,.config,.cs,.csproj,.vb,.vbproj,.webinfo,.licx,.resx,.resources,.mdb,.vjsproj,.java,.jsl,.ldb,.dsdgm,.ssdgm,.lsad,.ssmap,.cd,.dsprototype,.lsaprototype,.sdm,.sdmDocument,.mdf,.ldf,.ad,.dd,.ldd,.sd,.adprototype,.lddprototype,.exclude,.refresh,.compiled,.msgx,.vsdisco,.rules,.asa,.inc,.exe,.dll,.env,.config,.htaccess,.env.local,.json,.yaml,.yml"
	Error404Mode      = "default" // "default" or "IIS"
	COMProviderMode   = "auto"    // "auto" or "code"
	Version           = "0.0.0.0-dev"
)

// Global web.config parser
var webConfigParser *server.WebConfigParser

// executableDir is the directory containing the axonasp binary, resolved at startup.
// Used to locate bundled assets (e.g. errorpages/) when installed globally.
var executableDir string

func init() {
	// Resolve directory containing the executable (for locating bundled assets)
	if execPath, err := os.Executable(); err == nil {
		if resolved, err := filepath.EvalSymlinks(execPath); err == nil {
			executableDir = filepath.Dir(resolved)
		} else {
			executableDir = filepath.Dir(execPath)
		}
	}

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
		parts := strings.Split(val, ",")
		var pages []string
		for _, p := range parts {
			trimmed := strings.TrimSpace(p)
			if trimmed != "" {
				pages = append(pages, trimmed)
			}
		}
		if len(pages) > 0 {
			DefaultPages = pages
		}
	}
	if val := os.Getenv("SCRIPT_TIMEOUT"); val != "" {
		if i, err := strconv.Atoi(val); err == nil {
			ScriptTimeout = i
		}
	}
	if val := os.Getenv("DEBUG_ASP"); val == "TRUE" {
		DebugASP = true
	}
	if val := os.Getenv("CLEAN_SESSIONS"); val == "TRUE" {
		CleanupSessions = true
	}
	if val := os.Getenv("AXONASP_VM"); val == "TRUE" {
		UseVM = true
	}
	if val := os.Getenv("ASP_CACHE_TYPE"); val != "" {
		val = strings.ToLower(strings.TrimSpace(val))
		if val == "memory" || val == "disk" {
			ASTCacheType = val
		} else {
			fmt.Printf("Warning: Invalid ASP_CACHE_TYPE value '%s'. Using 'disk'. Valid values: 'memory', 'disk'\n", val)
			ASTCacheType = "disk"
		}
	}
	if val := os.Getenv("MEMORY_LIMIT_MB"); val != "" {
		if i, err := strconv.Atoi(val); err == nil && i >= 0 {
			MemoryLimitMB = i
		} else {
			fmt.Printf("Warning: Invalid MEMORY_LIMIT_MB value '%s'. Using 0 (no limit).\n", val)
			MemoryLimitMB = 0
		}
	}
	if val := os.Getenv("ASP_CACHE_TTL_MINUTES"); val != "" {
		if i, err := strconv.Atoi(val); err == nil && i >= 0 {
			ASTCacheTTLMin = i
		} else {
			fmt.Printf("Warning: Invalid ASP_CACHE_TTL_MINUTES value '%s'. Using 0 (keep forever).\n", val)
			ASTCacheTTLMin = 0
		}
	}
	if val := os.Getenv("VM_CACHE_TYPE"); val != "" {
		val = strings.ToLower(strings.TrimSpace(val))
		if val == "memory" || val == "disk" {
			VMCacheType = val
		} else {
			fmt.Printf("Warning: Invalid VM_CACHE_TYPE value '%s'. Using 'disk'. Valid values: 'memory', 'disk'\n", val)
			VMCacheType = "disk"
		}
	}
	if val := os.Getenv("VM_CACHE_TTL_MINUTES"); val != "" {
		if i, err := strconv.Atoi(val); err == nil && i >= 0 {
			VMCacheTTLMin = i
		} else {
			fmt.Printf("Warning: Invalid VM_CACHE_TTL_MINUTES value '%s'. Using 0 (keep forever).\n", val)
			VMCacheTTLMin = 0
		}
	}
	if val := os.Getenv("BLOCKED_EXTENSIONS"); val != "" {
		BlockedExtensions = val
	}
	if val := os.Getenv("ERROR_404_MODE"); val != "" {
		val = strings.ToLower(val)
		if val == "default" || val == "iis" {
			Error404Mode = val
		} else {
			fmt.Printf("Warning: Invalid ERROR_404_MODE value '%s'. Using 'default'. Valid values: 'default', 'IIS'\n", val)
			Error404Mode = "default"
		}
	}
	if val := os.Getenv("COM_PROVIDER"); val != "" {
		val = strings.ToLower(val)
		if val == "auto" || val == "code" {
			COMProviderMode = val
		} else {
			fmt.Printf("Warning: Invalid COM_PROVIDER value '%s'. Using 'auto'. Valid values: 'auto', 'code'\n", val)
			COMProviderMode = "auto"
		}
	}
	// Resolve WEB_ROOT to absolute path so the directory-traversal security
	// check works correctly regardless of how the server is started.
	if absRoot, err := filepath.Abs(RootDir); err == nil {
		RootDir = absRoot
	}

	server.SetEngineVersion(Version)
	server.SetCOMProviderMode(COMProviderMode)
	asp.ConfigureParseCache(ASTCacheType, RootDir)
	asp.SetParseCacheTTLMinutes(ASTCacheTTLMin)
	experimental.ConfigureBytecodeCache(VMCacheType, RootDir)
	experimental.SetBytecodeCacheTTLMinutes(VMCacheTTLMin)
	if MemoryLimitMB > 0 {
		debug.SetMemoryLimit(int64(MemoryLimitMB) * 1024 * 1024)
	}

	// Set timezone
	os.Setenv("TZ", DefaultTimezone)
	loc, err := time.LoadLocation(DefaultTimezone)
	if err != nil {
		fmt.Printf("Warning: Could not load timezone %s, using UTC: %v\n", DefaultTimezone, err)
		loc = time.UTC
	}
	time.Local = loc
}

func cleanupSessionFiles() {
	sessionDir := filepath.Join("temp", "session")
	entries, err := os.ReadDir(sessionDir)
	if err != nil {
		fmt.Printf("Warning: Failed to read session directory '%s': %v\n", sessionDir, err)
		return
	}

	removed := 0
	for _, entry := range entries {
		targetPath := filepath.Join(sessionDir, entry.Name())
		if entry.IsDir() {
			if err := os.RemoveAll(targetPath); err != nil {
				fmt.Printf("Warning: Failed to remove session folder '%s': %v\n", targetPath, err)
				continue
			}
			removed++
			continue
		}
		if err := os.Remove(targetPath); err != nil {
			fmt.Printf("Warning: Failed to remove session file '%s': %v\n", targetPath, err)
			continue
		}
		removed++
	}

	if removed > 0 {
		fmt.Printf("Info: Removed %d session item(s) from %s\n", removed, sessionDir)
	}
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
	fmt.Print("\033]0;G3Pix AxonASP Server\a")
	fmt.Printf("\033[48;5;240m\033[37mStarting G3pix AxonASP on http://localhost:%s ► \033[0m\n", Port)
	fmt.Printf("Serving files from %s\n", RootDir)
	fmt.Printf("Version: %s\n", Version)
	// Initialize web.config parser
	webConfigParser = server.NewWebConfigParser(RootDir)
	if err := webConfigParser.Load(); err != nil {
		webConfigParser = nil
		if strings.ToLower(Error404Mode) == "iis" {
			fmt.Printf("Warning: ERROR_404_MODE is 'IIS' but failed to load web.config: %v\n", err)
			fmt.Println("Falling back to default error page mode.")
			Error404Mode = "default"
		}
	}

	http.HandleFunc("/", handleRequest)
	setupShutdownHandlers()

	// Initialize session manager and start cleanup routine
	if CleanupSessions {
		cleanupSessionFiles()
	}
	sessionManager := server.GetSessionManager()
	sessionManager.StartCleanupRoutine(20 * time.Minute) // Cleanup every 20 minutes

	// Load Global.asa file
	globalASAManager := server.GetGlobalASAManager()
	err := globalASAManager.LoadGlobalASA(RootDir)
	if err != nil {
		fmt.Printf("Warning: Failed to load Global.asa: %v\n", err)
	}

	//fmt.Printf("Application_OnStart defined: %v\n", globalASAManager.HasApplicationOnStart())
	//fmt.Printf("Session_OnStart defined: %v\n", globalASAManager.HasSessionOnStart())

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
				UseVM:         UseVM,
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
			if DebugASP {
				if err := globalASAManager.ExecuteApplicationOnStart(executor, ctx); err != nil {
					fmt.Printf("[DEBUG] Error in Application_OnStart: %v\n", err)
				} else {
					fmt.Println("[DEBUG] Application_OnStart executed successfully")
				}
			}
		}()
	}

	if DebugASP {
		fmt.Println("[DEBUG] DEBUG_ASP mode is enabled")
		if UseVM {
			fmt.Print("[DEBUG] VM mode is enabled\n")
			fmt.Printf("[DEBUG] VM enabled: %v\n", UseVM)
			fmt.Printf("[DEBUG] VM cache type: %s\n", VMCacheType)
			fmt.Printf("[DEBUG] VM cache TTL minutes: %d\n", VMCacheTTLMin)
		} else {
			fmt.Print("[DEBUG] AST walker mode is enabled\n")
			fmt.Printf("[DEBUG] Cache type: %s\n", ASTCacheType)
			fmt.Printf("[DEBUG] Memory limit: %d MB\n", MemoryLimitMB)
			fmt.Printf("[DEBUG] Cache TTL minutes: %d\n", ASTCacheTTLMin)
		}
		//Display build info for debugging purposes
		if info, ok := debug.ReadBuildInfo(); ok {
			fmt.Printf("[DEBUG] Go Version: %s\n", info.GoVersion)
			for _, setting := range info.Settings {
				switch setting.Key {
				case "vcs.revision":
					fmt.Printf("[DEBUG] VCS Revision: %s\n", setting.Value)
				case "vcs.time":
					fmt.Printf("[DEBUG] Build Date: %s\n", setting.Value)
				case "GOARCH":
					fmt.Printf("[DEBUG] Architecture: %s\n", setting.Value)
				}
			}
		}
	}

	err = http.ListenAndServe(":"+Port, nil)
	if err != nil {
		fmt.Printf("Fatal error starting G3Pix AxonASP server:\n  %v\n", err)
		fmt.Print("Shutting down.\n")
	}
}

func setupShutdownHandlers() {
	if (ASTCacheType != "disk" || ASTCacheTTLMin <= 0) && (VMCacheType != "disk" || VMCacheTTLMin <= 0) {
		return
	}

	shutdownSignals := make(chan os.Signal, 1)
	signal.Notify(shutdownSignals, os.Interrupt, syscall.SIGTERM)

	go func() {
		<-shutdownSignals
		asp.CleanupParseCacheOnShutdown()
		experimental.CleanupBytecodeCacheOnShutdown()
		if CleanupSessions {
			cleanupSessionFiles()
		}
		os.Exit(0)
	}()
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
	if applyHTTPRedirect(w, r) {
		return
	}
	if applyRewriteRules(w, r) {
		return
	}
	path = r.URL.Path
	if path == "/" {
		// Try each default page
		found := false
		for _, page := range DefaultPages {
			fullPath := filepath.Join(RootDir, page)
			if _, err := os.Stat(fullPath); err == nil {
				path = "/" + page
				found = true
				break
			}
		}
		if !found {
			// If none found, fallback to the first one for the 404 to happen later
			if len(DefaultPages) > 0 {
				path = "/" + DefaultPages[0]
			} else {
				path = "/default.asp"
			}
		}
	}

	fullPath := filepath.Join(RootDir, path)

	// Security check: prevent directory traversal
	cleanPath := filepath.Clean(fullPath)
	cleanRoot := filepath.Clean(RootDir)
	if !strings.HasPrefix(cleanPath, cleanRoot) {
		serveErrorPage(w, r, 403)
		return
	}

	// Security check: block direct access to restricted file extensions, we use not found for safety
	fileExt := strings.ToLower(filepath.Ext(fullPath))
	if isBlockedExtension(fileExt) {
		serveErrorPage(w, r, 404)
		return
	}

	// Check if file exists
	info, err := os.Stat(fullPath)
	if os.IsNotExist(err) {
		serveErrorPage(w, r, 404)
		return
	}

	// If it's a directory, try to serve the default page
	if info.IsDir() {
		if !strings.HasSuffix(path, "/") {
			redirectPath := path + "/"
			if r.URL.RawQuery != "" {
				redirectPath += "?" + r.URL.RawQuery
			}
			http.Redirect(w, r, redirectPath, http.StatusMovedPermanently)
			return
		}

		found := false
		for _, page := range DefaultPages {
			checkPath := filepath.Join(fullPath, page)
			if _, err := os.Stat(checkPath); err == nil {
				fullPath = checkPath
				found = true
				break
			}
		}

		if !found {
			serveErrorPage(w, r, 404)
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
		serveErrorPage(w, r, 500)
		return
	}

	// Recover from panics to avoid crashing server
	defer func() {
		if rec := recover(); rec != nil {
			fmt.Printf("G3Pix AxonASP Runtime panic in %s: %v\n", path, rec)

			// Check if debug mode is enabled
			isDebug := os.Getenv("DEBUG_ASP") == "TRUE"

			if !isDebug {
				serveErrorPage(w, r, 500)
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
				fmt.Fprintf(w, "Error: %v<br>\n", rec)
				fmt.Fprintf(w, "<pre>%s</pre>\n", stack)
			} else {
				// Simple error output
				fmt.Fprintf(w, "<strong>G3pix AxonASP error</strong><br>\n")
				fmt.Fprintf(w, "Description: %v<br>\n", rec)
			}

			fmt.Fprintf(w, "</div>\n")
		}
	}()

	// Create ASP processor and execute
	processor := server.NewASPProcessor(&server.ASPProcessorConfig{
		RootDir:       RootDir,
		ScriptTimeout: ScriptTimeout,
		DebugASP:      DebugASP,
		UseVM:         UseVM,
	})

	err = processor.ExecuteASPFile(content, fullPath, w, r)
	if err != nil {
		fmt.Printf("[DEBUG] ASP processing error in %s: %v\n", path, err)
		if DebugASP {
			fmt.Fprintf(w, "<br><hr style='border-top: 1px dashed red;'>\n")
			fmt.Fprintf(w, "<div style='color: red; font-family: monospace; background: #ffe6e6; padding: 10px; border: 1px solid red;'>\n")
			fmt.Fprintf(w, "<strong>ASP Runtime Error</strong><br>\n")
			fmt.Fprintf(w, "File: %s<br>\n", html.EscapeString(path))
			fmt.Fprintf(w, "Error: %s<br>\n", html.EscapeString(err.Error()))
			fmt.Fprintf(w, "</div>\n")
		}
	}
}

func applyHTTPRedirect(w http.ResponseWriter, r *http.Request) bool {
	if webConfigParser == nil || !webConfigParser.IsLoaded() {
		return false
	}
	redirectConfig := webConfigParser.GetHTTPRedirectConfig()
	if redirectConfig == nil || !redirectConfig.Enabled || redirectConfig.Destination == "" {
		return false
	}
	if redirectConfig.ChildOnly && r.URL.Path == "/" {
		return false
	}
	location := buildHTTPRedirectLocation(redirectConfig, r)
	statusCode := redirectConfig.StatusCode
	if statusCode == 0 {
		statusCode = http.StatusFound
	}
	http.Redirect(w, r, location, statusCode)
	return true
}

func buildHTTPRedirectLocation(config *server.HTTPRedirectConfig, r *http.Request) string {
	location := config.Destination
	if config.ExactDestination {
		return location
	}
	path := r.URL.Path
	if !strings.HasSuffix(location, "/") && !strings.HasPrefix(path, "/") {
		location += "/"
	}
	location += strings.TrimPrefix(path, "/")
	if r.URL.RawQuery != "" {
		location += "?" + r.URL.RawQuery
	}
	return location
}

func applyRewriteRules(w http.ResponseWriter, r *http.Request) bool {
	if webConfigParser == nil || !webConfigParser.IsLoaded() {
		return false
	}
	result, ok := webConfigParser.ApplyRewriteRules(r.URL.Path, r.URL.RawQuery)
	if !ok || !result.Applied {
		return false
	}
	if result.ActionType == "redirect" {
		statusCode := result.RedirectStatus
		if statusCode == 0 {
			statusCode = http.StatusFound
		}
		http.Redirect(w, r, result.RedirectLocation, statusCode)
		return true
	}
	if result.Path != "" {
		r.URL.Path = result.Path
	}
	r.URL.RawQuery = result.RawQuery
	return false
}

// getErrorPagePath resolves the path for a static error page file.
// It checks three locations in order:
//  1. WEB_ROOT/errorpages/ — per-project custom error pages
//  2. <binary-dir>/errorpages/ — bundled defaults (global install)
//  3. errorpages/ relative to CWD — dev mode fallback
func getErrorPagePath(filename string) string {
	// 1. Check WEB_ROOT for per-project custom error pages
	projectPath := filepath.Join(RootDir, "errorpages", filename)
	if _, err := os.Stat(projectPath); err == nil {
		return projectPath
	}
	// 2. Check next to binary (global install)
	if executableDir != "" {
		bundledPath := filepath.Join(executableDir, "errorpages", filename)
		if _, err := os.Stat(bundledPath); err == nil {
			return bundledPath
		}
	}
	// 3. Fallback to CWD (dev mode)
	return filepath.Join("errorpages", filename)
}

// serveErrorPage serves a custom HTML error page from the errorpages directory
// or executes custom error handlers based on ERROR_404_MODE configuration
func serveErrorPage(w http.ResponseWriter, r *http.Request, statusCode int) {
	// Special handling for 404 errors with IIS mode
	if statusCode == 404 && strings.ToLower(Error404Mode) == "iis" && webConfigParser != nil && webConfigParser.IsLoaded() {
		handleCustom404(w, r)
		return
	}

	// Default behavior: serve static error page
	filename := fmt.Sprintf("%d.html", statusCode)
	filePath := getErrorPagePath(filename)
	content, err := os.ReadFile(filePath)
	if err != nil {
		// Fallback to default text if custom page is missing
		fmt.Printf("[DEBUG] Could not find error page %s: %s\n", filename, err)
		http.Error(w, fmt.Sprintf("G3Pix AxonASP Error: %d", statusCode), statusCode)
		return
	}
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	w.WriteHeader(statusCode)
	w.Write(content)
}

// handleCustom404 processes custom 404 error handling based on web.config
func handleCustom404(w http.ResponseWriter, r *http.Request) {
	fullPath, responseMode := webConfigParser.GetFullErrorHandlerPath(404)

	if fullPath == "" {
		// No custom handler configured, fall back to default
		fmt.Println("[DEBUG] No 404 handler found in web.config, using default")
		filePath := getErrorPagePath("404.html")
		content, err := os.ReadFile(filePath)
		if err != nil {
			http.Error(w, "G3Pix AxonASP Error: 404", http.StatusNotFound)
			return
		}
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		w.WriteHeader(http.StatusNotFound)
		w.Write(content)
		return
	}

	// Check if the error handler file exists
	if _, err := os.Stat(fullPath); os.IsNotExist(err) {
		fmt.Printf("[DEBUG] Custom 404 handler file not found: %s\n", fullPath)
		// Fall back to default error page
		filePath := getErrorPagePath("404.html")
		content, err := os.ReadFile(filePath)
		if err != nil {
			http.Error(w, "G3Pix AxonASP Error: 404", http.StatusNotFound)
			return
		}
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		w.WriteHeader(http.StatusNotFound)
		w.Write(content)
		return
	}

	// Handle based on responseMode
	if strings.ToLower(responseMode) == "executeurl" {
		// Execute the ASP file for 404 handling
		if strings.HasSuffix(strings.ToLower(fullPath), ".asp") {
			executeASPErrorHandler(w, r, fullPath)
		} else {
			// If not an ASP file, serve it as static content
			http.ServeFile(w, r, fullPath)
		}
	} else {
		// Default: serve the file as static content
		http.ServeFile(w, r, fullPath)
	}
}

// executeASPErrorHandler executes an ASP file as a 404 error handler
func executeASPErrorHandler(w http.ResponseWriter, r *http.Request, aspFilePath string) {
	content, err := asp.ReadFileText(aspFilePath)
	if err != nil {
		fmt.Printf("[DEBUG] Failed to read ASP error handler: %v\n", err)
		http.Error(w, "G3Pix AxonASP Error: 404", http.StatusNotFound)
		return
	}

	// Set 404 status before processing ASP
	w.WriteHeader(http.StatusNotFound)

	// Recover from panics to avoid crashing server
	defer func() {
		if r := recover(); r != nil {
			fmt.Printf("G3Pix AxonASP Runtime panic in error handler %s: %v\n", aspFilePath, r)
			// Error already written, just log it
		}
	}()

	// Create ASP processor and execute
	processor := server.NewASPProcessor(&server.ASPProcessorConfig{
		RootDir:       RootDir,
		ScriptTimeout: ScriptTimeout,
		DebugASP:      DebugASP,
		UseVM:         UseVM,
	})

	err = processor.ExecuteASPFile(content, aspFilePath, w, r)
	if err != nil {
		fmt.Printf("[DEBUG] ASP error handler processing error in %s: %v\n", aspFilePath, err)
		if DebugASP {
			fmt.Fprintf(w, "<br><hr style='border-top: 1px dashed red;'>\n")
			fmt.Fprintf(w, "<div style='color: red; font-family: monospace; background: #ffe6e6; padding: 10px; border: 1px solid red;'>\n")
			fmt.Fprintf(w, "<strong>ASP Runtime Error</strong><br>\n")
			fmt.Fprintf(w, "File: %s<br>\n", html.EscapeString(aspFilePath))
			fmt.Fprintf(w, "Error: %s<br>\n", html.EscapeString(err.Error()))
			fmt.Fprintf(w, "</div>\n")
		}
	}
}

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃƒÂ£es - G3pix Ltda
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
//go:generate goversioninfo -icon=icon_fcgi.ico -64=true
package main

import (
	"fmt"
	"html"
	"io"
	"log"
	"net"
	"net/http"
	"net/http/fcgi"
	"os"
	"os/signal"
	"path/filepath"
	"runtime"
	"runtime/debug"
	"slices"
	"strconv"
	"strings"
	"sync"
	"syscall"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
	"github.com/fsnotify/fsnotify"
	"github.com/joho/godotenv"
)

// FastCGI configuration values.
var (
	Version                  = "0.0.0.0"
	ListenNetwork            = "tcp"
	ListenAddr               = "127.0.0.1:9000"
	RootDir                  = "./www"
	DefaultPages             = []string{"default.asp", "default.htm", "index.asp", "index.html", "default.html"}
	ExecuteAsASPExtension    = []string{".asp"}
	DefaultErrorPagesDir     = "./www/error-pages"
	ScriptTimeout            = 60
	ResponseBufferLimitBytes = 4 * 1024 * 1024
	DebugASP                 = false
	CleanupSessions          = true
	CleanupCache             = true
	DefaultTimezone          = "UTC"
	MemoryLimitMB            = 128
	VMPoolSize               = 50
	BytecodeCachingMode      = "enabled"
	CacheMaxSizeMB           = 128
	SessionAutoFlushSeconds  = 15
	serverLocation           = time.UTC
	scriptCache              *axonvm.ScriptCache
)

// init loads environment variables and applies TOML-based configuration through Viper.
func init() {
	_ = godotenv.Load()
	axonvm.SetRuntimeVersion(strings.TrimSpace(Version))
	loadFastCGIConfig()
	applyRuntimeSettings()
}

// loadFastCGIConfig loads and applies fastcgi/global settings from config/axonasp.toml using Viper.
func loadFastCGIConfig() {
	v := axonconfig.NewViper()
	if strings.TrimSpace(v.ConfigFileUsed()) == "" {
		log.Printf("Warning: Failed to read configuration file, using defaults.\n")
	}
	axonconfig.EnableWatchIfConfigured(v, func(e fsnotify.Event) {
		fmt.Println("Config file changed:", e.Name)
		DebugASP = v.GetBool("global.enable_asp_debugging")
		axonvm.SetInternalErrorLogEnabled(v.GetBool("global.enable_error_log_file"))
	})
	if workingDir, err := os.Getwd(); err == nil {
		axonvm.SetInternalErrorLogRootPath(workingDir)
	}
	DebugASP = v.GetBool("global.enable_asp_debugging")
	axonvm.SetInternalErrorLogEnabled(v.GetBool("global.enable_error_log_file"))
	axonvm.SetDumpPreprocessedSourceEnabled(v.GetBool("global.dump_preprocessed_source"))

	if pages := v.GetStringSlice("fastcgi.default_pages"); len(pages) > 0 {
		DefaultPages = pages
	}
	if executeAsASP := v.GetStringSlice("global.execute_as_asp"); len(executeAsASP) > 0 {
		ExecuteAsASPExtension = normalizeExtensions(executeAsASP)
	}
	if errorPagesDir := strings.TrimSpace(v.GetString("server.default_error_pages_directory")); errorPagesDir != "" {
		DefaultErrorPagesDir = errorPagesDir
	}
	ScriptTimeout = v.GetInt("global.default_script_timeout")
	if ScriptTimeout <= 0 {
		ScriptTimeout = 60
	}
	if responseBufferLimitMB := v.GetInt("global.response_buffer_limit_mb"); responseBufferLimitMB > 0 {
		ResponseBufferLimitBytes = responseBufferLimitMB * 1024 * 1024
	}

	rawListenEndpoint := strings.TrimSpace(v.GetString("fastcgi.server_port"))
	if rawListenEndpoint != "" {
		network, address, err := parseFastCGIListenEndpoint(rawListenEndpoint)
		if err != nil {
			log.Printf("Warning: Invalid fastcgi.server_port value %q, keeping default %s://%s: %v\n", rawListenEndpoint, ListenNetwork, ListenAddr, err)
		} else {
			ListenNetwork = network
			ListenAddr = address
		}
	}

	CleanupSessions = v.GetBool("global.clean_sessions_on_startup")
	CleanupCache = v.GetBool("global.clean_cache_on_startup")
	if timezone := strings.TrimSpace(v.GetString("global.default_timezone")); timezone != "" {
		DefaultTimezone = timezone
	}
	if memoryMB := v.GetInt("global.golang_memory_limit_mb"); memoryMB >= 0 {
		MemoryLimitMB = memoryMB
	}
	if vmPoolSize := v.GetInt("global.vm_pool_size"); vmPoolSize >= 0 {
		VMPoolSize = vmPoolSize
	}
	BytecodeCachingMode = strings.TrimSpace(v.GetString("global.bytecode_caching_enabled"))
	if BytecodeCachingMode == "" {
		BytecodeCachingMode = strings.TrimSpace(v.GetString("global.bytecode_caching"))
	}
	if BytecodeCachingMode == "" {
		BytecodeCachingMode = "enabled"
	}
	if cacheSizeMB := v.GetInt("global.cache_max_size_mb"); cacheSizeMB > 0 {
		CacheMaxSizeMB = cacheSizeMB
	}
	if flushSeconds := v.GetInt("global.session_flush_interval_seconds"); flushSeconds > 0 {
		SessionAutoFlushSeconds = flushSeconds
	}
	axonvm.SetVMPoolSizeLimit(VMPoolSize)

	axonvm.InitGlobalAxonFunctions(v.GetBool("axfunctions.enable_global_ax"))
}

// parseFastCGIListenEndpoint parses FastCGI listener settings as TCP port/host:port or unix:/path socket endpoint.
func parseFastCGIListenEndpoint(raw string) (string, string, error) {
	value := strings.TrimSpace(raw)
	if value == "" {
		return "", "", fmt.Errorf("listen endpoint cannot be empty")
	}

	lower := strings.ToLower(value)
	if strings.HasPrefix(lower, "unix:") {
		socketPath := strings.TrimSpace(value[len("unix:"):])
		if socketPath == "" {
			return "", "", fmt.Errorf("unix socket path cannot be empty")
		}
		return "unix", socketPath, nil
	}

	if port, err := strconv.Atoi(value); err == nil {
		if port <= 0 || port > 65535 {
			return "", "", fmt.Errorf("tcp port must be between 1 and 65535")
		}
		return "tcp", "127.0.0.1:" + strconv.Itoa(port), nil
	}

	if strings.HasPrefix(value, ":") {
		return "tcp", "127.0.0.1" + value, nil
	}

	if _, _, err := net.SplitHostPort(value); err == nil {
		return "tcp", value, nil
	}

	return "", "", fmt.Errorf("invalid FastCGI listen endpoint")
}

// prepareFastCGIListener creates the configured FastCGI listener and prepares unix socket paths when needed.
func prepareFastCGIListener(network, address string) (net.Listener, error) {
	if network != "unix" {
		return net.Listen(network, address)
	}

	if runtime.GOOS == "windows" {
		return nil, fmt.Errorf("unix socket mode is not supported on Windows")
	}

	if err := ensureUnixSocketPathReady(address); err != nil {
		return nil, err
	}

	listener, err := net.Listen("unix", address)
	if err != nil {
		return nil, err
	}

	if chmodErr := os.Chmod(address, 0o660); chmodErr != nil {
		log.Printf("Warning: Failed to apply permissions to unix socket %s: %v\n", address, chmodErr)
	}

	return listener, nil
}

// ensureUnixSocketPathReady creates parent directories and removes stale unix socket files before binding.
func ensureUnixSocketPathReady(path string) error {
	dir := filepath.Dir(path)
	if dir != "" && dir != "." {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return err
		}
	}

	info, err := os.Lstat(path)
	if err != nil {
		if os.IsNotExist(err) {
			return nil
		}
		return err
	}

	if info.Mode()&os.ModeSocket == 0 {
		return fmt.Errorf("path exists and is not a unix socket: %s", path)
	}

	return os.Remove(path)
}

// cleanupFastCGIListenerArtifact removes unix socket files on graceful shutdown.
func cleanupFastCGIListenerArtifact(network, address string) {
	if network != "unix" {
		return
	}
	_ = os.Remove(address)
}

// applyRuntimeSettings applies timezone and Go memory limit based on loaded configuration.
func applyRuntimeSettings() {
	if MemoryLimitMB > 0 {
		debug.SetMemoryLimit(int64(MemoryLimitMB) * 1024 * 1024)
	}

	os.Setenv("TZ", DefaultTimezone)
	location, err := axonvm.ResolveTimezoneLocation(DefaultTimezone)
	if err != nil {
		fmt.Printf("Warning: Could not load timezone %s, using UTC: %v\n", DefaultTimezone, err)
		location = time.UTC
	}
	serverLocation = location
	time.Local = location
}

// normalizeExtensions normalizes file extensions to lowercase ".ext" values.
func normalizeExtensions(values []string) []string {
	cleaned := make([]string, 0, len(values))
	for _, value := range values {
		ext := strings.ToLower(strings.TrimSpace(value))
		if ext == "" {
			continue
		}
		if !strings.HasPrefix(ext, ".") {
			ext = "." + ext
		}
		if !slices.Contains(cleaned, ext) {
			cleaned = append(cleaned, ext)
		}
	}
	return cleaned
}

// cleanupSessionFiles removes all files and folders from temp/session.
func cleanupSessionFiles() {
	sessionDir := filepath.Join("temp", "session")
	entries, err := os.ReadDir(sessionDir)
	if err != nil {
		return
	}

	for _, entry := range entries {
		targetPath := filepath.Join(sessionDir, entry.Name())
		if entry.IsDir() {
			_ = os.RemoveAll(targetPath)
			continue
		}
		_ = os.Remove(targetPath)
	}
}

// cleanupCacheFiles removes all files and folders from temp/cache.
func cleanupCacheFiles() {
	cacheDir := filepath.Join("temp", "cache")
	entries, err := os.ReadDir(cacheDir)
	if err != nil {
		return
	}

	for _, entry := range entries {
		targetPath := filepath.Join(cacheDir, entry.Name())
		if entry.IsDir() {
			_ = os.RemoveAll(targetPath)
			continue
		}
		_ = os.Remove(targetPath)
	}
}

// dummyResponseWriter provides a no-op http.ResponseWriter for global.asa events.
type dummyResponseWriter struct{}

func (d *dummyResponseWriter) Header() http.Header         { return make(http.Header) }
func (d *dummyResponseWriter) Write(b []byte) (int, error) { return len(b), nil }
func (d *dummyResponseWriter) WriteHeader(statusCode int)  {}

// main starts the FastCGI listener and serves ASP requests.
func main() {
	if CleanupSessions {
		cleanupSessionFiles()
	}
	if CleanupCache {
		cleanupCacheFiles()
	}

	if _, err := os.Stat(RootDir); os.IsNotExist(err) {
		axonvm.ReportInternalError(axonvm.ErrRootDirectoryDoesNotExist, err, "Creating missing root directory.", RootDir, 0)
		if mkdirErr := os.MkdirAll(RootDir, 0o755); mkdirErr != nil {
			axonvm.ReportInternalError(axonvm.ErrRootDirInvalid, mkdirErr, "Failed to create the configured root directory.", RootDir, 0)
			os.Exit(1)
		}
	}

	cacheRoot, cacheRootErr := filepath.Abs(RootDir)
	if cacheRootErr != nil {
		cacheRoot = RootDir
	}
	scriptCache = axonvm.NewScriptCache(
		axonvm.ParseBytecodeCacheMode(BytecodeCachingMode),
		filepath.Join("temp", "cache"),
		CacheMaxSizeMB,
	)
	if err := scriptCache.StartInvalidator([]string{cacheRoot}); err != nil {
		log.Printf("Warning: Failed to start bytecode invalidator: %v\n", err)
	}
	defer scriptCache.StopInvalidator()

	asp.StartSessionAutoFlush(time.Duration(SessionAutoFlushSeconds) * time.Second)
	defer asp.StopSessionAutoFlush()

	// Load and compile global.asa
	if err := axonvm.GetGlobalASA().LoadAndCompile(RootDir, GetSharedApplication()); err != nil {
		fmt.Printf("Warning: Failed to load global.asa: %v\n", err)
	} else if axonvm.GetGlobalASA().IsLoaded() {
		// Execute Application_OnStart using a dummy host
		req, _ := http.NewRequest("GET", "http://localhost/", nil)
		dummyHost := NewFastCGIHost(&dummyResponseWriter{}, req)
		_ = axonvm.GetGlobalASA().ExecuteApplicationOnStart(dummyHost)
	}

	mux := http.NewServeMux()
	mux.HandleFunc("/", fastCGIMiddleware(handleRequest))

	listener, err := prepareFastCGIListener(ListenNetwork, ListenAddr)
	if err != nil {
		axonvm.ReportInternalError(axonvm.ErrCouldNotListenOn, err, "FastCGI listener could not start.", ListenNetwork+"://"+ListenAddr, 0)
		os.Exit(1)
	}
	defer listener.Close()
	defer cleanupFastCGIListenerArtifact(ListenNetwork, ListenAddr)

	fmt.Printf("AxonASP FastCGI %s started on %s://%s\n", Version, ListenNetwork, ListenAddr)
	fmt.Printf("Serving files from: %s\n", RootDir)

	stop := make(chan os.Signal, 1)
	signal.Notify(stop, os.Interrupt, syscall.SIGTERM)

	go func() {
		if err := fcgi.Serve(listener, mux); err != nil {
			axonvm.ReportInternalError(axonvm.ErrFastCGIProtocolError, err, "FastCGI server returned an execution error.", ListenNetwork+"://"+ListenAddr, 0)
			os.Exit(1)
		}
	}()

	<-stop
	fmt.Println("\nShutting down FastCGI server...")

	if axonvm.GetGlobalASA().IsLoaded() {
		req, _ := http.NewRequest("GET", "http://localhost/", nil)
		dummyHost := NewFastCGIHost(&dummyResponseWriter{}, req)
		_ = axonvm.GetGlobalASA().ExecuteApplicationOnEnd(dummyHost)
	}
}

// fastCGIMiddleware wraps handleRequest to properly extract and pass FastCGI parameters.
// The fcgi package stores CGI environment variables in request headers, and this
// middleware also ensures FastCGI responses advertise AxonASP via X-Powered-By.
func fastCGIMiddleware(next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("X-Powered-By", "AxonASP")
		next(w, r)
	}
}

// getFastCGIParam retrieves a FastCGI CGI environment parameter from the request.
// fcgi.Serve stores non-HTTP CGI variables (DOCUMENT_ROOT, SCRIPT_FILENAME, etc.)
// in the request context via fcgi.ProcessEnv. Standard vars consumed by
// cgi.RequestFromMap (SCRIPT_NAME → r.URL.Path, HTTP_* → r.Header, etc.) are
// NOT stored there and must be read from the appropriate http.Request fields.
func getFastCGIParam(r *http.Request, paramName string) string {
	if r == nil {
		return ""
	}
	env := fcgi.ProcessEnv(r)
	return env[paramName]
}

// extractFastCGIParams returns all non-HTTP CGI environment variables stored in
// the request context by fcgi.Serve (via fcgi.ProcessEnv). HTTP_* browser headers
// and standard request fields (SCRIPT_NAME, REQUEST_METHOD, etc.) are not included
// — those are already accessible via r.Header and r.URL/r.Method respectively.
func extractFastCGIParams(r *http.Request) map[string]string {
	return fcgi.ProcessEnv(r)
}

// handleRequest resolves the target path and serves static or ASP content.
// This function supports FastCGI operation with multiple document roots when
// DOCUMENT_ROOT is provided by the reverse proxy (nginx/Apache). It resolves
// the absolute file path using:
//  1. DOCUMENT_ROOT from FastCGI params (if available) - base directory
//  2. SCRIPT_NAME from FastCGI params or URL.Path - relative script path
//  3. Falls back to RootDir and URL.Path if FastCGI params are absent
//
// The resolved path is validated to prevent directory traversal attacks.
func handleRequest(w http.ResponseWriter, r *http.Request) {
	// DEBUG: Log all headers to understand FastCGI parameter passing format
	fcgiParams := extractFastCGIParams(r)
	if DebugASP {
		log.Printf("[FastCGI] Request: %s %s\n", r.Method, r.URL.Path)
		log.Printf("[FastCGI] All headers/params:\n")
		for k, v := range fcgiParams {
			// Truncate long values for logging
			val := v
			if len(val) > 100 {
				val = val[:100] + "..."
			}
			log.Printf("  %s = %s\n", k, val)
		}
	}

	// Extract FastCGI parameters from reverse proxy (nginx/Apache).
	// DOCUMENT_ROOT is provided by nginx and stored in the request context by fcgi.Serve.
	// SCRIPT_NAME is consumed by cgi.RequestFromMap to build r.URL.Path, so it is
	// already in r.URL.Path and excluded from fcgi.ProcessEnv.
	documentRoot := strings.TrimSpace(getFastCGIParam(r, "DOCUMENT_ROOT"))

	// Determine the effective document root: use DOCUMENT_ROOT when provided,
	// otherwise fall back to the configured RootDir (backward compatible).
	effectiveRoot := RootDir
	if documentRoot != "" {
		effectiveRoot = documentRoot
	}

	// Script path comes from r.URL.Path, which cgi.RequestFromMap sets from SCRIPT_NAME
	// (or REQUEST_URI) as sent by the proxy. No extra extraction needed.
	scriptPath := r.URL.Path
	if scriptPath == "" {
		scriptPath = "/"
	}

	// Construct the full file path
	relativePath := strings.TrimPrefix(scriptPath, "/")
	fullPath := filepath.Join(effectiveRoot, filepath.FromSlash(relativePath))
	cleanPath := filepath.Clean(fullPath)

	// Resolve root directory to absolute path for security validation
	absRoot, err := filepath.Abs(effectiveRoot)
	if err != nil {
		respondInternalHTTPError(w, axonvm.ErrCouldNotResolveCurrentDir, err, "Failed to resolve the effective root directory.", effectiveRoot)
		return
	}

	// Resolve requested path to absolute path for security validation
	absPath, err := filepath.Abs(cleanPath)
	if err != nil {
		respondInternalHTTPError(w, axonvm.ErrBadFileName, err, "Failed to resolve the requested path.", cleanPath)
		return
	}

	// Directory traversal prevention: ensure the resolved path is within the document root
	if !strings.HasPrefix(absPath, absRoot) {
		serveErrorPage(w, r, http.StatusForbidden)
		return
	}

	info, err := os.Stat(fullPath)
	if os.IsNotExist(err) {
		serveErrorPage(w, r, http.StatusNotFound)
		return
	}
	if err != nil {
		respondInternalHTTPError(w, axonvm.ErrCouldNotReadFile, err, "Failed to inspect the requested path.", fullPath)
		return
	}
	if info.IsDir() {
		if !strings.HasSuffix(scriptPath, "/") {
			redirectPath := scriptPath + "/"
			if r.URL.RawQuery != "" {
				redirectPath += "?" + r.URL.RawQuery
			}
			http.Redirect(w, r, redirectPath, http.StatusMovedPermanently)
			return
		}

		foundDefault := false
		for _, page := range DefaultPages {
			candidate := filepath.Join(fullPath, page)
			candidateInfo, candidateErr := os.Stat(candidate)
			if candidateErr == nil && !candidateInfo.IsDir() {
				fullPath = candidate
				foundDefault = true
				break
			}
		}

		if !foundDefault {
			serveErrorPage(w, r, http.StatusNotFound)
			return
		}
	}

	if !isASPExecutionExtension(fullPath) {
		http.ServeFile(w, r, fullPath)
		return
	}

	executeASP(w, r, fullPath)
}

// isASPExecutionExtension reports whether a file should be executed as ASP based on configured extensions.
func isASPExecutionExtension(filePath string) bool {
	ext := strings.ToLower(filepath.Ext(filePath))
	return slices.Contains(ExecuteAsASPExtension, ext)
}

// singleHeaderResponseWriter prevents duplicate WriteHeader calls and can apply a default status code.
type singleHeaderResponseWriter struct {
	http.ResponseWriter
	wroteHeader   bool
	defaultStatus int
}

// newSingleHeaderResponseWriter wraps a response writer with duplicate WriteHeader protection.
func newSingleHeaderResponseWriter(w http.ResponseWriter, defaultStatus int) *singleHeaderResponseWriter {
	return &singleHeaderResponseWriter{ResponseWriter: w, defaultStatus: defaultStatus}
}

// WriteHeader writes the status code only once.
func (w *singleHeaderResponseWriter) WriteHeader(statusCode int) {
	if w.wroteHeader {
		return
	}
	w.wroteHeader = true
	w.ResponseWriter.WriteHeader(statusCode)
}

// Write applies the default status if needed and writes the response body.
func (w *singleHeaderResponseWriter) Write(data []byte) (int, error) {
	if !w.wroteHeader {
		if w.defaultStatus > 0 {
			w.WriteHeader(w.defaultStatus)
		} else {
			w.wroteHeader = true
		}
	}
	return w.ResponseWriter.Write(data)
}

// Flush forwards flush operations to the underlying writer.
func (w *singleHeaderResponseWriter) Flush() {
	if flusher, ok := w.ResponseWriter.(http.Flusher); ok {
		flusher.Flush()
	}
}

// ReadFrom forwards optimized copy operations to the underlying writer when available.
func (w *singleHeaderResponseWriter) ReadFrom(reader io.Reader) (int64, error) {
	if !w.wroteHeader {
		if w.defaultStatus > 0 {
			w.WriteHeader(w.defaultStatus)
		} else {
			w.wroteHeader = true
		}
	}
	if readFrom, ok := w.ResponseWriter.(io.ReaderFrom); ok {
		return readFrom.ReadFrom(reader)
	}
	return io.Copy(w.ResponseWriter, reader)
}

// executeASP compiles and executes an ASP file using the VM and FastCGI host.
func executeASP(w http.ResponseWriter, r *http.Request, filePath string) {
	executeASPWithStatus(w, r, filePath, 0)
}

// resolveRequestScriptTimeout returns the effective timeout in seconds for one request,
// preferring the current ASP Server.ScriptTimeout value over the bootstrap fallback.
func resolveRequestScriptTimeout(host axonvm.ASPHostEnvironment, fallback int) int {
	if host != nil {
		if server := host.Server(); server != nil {
			timeout := server.GetScriptTimeout()
			if timeout > 0 {
				return timeout
			}
		}
	}
	if fallback > 0 {
		return fallback
	}
	return 60
}

// executeASPWithStatus compiles and executes an ASP file, running vm.Run() in a
// goroutine so that a blocking CGO/COM call cannot hold the FastCGI handler
// indefinitely. On timeout a 503 is returned and the goroutine is detached.
func executeASPWithStatus(w http.ResponseWriter, r *http.Request, filePath string, defaultStatus int) {
	single := newSingleHeaderResponseWriter(w, defaultStatus)
	cw := newCancellableWriter(single)
	host := NewFastCGIHost(cw, r)

	if scriptCache == nil {
		scriptCache = axonvm.NewScriptCache(axonvm.BytecodeCacheDisabled, filepath.Join("temp", "cache"), 1)
	}
	program, err := scriptCache.LoadOrCompile(filePath)
	if err != nil {
		aspErr := axonvm.CompilerErrorToASPError(err, filePath)
		host.Server().SetLastError(aspErr)
		axonvm.LogASPProcessedError(aspErr, "fastcgi.executeASP.compile")
		if !DebugASP {
			if isErrorPageHandlerPath(filePath) {
				http.Error(w, http.StatusText(http.StatusInternalServerError), http.StatusInternalServerError)
				return
			}
			serveErrorPage(w, r, http.StatusInternalServerError)
			return
		}
		renderClassicASPDebugError(w, http.StatusInternalServerError, "Compilation Error", aspErr)
		return
	}

	vm := axonvm.AcquireVMFromCachedProgram(program)
	vm.SetHost(host)

	timeoutSec := resolveRequestScriptTimeout(host, ScriptTimeout)

	type vmResult struct{ err error }
	done := make(chan vmResult, 1)
	go func() {
		defer vm.Release()
		done <- vmResult{vm.Run()}
	}()

	start := time.Now()
	watchdog := time.NewTicker(250 * time.Millisecond)
	defer watchdog.Stop()

	for {
		select {
		case res := <-done:
			if res.err != nil {
				aspErr := axonvm.RuntimeErrorToASPError(res.err, filePath)
				host.Server().SetLastError(aspErr)
				axonvm.LogASPProcessedError(aspErr, "fastcgi.executeASP.runtime")
				if !DebugASP {
					if isErrorPageHandlerPath(filePath) {
						http.Error(w, http.StatusText(http.StatusInternalServerError), http.StatusInternalServerError)
						return
					}
					serveErrorPage(w, r, http.StatusInternalServerError)
					return
				}
				renderClassicASPDebugError(w, http.StatusInternalServerError, "Runtime Error", aspErr)
				return
			}
			host.PersistSession()
			host.Response().Flush()
			return

		case <-watchdog.C:
			effectiveTimeout := resolveRequestScriptTimeout(host, timeoutSec)
			if time.Since(start) >= time.Duration(effectiveTimeout)*time.Second {
				cw.cancel()
				log.Printf("[axonasp] script timeout (%ds): %s - goroutine detached\n", effectiveTimeout, filePath)
				http.Error(w, "Script execution timed out", http.StatusServiceUnavailable)
				return
			}
		}
	}
}

// cancellableWriter wraps an http.ResponseWriter and silently discards all writes
// after cancel() is called, protecting the real writer from concurrent access after
// a script timeout detaches the VM goroutine from the FastCGI handler.
type cancellableWriter struct {
	mu       sync.Mutex
	inner    http.ResponseWriter
	canceled bool
}

func newCancellableWriter(w http.ResponseWriter) *cancellableWriter {
	return &cancellableWriter{inner: w}
}

func (c *cancellableWriter) Header() http.Header {
	c.mu.Lock()
	defer c.mu.Unlock()
	if c.canceled {
		return make(http.Header)
	}
	return c.inner.Header()
}

func (c *cancellableWriter) Write(p []byte) (int, error) {
	c.mu.Lock()
	defer c.mu.Unlock()
	if c.canceled {
		return len(p), nil
	}
	return c.inner.Write(p)
}

func (c *cancellableWriter) WriteHeader(status int) {
	c.mu.Lock()
	defer c.mu.Unlock()
	if c.canceled {
		return
	}
	c.inner.WriteHeader(status)
}

func (c *cancellableWriter) cancel() {
	c.mu.Lock()
	c.canceled = true
	c.mu.Unlock()
}

// isErrorPageHandlerPath reports whether the current execution target is an error page handler itself.
func isErrorPageHandlerPath(filePath string) bool {
	errorDirAbs, err := filepath.Abs(DefaultErrorPagesDir)
	if err != nil {
		return false
	}

	fileAbs, err := filepath.Abs(filePath)
	if err != nil {
		return false
	}

	if filepath.Dir(fileAbs) != errorDirAbs {
		return false
	}

	name := strings.ToLower(filepath.Base(fileAbs))
	return strings.HasSuffix(name, ".asp") || strings.HasSuffix(name, ".html")
}

// renderClassicASPDebugError renders an ASP/VBScript-style debug page while preserving HTTP status.
func renderClassicASPDebugError(w http.ResponseWriter, statusCode int, stage string, err *asp.ASPError) {
	if err == nil {
		http.Error(w, http.StatusText(statusCode), statusCode)
		return
	}

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	w.WriteHeader(statusCode)

	source := html.EscapeString(strings.TrimSpace(err.Source))
	if source == "" {
		source = "Microsoft VBScript runtime"
	}
	description := html.EscapeString(strings.TrimSpace(err.Description))
	if description == "" {
		description = "Unknown runtime error"
	}
	fileName := html.EscapeString(strings.TrimSpace(err.File))
	if fileName == "" {
		fileName = "unknown"
	}

	fmt.Fprintf(w, "<html><head><title>500 - Internal Server Error</title></head><body>")
	fmt.Fprintf(w, "<h2>Server Error in '/' Application.</h2>")
	fmt.Fprintf(w, "<hr><p><b>%s error '%08X'</b></p>", source, uint32(int32(err.Number)))
	fmt.Fprintf(w, "<p>%s</p>", description)
	fmt.Fprintf(w, "<p><b>File:</b> %s<br><b>Line:</b> %d<br><b>Column:</b> %d</p>", fileName, err.Line, err.Column)
	fmt.Fprintf(w, "<p><b>Stage:</b> %s</p>", html.EscapeString(stage))
	fmt.Fprintf(w, "<hr><p><i>AxonASP Debug Mode (global.enable_asp_debugging=true)</i></p></body></html>")
}

// serveErrorPage serves configured error pages using .asp or .html handlers for a given HTTP status code.
func serveErrorPage(w http.ResponseWriter, r *http.Request, statusCode int) {
	aspPagePath := filepath.Join(DefaultErrorPagesDir, fmt.Sprintf("%d.asp", statusCode))
	if pageInfo, err := os.Stat(aspPagePath); err == nil && !pageInfo.IsDir() {
		executeASPWithStatus(w, r, aspPagePath, statusCode)
		return
	}

	htmlPagePath := filepath.Join(DefaultErrorPagesDir, fmt.Sprintf("%d.html", statusCode))
	if pageInfo, err := os.Stat(htmlPagePath); err == nil && !pageInfo.IsDir() {
		http.ServeFile(newSingleHeaderResponseWriter(w, statusCode), r, htmlPagePath)
		return
	}

	http.Error(w, http.StatusText(statusCode), statusCode)
}

// respondInternalHTTPError logs an internal AxonASP error and writes a compact HTTP response.
func respondInternalHTTPError(w http.ResponseWriter, code axonvm.AxonASPErrorCode, err error, description string, fileName string) {
	axonvm.ReportInternalError(code, err, description, fileName, 0)
	http.Error(w, "AxonASP Error ["+strconv.Itoa(int(code))+"] "+code.String(), httpStatusFromAxonCode(code))
}

// httpStatusFromAxonCode maps AxonASP HTTP codes to HTTP status values.
func httpStatusFromAxonCode(code axonvm.AxonASPErrorCode) int {
	if code >= 400 && code <= 599 {
		return int(code)
	}

	return int(axonvm.HTTPInternalServerError)
}

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
//go:generate goversioninfo -icon=icon_server.ico -64=true
package main

import (
	"context"
	"fmt"
	"html"
	"io"
	"log"
	"mime"
	"net/http"
	"net/http/pprof"
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

// Configuration variables.
var (
	Version                    = "0.0.0.0"
	Port                       = "8801"
	RootDir                    = "./www"
	EnableWebConfig            = true
	EnableDirectoryListing     = false
	DirectoryListingTemplate   = "./www/axonasp-pages/directory-listing.html"
	DefaultPages               = []string{"index.asp", "default.asp", "index.html", "default.html", "default.asp"}
	ExecuteAsASPExtensions     = []string{".asp"}
	BlockedExtensions          = []string{}
	BlockedFiles               = []string{}
	BlockedDirs                = []string{}
	DefaultErrorPagesDirectory = "./www/error-pages"
	ScriptTimeout              = 60 // in seconds
	ResponseBufferLimitBytes   = 4 * 1024 * 1024
	DebugASP                   = false
	CleanupSessions            = true
	CleanupCache               = true
	DefaultTimezone            = "UTC"
	MemoryLimitMB              = 128
	VMPoolSize                 = 50
	BytecodeCachingMode        = "enabled"
	CacheMaxSizeMB             = 128
	SessionAutoFlushSeconds    = 15
	serverLocation             = time.UTC
	blockedDirPrefixes         = []string{}
	scriptCache                *axonvm.ScriptCache
	activeWebConfig            *WebConfigProcessor
	directoryListingRenderer   *DirectoryListingRenderer
)

// init loads environment variables and applies TOML-based configuration through Viper.
func init() {
	_ = godotenv.Load()
	registerFixedMIMETypes()
	axonvm.SetRuntimeVersion(strings.TrimSpace(Version))
	loadServerConfig()
	applyRuntimeSettings()
}

// registerFixedMIMETypes ensures critical content types are available even when host MIME tables are incomplete.
func registerFixedMIMETypes() {
	_ = mime.AddExtensionType(".svg", "image/svg+xml; charset=utf-8;")
}

// loadServerConfig loads and applies server/global settings from config/axonasp.toml using Viper.
func loadServerConfig() {
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

	if port := v.GetInt("server.server_port"); port > 0 {
		Port = strconv.Itoa(port)
	}
	if rootDir := strings.TrimSpace(v.GetString("server.web_root")); rootDir != "" {
		RootDir = rootDir
	}
	if pages := v.GetStringSlice("server.default_pages"); len(pages) > 0 {
		DefaultPages = pages
	}
	EnableWebConfig = v.GetBool("server.enable_webconfig")
	EnableDirectoryListing = v.GetBool("server.enable_directory_listing")
	if templatePath := strings.TrimSpace(v.GetString("server.directory_listing_template")); templatePath != "" {
		DirectoryListingTemplate = templatePath
	}
	if executeAsASP := v.GetStringSlice("global.execute_as_asp"); len(executeAsASP) > 0 {
		ExecuteAsASPExtensions = normalizeExtensions(executeAsASP)
	}
	if blocked := v.GetStringSlice("server.blocked_extensions"); len(blocked) > 0 {
		BlockedExtensions = normalizeExtensions(blocked)
	}
	if blockedFiles := v.GetStringSlice("server.blocked_files"); len(blockedFiles) > 0 {
		BlockedFiles = normalizeNames(blockedFiles)
	}
	if blockedDirs := v.GetStringSlice("server.blocked_dirs"); len(blockedDirs) > 0 {
		BlockedDirs = blockedDirs
	}
	if errorPagesDir := strings.TrimSpace(v.GetString("server.default_error_pages_directory")); errorPagesDir != "" {
		DefaultErrorPagesDirectory = errorPagesDir
	}

	ScriptTimeout = v.GetInt("global.default_script_timeout")
	if ScriptTimeout <= 0 {
		ScriptTimeout = 60
	}
	if responseBufferLimitMB := v.GetInt("global.response_buffer_limit_mb"); responseBufferLimitMB > 0 {
		ResponseBufferLimitBytes = responseBufferLimitMB * 1024 * 1024
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

	blockedDirPrefixes = buildBlockedDirPrefixes(BlockedDirs)
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

// normalizeNames normalizes names for case-insensitive matching.
func normalizeNames(values []string) []string {
	cleaned := make([]string, 0, len(values))
	for _, value := range values {
		name := strings.ToLower(strings.TrimSpace(value))
		if name == "" {
			continue
		}
		if !slices.Contains(cleaned, name) {
			cleaned = append(cleaned, name)
		}
	}
	return cleaned
}

// buildBlockedDirPrefixes resolves blocked directories into absolute normalized paths.
func buildBlockedDirPrefixes(blockedDirs []string) []string {
	resolved := make([]string, 0, len(blockedDirs))
	for _, blockedDir := range blockedDirs {
		candidate := strings.TrimSpace(blockedDir)
		if candidate == "" {
			continue
		}
		absPath, err := filepath.Abs(candidate)
		if err != nil {
			continue
		}
		if !slices.Contains(resolved, absPath) {
			resolved = append(resolved, absPath)
		}
	}
	return resolved
}

// dummyResponseWriter provides a no-op http.ResponseWriter for global.asa events.
type dummyResponseWriter struct{}

func (d *dummyResponseWriter) Header() http.Header         { return make(http.Header) }
func (d *dummyResponseWriter) Write(b []byte) (int, error) { return len(b), nil }
func (d *dummyResponseWriter) WriteHeader(statusCode int)  {}

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

// main starts the HTTP server and handles graceful shutdown.
func main() {
	runtime.SetBlockProfileRate(1)
	runtime.SetMutexProfileFraction(5)

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

	initializeServerOptionalFeatures()

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
		dummyHost := NewWebHost(&dummyResponseWriter{}, req)
		_ = axonvm.GetGlobalASA().ExecuteApplicationOnStart(dummyHost)
	}

	mux := http.NewServeMux()
	registerPprofHandlers(mux)
	mux.HandleFunc("/", handleRequest)

	httpServer := &http.Server{
		Addr:    ":" + Port,
		Handler: mux,
	}

	stop := make(chan os.Signal, 1)
	signal.Notify(stop, os.Interrupt, syscall.SIGTERM)

	go func() {
		fmt.Println("\033[H\033[2J\033[1mG3Pix AxonASP Server >\033[0m")
		fmt.Print("\033]11;#3b6ea5\007\033[1;37m")
		fmt.Println("--------------------")
		fmt.Printf("HTTP Server starting on port %s\n", Port)
		fmt.Printf("Serving files from: %s\n", RootDir)
		if err := httpServer.ListenAndServe(); err != nil && err != http.ErrServerClosed {
			axonvm.ReportInternalError(axonvm.ErrCouldNotListenOn, err, "HTTP server could not start listening.", Port, 0)
			os.Exit(1)
		}
	}()

	<-stop
	fmt.Println("\nShutting down server...")

	if axonvm.GetGlobalASA().IsLoaded() {
		// Execute Application_OnEnd using a dummy host
		req, _ := http.NewRequest("GET", "http://localhost/", nil)
		dummyHost := NewWebHost(&dummyResponseWriter{}, req)
		_ = axonvm.GetGlobalASA().ExecuteApplicationOnEnd(dummyHost)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()

	if err := httpServer.Shutdown(ctx); err != nil {
		axonvm.ReportInternalError(axonvm.ErrServerForcedToShutdown, err, "HTTP server shutdown failed.", "", 0)
		os.Exit(1)
	}

	fmt.Println("Server exited gracefully.")
}

// registerPprofHandlers exposes runtime profiling endpoints on the main server mux.
func registerPprofHandlers(mux *http.ServeMux) {
	mux.HandleFunc("/debug/pprof/", pprof.Index)
	mux.HandleFunc("/debug/pprof/cmdline", pprof.Cmdline)
	mux.HandleFunc("/debug/pprof/profile", pprof.Profile)
	mux.HandleFunc("/debug/pprof/symbol", pprof.Symbol)
	mux.HandleFunc("/debug/pprof/trace", pprof.Trace)
	mux.Handle("/debug/pprof/allocs", pprof.Handler("allocs"))
	mux.Handle("/debug/pprof/block", pprof.Handler("block"))
	mux.Handle("/debug/pprof/goroutine", pprof.Handler("goroutine"))
	mux.Handle("/debug/pprof/heap", pprof.Handler("heap"))
	mux.Handle("/debug/pprof/mutex", pprof.Handler("mutex"))
	mux.Handle("/debug/pprof/threadcreate", pprof.Handler("threadcreate"))
}

// handleRequest resolves the target path and serves static or ASP content.
func handleRequest(w http.ResponseWriter, r *http.Request) {
	path := r.URL.Path
	if path == "" {
		path = "/"
	}

	if activeWebConfig != nil {
		result, ok := activeWebConfig.Apply(path, r.URL.RawQuery)
		if ok {
			switch result.ActionType {
			case "redirect":
				http.Redirect(w, r, result.RedirectLocation, result.RedirectStatus)
				return
			case "rewrite":
				path = result.Path
				r.URL.Path = result.Path
				r.URL.RawQuery = result.RawQuery
			}
		}
	}

	relativePath := strings.TrimPrefix(path, "/")
	fullPath := filepath.Join(RootDir, filepath.FromSlash(relativePath))
	cleanPath := filepath.Clean(fullPath)
	requestedExt := strings.ToLower(filepath.Ext(cleanPath))
	requestedName := strings.ToLower(filepath.Base(cleanPath))

	absRoot, err := filepath.Abs(RootDir)
	if err != nil {
		respondInternalHTTPError(w, axonvm.ErrCouldNotResolveCurrentDir, err, "Failed to resolve the configured root directory.", RootDir)
		return
	}

	absPath, err := filepath.Abs(cleanPath)
	if err != nil {
		respondInternalHTTPError(w, axonvm.ErrBadFileName, err, "Failed to resolve the requested path.", cleanPath)
		return
	}

	if !strings.HasPrefix(absPath, absRoot) {
		serveErrorPage(w, r, http.StatusForbidden)
		return
	}

	if isBlockedExtension(requestedExt) || isBlockedFile(requestedName) || isBlockedDirectory(absPath) {
		serveErrorPage(w, r, http.StatusNotFound)
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
		if !strings.HasSuffix(path, "/") {
			redirectPath := path + "/"
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
				if isBlockedExtension(strings.ToLower(filepath.Ext(candidate))) || isBlockedFile(strings.ToLower(filepath.Base(candidate))) {
					continue
				}
				fullPath = candidate
				foundDefault = true
				break
			}
		}

		if !foundDefault {
			if EnableDirectoryListing && directoryListingRenderer != nil {
				if err := directoryListingRenderer.Render(w, r, fullPath, path); err == nil {
					return
				}
			}
			serveErrorPage(w, r, http.StatusNotFound)
			return
		}
	}

	if !isASPExecutionExtension(fullPath) {
		serveStaticFileWithMIME(w, r, fullPath)
		return
	}

	executeASP(w, r, fullPath)
}

// isASPExecutionExtension reports whether a file should be executed as ASP based on configured extensions.
func isASPExecutionExtension(filePath string) bool {
	ext := strings.ToLower(filepath.Ext(filePath))
	return slices.Contains(ExecuteAsASPExtensions, ext)
}

// isBlockedExtension reports whether a file extension is blocked for direct requests.
func isBlockedExtension(ext string) bool {
	if ext == "" {
		return false
	}
	return slices.Contains(BlockedExtensions, strings.ToLower(ext))
}

// isBlockedFile reports whether a requested base filename is blocked.
func isBlockedFile(baseName string) bool {
	if baseName == "" {
		return false
	}
	return slices.Contains(BlockedFiles, strings.ToLower(baseName))
}

// isBlockedDirectory reports whether a requested absolute path belongs to any blocked directory.
func isBlockedDirectory(absPath string) bool {
	for _, prefix := range blockedDirPrefixes {
		if absPath == prefix || strings.HasPrefix(absPath, prefix+string(os.PathSeparator)) {
			return true
		}
	}
	return false
}

// serveStaticFileWithMIME resolves and sets Content-Type using mime.TypeByExtension before serving files.
func serveStaticFileWithMIME(w http.ResponseWriter, r *http.Request, filePath string) {
	ext := strings.ToLower(filepath.Ext(filePath))
	if ext != "" {
		if contentType := mime.TypeByExtension(ext); contentType != "" {
			w.Header().Set("Content-Type", contentType)
		}
	}
	http.ServeFile(w, r, filePath)
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
	if w.defaultStatus > 0 && (statusCode <= 0 || statusCode == http.StatusOK) {
		statusCode = w.defaultStatus
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

// cancellableWriter wraps an http.ResponseWriter and silently discards all writes
// after cancel() is called. This is used to safely detach a goroutine that is still
// executing ASP (typically stuck inside a blocking CGO/COM call) from the real
// http.ResponseWriter after the script timeout has fired.
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
// goroutine so that a blocking CGO/COM call (e.g. OLE ADODB.Execute) cannot hold
// the HTTP handler indefinitely. If the goroutine does not complete within
// ScriptTimeout seconds the handler cancels the response writer, writes a 503 to
// the client, and returns. The goroutine continues until the CGO call unblocks,
// then its deferred CleanupRequestResources drains OLE objects.
func executeASPWithStatus(w http.ResponseWriter, r *http.Request, filePath string, defaultStatus int) {
	single := newSingleHeaderResponseWriter(w, defaultStatus)
	cw := newCancellableWriter(single)
	host := NewWebHost(cw, r)

	if scriptCache == nil {
		scriptCache = axonvm.NewScriptCache(axonvm.BytecodeCacheDisabled, filepath.Join("temp", "cache"), 1)
	}
	program, err := scriptCache.LoadOrCompile(filePath)
	if err != nil {
		aspErr := axonvm.CompilerErrorToASPError(err, filePath)
		host.Server().SetLastError(aspErr)
		axonvm.LogASPProcessedError(aspErr, "server.executeASP.compile")
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
				axonvm.LogASPProcessedError(aspErr, "server.executeASP.runtime")
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
				timeoutErr := fmt.Errorf("script timeout reached after %ds", effectiveTimeout)
				respondInternalHTTPError(
					w,
					axonvm.ErrScriptTimeoutDetachedGoroutine,
					timeoutErr,
					fmt.Sprintf("Detached blocked ASP execution goroutine after script timeout (%ds).", effectiveTimeout),
					filePath,
				)
				return
			}
		}
	}
}

// isErrorPageHandlerPath reports whether the current execution target is an error page handler itself.
func isErrorPageHandlerPath(filePath string) bool {
	errorDirAbs, err := filepath.Abs(DefaultErrorPagesDirectory)
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
		source = "VBScript runtime"
	}
	description := html.EscapeString(strings.TrimSpace(err.Description))
	if description == "" {
		description = "Unknown runtime error"
	}
	fileName := html.EscapeString(strings.TrimSpace(err.File))
	if fileName == "" {
		fileName = "unknown"
	}

	fmt.Fprintf(w, "<html><head><title>500 - Internal Server Error</title></head><body style=\"font-family:'Cascadia Mono','SF Mono Regular','Ubuntu Mono',consolas,monospace;\">")
	fmt.Fprintf(w, "<h2>Server Error in AxonASP Application.</h2>")
	fmt.Fprintf(w, "<hr><p><b>%s error '%08X'</b></p>", source, uint32(int32(err.Number)))
	fmt.Fprintf(w, "<p>%s</p>", description)
	fmt.Fprintf(w, "<p><b>File:</b> %s<br><b>Line:</b> %d<br><b>Column:</b> %d</p>", fileName, err.Line, err.Column)
	fmt.Fprintf(w, "<p><b>Stage:</b> %s</p>", html.EscapeString(stage))
	fmt.Fprintf(w, "<hr><p><i>G3Pix AxonASP Debug Mode</i></p></body></html>")
}

// serveErrorPage serves configured error pages using .asp or .html handlers for a given HTTP status code.
func serveErrorPage(w http.ResponseWriter, r *http.Request, statusCode int) {
	if activeWebConfig != nil {
		if customError, ok := activeWebConfig.GetCustomError(statusCode); ok {
			if serveWebConfigCustomError(w, r, statusCode, customError) {
				return
			}
		}
	}

	aspPagePath := filepath.Join(DefaultErrorPagesDirectory, fmt.Sprintf("%d.asp", statusCode))
	if pageInfo, err := os.Stat(aspPagePath); err == nil && !pageInfo.IsDir() {
		executeASPWithStatus(w, r, aspPagePath, statusCode)
		return
	}

	htmlPagePath := filepath.Join(DefaultErrorPagesDirectory, fmt.Sprintf("%d.html", statusCode))
	if pageInfo, err := os.Stat(htmlPagePath); err == nil && !pageInfo.IsDir() {
		serveStaticFileWithMIME(newSingleHeaderResponseWriter(w, statusCode), r, htmlPagePath)
		return
	}

	http.Error(w, http.StatusText(statusCode), statusCode)
}

// initializeServerOptionalFeatures configures server-only routing helpers.
func initializeServerOptionalFeatures() {
	activeWebConfig = nil
	directoryListingRenderer = nil

	if EnableWebConfig {
		processor, err := NewWebConfigProcessor(RootDir)
		if err != nil {
			log.Printf("Warning: Failed to load web.config, using default routing: %v\n", err)
		} else {
			activeWebConfig = processor
		}
	}

	if EnableDirectoryListing {
		renderer, err := NewDirectoryListingRenderer(RootDir, DirectoryListingTemplate)
		if err != nil {
			log.Printf("Warning: Failed to load directory listing template, disabling listing: %v\n", err)
		} else {
			directoryListingRenderer = renderer
		}
	}
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

	if code == axonvm.ErrScriptTimeoutDetachedGoroutine {
		return int(axonvm.HTTPServiceUnavailable)
	}

	return int(axonvm.HTTPInternalServerError)
}

package caddy

import (
	"fmt"
	"html"
	"io"
	"mime"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"runtime/debug"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/caddyserver/caddy/v2"
	"github.com/caddyserver/caddy/v2/caddyconfig/caddyfile"
	"github.com/caddyserver/caddy/v2/caddyconfig/httpcaddyfile"
	"github.com/caddyserver/caddy/v2/modules/caddyhttp"
	"github.com/spf13/viper"
	"go.uber.org/zap"

	"g3pix.com.br/axonasp/v2/axonconfig"
	"g3pix.com.br/axonasp/v2/axonvm"
	"g3pix.com.br/axonasp/v2/axonvm/asp"
)

// Interface guards to verify implementation of standard Caddy v2 interfaces.
var (
	_ caddy.Module                = (*AxonASP)(nil)
	_ caddy.Provisioner           = (*AxonASP)(nil)
	_ caddy.Validator             = (*AxonASP)(nil)
	_ caddy.CleanerUpper          = (*AxonASP)(nil)
	_ caddyhttp.MiddlewareHandler = (*AxonASP)(nil)
	_ caddyfile.Unmarshaler       = (*AxonASP)(nil)
)

func init() {
	caddy.RegisterModule(AxonASP{})
	httpcaddyfile.RegisterHandlerDirective("axonasp", parseCaddyfile)
	httpcaddyfile.RegisterDirectiveOrder("axonasp", httpcaddyfile.Before, "file_server")
	_ = mime.AddExtensionType(".svg", "image/svg+xml")
}

// AxonASP implements a Caddy HTTP handler for the AxonASP runtime.
type AxonASP struct {
	SiteName      string `json:"site_name,omitempty"`
	ConfigFile    string `json:"config_file,omitempty"`
	GlobalAsaPath string `json:"global_asa_path,omitempty"`

	logger      *zap.Logger
	scriptCache *axonvm.ScriptCache
	globalASA   *axonvm.GlobalASA
	application *asp.Application
	config      *viper.Viper

	resolvedConfigPath string
}

// Version is the AxonASP Caddy module version. Build scripts may stamp it with
// -X g3pix.com.br/axonasp/caddy.Version=<version>; otherwise the module resolves
// the version from the Go build info at provision time.
var Version = ""

const (
	// axonaspVersionLine is the AxonASP release line used when the build does not
	// provide an explicit version.
	axonaspVersionLine = "2.3"
	// axonaspVersionFallback is used when no build metadata is available at all.
	axonaspVersionFallback = axonaspVersionLine + ".0"
	// axonaspModulePath identifies this module inside the Go build info.
	axonaspModulePath = "g3pix.com.br/axonasp/caddy"
	// caddyCoreModulePath identifies the Caddy core module inside the Go build info.
	caddyCoreModulePath = "github.com/caddyserver/caddy/v2"
	// shortRevisionLength is the number of commit characters kept in a version.
	shortRevisionLength = 7
)

// resolveRuntimeVersion returns the version reported by AxVersion() and by the
// runtime banner inside the Caddy host. xcaddy does not accept linker flags, so
// resolution falls back to the Go build info in this order:
//  1. Version injected with -X g3pix.com.br/axonasp/caddy.Version=...
//  2. the module version of this package, which is a release tag for tagged
//     builds and a pseudo-version (timestamp plus commit) for source builds
//  3. the VCS revision stamped by the Go toolchain
//  4. axonaspVersionFallback
//
// The Caddy core version is appended as build metadata, so the reported engine
// version identifies the exact AxonASP and Caddy pair in use.
func resolveRuntimeVersion() string {
	if injected := strings.TrimSpace(Version); injected != "" {
		return injected
	}

	buildInfo, ok := debug.ReadBuildInfo()
	if !ok {
		return axonaspVersionFallback
	}

	resolved := versionFromModuleInfo(buildInfo)
	if resolved == "" {
		resolved = versionFromVCSStamp(buildInfo)
	}
	if resolved == "" {
		resolved = axonaspVersionFallback
	}

	if caddyVersion := coreCaddyVersion(buildInfo); caddyVersion != "" && !strings.Contains(resolved, "+") {
		resolved += "+caddy." + caddyVersion
	}
	return resolved
}

// versionFromModuleInfo derives the AxonASP version from the module metadata
// recorded for this package in the Go build info.
func versionFromModuleInfo(buildInfo *debug.BuildInfo) string {
	if buildInfo == nil {
		return ""
	}

	for _, dep := range buildInfo.Deps {
		if dep == nil || dep.Path != axonaspModulePath {
			continue
		}
		moduleVersion := strings.TrimSpace(dep.Version)
		if moduleVersion == "" && dep.Replace != nil {
			moduleVersion = strings.TrimSpace(dep.Replace.Version)
		}
		return normalizeBuildVersion(moduleVersion)
	}
	return ""
}

// normalizeBuildVersion converts a Go module version into the AxonASP version
// format: release tags keep their numeric form (2.3.22), while pseudo-versions
// (v0.0.0-20260907212856-9ded92f3cc7a) degrade to line plus short revision
// (2.3.0.9ded92f).
func normalizeBuildVersion(moduleVersion string) string {
	moduleVersion = strings.TrimPrefix(strings.TrimSpace(moduleVersion), "v")
	if moduleVersion == "" || moduleVersion == "0.0.0" {
		return ""
	}

	if isReleaseVersion(moduleVersion) {
		return moduleVersion
	}
	if strings.HasPrefix(moduleVersion, axonaspVersionFallback+"-") || strings.HasPrefix(moduleVersion, "0.0.0-") {
		revision := pseudoVersionRevision(moduleVersion)
		if revision == "" {
			return ""
		}
		return axonaspVersionFallback + "." + revision
	}
	return moduleVersion
}

// isReleaseVersion reports whether a module version is a plain MAJOR.MINOR.PATCH tag.
func isReleaseVersion(version string) bool {
	parts := strings.Split(version, ".")
	if len(parts) != 3 {
		return false
	}

	for _, part := range parts {
		if part == "" {
			return false
		}
		for i := range len(part) {
			if part[i] < '0' || part[i] > '9' {
				return false
			}
		}
	}
	return true
}

// pseudoVersionRevision extracts the short commit hash from a Go pseudo-version.
func pseudoVersionRevision(version string) string {
	lastDash := strings.LastIndexByte(version, '-')
	if lastDash < 0 || lastDash == len(version)-1 {
		return ""
	}

	revision := version[lastDash+1:]
	if len(revision) > shortRevisionLength {
		revision = revision[:shortRevisionLength]
	}
	return revision
}

// versionFromVCSStamp derives a version from the VCS metadata the Go toolchain
// stamps into the binary when the build runs from a checkout.
func versionFromVCSStamp(buildInfo *debug.BuildInfo) string {
	if buildInfo == nil {
		return ""
	}

	revision := ""
	for _, setting := range buildInfo.Settings {
		if setting.Key == "vcs.revision" {
			revision = strings.TrimSpace(setting.Value)
			break
		}
	}
	if revision == "" {
		return ""
	}
	if len(revision) > shortRevisionLength {
		revision = revision[:shortRevisionLength]
	}
	return axonaspVersionFallback + "." + revision
}

// coreCaddyVersion returns the Caddy core version without its leading "v".
func coreCaddyVersion(buildInfo *debug.BuildInfo) string {
	if buildInfo == nil {
		return ""
	}

	for _, dep := range buildInfo.Deps {
		if dep == nil || dep.Path != caddyCoreModulePath {
			continue
		}
		version := strings.TrimSpace(dep.Version)
		if version == "" && dep.Replace != nil {
			version = strings.TrimSpace(dep.Replace.Version)
		}
		return strings.TrimPrefix(version, "v")
	}
	return ""
}

// AcquireVM borrows a VM instance from the shared AxonASP interpreter pool.
// The axonvm runtime pools instances per compiled program, so recycling is
// owned by the VM package and this module never touches VM internals.
func (a *AxonASP) AcquireVM(program axonvm.CachedProgram) *axonvm.VM {
	return axonvm.AcquireVMFromCachedProgram(program)
}

// CaddyModule returns the Caddy module information.
func (AxonASP) CaddyModule() caddy.ModuleInfo {
	return caddy.ModuleInfo{
		ID:  "http.handlers.axonasp",
		New: func() caddy.Module { return new(AxonASP) },
	}
}

// Provision sets up the AxonASP Caddy module.
func (a *AxonASP) Provision(ctx caddy.Context) error {
	a.logger = ctx.Logger(a)

	// xcaddy does not accept linker flags, so the engine version is resolved from
	// the build info recorded for this module instead of relying on -X injection.
	runtimeVersion := resolveRuntimeVersion()
	axonvm.SetRuntimeVersion(runtimeVersion)

	if strings.TrimSpace(a.ConfigFile) != "" {
		resolved, err := filepath.Abs(a.ConfigFile)
		if err != nil {
			return fmt.Errorf("invalid config_file path: %w", err)
		}
		axonconfig.SetCustomConfigPath(resolved)
	}

	active := axonconfig.NewViper()
	configPath := active.ConfigFileUsed()
	if configPath != "" {
		if absConfigPath, err := filepath.Abs(configPath); err == nil {
			configPath = absConfigPath
		}
	}
	if configPath == "" {
		// Fallback to custom candidate search
		var configErr error
		configPath, configErr = a.resolveConfigFilePath()
		if configErr != nil {
			return fmt.Errorf("failed to locate axonasp.toml using loader or fallback: %w", configErr)
		}
		axonconfig.SetCustomConfigPath(configPath)
		active = axonconfig.NewViper()
		configPath = active.ConfigFileUsed()
		if absConfigPath, err := filepath.Abs(configPath); err == nil {
			configPath = absConfigPath
		}
	}

	a.resolvedConfigPath = configPath
	if normalizeErr := normalizeAxFunctionResourcePaths(active, configPath); normalizeErr != nil {
		return fmt.Errorf("failed to normalize axfunctions resource paths: %w", normalizeErr)
	}

	v := viper.New()
	v.SetConfigType("toml")
	v.SetConfigFile(configPath)
	if err := v.ReadInConfig(); err != nil {
		return fmt.Errorf("failed to read config file %q: %w", configPath, err)
	}

	if normalizeErr := normalizeAxFunctionResourcePaths(v, configPath); normalizeErr != nil {
		return fmt.Errorf("failed to normalize local config resource paths: %w", normalizeErr)
	}
	a.config = v
	a.logger.Info("Loaded AxonASP config for Caddy", zap.String("path", configPath), zap.String("version", runtimeVersion))

	siteTemp := a.resolveSiteTempDir(nil)
	a.setupSiteTempDir(siteTemp)
	a.logger.Info("Provision: Intercepted AxonASP TempDir to Caddy native temp storage", zap.String("path", siteTemp))

	// Verify logo path exists if configured
	logoPath := active.GetString("axfunctions.ax_default_logo_path")
	if logoPath != "" {
		a.logger.Info("Provision: AxonASP logo path configured to", zap.String("path", logoPath))
		if _, err := os.Stat(logoPath); err != nil {
			return fmt.Errorf("configured ax_default_logo_path %q does not exist: %w", logoPath, err)
		}
		a.logger.Info("Provision: AxonASP logo file exists on disk", zap.String("path", logoPath))
	} else {
		a.logger.Info("Provision: AxonASP logo path is not configured")
	}

	// Verify CSS path exists if configured
	cssPath := active.GetString("axfunctions.ax_default_css_path")
	if cssPath != "" {
		a.logger.Info("Provision: AxonASP CSS path configured to", zap.String("path", cssPath))
		if _, err := os.Stat(cssPath); err != nil {
			return fmt.Errorf("configured ax_default_css_path %q does not exist: %w", cssPath, err)
		}
		a.logger.Info("Provision: AxonASP CSS file exists on disk", zap.String("path", cssPath))
	} else {
		a.logger.Info("Provision: AxonASP CSS path is not configured")
	}

	a.scriptCache = axonvm.NewScriptCache(axonvm.BytecodeCacheMemoryOnly, "", 64)
	a.application = asp.NewApplication()
	a.globalASA = &axonvm.GlobalASA{}

	if a.GlobalAsaPath != "" {
		webRoot := filepath.Dir(a.GlobalAsaPath)
		err := a.globalASA.LoadAndCompile(webRoot, a.application)
		if err != nil {
			a.logger.Warn("Failed to load global.asa", zap.Error(err), zap.String("path", a.GlobalAsaPath))
		} else {
			a.logger.Info("Loaded global.asa", zap.String("path", a.GlobalAsaPath))
			// Execute Application_OnStart using a dummy host to initialize state
			req, _ := http.NewRequest("GET", "http://localhost/", nil)
			dummyHost := NewCaddyWebHost(&dummyResponseWriter{}, req, a, webRoot)
			_ = a.globalASA.ExecuteApplicationOnStart(dummyHost)
		}
	}

	if a.isG3AxonLiveActive() {
		axonvm.G3ALStartCleanup(30)
	}

	return nil
}

// Validate ensures the module is properly configured.
func (a *AxonASP) Validate() error {
	if a.ConfigFile != "" {
		if _, err := os.Stat(a.ConfigFile); os.IsNotExist(err) {
			return fmt.Errorf("configured config_file path does not exist: %s", a.ConfigFile)
		}
	}
	if a.GlobalAsaPath != "" {
		if _, err := os.Stat(a.GlobalAsaPath); os.IsNotExist(err) {
			return fmt.Errorf("configured global_asa_path does not exist: %s", a.GlobalAsaPath)
		}
	}
	return nil
}

// Cleanup releases every resource owned by this handler when Caddy unloads or
// reloads the module: the bytecode cache invalidator, the pooled VM instances
// and the G3AxonLive background sweeper. This keeps configuration reloads from
// stacking idle VM workers and goroutines in the process.
func (a *AxonASP) Cleanup() error {
	if a.scriptCache != nil {
		a.scriptCache.StopInvalidator()
	}
	axonvm.PurgeVMProgramPools()
	axonvm.G3ALStopCleanup()
	return nil
}

func (a *AxonASP) resolveConfigFilePath() (string, error) {
	if strings.TrimSpace(a.ConfigFile) != "" {
		resolved, err := filepath.Abs(a.ConfigFile)
		if err != nil {
			return "", err
		}
		if _, statErr := os.Stat(resolved); statErr != nil {
			return "", statErr
		}
		return resolved, nil
	}

	candidates := make([]string, 0, 16)

	cwd, cwdErr := os.Getwd()
	if cwdErr == nil {
		candidates = append(candidates,
			filepath.Join(cwd, "config", "axonasp.toml"),
			filepath.Join(cwd, "axonasp.toml"),
		)
	}

	if a.GlobalAsaPath != "" {
		webRoot := filepath.Dir(a.GlobalAsaPath)
		projectRoot := filepath.Dir(webRoot)
		candidates = append(candidates,
			filepath.Join(projectRoot, "config", "axonasp.toml"),
			filepath.Join(webRoot, "config", "axonasp.toml"),
		)
	}

	if executablePath, err := os.Executable(); err == nil {
		exeDir := filepath.Dir(executablePath)
		candidates = append(candidates,
			filepath.Join(exeDir, "config", "axonasp.toml"),
			filepath.Join(exeDir, "axonasp.toml"),
		)
	}

	if cwdErr == nil {
		walk := cwd
		for range 8 {
			candidates = append(candidates,
				filepath.Join(walk, "config", "axonasp.toml"),
				filepath.Join(walk, "axonasp.toml"),
			)
			parent := filepath.Dir(walk)
			if parent == walk {
				break
			}
			walk = parent
		}
	}

	seen := make(map[string]struct{}, len(candidates))
	for _, candidate := range candidates {
		if strings.TrimSpace(candidate) == "" {
			continue
		}
		abs, err := filepath.Abs(candidate)
		if err != nil {
			continue
		}
		if _, ok := seen[abs]; ok {
			continue
		}
		seen[abs] = struct{}{}
		if _, statErr := os.Stat(abs); statErr == nil {
			return abs, nil
		}
	}

	return "", fmt.Errorf("unable to locate axonasp.toml")
}

func normalizeAxFunctionResourcePaths(v *viper.Viper, configPath string) error {
	if v == nil {
		return fmt.Errorf("nil viper instance")
	}

	absConfigPath, err := filepath.Abs(configPath)
	if err != nil {
		absConfigPath = configPath
	}
	configDir := filepath.Dir(absConfigPath)
	normalize := func(key string) {
		val := strings.TrimSpace(v.GetString(key))
		if val == "" || filepath.IsAbs(val) {
			return
		}
		// Try relative to configDir first
		p := filepath.Clean(filepath.Join(configDir, val))
		if _, err := os.Stat(p); err != nil {
			// Try relative to the parent of configDir (project root)
			p2 := filepath.Clean(filepath.Join(filepath.Dir(configDir), val))
			if _, err2 := os.Stat(p2); err2 == nil {
				p = p2
			}
		}
		if absPath, err := filepath.Abs(p); err == nil {
			p = absPath
		}
		v.Set(key, p)
	}

	normalize("axfunctions.ax_default_logo_path")
	normalize("axfunctions.ax_default_css_path")
	return nil
}

// ServeHTTP bridges the Caddy HTTP request and serves ASP files.
func (a *AxonASP) ServeHTTP(w http.ResponseWriter, r *http.Request, next caddyhttp.Handler) error {
	path := r.URL.Path
	if path == "" {
		path = "/"
	}

	if isAxonLiveEndpoint(path) {
		if a.isG3AxonLiveActive() {
			siteTemp := a.resolveSiteTempDir(r)
			a.setupSiteTempDir(siteTemp)
			return a.handleAxonLive(w, r)
		}
	}

	cleanReqPath := strings.TrimRight(path, "/")
	reqFile := strings.ToLower(filepath.Base(cleanReqPath))
	if reqFile == "global.asa" || reqFile == "myinfo.xml" {
		http.Error(w, "404 page not found", http.StatusNotFound)
		return nil
	}

	webRoot := a.resolveWebRoot(r)

	aspPath, shouldExecuteASP, shouldRedirect := a.matchASPRequest(webRoot, path)

	if !shouldExecuteASP {
		return next.ServeHTTP(w, r)
	}

	if shouldRedirect {
		redirectPath := path + "/"
		if r.URL.RawQuery != "" {
			redirectPath += "?" + r.URL.RawQuery
		}
		http.Redirect(w, r, redirectPath, http.StatusMovedPermanently)
		return nil
	}

	siteTemp := a.resolveSiteTempDir(r)
	a.setupSiteTempDir(siteTemp)

	return a.executeASPFile(w, r, webRoot, aspPath)
}

func (a *AxonASP) matchASPRequest(webRoot, path string) (aspPath string, shouldExecuteASP bool, shouldRedirect bool) {
	cleanReqPath := strings.TrimPrefix(path, "/")
	filePath := filepath.Clean(filepath.Join(webRoot, filepath.FromSlash(cleanReqPath)))

	// Security check: ensure target path is within webRoot
	absRoot, rootErr := filepath.Abs(webRoot)
	absFile, fileErr := filepath.Abs(filePath)
	if rootErr == nil && fileErr == nil {
		if !strings.HasPrefix(absFile+string(filepath.Separator), absRoot+string(filepath.Separator)) && absFile != absRoot {
			return "", false, false
		}
	}

	ext := strings.ToLower(filepath.Ext(filePath))

	// Direct ASP file request
	if ext == ".asp" {
		if info, err := os.Stat(filePath); err == nil && !info.IsDir() {
			return filePath, true, false
		}
		return "", false, false
	}

	// Non-ASP extensions (.php, .html, .js, .css, etc.) bypass ASP execution
	if ext != "" && ext != "." {
		return "", false, false
	}

	// Directory or root path request
	isDirRequest := path == "/" || strings.HasSuffix(path, "/")
	if !isDirRequest {
		if info, err := os.Stat(filePath); err == nil && info.IsDir() {
			if resolved, ok := resolveDefaultASPPath(webRoot, path+"/"); ok {
				resolvedFilePath := filepath.Clean(filepath.Join(webRoot, filepath.FromSlash(strings.TrimPrefix(resolved, "/"))))
				if infoRes, errRes := os.Stat(resolvedFilePath); errRes == nil && !infoRes.IsDir() {
					return canonicalDefaultASPPath(resolvedFilePath), true, true
				}
			}
		}
		return "", false, false
	}

	// Directory request ending with "/"
	if resolved, ok := resolveDefaultASPPath(webRoot, path); ok {
		resolvedFilePath := filepath.Clean(filepath.Join(webRoot, filepath.FromSlash(strings.TrimPrefix(resolved, "/"))))
		if infoRes, errRes := os.Stat(resolvedFilePath); errRes == nil && !infoRes.IsDir() {
			return canonicalDefaultASPPath(resolvedFilePath), true, false
		}
	}

	return "", false, false
}

// canonicalDefaultASPPath resolves one default-document candidate to its real
// physical form on disk. Default documents are synthesized from a fixed list, so
// the on-disk casing must be recovered for implicit directory routes to compile
// the exact same file as an explicit request.
func canonicalDefaultASPPath(path string) string {
	if canonical, err := axonvm.CanonicalizePhysicalPath(path); err == nil {
		return canonical
	}
	return path
}

func (a *AxonASP) executeASPFile(w http.ResponseWriter, r *http.Request, webRoot string, cleanPath string) error {
	w.Header().Set("X-Powered-By", "AxonASP")

	single := newSingleHeaderResponseWriter(w, 0)
	cw := newCancellableWriter(single)
	host := NewCaddyWebHost(cw, r, a, webRoot)

	program, err := a.scriptCache.LoadOrCompileWithOptions(cleanPath, axonvm.ScriptCompileOptions{IncludeSiteRoot: host.Server().MapPath("/")})
	if err != nil {
		aspErr := axonvm.CompilerErrorToASPError(err, cleanPath)
		host.Server().SetLastError(aspErr)

		a.logger.Error("Compilation Error",
			zap.String("site_name", a.SiteName),
			zap.String("file", cleanPath),
			zap.String("description", aspErr.Description),
			zap.Int("line", aspErr.Line),
			zap.Int("column", aspErr.Column),
			zap.String("source", aspErr.Source),
		)

		renderClassicASPDebugError(w, http.StatusInternalServerError, "Compilation Error", aspErr)
		return nil
	}

	vm := a.AcquireVM(program)
	vm.SetHost(host)

	timeoutSec := 60
	if srv := host.Server(); srv != nil {
		if t := srv.GetScriptTimeout(); t > 0 {
			timeoutSec = t
		}
	}

	type vmResult struct{ err error }
	done := make(chan vmResult, 1)
	go func() {
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

	start := time.Now()
	watchdog := time.NewTicker(250 * time.Millisecond)
	defer watchdog.Stop()

	for {
		select {
		case res := <-done:
			if res.err != nil {
				aspErr := axonvm.RuntimeErrorToASPError(res.err, cleanPath)
				host.Server().SetLastError(aspErr)

				a.logger.Error("Runtime Error",
					zap.String("site_name", a.SiteName),
					zap.String("file", cleanPath),
					zap.String("description", aspErr.Description),
					zap.Int("line", aspErr.Line),
					zap.Int("column", aspErr.Column),
					zap.String("source", aspErr.Source),
				)

				renderClassicASPDebugError(w, http.StatusInternalServerError, "Runtime Error", aspErr)
				return nil
			}
			host.PersistSession()
			host.Response().Flush()
			host.Response().ReleaseBuffer()
			return nil

		case <-watchdog.C:
			effectiveTimeout := timeoutSec
			if srv := host.Server(); srv != nil {
				if t := srv.GetScriptTimeout(); t > 0 {
					effectiveTimeout = t
				}
			}
			if time.Since(start) >= time.Duration(effectiveTimeout)*time.Second {
				cw.cancel()
				timeoutErr := fmt.Errorf("script timeout reached after %ds", effectiveTimeout)

				a.logger.Error("Script Timeout",
					zap.String("site_name", a.SiteName),
					zap.String("file", cleanPath),
					zap.Error(timeoutErr),
				)

				http.Error(w, "Script Timeout: "+timeoutErr.Error(), http.StatusGatewayTimeout)
				return nil
			}
		}
	}
}

// CaddyWebHost implements axonvm.ASPHostEnvironment for Caddy requests.
type CaddyWebHost struct {
	response       *asp.Response
	request        *asp.Request
	server         *asp.Server
	session        *asp.Session
	application    *asp.Application
	sessionEnabled bool
	engineMode     axonvm.EngineMode
	site           *AxonASP
}

func (h *CaddyWebHost) Response() *asp.Response        { return h.response }
func (h *CaddyWebHost) Request() *asp.Request          { return h.request }
func (h *CaddyWebHost) Server() *asp.Server            { return h.server }
func (h *CaddyWebHost) Session() *asp.Session          { return h.session }
func (h *CaddyWebHost) Application() *asp.Application  { return h.application }
func (h *CaddyWebHost) SetSessionEnabled(enabled bool) { h.sessionEnabled = enabled }
func (h *CaddyWebHost) SessionEnabled() bool           { return h.sessionEnabled }
func (h *CaddyWebHost) EngineMode() axonvm.EngineMode  { return h.engineMode }

func (h *CaddyWebHost) Write(p []byte) (int, error) {
	h.response.Write(string(p))
	return len(p), nil
}

func (h *CaddyWebHost) WriteString(s string) {
	h.response.Write(s)
}

func (h *CaddyWebHost) ExecuteASPFile(absPath string) error {
	previousRequestPath := h.server.GetRequestPath()
	h.server.SetRequestPath(h.server.VirtualPathFromAbsolutePath(absPath))
	defer h.server.SetRequestPath(previousRequestPath)

	cache := h.site.scriptCache
	program := axonvm.CachedProgram{}
	if cache != nil {
		if cached, found := cache.Get(absPath); found {
			program = cached
		} else {
			compiled, compileErr := cache.LoadOrCompileWithOptions(absPath, axonvm.ScriptCompileOptions{IncludeSiteRoot: h.server.MapPath("/")})
			if compileErr != nil {
				return compileErr
			}
			program = compiled
		}
	} else {
		content, err := os.ReadFile(absPath)
		if err != nil {
			return err
		}
		if len(content) >= 3 && content[0] == 0xEF && content[1] == 0xBB && content[2] == 0xBF {
			content = content[3:]
		}

		compiler := axonvm.NewASPCompiler(string(content))

		compiler.SetSourceName(absPath)
		compiler.SetIncludeSiteRoot(h.server.MapPath("/"))
		if err := compiler.Compile(); err != nil {
			return err
		}
		childVM := axonvm.NewVMFromCompiler(compiler)
		childVM.SetHost(h)
		defer childVM.Release()
		return childVM.Run()
	}

	childVM := h.site.AcquireVM(program)
	childVM.SetHost(h)
	defer childVM.Release()
	return childVM.Run()
}

func NewCaddyWebHost(w http.ResponseWriter, r *http.Request, site *AxonASP, webRoot string) *CaddyWebHost {
	session, isNew := loadOrCreateSession(r)

	host := &CaddyWebHost{
		response:       asp.NewResponse(w),
		request:        asp.NewRequest(),
		server:         asp.NewServer(),
		session:        session,
		application:    site.application,
		sessionEnabled: true,
		engineMode:     axonvm.EngineModeDefault,
		site:           site,
	}

	host.response.SetRequest(r)
	host.response.SetMaxBufferBytes(4 * 1024 * 1024)
	host.request.SetHTTPRequest(r)
	host.server.SetRootDir(webRoot)
	host.server.SetRequestPath(r.URL.Path)
	_ = host.server.SetScriptTimeout(60)

	if len(r.URL.RawQuery) > 0 {
		host.request.QueryString.SetLazyPayload([]byte(r.URL.RawQuery))
	}

	host.request.SetBodyLoader(func() ([]byte, error) {
		if r.Body == nil {
			return []byte{}, nil
		}
		loadedBody, err := io.ReadAll(r.Body)
		if err != nil {
			return nil, err
		}
		return loadedBody, nil
	})

	for _, cookie := range r.Cookies() {
		host.request.Cookies.AddCookie(cookie.Name, cookie.Value)
	}

	hostName := requestServerName(r)
	port := requestServerPort(r)
	queryString := ""
	requestURI := r.URL.Path
	if r.URL != nil {
		queryString = r.URL.RawQuery
		requestURI = r.URL.RequestURI()
	}
	httpsValue := "off"
	if r.TLS != nil {
		httpsValue = "on"
	}
	host.request.ServerVars.Add("QUERY_STRING", queryString)
	host.request.ServerVars.Add("HTTP_HOST", r.Host)
	host.request.ServerVars.Add("HTTP_CONTENT_TYPE", r.Header.Get("Content-Type"))
	host.request.ServerVars.Add("HTTPS", httpsValue)
	host.request.ServerVars.Add("AUTH_TYPE", requestAuthType(r))
	host.request.ServerVars.Add("SERVER_ADDR", requestServerAddr(r))
	host.request.ServerVars.Add("GATEWAY_INTERFACE", "CGI/1.1")
	host.request.ServerVars.Add("SERVER_SOFTWARE", "G3pix-AxonASP-Caddy")
	host.request.ServerVars.Add("SERVER_PROTOCOL", r.Proto)
	host.request.ServerVars.Add("REQUEST_URI", requestURI)
	host.request.ServerVars.Add("PATH_INFO", r.URL.Path)
	host.request.ServerVars.Add("PATH_TRANSLATED", host.server.MapPath(r.URL.Path))
	host.request.ServerVars.Add("APPL_PHYSICAL_PATH", host.server.MapPath("/"))
	host.request.ServerVars.Add("REMOTE_ADDR", requestRemoteAddr(r.RemoteAddr))
	host.request.ServerVars.Add("REQUEST_METHOD", r.Method)
	host.request.ServerVars.Add("SERVER_NAME", hostName)
	host.request.ServerVars.Add("SERVER_PORT", port)
	host.request.ServerVars.Add("SCRIPT_NAME", r.URL.Path)
	host.request.ServerVars.Add("URL", r.URL.Path)
	host.request.ServerVars.Add("HTTP_USER_AGENT", r.UserAgent())
	host.request.ServerVars.Add("HTTP_ACCEPT_LANGUAGE", r.Header.Get("Accept-Language"))
	host.request.ServerVars.Add("CONTENT_LENGTH", strconv.FormatInt(host.request.TotalBytes(), 10))
	host.request.ServerVars.Add("CONTENT_TYPE", r.Header.Get("Content-Type"))
	allHTTP, allRaw := buildAggregateHeaderServerVariables(r.Header)
	host.request.ServerVars.Add("ALL_HTTP", allHTTP)
	host.request.ServerVars.Add("ALL_RAW", allRaw)

	for headerName, values := range r.Header {
		if len(values) == 0 {
			continue
		}
		host.request.ServerVars.Add(serverVariableFromHeader(headerName), strings.Join(values, ","))
	}

	host.setSessionCookie()

	if isNew && site.globalASA != nil && site.globalASA.IsLoaded() {
		site.globalASA.PopulateSessionStaticObjects(session)
		_ = site.globalASA.ExecuteSessionOnStart(host)
		if host.response.IsEnded() {
			host.response.ResetEnded()
		}
		host.response.Clear()
	}

	return host
}

func (h *CaddyWebHost) setSessionCookie() {
	if h.response == nil || h.response.Output == nil || h.session == nil {
		return
	}
	writer, ok := h.response.Output.(http.ResponseWriter)
	if !ok {
		return
	}
	replaceResponseCookie(writer, "ASPSESSIONID")
	http.SetCookie(writer, &http.Cookie{
		Name:     "ASPSESSIONID",
		Value:    h.session.ID,
		Path:     "/",
		HttpOnly: true,
		SameSite: http.SameSiteLaxMode,
	})
}

func (h *CaddyWebHost) PersistSession() {
	if h.session == nil || !h.sessionEnabled {
		return
	}
	if h.session.IsAbandoned() {
		_ = h.session.Delete()
		newSession, err := asp.CreateSession()
		if err == nil {
			h.session = newSession
			h.setSessionCookie()
		}
		return
	}
	h.session.QueueSaveIfDirty()
	h.setSessionCookie()
}

// Caddy log redirection helper
type zapLogWriter struct {
	logger *zap.Logger
}

func (z *zapLogWriter) Write(p []byte) (n int, err error) {
	msg := string(p)
	msg = strings.TrimSuffix(msg, "\n")
	z.logger.Info(msg)
	return len(p), nil
}

// Response and request helpers

func loadOrCreateSession(r *http.Request) (*asp.Session, bool) {
	var sessionID string
	if cookie, err := r.Cookie("ASPSESSIONID"); err == nil && cookie != nil {
		sessionID = cookie.Value
	}
	session, isNew, err := asp.GetOrCreateSession(sessionID)
	if err != nil {
		return asp.NewSession(), true
	}
	return session, isNew
}

func replaceResponseCookie(writer http.ResponseWriter, cookieName string) {
	if writer == nil {
		return
	}
	headers := writer.Header()
	if headers == nil {
		return
	}
	existing := headers.Values("Set-Cookie")
	if len(existing) == 0 {
		return
	}
	prefix := cookieName + "="
	filtered := make([]string, 0, len(existing))
	for _, value := range existing {
		if strings.HasPrefix(value, prefix) {
			continue
		}
		filtered = append(filtered, value)
	}
	headers.Del("Set-Cookie")
	for _, value := range filtered {
		headers.Add("Set-Cookie", value)
	}
}

func buildAggregateHeaderServerVariables(header http.Header) (string, string) {
	if len(header) == 0 {
		return "", ""
	}
	names := make([]string, 0, len(header))
	for name := range header {
		names = append(names, name)
	}
	sort.Strings(names)

	var allHTTP strings.Builder
	var allRaw strings.Builder
	for _, name := range names {
		values := header.Values(name)
		if len(values) == 0 {
			continue
		}
		joined := strings.Join(values, ",")
		if allHTTP.Len() > 0 {
			allHTTP.WriteString("\r\n")
		}
		allHTTP.WriteString(serverVariableFromHeader(name))
		allHTTP.WriteString(":")
		allHTTP.WriteString(joined)

		if allRaw.Len() > 0 {
			allRaw.WriteString("\r\n")
		}
		allRaw.WriteString(name)
		allRaw.WriteString(":")
		allRaw.WriteString(joined)
	}
	return allHTTP.String(), allRaw.String()
}

func requestRemoteAddr(remoteAddr string) string {
	host, _, err := net.SplitHostPort(remoteAddr)
	if err == nil && host != "" {
		return host
	}
	return remoteAddr
}

func requestServerName(r *http.Request) string {
	if r == nil {
		return ""
	}
	if host := r.URL.Hostname(); host != "" {
		return host
	}
	host, _, err := net.SplitHostPort(r.Host)
	if err == nil && host != "" {
		return host
	}
	return r.Host
}

func requestServerPort(r *http.Request) string {
	if r == nil {
		return ""
	}
	if port := r.URL.Port(); port != "" {
		return port
	}
	_, port, err := net.SplitHostPort(r.Host)
	if err == nil && port != "" {
		return port
	}
	if r.TLS != nil {
		return "443"
	}
	return "80"
}

func requestServerAddr(r *http.Request) string {
	if r == nil {
		return ""
	}
	if host := requestServerName(r); host != "" {
		return host
	}
	return requestRemoteAddr(r.RemoteAddr)
}

func requestAuthType(r *http.Request) string {
	if r == nil {
		return ""
	}
	authorization := strings.TrimSpace(r.Header.Get("Authorization"))
	if authorization == "" {
		return ""
	}
	if space := strings.IndexByte(authorization, ' '); space > 0 {
		return authorization[:space]
	}
	return authorization
}

func serverVariableFromHeader(headerName string) string {
	normalized := strings.ToUpper(strings.ReplaceAll(headerName, "-", "_"))
	return "HTTP_" + normalized
}

func relativePathToDoc(path, doc string) string {
	path = strings.TrimPrefix(path, "/")
	if path == "" {
		return doc
	}
	return filepath.Join(filepath.FromSlash(path), doc)
}

func caddyNativeTempDir() string {
	if dir := os.Getenv("CADDY_DATA_DIR"); dir != "" {
		return filepath.Join(dir, "temp")
	}
	if dir := os.Getenv("XDG_DATA_HOME"); dir != "" {
		return filepath.Join(dir, "caddy", "temp")
	}
	if userConfig, err := os.UserConfigDir(); err == nil && userConfig != "" {
		return filepath.Join(userConfig, "Caddy", "temp")
	}
	if home, err := os.UserHomeDir(); err == nil && home != "" {
		return filepath.Join(home, ".local", "share", "caddy", "temp")
	}
	return filepath.Join(os.TempDir(), "caddy", "temp")
}

func (a *AxonASP) resolveSiteTempDir(r *http.Request) string {
	siteKey := strings.TrimSpace(a.SiteName)
	if siteKey == "" && r != nil {
		siteKey = requestServerName(r)
	}
	if siteKey == "" && a.GlobalAsaPath != "" {
		siteKey = filepath.Base(filepath.Dir(a.GlobalAsaPath))
	}
	if siteKey == "" {
		siteKey = "default"
	}

	siteKey = sanitizeDirName(siteKey)
	return filepath.Join(caddyNativeTempDir(), "axonasp", siteKey)
}

func sanitizeDirName(name string) string {
	name = strings.TrimSpace(name)
	var sb strings.Builder
	for _, r := range name {
		if (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') || (r >= '0' && r <= '9') || r == '-' || r == '_' || r == '.' {
			sb.WriteRune(r)
		} else {
			sb.WriteRune('_')
		}
	}
	res := sb.String()
	if res == "" || res == "." || res == ".." {
		return "default"
	}
	return res
}

func (a *AxonASP) setupSiteTempDir(siteTemp string) {
	cacheDir := filepath.Join(siteTemp, "cache")
	sessionsDir := filepath.Join(siteTemp, "sessions")
	sessionDir := filepath.Join(siteTemp, "session")

	_ = os.MkdirAll(siteTemp, 0o755)
	_ = os.MkdirAll(cacheDir, 0o755)
	_ = os.MkdirAll(sessionsDir, 0o755)
	_ = os.MkdirAll(sessionDir, 0o755)

	axonconfig.NewViper().Set("global.temp_dir", siteTemp)
	if a.config != nil {
		a.config.Set("global.temp_dir", siteTemp)
	}
	asp.SetSessionStorageDir(sessionsDir)
}

func resolveDefaultASPPath(webRoot, requestPath string) (string, bool) {
	directoryPath := requestPath
	if directoryPath == "" {
		directoryPath = "/"
	}

	shouldResolve := directoryPath == "/" || strings.HasSuffix(directoryPath, "/")
	if !shouldResolve {
		relativePath := strings.TrimPrefix(directoryPath, "/")
		candidatePath := filepath.Clean(filepath.Join(webRoot, filepath.FromSlash(relativePath)))
		info, err := os.Stat(candidatePath)
		if err != nil || !info.IsDir() {
			return requestPath, false
		}
		directoryPath += "/"
		shouldResolve = true
	}

	if !shouldResolve {
		return requestPath, false
	}

	defaultDocs := []string{"index.asp", "default.asp", "home.asp", "main.asp"}
	for _, doc := range defaultDocs {
		testPath := filepath.Join(webRoot, relativePathToDoc(directoryPath, doc))
		if info, err := os.Stat(testPath); err == nil && !info.IsDir() {
			return strings.TrimSuffix(directoryPath, "/") + "/" + doc, true
		}
	}

	return requestPath, false
}

// singleHeaderResponseWriter prevents duplicate WriteHeader calls
type singleHeaderResponseWriter struct {
	http.ResponseWriter
	wroteHeader   bool
	defaultStatus int
}

func newSingleHeaderResponseWriter(w http.ResponseWriter, defaultStatus int) *singleHeaderResponseWriter {
	return &singleHeaderResponseWriter{ResponseWriter: w, defaultStatus: defaultStatus}
}

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

func (w *singleHeaderResponseWriter) Flush() {
	if flusher, ok := w.ResponseWriter.(http.Flusher); ok {
		flusher.Flush()
	}
}

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

type dummyResponseWriter struct{}

func (d *dummyResponseWriter) Header() http.Header         { return make(http.Header) }
func (d *dummyResponseWriter) Write(b []byte) (int, error) { return len(b), nil }
func (d *dummyResponseWriter) WriteHeader(statusCode int)  {}

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
	fmt.Fprintf(w, "<!doctype html><html><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>500 - Internal Server Error - AxonASP Server</title><style>body{margin:0;background:#f4f4f4;font-family:\"IBM Plex Sans\",Helvetica,sans-serif;color:#161616;font-size:13px}#h{height:60px;padding:0 15px;font-size:24px;display:flex;align-items:center;font-weight:600;border-bottom:1px solid #d9d9d9;background:#f4f4f4}.shell{padding:40px 20px;display:flex;justify-content:center}.card{background:#fff;border:1px solid #d9d9d9;max-width:760px;width:100%%;padding:28px;box-shadow:0 10px 20px rgba(22,22,22,.06)}h1{margin:0 0 16px;font-size:24px;border-bottom:1px solid #d9d9d9;padding-bottom:8px}p{margin:0 0 12px}table{width:100%%;border-collapse:collapse;border:1px solid #d9d9d9;margin:14px 0}td{border:1px solid #d9d9d9;padding:7px 10px;font-size:12px}td.k{width:120px;background:#f8f8f8;font-weight:600}.ft{margin-top:24px;border-top:1px solid #d9d9d9;padding-top:10px;font-size:11px;color:#525252}</style></head><body><div id=\"h\">❖ AxonASP Server</div><div class=\"shell\"><div class=\"card\"><h1>Application error</h1><p><b>%s error '%08X'</b></p><p>%s</p><table><tr><td class=\"k\">File</td><td>%s</td></tr><tr><td class=\"k\">Line</td><td>%d</td></tr><tr><td class=\"k\">Column</td><td>%d</td></tr><tr><td class=\"k\">Stage</td><td>%s</td></tr></table><div class=\"ft\">G3Pix ❖ AxonASP</div></div></div></body></html>", source, uint32(int32(err.Number)), description, fileName, err.Line, err.Column, html.EscapeString(stage))
}

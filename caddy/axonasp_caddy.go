package caddy

import (
	"fmt"
	"html"
	"io"
	"log"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
	"unsafe"

	"github.com/caddyserver/caddy/v2"
	"github.com/caddyserver/caddy/v2/caddyconfig/httpcaddyfile"
	"github.com/caddyserver/caddy/v2/modules/caddyhttp"
	"github.com/spf13/viper"
	"go.uber.org/zap"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
)

func init() {
	caddy.RegisterModule(AxonASP{})
	httpcaddyfile.RegisterHandlerDirective("axonasp", parseCaddyfile)
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

	vmPools *vmPoolManager

	resolvedConfigPath string
}

type vmPoolManager struct {
	mu    sync.Mutex
	pools map[string]unsafe.Pointer
}

type localVMProgramPool struct {
	mu          sync.Mutex
	items       []*axonvm.VM
	maxRetained int
	program     axonvm.CachedProgram
}

var (
	vmPooledFromOffset uintptr
	vmPooledSlotOffset uintptr
	vmPooledFromFound  bool
	vmPooledSlotFound  bool
	offsetOnce         sync.Once
)

func initOffsets(vmType reflect.Type) {
	offsetOnce.Do(func() {
		for f := range vmType.Fields() {
			f := f
			switch f.Name {
			case "pooledFrom":
				vmPooledFromOffset = f.Offset
				vmPooledFromFound = true
			case "pooledSlot":
				vmPooledSlotOffset = f.Offset
				vmPooledSlotFound = true
			}
		}
	})
}

func setVMPrivateFields(vm *axonvm.VM, pool unsafe.Pointer, slot chan struct{}) {
	initOffsets(reflect.TypeFor[axonvm.VM]())

	if vmPooledFromFound {
		ptr := unsafe.Pointer(uintptr(unsafe.Pointer(vm)) + vmPooledFromOffset)
		*(*unsafe.Pointer)(ptr) = pool
	}
	if vmPooledSlotFound {
		ptr := unsafe.Pointer(uintptr(unsafe.Pointer(vm)) + vmPooledSlotOffset)
		*(*chan struct{})(ptr) = slot
	}
}

func (a *AxonASP) AcquireVM(program axonvm.CachedProgram) *axonvm.VM {
	key := program.SourceName
	if key == "" {
		key = fmt.Sprintf("hash:%d", program.ProgramHash)
	}

	a.vmPools.mu.Lock()
	poolPtr, ok := a.vmPools.pools[key]
	if !ok {
		pool := &localVMProgramPool{
			maxRetained: 250,
			program:     program,
		}
		// Pre-warm the pool
		for range 5 {
			vm := axonvm.NewVMFromCachedProgram(program)
			setVMPrivateFields(vm, unsafe.Pointer(pool), nil)
			pool.items = append(pool.items, vm)
		}
		poolPtr = unsafe.Pointer(pool)
		a.vmPools.pools[key] = poolPtr
	}
	a.vmPools.mu.Unlock()

	pool := (*localVMProgramPool)(poolPtr)
	pool.mu.Lock()
	var vm *axonvm.VM
	if len(pool.items) > 0 {
		vm = pool.items[len(pool.items)-1]
		pool.items = pool.items[:len(pool.items)-1]
	}
	pool.mu.Unlock()

	if vm == nil {
		vm = axonvm.NewVMFromCachedProgram(program)
		setVMPrivateFields(vm, poolPtr, nil)
	}

	return vm
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

	// Redirect standard logger to Caddy zap logger
	log.SetOutput(&zapLogWriter{logger: a.logger})

	a.vmPools = &vmPoolManager{pools: make(map[string]unsafe.Pointer)}

	configPath, configErr := a.resolveConfigFilePath()
	if configErr == nil {
		a.resolvedConfigPath = configPath
		axonconfig.SetCustomConfigPath(configPath)
		// Force config loading with the path used by this Caddy module.
		active := axonconfig.NewViper()
		if normalizeErr := normalizeAxFunctionResourcePaths(active, configPath); normalizeErr != nil {
			a.logger.Warn("Failed to normalize axfunctions resource paths", zap.Error(normalizeErr), zap.String("config", configPath))
		}

		v := viper.New()
		v.SetConfigType("toml")
		v.SetConfigFile(configPath)
		if err := v.ReadInConfig(); err == nil {
			_ = normalizeAxFunctionResourcePaths(v, configPath)
			a.config = v
			a.logger.Info("Loaded AxonASP config for Caddy", zap.String("path", configPath))
		} else {
			a.logger.Warn("Failed to read config file", zap.Error(err), zap.String("path", configPath))
		}
	} else {
		a.logger.Warn("No axonasp.toml found for Caddy module", zap.Error(configErr))
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

	configDir := filepath.Dir(configPath)
	normalize := func(key string) {
		val := strings.TrimSpace(v.GetString(key))
		if val == "" || filepath.IsAbs(val) {
			return
		}
		v.Set(key, filepath.Clean(filepath.Join(configDir, val)))
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

	webRoot := "."
	if a.GlobalAsaPath != "" {
		webRoot = filepath.Dir(a.GlobalAsaPath)
	} else if docRoot, ok := caddyhttp.GetVar(r.Context(), "root").(string); ok && docRoot != "" {
		webRoot = docRoot
	}

	if path != "/" && !strings.HasSuffix(path, "/") {
		relativePath := strings.TrimPrefix(path, "/")
		dirPath := filepath.Clean(filepath.Join(webRoot, filepath.FromSlash(relativePath)))
		if info, err := os.Stat(dirPath); err == nil && info.IsDir() {
			redirectPath := path + "/"
			if r.URL.RawQuery != "" {
				redirectPath += "?" + r.URL.RawQuery
			}
			http.Redirect(w, r, redirectPath, http.StatusMovedPermanently)
			return nil
		}
	}

	if resolvedPath, ok := resolveDefaultASPPath(webRoot, path); ok {
		path = resolvedPath
	}

	relativePath := strings.TrimPrefix(path, "/")
	filePath := filepath.Join(webRoot, filepath.FromSlash(relativePath))
	cleanPath := filepath.Clean(filePath)

	if _, err := os.Stat(cleanPath); os.IsNotExist(err) {
		return next.ServeHTTP(w, r)
	}

	ext := strings.ToLower(filepath.Ext(cleanPath))
	if ext != ".asp" {
		return next.ServeHTTP(w, r)
	}

	w.Header().Set("Server", "AxonASP-Caddy")
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

	defaultDocs := []string{"default.asp", "index.asp"}
	for _, doc := range defaultDocs {
		testPath := filepath.Join(webRoot, relativePathToDoc(directoryPath, doc))
		if info, err := os.Stat(testPath); err == nil && !info.IsDir() {
			return strings.TrimSuffix(directoryPath, "/") + "/" + doc, true
		}
	}

	return strings.TrimSuffix(directoryPath, "/") + "/index.asp", true
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

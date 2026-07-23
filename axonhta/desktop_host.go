package main

import (
	"io"
	"net/http"

	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
)

type DesktopHost struct {
	response       *asp.Response
	request        *asp.Request
	server         *asp.Server
	session        *asp.Session
	application    *asp.Application
	sessionEnabled bool
	engineMode     axonvm.EngineMode
}

func NewDesktopHost(w http.ResponseWriter, r *http.Request, appDir string, sharedApp *asp.Application) *DesktopHost {
	host := &DesktopHost{
		response:       asp.NewResponse(w),
		request:        asp.NewRequest(),
		server:         asp.NewServer(),
		session:        asp.NewSession(),
		application:    sharedApp,
		sessionEnabled: true,
		engineMode:     axonvm.EngineModeDefault,
	}

	host.request.SetHTTPRequest(r)
	host.server.SetRootDir(appDir)
	host.server.SetRequestPath(r.URL.Path)
	// HTA apps are trusted desktop applications — allow FSO to access
	// paths outside the web root (e.g., user-selected music directories).
	host.server.SetUnrestrictedFS(true)

	if len(r.URL.RawQuery) > 0 {
		host.request.QueryString.SetLazyPayload([]byte(r.URL.RawQuery))
	}

	host.request.SetBodyLoader(func() ([]byte, error) {
		if r.Body == nil {
			return []byte{}, nil
		}
		return io.ReadAll(r.Body)
	})

	host.request.ServerVars.Add("REQUEST_METHOD", r.Method)
	host.request.ServerVars.Add("URL", r.URL.Path)
	host.request.ServerVars.Add("SCRIPT_NAME", r.URL.Path)
	host.request.ServerVars.Add("PATH_INFO", r.URL.Path)
	host.request.ServerVars.Add("QUERY_STRING", r.URL.RawQuery)
	host.request.ServerVars.Add("HTTP_USER_AGENT", r.UserAgent())
	host.request.ServerVars.Add("SERVER_NAME", "localhost")
	host.request.ServerVars.Add("SERVER_PORT", "0")
	host.request.ServerVars.Add("REMOTE_ADDR", "127.0.0.1")
	host.request.ServerVars.Add("SERVER_SOFTWARE", "AxonHTA")
	host.request.ServerVars.Add("APPL_PHYSICAL_PATH", appDir)
	host.request.ServerVars.Add("PATH_TRANSLATED", host.server.MapPath(r.URL.Path))

	return host
}

func (h *DesktopHost) Response() *asp.Response        { return h.response }
func (h *DesktopHost) Request() *asp.Request          { return h.request }
func (h *DesktopHost) Server() *asp.Server            { return h.server }
func (h *DesktopHost) Session() *asp.Session          { return h.session }
func (h *DesktopHost) Application() *asp.Application  { return h.application }
func (h *DesktopHost) SetSessionEnabled(enabled bool) { h.sessionEnabled = enabled }
func (h *DesktopHost) SessionEnabled() bool           { return h.sessionEnabled }
func (h *DesktopHost) EngineMode() axonvm.EngineMode  { return h.engineMode }

func (h *DesktopHost) Write(p []byte) (int, error) {
	h.response.Write(string(p))
	return len(p), nil
}

func (h *DesktopHost) WriteString(s string) {
	h.response.Write(s)
}

func (h *DesktopHost) ExecuteASPFile(absPath string) error {
	previousRequestPath := h.server.GetRequestPath()
	h.server.SetRequestPath(h.server.VirtualPathFromAbsolutePath(absPath))
	defer h.server.SetRequestPath(previousRequestPath)

	program, err := scriptCache.LoadOrCompileWithOptions(absPath, axonvm.ScriptCompileOptions{
		IncludeSiteRoot: h.server.MapPath("/"),
	})
	if err != nil {
		return err
	}

	vm := axonvm.AcquireVMFromCachedProgram(program)
	vm.SetHost(h)
	defer vm.Release()
	return vm.Run()
}

func (h *DesktopHost) PersistSession() {
	if h.session != nil && h.sessionEnabled {
		h.session.QueueSaveIfDirty()
	}
}

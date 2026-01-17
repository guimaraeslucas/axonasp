package asp

import (
	"fmt"
	"net/http"
	"net/url"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

// Global Application State
var AppState = &Application{
	store: sync.Map{},
}

// SessionStorage holds all sessions (simplified in-memory storage)
var SessionStorage = &SessionManager{
	sessions: sync.Map{},
}

type Application struct {
	store         sync.Map
	mutex         sync.Mutex
	GlobalASACode string // Stores raw code from global.asa
}

func (a *Application) Lock()   { a.mutex.Lock() }
func (a *Application) Unlock() { a.mutex.Unlock() }

// CORREÇÃO: Normalizar chaves para lowercase
func (a *Application) Set(key string, value interface{}) {
	a.store.Store(strings.ToLower(key), value)
}

func (a *Application) Get(key string) interface{} {
	val, ok := a.store.Load(strings.ToLower(key))
	if !ok {
		return nil
	}
	return val
}

// StaticObjects returns all keys in Application for enumeration
func (a *Application) StaticObjects() []string {
	var keys []string
	a.store.Range(func(k, v interface{}) bool {
		if keyStr, ok := k.(string); ok {
			keys = append(keys, keyStr)
		}
		return true
	})
	return keys
}

// ContentsCollection provides collection-like access for Session and Application
type ContentsCollection struct {
	ownerSet func(key string, value interface{})
	ownerGet func(key string) interface{}
	ownerDel func(key string)
}

func (c *ContentsCollection) Item(key string) interface{} {
	return c.ownerGet(key)
}

func (c *ContentsCollection) Set(key string, value interface{}) {
	c.ownerSet(key, value)
}

func (c *ContentsCollection) Remove(key string) {
	c.ownerDel(key)
}

func (c *ContentsCollection) RemoveAll() {
	// No direct access to underlying map here; caller should recreate map when needed.
	// We'll call Remove on keys until empty by assuming ownerGet returns nil when absent.
	// This is a best-effort for current in-memory implementation.
	// Not the most efficient but sufficient for tests.
	// Iterate via a temporary list of keys from ownerGet cannot be obtained; so implement specialized methods on Session/Application instead.
}

type SessionManager struct {
	sessions sync.Map
}

func (sm *SessionManager) Range(f func(key, value interface{}) bool) {
	sm.sessions.Range(f)
}

func (sm *SessionManager) Delete(key string) {
	sm.sessions.Delete(key)
}

type Session struct {
	ID           string
	store        sync.Map
	CreatedAt    time.Time
	LastAccessed time.Time
	Timeout      int // Minutes, default 20
	Abandoned    bool
	IsNew        bool // Flag to trigger Session_OnStart
	// Session-level properties
	CodePage int
	LCID     int
}

// CORREÇÃO: Normalizar chaves para lowercase na Session também
func (s *Session) Set(key string, value interface{}) {
	if s.Abandoned {
		return
	}
	s.store.Store(strings.ToLower(key), value)
}

func (s *Session) Get(key string) interface{} {
	if s.Abandoned {
		return nil
	}
	val, ok := s.store.Load(strings.ToLower(key))
	if !ok {
		return nil
	}
	return val
}

func (s *Session) Abandon() {
	s.Abandoned = true
	s.store = sync.Map{} // Clear data
}

func GetSession(w http.ResponseWriter, r *http.Request) *Session {
	cookie, err := r.Cookie("ASPSESSIONID")
	var sessionID string
	var session *Session
	now := time.Now()

	if err == nil && cookie.Value != "" {
		sessionID = cookie.Value
		if val, ok := SessionStorage.sessions.Load(sessionID); ok {
			s := val.(*Session)
			// Check Expiration
			limit := time.Duration(s.Timeout) * time.Minute
			if s.Abandoned || now.Sub(s.LastAccessed) > limit {
				// Expired or Abandoned
				SessionStorage.sessions.Delete(sessionID)
				sessionID = "" // Force new
			} else {
				session = s
				session.LastAccessed = now // Update access time
				session.IsNew = false      // Not new anymore
			}
		} else {
			sessionID = "" // Not found (server restart or invalid), force new
		}
	}

	if sessionID == "" || session == nil {
		// New Session
		sessionID = fmt.Sprintf("AxonASP_%d", now.UnixNano())
		session = &Session{
			ID:           sessionID,
			CreatedAt:    now,
			LastAccessed: now,
			Timeout:      20, // Default ASP timeout
			IsNew:        true,
		}
		SessionStorage.sessions.Store(sessionID, session)

		http.SetCookie(w, &http.Cookie{
			Name:     "ASPSESSIONID",
			Value:    sessionID,
			Path:     "/",
			HttpOnly: false,
			Secure:   false,
			MaxAge:   0,
			SameSite: http.SameSiteStrictMode,
		})
	}

	return session
}

// Contents and StaticObjects helpers for Session and Application
func (s *Session) Contents() *ContentsCollection {
	return &ContentsCollection{
		ownerSet: func(k string, v interface{}) { s.Set(k, v) },
		ownerGet: func(k string) interface{} { return s.Get(k) },
		ownerDel: func(k string) { s.store.Delete(strings.ToLower(k)) },
	}
}

func (a *Application) Contents() *ContentsCollection {
	return &ContentsCollection{
		ownerSet: func(k string, v interface{}) { a.Set(k, v) },
		ownerGet: func(k string) interface{} { return a.Get(k) },
		ownerDel: func(k string) { a.store.Delete(strings.ToLower(k)) },
	}
}

// RemoveAll implementations for Session and Application
func (s *Session) RemoveAllContents() {
	s.store = sync.Map{}
}

func (a *Application) RemoveAllContents() {
	a.store = sync.Map{}
}

// ResponseState tracks ASP Response properties
type ResponseState struct {
	Buffer          bool
	ContentType     string
	Status          string // e.g. "200 OK"
	Cookies         map[string]string
	Headers         map[string]string
	IsEnded         bool
	Expires         int       // Minutes
	ExpiresAbsolute time.Time // Absolute time
	CacheControl    string    // e.g., "Private", "Public"
	Charset         string
	PICS            string
}

// ASPError represents the VBScript Err object
type ASPError struct {
	Number      int
	Description string
	Source      string
	HelpFile    string
	HelpContext int
}

func (e *ASPError) Clear() {
	e.Number = 0
	e.Description = ""
	e.Source = ""
	e.HelpFile = ""
	e.HelpContext = 0
}

// Raise populates the Err object fields. Does not panic by itself.
func (e *ASPError) Raise(number int, source, description, helpFile string, helpContext int) {
	e.Number = number
	e.Source = source
	e.Description = description
	e.HelpFile = helpFile
	e.HelpContext = helpContext
}

// ObjectContext represents MTS/COM+ transaction context
// This is a stub implementation for compatibility
type ObjectContext struct {
	TransactionAborted  bool
	TransactionComplete bool
}

// SetComplete signals successful completion of transactional script
func (oc *ObjectContext) SetComplete() {
	oc.TransactionComplete = true
	oc.TransactionAborted = false
}

// SetAbort signals transaction should be aborted
func (oc *ObjectContext) SetAbort() {
	oc.TransactionAborted = true
	oc.TransactionComplete = false
}

// Context holds the state of a single request execution
type ExecutionContext struct {
	Response              http.ResponseWriter
	Request               *http.Request
	Session               *Session
	Application           *Application
	ObjectContext         *ObjectContext
	Variables             map[string]interface{}
	GlobalVariables       map[string]interface{} // Global Scope
	Constants             map[string]interface{} // Constants
	Output                []byte                 // Buffer for response
	ResponseState         *ResponseState
	RootDir               string
	ScriptTimeout         int // Seconds
	OnErrorResumeNext     bool
	OptionExplicitEnabled bool
	Err                   *ASPError

	// Class Support
	Engine            *Engine
	GlobalClasses     map[string]*ClassDef
	CurrentInstance   *ClassInstance
	CurrentMethodName string

	// Function call stack for preventing recursion
	ExecutingFunctions map[string]bool // Track which functions are currently executing
}

func NewExecutionContext(w http.ResponseWriter, r *http.Request, rootDir string) *ExecutionContext {
	// Parse Form Data immediately
	r.ParseForm()

	vars := make(map[string]interface{})

	return &ExecutionContext{
		Response:        w,
		Request:         r,
		Session:         GetSession(w, r),
		Application:     AppState,
		ObjectContext:   &ObjectContext{},
		Variables:       vars,
		GlobalVariables: vars,
		Constants:       make(map[string]interface{}),
		ResponseState: &ResponseState{
			Buffer:       true,
			ContentType:  "text/html",
			Status:       "200 OK",
			Cookies:      make(map[string]string),
			Headers:      make(map[string]string),
			CacheControl: "Private", // Default ASP
			Charset:      "utf-8",   // Default for this interpreter
		},
		RootDir:            rootDir,
		ScriptTimeout:      90, // Default
		Err:                &ASPError{},
		GlobalClasses:      make(map[string]*ClassDef),
		ExecutingFunctions: make(map[string]bool),
	}
}

// Collection represents a simple iterable collection (keys) for Request.Form/QueryString
type Collection struct {
	keys []string
	data map[string][]string
}

func NewCollectionFromValues(v url.Values) *Collection {
	c := &Collection{data: map[string][]string{}, keys: []string{}}
	if v == nil {
		return c
	}
	for k, vals := range v {
		c.data[k] = vals
		c.keys = append(c.keys, k)
	}
	return c
}

// Keys returns the list of keys in the collection
func (c *Collection) Keys() []string {
	return c.keys
}

// Items returns values as []string concatenated (first value per key)
func (c *Collection) Items() []string {
	var out []string
	for _, k := range c.keys {
		if vals, ok := c.data[k]; ok && len(vals) > 0 {
			out = append(out, vals[0])
		} else {
			out = append(out, "")
		}
	}
	return out
}

// Get returns the value for a given key (first element)
func (c *Collection) Get(key string) string {
	if c == nil {
		return ""
	}
	if vals, ok := c.data[key]; ok && len(vals) > 0 {
		return vals[0]
	}
	return ""
}

func (ctx *ExecutionContext) Write(str string) {
	// If Buffer is true, append to Output.
	// If false, write directly to Response?
	// For simplicity in this interpreter, we always write to Output buffer
	// and Flush it at the end or when Response.Flush is called.
	ctx.Output = append(ctx.Output, []byte(str)...)
}

func (ctx *ExecutionContext) BinaryWrite(data []byte) {
	ctx.Output = append(ctx.Output, data...)
}

func (ctx *ExecutionContext) AddHeader(name, value string) {
	ctx.ResponseState.Headers[name] = value
}

func (ctx *ExecutionContext) AppendToLog(str string) {
	fmt.Printf("[AxonASP Log] %s\n", str)
}

func (ctx *ExecutionContext) Flush() {
	if ctx.ResponseState.IsEnded {
		return
	}

	//Set Server Header with software name
	ctx.Response.Header().Set("Server", "AxonASP")

	// Apply headers
	for k, v := range ctx.ResponseState.Headers {
		ctx.Response.Header().Set(k, v)
	}

	// Apply ContentType and Charset
	ct := ctx.ResponseState.ContentType
	if ctx.ResponseState.Charset != "" {
		if !strings.Contains(strings.ToLower(ct), "charset") {
			ct += "; charset=" + ctx.ResponseState.Charset
		}
	}
	if ct != "" {
		ctx.Response.Header().Set("Content-Type", ct)
	}

	// Apply Cache Control
	if ctx.ResponseState.CacheControl != "" {
		ctx.Response.Header().Set("Cache-Control", ctx.ResponseState.CacheControl)
		// Pragma: no-cache logic
		if (strings.ToLower(ctx.ResponseState.CacheControl) == "no-cache" || strings.ToLower(ctx.ResponseState.CacheControl) == "private") && ctx.ResponseState.Expires <= 0 {
			ctx.Response.Header().Set("Pragma", "no-cache")
		}
	}

	// Apply PICS
	if ctx.ResponseState.PICS != "" {
		ctx.Response.Header().Set("PICS-Label", ctx.ResponseState.PICS)
	}

	// Apply Expires
	if ctx.ResponseState.Expires > 0 {
		// Expires is in minutes from now
		exp := time.Now().Add(time.Duration(ctx.ResponseState.Expires) * time.Minute)
		ctx.Response.Header().Set("Expires", exp.Format(http.TimeFormat))
	} else if ctx.ResponseState.Expires == 0 {
		ctx.Response.Header().Set("Expires", "0")
	}

	// Apply ExpiresAbsolute
	if !ctx.ResponseState.ExpiresAbsolute.IsZero() {
		ctx.Response.Header().Set("Expires", ctx.ResponseState.ExpiresAbsolute.Format(http.TimeFormat))
	}

	// Apply Cookies
	for k, v := range ctx.ResponseState.Cookies {
		http.SetCookie(ctx.Response, &http.Cookie{Name: k, Value: v, Path: "/"})
	}

	// Write Buffer
	if len(ctx.Output) > 0 {
		ctx.Response.Write(ctx.Output)
		ctx.Output = nil // Clear buffer after flush
	}
}

func (ctx *ExecutionContext) Clear() {
	ctx.Output = nil
}

func (ctx *ExecutionContext) End() {
	ctx.Flush()
	ctx.ResponseState.IsEnded = true
}

func (ctx *ExecutionContext) Redirect(url string) {
	ctx.Clear() // Clear buffer logic for Redirect?
	http.Redirect(ctx.Response, ctx.Request, url, http.StatusFound)
	ctx.ResponseState.IsEnded = true
}

// Server Helper Methods

func (ctx *ExecutionContext) Server_MapPath(path string) string {
	// Simple mapping relative to current working dir or www root
	// Assuming "./www" is root for now or ctx defined root
	rootDir := ctx.RootDir
	if rootDir == "" {
		rootDir = "./www"
	}
	fullPath := filepath.Join(rootDir, path)
	absPath, err := filepath.Abs(fullPath)
	if err != nil {
		return fullPath
	}
	return absPath
}

func (ctx *ExecutionContext) Server_URLEncode(str string) string {
	return url.QueryEscape(str)
}

func (ctx *ExecutionContext) Server_HTMLEncode(str string) string {
	// Simple HTML escape
	s := strings.ReplaceAll(str, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	return s
}

// Server_GetLastError returns the last error object for the current context
func (ctx *ExecutionContext) Server_GetLastError() *ASPError {
	return ctx.Err
}

func (ctx *ExecutionContext) IsClientConnected() bool {
	select {
	case <-ctx.Request.Context().Done():
		return false
	default:
		return true
	}
}

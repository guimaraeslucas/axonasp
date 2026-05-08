//go:build !lib_g3axonlive_disabled

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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
package axonvm

import (
	"encoding/json"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/vbscript"
	"github.com/spf13/viper"
)

// ---------------------------------------------------------------------------
// JSON response types — shared between the G3AXONLIVE VM object and the
// handler packages. Kept here so the VM can write the response directly.
// ---------------------------------------------------------------------------

// G3ALPatch carries the updated outer-HTML for one reactive component.
type G3ALPatch struct {
	ComponentID string `json:"componentId"`
	HTML        string `json:"html"`
}

// G3ALAction is a server-triggered client instruction embedded in the response.
// The "type" field determines what the client engine does.
//
// Supported types:
//   - "set_timer"        — fire eventName on componentId after delay ms
//   - "redirect"         — navigate the browser to url
//   - "trigger"          — immediately fire eventName on componentId
//   - "add_attribute"    — set attr on componentId element
//   - "set_property"     — assign DOM property (value, disabled, checked, etc.)
//   - "set_style"        — update element.style[name]
//   - "add_class"        — add CSS class to element.classList
//   - "remove_class"     — remove CSS class from element.classList
//   - "remove_attribute" — remove DOM attribute
//   - "add_title"        — add tooltip title
//   - "remove_title"     — remove tooltip title
//   - "set_value"        — set form field value
type G3ALAction struct {
	Type        string `json:"type"`
	ComponentID string `json:"componentId,omitempty"`
	EventName   string `json:"eventName,omitempty"`
	DelayMS     int    `json:"delay,omitempty"`
	URL         string `json:"url,omitempty"`
	AttrName    string `json:"name,omitempty"`
	AttrValue   string `json:"value,omitempty"`
}

// G3ALResponse is the complete JSON envelope sent to the browser.
type G3ALResponse struct {
	Success    bool         `json:"success"`
	Components []G3ALPatch  `json:"components,omitempty"`
	Actions    []G3ALAction `json:"actions,omitempty"`
	Error      string       `json:"error,omitempty"`
}

// g3alMaxBodyBytes is the maximum POST body size for a G3AxonLive fetch request.
const g3alMaxBodyBytes int64 = 256 * 1024

// g3alMaxPatchesPerResponse returns the maximum number of component patches
// that a single EndAsyncResponse call may include, read from viper config.
func g3alMaxPatchesPerResponse() int {
	limit := viper.GetInt("g3axonlive.max_components_per_response")
	if limit <= 0 {
		limit = 200
	}
	return limit
}

// ---------------------------------------------------------------------------
// Process-wide singleton — persists across all HTTP requests
// ---------------------------------------------------------------------------

// g3alComponentEntry holds the string value of one component property and its
// last-update timestamp. Using string storage avoids Value GC pressure.
type g3alComponentEntry struct {
	value     string
	updatedAt time.Time
}

// g3alStore is the process-wide singleton that holds component state and page
// registrations for every active G3AxonLive session. All fields are protected
// by a single RWMutex to keep the locking model simple and predictable.
type g3alStore struct {
	// componentValues maps "sessionID\x00componentID\x00propertyName_lower" -> entry.
	// The null-byte separator prevents collisions without nested maps.
	componentValues map[string]g3alComponentEntry
	// pageRegistry maps sessionID -> script URL (e.g. "/axonlive/counter.asp").
	pageRegistry map[string]string
	// lastAccess maps sessionID -> last access time, used by the cleanup goroutine.
	lastAccess map[string]time.Time
	mu         sync.RWMutex
}

var (
	g3alSingleton     *g3alStore
	g3alSingletonOnce sync.Once
	// g3alCleanupStop receives a signal when G3ALStopCleanup is called.
	g3alCleanupStop chan struct{}
)

// getG3ALStore returns the singleton, creating it on first call.
func getG3ALStore() *g3alStore {
	g3alSingletonOnce.Do(func() {
		g3alSingleton = &g3alStore{
			componentValues: make(map[string]g3alComponentEntry),
			pageRegistry:    make(map[string]string),
			lastAccess:      make(map[string]time.Time),
		}
	})
	return g3alSingleton
}

// ---------------------------------------------------------------------------
// Package-level exported helpers (called by the server/fastcgi entry points)
// ---------------------------------------------------------------------------

// G3ALRegisterPage stores the mapping from sessionID to the ASP script URL so
// the /g3al/ endpoint knows which file to re-execute when an event arrives.
func G3ALRegisterPage(sessionID, scriptURL string) {
	if sessionID == "" || scriptURL == "" {
		return
	}
	s := getG3ALStore()
	s.mu.Lock()
	s.pageRegistry[sessionID] = scriptURL
	s.lastAccess[sessionID] = time.Now()
	s.mu.Unlock()
}

// G3ALGetPageForSession returns the script URL registered for a session, or an
// empty string when the session is unknown or has expired.
func G3ALGetPageForSession(sessionID string) string {
	if sessionID == "" {
		return ""
	}
	s := getG3ALStore()
	s.mu.RLock()
	url := s.pageRegistry[sessionID]
	s.mu.RUnlock()
	return url
}

// G3ALStartCleanup begins a background goroutine that removes sessions which
// have been idle longer than the configured TTL. Call this once from server
// startup. Safe to call multiple times; subsequent calls are ignored.
// The idle TTL is derived from the viper config:
// 20× "global.default_script_timeout" seconds, with a floor of 30 minutes.
func G3ALStartCleanup(idleMinutes int) {
	// Prefer TTL from viper config; fall back to caller-supplied minutes.
	scriptTimeout := viper.GetInt("global.default_script_timeout")
	var idle time.Duration
	if scriptTimeout > 0 {
		idle = time.Duration(scriptTimeout*20) * time.Second
		if idle < 30*time.Minute {
			idle = 30 * time.Minute
		}
	} else {
		if idleMinutes <= 0 {
			idleMinutes = 30
		}
		idle = time.Duration(idleMinutes) * time.Minute
	}

	// Ensure singleton is initialized.
	getG3ALStore()

	// Use a package-level stop channel so only one goroutine ever runs.
	if g3alCleanupStop != nil {
		return // already running
	}
	g3alCleanupStop = make(chan struct{})

	// Read cleanup interval from config; fall back to 5 minutes.
	intervalMin := viper.GetInt("g3axonlive.g3axonlive_cleanup_interval_minutes")
	if intervalMin <= 0 {
		intervalMin = 5
	}
	ticker := time.NewTicker(time.Duration(intervalMin) * time.Minute)

	go func() {
		defer ticker.Stop()
		for {
			select {
			case <-ticker.C:
				g3alPerformCleanup(idle)
			case <-g3alCleanupStop:
				return
			}
		}
	}()
}

// Call this during server shutdown.
func G3ALStopCleanup() {
	if g3alCleanupStop != nil {
		close(g3alCleanupStop)
		g3alCleanupStop = nil
	}
}

// g3alPerformCleanup removes sessions that have been idle longer than the given
// duration. It is called periodically by the cleanup goroutine.
func g3alPerformCleanup(idle time.Duration) {
	s := getG3ALStore()
	now := time.Now()

	s.mu.Lock()
	defer s.mu.Unlock()

	// Collect sessions that have exceeded the idle limit.
	var expired []string
	for sessionID, t := range s.lastAccess {
		if now.Sub(t) > idle {
			expired = append(expired, sessionID)
		}
	}

	// Remove all state related to each expired session.
	for _, sessionID := range expired {
		prefix := sessionID + "\x00"
		for key := range s.componentValues {
			if strings.HasPrefix(key, prefix) {
				delete(s.componentValues, key)
			}
		}
		delete(s.pageRegistry, sessionID)
		delete(s.lastAccess, sessionID)
	}
}

// ---------------------------------------------------------------------------
// Component Proxy Object — Granular DOM manipulation
// ---------------------------------------------------------------------------

// G3ALComponentProxy is a native object that represents a specific reactive component.
// It allows ASP code to modify component properties, styles, and classes directly
// without re-rendering the entire HTML block.
type G3ALComponentProxy struct {
	parent      *G3AXONLIVE
	componentID string
}

// DispatchPropertyGet retrieves a property value from the persistent g3alStore.
func (p *G3ALComponentProxy) DispatchPropertyGet(propertyName string) Value {
	method := strings.ToLower(strings.TrimSpace(propertyName))

	// Determine the active session ID.
	sessionID := ""
	if p.parent.vm.host.Session() != nil {
		sessionID = p.parent.vm.host.Session().ID
	}
	if sessionID == "" && p.parent.eventSessionID != "" {
		sessionID = p.parent.eventSessionID
	}

	if sessionID == "" {
		return Value{Type: VTEmpty}
	}

	s := getG3ALStore()
	key := componentKey(sessionID, p.componentID, method)
	s.mu.RLock()
	entry, ok := s.componentValues[key]
	s.mu.RUnlock()

	if !ok {
		return Value{Type: VTEmpty}
	}

	// Simple type coercion: try to return bool for "true"/"false" strings.
	if strings.EqualFold(entry.value, "true") {
		return Value{Type: VTBool, Num: 1}
	}
	if strings.EqualFold(entry.value, "false") {
		return Value{Type: VTBool, Num: 0}
	}

	return NewString(entry.value)
}

// DispatchPropertySet queues a "set_property" action and updates the persistent store.
func (p *G3ALComponentProxy) DispatchPropertySet(propertyName string, args []Value) {
	if len(args) < 1 {
		return
	}
	method := strings.ToLower(strings.TrimSpace(propertyName))
	val := args[0]
	valStr := val.String()

	// 1. Queue the client-side mutation action.
	p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
		Type:        "set_property",
		ComponentID: p.componentID,
		AttrName:    method,
		AttrValue:   valStr,
	})

	// 2. Persistence: Update the global state so future reads reflect this change.
	sessionID := ""
	if p.parent.vm.host.Session() != nil {
		sessionID = p.parent.vm.host.Session().ID
	}
	if sessionID == "" && p.parent.eventSessionID != "" {
		sessionID = p.parent.eventSessionID
	}

	if sessionID != "" {
		s := getG3ALStore()
		key := componentKey(sessionID, p.componentID, method)
		s.mu.Lock()
		s.componentValues[key] = g3alComponentEntry{value: valStr, updatedAt: time.Now()}
		s.lastAccess[sessionID] = time.Now()
		s.mu.Unlock()
	}
}

// DispatchMethod implements granular manipulation methods (SetStyle, AddClass, etc.)
func (p *G3ALComponentProxy) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(strings.TrimSpace(methodName))
	switch method {
	case "setstyle":
		if len(args) >= 2 {
			p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
				Type:        "set_style",
				ComponentID: p.componentID,
				AttrName:    args[0].String(),
				AttrValue:   args[1].String(),
			})
		}
	case "addclass":
		if len(args) >= 1 {
			p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
				Type:        "add_class",
				ComponentID: p.componentID,
				AttrValue:   args[0].String(),
			})
		}
	case "removeclass":
		if len(args) >= 1 {
			p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
				Type:        "remove_class",
				ComponentID: p.componentID,
				AttrValue:   args[0].String(),
			})
		}
	case "setattribute":
		if len(args) >= 2 {
			p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
				Type:        "add_attribute",
				ComponentID: p.componentID,
				AttrName:    args[0].String(),
				AttrValue:   args[1].String(),
			})
		}
	case "removeattribute":
		if len(args) >= 1 {
			p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
				Type:        "remove_attribute",
				ComponentID: p.componentID,
				AttrName:    args[0].String(),
			})
		}
	case "addtitle":
		if len(args) >= 1 {
			p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
				Type:        "add_title",
				ComponentID: p.componentID,
				AttrValue:   args[0].String(),
			})
		}
	case "removetitle":
		p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
			Type:        "remove_title",
			ComponentID: p.componentID,
		})
	case "setvalue":
		if len(args) >= 1 {
			p.parent.pendingActions = append(p.parent.pendingActions, G3ALAction{
				Type:        "set_value",
				ComponentID: p.componentID,
				AttrValue:   args[0].String(),
			})
		}
	default:
		// Fallback for zero-arg property reads (VBScript compatibility).
		if len(args) == 0 {
			return p.DispatchPropertyGet(methodName)
		}
	}
	return Value{Type: VTEmpty}
}

// ---------------------------------------------------------------------------
// Per-request G3AXONLIVE struct — thin proxy to the singleton + request state
// ---------------------------------------------------------------------------

// G3AXONLIVE is the native library object exposed to ASP code via
// Server.CreateObject("G3AXONLIVE"). Each request gets a fresh instance.
// Persistent cross-request state lives in the process-wide g3alStore.
// Per-request event data (parsed from the incoming JSON body) lives here.
type G3AXONLIVE struct {
	vm *VM

	// Per-request async state, populated by InitPage().
	initiated        bool
	isAsyncRequest   bool
	eventSessionID   string
	eventComponentID string
	eventName        string
	eventArgs        map[string]string

	// Component HTML patches collected for EndAsyncResponse().
	pendingPatches []G3ALPatch

	// Server-triggered client actions (timer, redirect, trigger, addAttribute).
	pendingActions []G3ALAction

	// responseEnded guards against double-calls to EndAsyncResponse.
	responseEnded bool
}

// newG3AxonLiveObject creates a new per-request G3AXONLIVE proxy and registers it
// with the VM's native-object table.
func (vm *VM) newG3AxonLiveObject() Value {
	obj := &G3AXONLIVE{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3axonliveItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet routes property reads to the method dispatch table.
func (g *G3AXONLIVE) DispatchPropertyGet(propertyName string) Value {
	return g.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet is a no-op; all state is written via method calls.
func (g *G3AXONLIVE) DispatchPropertySet(_ string, _ []Value) {}

// DispatchMethod resolves all G3AXONLIVE methods and properties.
//
// Core lifecycle methods:
//   - InitPage()                      — parse async request body; register page
//   - IsAsyncRequest                  — boolean: was this a /g3al/ POST?
//   - EventComponentID                — string: which component fired?
//   - EventName                       — string: which event fired?
//   - EventArgs                       — JSON string of all event arguments
//   - GetEventArg(name)               — get a single event arg by name
//   - RegisterComponent(id, html)     — queue an HTML patch for the response
//   - EndAsyncResponse()              — flush JSON patches + halt execution
//
// Server-triggered client actions:
//   - SetTimer(cmpId, event, delayMs) — schedule client-side event dispatch
//   - Redirect(url)                   — navigate the browser to url
//   - Trigger(cmpId, event)           — immediately fire a client-side event
//   - AddAttribute(cmpId, name, val)  — set a DOM attribute on a component
//
// State management (backward-compatible):
//   - SetComponentProperty(sid, cid, prop, val)
//   - GetComponentProperty(sid, cid, prop)
//   - RemoveComponentProperty(sid, cid, prop)
//   - ClearComponentState(sid, cid)
//   - GetComponentState(sid, cid)
//   - RegisterPage(sid, url)
//   - RemoveSession(sid)
//   - StartCleanup / StopCleanup
//   - Version
func (g *G3AXONLIVE) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(strings.TrimSpace(methodName))
	switch method {
	// --- Core lifecycle ---
	case "initpage":
		return g.initPage()
	case "isasyncrequest":
		g.ensureAsyncStateFromRequest()
		if g.isAsyncRequest {
			return Value{Type: VTBool, Num: 1}
		}
		return Value{Type: VTBool, Num: 0}
	case "eventcomponentid":
		g.ensureAsyncStateFromRequest()
		return NewString(g.eventComponentID)
	case "eventname":
		g.ensureAsyncStateFromRequest()
		return NewString(g.eventName)
	case "eventargs":
		g.ensureAsyncStateFromRequest()
		return g.getEventArgsJSON()
	case "geteventarg":
		g.ensureAsyncStateFromRequest()
		return g.getEventArg(args)
	case "registercomponent":
		return g.registerComponent(args)
	case "getcomponent":
		return g.getComponent(args)
	case "endasyncresponse":
		return g.endAsyncResponse()

	// --- Server-triggered client actions ---
	case "settimer":
		return g.setTimer(args)
	case "redirect":
		return g.addRedirectAction(args)
	case "trigger":
		return g.addTriggerAction(args)
	case "addattribute":
		return g.addAttributeAction(args)

	// --- State management (backward-compatible) ---
	case "setcomponentproperty":
		return g.setComponentProperty(args)
	case "getcomponentproperty":
		return g.getComponentProperty(args)
	case "removecomponentproperty":
		return g.removeComponentProperty(args)
	case "clearcomponentstate":
		return g.clearComponentState(args)
	case "getcomponentstate":
		return g.getComponentState(args)
	case "registerpage":
		return g.registerPage(args)
	case "removesession":
		return g.removeSession(args)
	case "version":
		return NewString("2.0.0")
	case "startcleanup":
		G3ALStartCleanup(30)
		return Value{Type: VTEmpty}
	case "stopcleanup":
		G3ALStopCleanup()
		return Value{Type: VTEmpty}
	default:
		return Value{Type: VTEmpty}
	}
}

// ---------------------------------------------------------------------------
// Core lifecycle method implementations
// ---------------------------------------------------------------------------

// initPage parses the incoming request to determine whether this is an async
// G3AxonLive POST. On a normal page load it registers the session→script mapping
// for future async calls. On an async call it extracts the event payload.
// Returns True when IsAsyncRequest, False on a normal page load.
func (g *G3AXONLIVE) initPage() Value {
	if g.initiated {
		// Idempotent — calling InitPage more than once is a no-op.
		if g.isAsyncRequest {
			return Value{Type: VTBool, Num: 1}
		}
		return Value{Type: VTBool, Num: 0}
	}
	g.initiated = true

	req := g.vm.host.Request()
	sess := g.vm.host.Session()

	// Determine whether this is an async G3AxonLive request by inspecting
	// the HTTP_X_G3AXONLIVE server variable (mapped from X-G3AxonLive header).
	headerVal := req.ServerVars.Get("HTTP_X_G3AXONLIVE")
	methodVal := strings.TrimSpace(req.ServerVars.Get("REQUEST_METHOD"))
	contentTypeVal := strings.ToLower(strings.TrimSpace(req.ServerVars.Get("CONTENT_TYPE")))
	isJSONPost := strings.EqualFold(methodVal, "POST") && strings.Contains(contentTypeVal, "application/json")
	hasForwardedEvent := strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_SESSIONID")) != "" ||
		strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_COMPONENTID")) != "" ||
		strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_EVENTNAME")) != ""
	if !strings.EqualFold(strings.TrimSpace(headerVal), "true") && !hasForwardedEvent && !isJSONPost {
		// Regular full-page load — register this page for future async calls.
		if sess != nil && sess.ID != "" {
			scriptURL := req.ServerVars.Get("SCRIPT_NAME")
			G3ALRegisterPage(sess.ID, scriptURL)
		}
		g.isAsyncRequest = false
		return Value{Type: VTBool, Num: 0}
	}

	// Async request — read the JSON body (buffered by the /g3al/ handler).
	g.isAsyncRequest = true
	total := req.TotalBytes()
	if total <= 0 {
		total = g3alMaxBodyBytes
	}
	body := req.BinaryRead(total)
	if len(body) == 0 {
		g.loadAsyncEventFromForwardHeaders()
		return Value{Type: VTBool, Num: 1}
	}

	// Parse the JSON event payload.
	var event struct {
		SessionID   string            `json:"sessionId"`
		ComponentID string            `json:"componentId"`
		EventName   string            `json:"eventName"`
		EventArgs   map[string]string `json:"eventArgs"`
	}
	if err := json.Unmarshal(body, &event); err != nil {
		g.loadAsyncEventFromForwardHeaders()
		return Value{Type: VTBool, Num: 1}
	}

	g.eventSessionID = strings.TrimSpace(event.SessionID)
	g.eventComponentID = strings.TrimSpace(event.ComponentID)
	g.eventName = strings.TrimSpace(event.EventName)
	if event.EventArgs != nil {
		g.eventArgs = event.EventArgs
	} else {
		g.eventArgs = map[string]string{}
	}

	// Refresh the page registration so the session stays alive.
	if g.eventSessionID != "" {
		scriptURL := req.ServerVars.Get("SCRIPT_NAME")
		G3ALRegisterPage(g.eventSessionID, scriptURL)
	}

	return Value{Type: VTBool, Num: 1}
}

// loadAsyncEventFromForwardHeaders fills async event fields from handler-forwarded
// headers when BinaryRead is unavailable or JSON parsing fails.
func (g *G3AXONLIVE) loadAsyncEventFromForwardHeaders() {
	req := g.vm.host.Request()
	g.eventSessionID = strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_SESSIONID"))
	g.eventComponentID = strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_COMPONENTID"))
	g.eventName = strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_EVENTNAME"))

	argsJSON := strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_EVENTARGS"))
	if argsJSON == "" {
		g.eventArgs = map[string]string{}
	} else {
		var args map[string]string
		if err := json.Unmarshal([]byte(argsJSON), &args); err != nil || args == nil {
			g.eventArgs = map[string]string{}
		} else {
			g.eventArgs = args
		}
	}

	if g.eventSessionID != "" {
		scriptURL := req.ServerVars.Get("SCRIPT_NAME")
		G3ALRegisterPage(g.eventSessionID, scriptURL)
	}
}

// ensureAsyncStateFromRequest lazily recovers async request state from request
// headers when per-object state was not initialized or not retained.
func (g *G3AXONLIVE) ensureAsyncStateFromRequest() {
	if g.isAsyncRequest {
		return
	}
	req := g.vm.host.Request()
	headerVal := req.ServerVars.Get("HTTP_X_G3AXONLIVE")
	methodVal := strings.TrimSpace(req.ServerVars.Get("REQUEST_METHOD"))
	contentTypeVal := strings.ToLower(strings.TrimSpace(req.ServerVars.Get("CONTENT_TYPE")))
	isJSONPost := strings.EqualFold(methodVal, "POST") && strings.Contains(contentTypeVal, "application/json")
	hasForwardedEvent := strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_SESSIONID")) != "" ||
		strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_COMPONENTID")) != "" ||
		strings.TrimSpace(req.ServerVars.Get("HTTP_X_G3AXONLIVE_EVENTNAME")) != ""
	if strings.EqualFold(strings.TrimSpace(headerVal), "true") || hasForwardedEvent || isJSONPost {
		g.isAsyncRequest = true
		if g.eventSessionID == "" && g.eventComponentID == "" && g.eventName == "" {
			g.loadAsyncEventFromForwardHeaders()
		}
	}
}

// getEventArgsJSON returns the entire eventArgs map encoded as a JSON string.
// Useful for debugging or passing to G3JSON.Parse on the ASP side.
func (g *G3AXONLIVE) getEventArgsJSON() Value {
	if !g.isAsyncRequest || len(g.eventArgs) == 0 {
		return NewString("{}")
	}
	data, err := json.Marshal(g.eventArgs)
	if err != nil {
		return NewString("{}")
	}
	return NewString(string(data))
}

// getEventArg retrieves a single named event argument.
// Signature: GetEventArg(argName) -> string
func (g *G3AXONLIVE) getEventArg(args []Value) Value {
	if len(args) < 1 || !g.isAsyncRequest {
		return Value{Type: VTEmpty}
	}
	name := strings.TrimSpace(args[0].String())
	if name == "" {
		return Value{Type: VTEmpty}
	}
	if val, ok := g.eventArgs[name]; ok {
		return NewString(val)
	}
	return Value{Type: VTEmpty}
}

// registerComponent queues an HTML patch for inclusion in EndAsyncResponse.
// Signature: RegisterComponent(componentId, html)
func (g *G3AXONLIVE) registerComponent(args []Value) Value {
	if len(args) < 2 {
		return Value{Type: VTEmpty}
	}
	componentID := strings.TrimSpace(args[0].String())
	html := args[1].String()
	if componentID == "" {
		return Value{Type: VTEmpty}
	}

	// Enforce per-response component patch limit.
	limit := g3alMaxPatchesPerResponse()
	if len(g.pendingPatches) >= limit {
		g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ALComponentLimitExceeded, nil, AxonASPErrorMessages[ErrG3ALComponentLimitExceeded], "axonvm/lib_g3axonlive.go", 0).Error())
		return Value{Type: VTEmpty}
	}

	// Validate that the component ID contains only safe characters.
	if !g3alIsValidComponentID(componentID) {
		g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ALInvalidComponentID, nil, AxonASPErrorMessages[ErrG3ALInvalidComponentID], "axonvm/lib_g3axonlive.go", 0).Error())
		return Value{Type: VTEmpty}
	}

	g.pendingPatches = append(g.pendingPatches, G3ALPatch{
		ComponentID: componentID,
		HTML:        html,
	})
	return Value{Type: VTEmpty}
}

// getComponent returns a G3ALComponentProxy native object for granular DOM manipulation.
// Signature: GetComponent(componentId)
func (g *G3AXONLIVE) getComponent(args []Value) Value {
	if len(args) < 1 {
		return Value{Type: VTEmpty}
	}
	componentID := strings.TrimSpace(args[0].String())
	if componentID == "" {
		return Value{Type: VTEmpty}
	}

	// Validate that the component ID contains only safe characters.
	if !g3alIsValidComponentID(componentID) {
		g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ALInvalidComponentID, nil, AxonASPErrorMessages[ErrG3ALInvalidComponentID], "axonvm/lib_g3axonlive.go", 0).Error())
		return Value{Type: VTEmpty}
	}

	proxy := &G3ALComponentProxy{
		parent:      g,
		componentID: componentID,
	}
	id := g.vm.nextDynamicNativeID
	g.vm.nextDynamicNativeID++
	if g.vm.g3axonliveProxyItems == nil {
		g.vm.g3axonliveProxyItems = make(map[int64]*G3ALComponentProxy)
	}
	g.vm.g3axonliveProxyItems[id] = proxy
	return Value{Type: VTNativeObject, Num: id}
}

// endAsyncResponse serializes all pending patches and client actions into a
// JSON response, writes it with the correct headers, and calls Response.End()
// to halt ASP execution. Must be called at the end of async event handling.
// Signature: EndAsyncResponse()
func (g *G3AXONLIVE) endAsyncResponse() Value {
	if g.responseEnded {
		g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ALResponseAlreadyEnded, nil, AxonASPErrorMessages[ErrG3ALResponseAlreadyEnded], "axonvm/lib_g3axonlive.go", 0).Error())
		return Value{Type: VTEmpty}
	}
	g.responseEnded = true

	resp := g.vm.host.Response()

	envelope := G3ALResponse{
		Success:    true,
		Components: g.pendingPatches,
		Actions:    g.pendingActions,
	}
	data, err := json.Marshal(envelope)
	if err != nil {
		errEnvelope := G3ALResponse{Success: false, Error: "internal: failed to serialize response"}
		data, _ = json.Marshal(errEnvelope)
	}

	// Override content type, clear any buffered HTML, write JSON, and halt.
	resp.SetContentType("application/json; charset=utf-8")
	resp.Clear()
	resp.Write(string(data))
	resp.End() // panics with asp.ResponseEndSignal — stops execution cleanly
	return Value{Type: VTEmpty}
}

// ---------------------------------------------------------------------------
// Server-triggered client action helpers
// ---------------------------------------------------------------------------

// setTimer queues a "set_timer" action that instructs the browser to send the
// named event after delayMs milliseconds.
// Signature: SetTimer(componentId, eventName, delayMs)
func (g *G3AXONLIVE) setTimer(args []Value) Value {
	if len(args) < 3 {
		return Value{Type: VTEmpty}
	}
	componentID := strings.TrimSpace(args[0].String())
	evtName := strings.TrimSpace(args[1].String())
	delay := g.vm.asInt(args[2])

	if componentID == "" || evtName == "" {
		return Value{Type: VTEmpty}
	}
	if delay <= 0 {
		g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ALTimerDelayInvalid, nil, AxonASPErrorMessages[ErrG3ALTimerDelayInvalid], "axonvm/lib_g3axonlive.go", 0).Error())
		return Value{Type: VTEmpty}
	}

	g.pendingActions = append(g.pendingActions, G3ALAction{
		Type:        "set_timer",
		ComponentID: componentID,
		EventName:   evtName,
		DelayMS:     delay,
	})
	return Value{Type: VTEmpty}
}

// addRedirectAction queues a "redirect" action that navigates the browser.
// Signature: Redirect(url)
func (g *G3AXONLIVE) addRedirectAction(args []Value) Value {
	if len(args) < 1 {
		return Value{Type: VTEmpty}
	}
	url := strings.TrimSpace(args[0].String())
	if url == "" {
		return Value{Type: VTEmpty}
	}
	g.pendingActions = append(g.pendingActions, G3ALAction{
		Type: "redirect",
		URL:  url,
	})
	return Value{Type: VTEmpty}
}

// addTriggerAction queues a "trigger" action that immediately fires a
// client-side event without a user interaction.
// Signature: Trigger(componentId, eventName)
func (g *G3AXONLIVE) addTriggerAction(args []Value) Value {
	if len(args) < 2 {
		return Value{Type: VTEmpty}
	}
	componentID := strings.TrimSpace(args[0].String())
	evtName := strings.TrimSpace(args[1].String())
	if componentID == "" || evtName == "" {
		return Value{Type: VTEmpty}
	}
	g.pendingActions = append(g.pendingActions, G3ALAction{
		Type:        "trigger",
		ComponentID: componentID,
		EventName:   evtName,
	})
	return Value{Type: VTEmpty}
}

// addAttributeAction queues an "add_attribute" action that sets a DOM attribute
// on the specified component element in the browser.
// Signature: AddAttribute(componentId, attributeName, attributeValue)
func (g *G3AXONLIVE) addAttributeAction(args []Value) Value {
	if len(args) < 3 {
		return Value{Type: VTEmpty}
	}
	componentID := strings.TrimSpace(args[0].String())
	attrName := strings.TrimSpace(args[1].String())
	attrValue := args[2].String()
	if componentID == "" || attrName == "" {
		return Value{Type: VTEmpty}
	}
	g.pendingActions = append(g.pendingActions, G3ALAction{
		Type:        "add_attribute",
		ComponentID: componentID,
		AttrName:    attrName,
		AttrValue:   attrValue,
	})
	return Value{Type: VTEmpty}
}

// g3alIsValidComponentID returns true when id contains only characters safe
// for use as a DOM id attribute: alphanumeric, hyphen, underscore, dot, colon.
func g3alIsValidComponentID(id string) bool {
	if id == "" {
		return false
	}
	for _, c := range id {
		if !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') ||
			(c >= '0' && c <= '9') || c == '-' || c == '_' || c == '.' || c == ':') {
			return false
		}
	}
	return true
}

// ---------------------------------------------------------------------------
// Method implementations — all delegate to the process-wide g3alStore
// ---------------------------------------------------------------------------

// componentKey builds the flat map key for a component property.
// Format: "sessionID\x00componentID\x00propertyName_lower"
func componentKey(sessionID, componentID, propertyName string) string {
	return sessionID + "\x00" + componentID + "\x00" + strings.ToLower(propertyName)
}

// componentPrefix returns the key prefix for all properties of a session/component.
func componentPrefix(sessionID, componentID string) string {
	return sessionID + "\x00" + componentID + "\x00"
}

// sessionPrefix returns the key prefix for all properties belonging to a session.
func sessionPrefix(sessionID string) string {
	return sessionID + "\x00"
}

// setComponentProperty stores a property value in the global singleton.
// Signature: SetComponentProperty(sessionID, componentID, propertyName, propertyValue)
func (g *G3AXONLIVE) setComponentProperty(args []Value) Value {
	if len(args) < 4 {
		return Value{Type: VTEmpty}
	}
	sessionID := strings.TrimSpace(args[0].String())
	componentID := strings.TrimSpace(args[1].String())
	propertyName := strings.TrimSpace(args[2].String())
	if sessionID == "" || componentID == "" || propertyName == "" {
		return Value{Type: VTEmpty}
	}
	s := getG3ALStore()
	key := componentKey(sessionID, componentID, propertyName)
	s.mu.Lock()
	s.componentValues[key] = g3alComponentEntry{value: args[3].String(), updatedAt: time.Now()}
	s.lastAccess[sessionID] = time.Now()
	s.mu.Unlock()
	return Value{Type: VTEmpty}
}

// getComponentProperty retrieves a property value from the global singleton.
// Signature: GetComponentProperty(sessionID, componentID, propertyName)
// Returns the stored value as a string, or Empty if not found.
func (g *G3AXONLIVE) getComponentProperty(args []Value) Value {
	if len(args) < 3 {
		return Value{Type: VTEmpty}
	}
	sessionID := strings.TrimSpace(args[0].String())
	componentID := strings.TrimSpace(args[1].String())
	propertyName := strings.TrimSpace(args[2].String())
	if sessionID == "" || componentID == "" || propertyName == "" {
		return Value{Type: VTEmpty}
	}
	s := getG3ALStore()
	key := componentKey(sessionID, componentID, propertyName)
	s.mu.RLock()
	entry, ok := s.componentValues[key]
	s.mu.RUnlock()
	if !ok {
		return Value{Type: VTEmpty}
	}
	// Refresh last-access timestamp.
	s.mu.Lock()
	s.lastAccess[sessionID] = time.Now()
	s.mu.Unlock()
	return NewString(entry.value)
}

// removeComponentProperty deletes one property entry from the global singleton.
// Signature: RemoveComponentProperty(sessionID, componentID, propertyName)
func (g *G3AXONLIVE) removeComponentProperty(args []Value) Value {
	if len(args) < 3 {
		return Value{Type: VTEmpty}
	}
	sessionID := strings.TrimSpace(args[0].String())
	componentID := strings.TrimSpace(args[1].String())
	propertyName := strings.TrimSpace(args[2].String())
	if sessionID == "" || componentID == "" || propertyName == "" {
		return Value{Type: VTEmpty}
	}
	s := getG3ALStore()
	key := componentKey(sessionID, componentID, propertyName)
	s.mu.Lock()
	delete(s.componentValues, key)
	s.mu.Unlock()
	return Value{Type: VTEmpty}
}

// clearComponentState removes all property entries for a session/component pair.
// Signature: ClearComponentState(sessionID, componentID)
func (g *G3AXONLIVE) clearComponentState(args []Value) Value {
	if len(args) < 2 {
		return Value{Type: VTEmpty}
	}
	sessionID := strings.TrimSpace(args[0].String())
	componentID := strings.TrimSpace(args[1].String())
	if sessionID == "" || componentID == "" {
		return Value{Type: VTEmpty}
	}
	prefix := componentPrefix(sessionID, componentID)
	s := getG3ALStore()
	s.mu.Lock()
	for key := range s.componentValues {
		if strings.HasPrefix(key, prefix) {
			delete(s.componentValues, key)
		}
	}
	s.mu.Unlock()
	return Value{Type: VTEmpty}
}

// getComponentState returns a diagnostic string listing all stored properties.
// Signature: GetComponentState(sessionID, componentID)
func (g *G3AXONLIVE) getComponentState(args []Value) Value {
	if len(args) < 2 {
		return NewString("")
	}
	sessionID := strings.TrimSpace(args[0].String())
	componentID := strings.TrimSpace(args[1].String())
	if sessionID == "" || componentID == "" {
		return NewString("")
	}
	prefix := componentPrefix(sessionID, componentID)
	s := getG3ALStore()
	s.mu.RLock()
	var result strings.Builder
	result.WriteString("Component: ")
	result.WriteString(componentID)
	result.WriteString(" (Session: ")
	result.WriteString(sessionID)
	result.WriteString(")\n")
	for key, entry := range s.componentValues {
		if strings.HasPrefix(key, prefix) {
			prop := key[len(prefix):]
			result.WriteString("  ")
			result.WriteString(prop)
			result.WriteString(": ")
			result.WriteString(entry.value)
			result.WriteString(" (Updated: ")
			result.WriteString(entry.updatedAt.Format(time.RFC3339))
			result.WriteString(")\n")
		}
	}
	s.mu.RUnlock()
	return NewString(result.String())
}

// registerPage records the ASP script URL for a session so the /g3al/ endpoint
// can locate and re-execute the correct page when an async event arrives.
// Signature: RegisterPage(sessionID, scriptURL)
func (g *G3AXONLIVE) registerPage(args []Value) Value {
	if len(args) < 2 {
		return Value{Type: VTEmpty}
	}
	sessionID := strings.TrimSpace(args[0].String())
	scriptURL := strings.TrimSpace(args[1].String())
	G3ALRegisterPage(sessionID, scriptURL)
	return Value{Type: VTEmpty}
}

// removeSession deletes all state (component values, page registration, access time)
// for the given session. Called when a user's session ends or times out.
// Signature: RemoveSession(sessionID)
func (g *G3AXONLIVE) removeSession(args []Value) Value {
	if len(args) < 1 {
		return Value{Type: VTEmpty}
	}
	sessionID := strings.TrimSpace(args[0].String())
	if sessionID == "" {
		return Value{Type: VTEmpty}
	}
	prefix := sessionPrefix(sessionID)
	s := getG3ALStore()
	s.mu.Lock()
	for key := range s.componentValues {
		if strings.HasPrefix(key, prefix) {
			delete(s.componentValues, key)
		}
	}
	delete(s.pageRegistry, sessionID)
	delete(s.lastAccess, sessionID)
	s.mu.Unlock()
	return Value{Type: VTEmpty}
}

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Jeffrey He (@jeffreyheping)
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
	"bytes"
	"net/http"
	"strconv"
	"strings"
)

// appRuntimeJS is injected into every HTML response to provide a desktop-like
// experience in Chrome app mode: disable the browser context menu and
// drag-and-drop of page elements, and send periodic heartbeats so the
// server knows when the browser window is still open.
const appRuntimeJS = `<script>(function(){
document.addEventListener("contextmenu",function(e){e.preventDefault();return false;});
document.addEventListener("dragstart",function(e){e.preventDefault();return false;});
setInterval(function(){fetch('/__heartbeat__',{method:'HEAD',cache:'no-store'}).catch(function(){})},5000);
// Send a final heartbeat when the page is about to unload (form submission,
// redirect navigation, or window close). This resets the server-side timeout
// timer so the process is not killed during the navigation gap before the
// new page's JS starts sending heartbeats.
// sendBeacon is fire-and-forget, so it works in beforeunload/pagehide where
// async fetch would be cancelled.
window.addEventListener('pagehide',function(){navigator.sendBeacon('/__heartbeat__');});
window.addEventListener('beforeunload',function(){navigator.sendBeacon('/__heartbeat__');});
})();</script>`

// injectAppScript injects appRuntimeJS before </body>, </head>, or at the
// end of the HTML content. The search is case-insensitive.
func injectAppScript(html []byte) []byte {
	script := []byte(appRuntimeJS)
	lower := bytes.ToLower(html)

	if idx := bytes.Index(lower, []byte("</body>")); idx >= 0 {
		result := make([]byte, 0, len(html)+len(script))
		result = append(result, html[:idx]...)
		result = append(result, script...)
		result = append(result, html[idx:]...)
		return result
	}
	if idx := bytes.Index(lower, []byte("</head>")); idx >= 0 {
		result := make([]byte, 0, len(html)+len(script))
		result = append(result, html[:idx]...)
		result = append(result, script...)
		result = append(result, html[idx:]...)
		return result
	}
	result := make([]byte, 0, len(html)+len(script))
	result = append(result, html...)
	result = append(result, script...)
	return result
}

// htmlInjectWriter wraps http.ResponseWriter to buffer HTML responses and
// inject appRuntimeJS before sending to the client.
//
// Crucially, this type does NOT implement http.Flusher. The ASP Response
// object checks for http.Flusher and calls Flush() to push content to the
// client. By not implementing it, the ASP Response's Flush() still writes
// content to our buffer via Write(), but cannot force an early flush.
// All content stays buffered until FinalFlush() injects the script and
// writes everything to the client in one shot.
type htmlInjectWriter struct {
	http.ResponseWriter
	buf        bytes.Buffer
	isHTML     bool
	checked    bool
	statusCode int
	headerSent bool
}

// newHTMLInjectWriter creates a response writer that injects appRuntimeJS
// into HTML responses.
func newHTMLInjectWriter(w http.ResponseWriter) *htmlInjectWriter {
	return &htmlInjectWriter{
		ResponseWriter: w,
		statusCode:     http.StatusOK,
	}
}

// WriteHeader stores the status code without forwarding it. The header is
// sent later in FinalFlush, after the script has been injected and
// Content-Length updated.
func (w *htmlInjectWriter) WriteHeader(code int) {
	w.statusCode = code
}

// Write buffers HTML content for later injection, or passes through
// non-HTML content directly to the underlying writer.
func (w *htmlInjectWriter) Write(b []byte) (int, error) {
	if !w.checked {
		w.checked = true
		ct := w.Header().Get("Content-Type")
		if strings.Contains(ct, "text/html") {
			w.isHTML = true
		} else if ct == "" {
			detected := http.DetectContentType(b)
			if strings.HasPrefix(detected, "text/html") {
				w.isHTML = true
				w.Header().Set("Content-Type", detected)
			}
		}
	}
	if w.isHTML {
		return w.buf.Write(b)
	}
	// Non-HTML: pass through immediately.
	if !w.headerSent {
		w.ResponseWriter.WriteHeader(w.statusCode)
		w.headerSent = true
	}
	return w.ResponseWriter.Write(b)
}

// FinalFlush writes buffered HTML content with the injected script, or
// forwards non-HTML / empty responses. Must be called via defer after the
// handler completes.
func (w *htmlInjectWriter) FinalFlush() {
	if w.isHTML && w.buf.Len() > 0 {
		injected := injectAppScript(w.buf.Bytes())
		w.Header().Del("Accept-Ranges")
		w.Header().Set("Content-Length", strconv.Itoa(len(injected)))
		if !w.headerSent {
			w.ResponseWriter.WriteHeader(w.statusCode)
			w.headerSent = true
		}
		w.ResponseWriter.Write(injected)
		return
	}
	// Non-HTML buffered content (rare, but handle it).
	if w.buf.Len() > 0 {
		if !w.headerSent {
			w.ResponseWriter.WriteHeader(w.statusCode)
			w.headerSent = true
		}
		w.ResponseWriter.Write(w.buf.Bytes())
		return
	}
	// Empty body (e.g. redirect 302): still need to send headers.
	if !w.headerSent {
		w.ResponseWriter.WriteHeader(w.statusCode)
		w.headerSent = true
	}
}

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
package asp

import (
	"strings"
	"testing"

	"g3pix.com.br/axonasp/vbscript"
)

// TestServerEncoding verifies Server HTML and URL encoding helper methods.
func TestServerEncoding(t *testing.T) {
	server := NewServer()

	htmlEncoded := server.HTMLEncode("<b>&\"</b>")
	if htmlEncoded != "&lt;b&gt;&amp;&#34;&lt;/b&gt;" {
		t.Fatalf("unexpected HTMLEncode output: %s", htmlEncoded)
	}

	urlEncoded := server.URLEncode("a b&c")
	if urlEncoded != "a+b%26c" {
		t.Fatalf("unexpected URLEncode output: %s", urlEncoded)
	}

	pathEncoded := server.URLPathEncode("a b/c+d")
	if pathEncoded != "a%20b/c+d" {
		t.Fatalf("unexpected URLPathEncode output: %s", pathEncoded)
	}
}

// TestServerMapPath verifies absolute and relative path mapping behavior.
func TestServerMapPath(t *testing.T) {
	server := NewServer()
	server.SetRootDir("./www")
	server.SetRequestPath("/folder/page.asp")

	absolute := server.MapPath("/index.asp")
	if !strings.HasSuffix(strings.ReplaceAll(absolute, "\\", "/"), "/www/index.asp") {
		t.Fatalf("unexpected absolute map path: %s", absolute)
	}

	relative := server.MapPath("local.asp")
	if !strings.HasSuffix(strings.ReplaceAll(relative, "\\", "/"), "/www/folder/local.asp") {
		t.Fatalf("unexpected relative map path: %s", relative)
	}
}

// TestServerCreateObjectError verifies CreateObject unsupported behavior and last error tracking.
func TestServerCreateObjectError(t *testing.T) {
	server := NewServer()

	_, err := server.CreateObject("ADODB.Connection")
	if err == nil {
		t.Fatalf("expected CreateObject error")
	}

	lastError := server.GetLastError()
	if lastError.Number != vbscript.HRESULTFromVBScriptCode(vbscript.ActiveXCannotCreateObject) {
		t.Fatalf("expected error number 429, got %d", lastError.Number)
	}
	if lastError.ASPCode != int(vbscript.ActiveXCannotCreateObject) {
		t.Fatalf("expected ASP code 429, got %d", lastError.ASPCode)
	}
	if lastError.Source != "Server.CreateObject" {
		t.Fatalf("unexpected error source: %s", lastError.Source)
	}
}

// TestServerScriptTimeout verifies timeout validation and getter behavior.
func TestServerScriptTimeout(t *testing.T) {
	server := NewServer()

	if server.GetScriptTimeout() != 90 {
		t.Fatalf("expected default timeout 90, got %d", server.GetScriptTimeout())
	}

	if err := server.SetScriptTimeout(120); err != nil {
		t.Fatalf("unexpected timeout set error: %v", err)
	}
	if server.GetScriptTimeout() != 120 {
		t.Fatalf("expected timeout 120, got %d", server.GetScriptTimeout())
	}

	if err := server.SetScriptTimeout(0); err == nil {
		t.Fatalf("expected timeout validation error")
	}
}

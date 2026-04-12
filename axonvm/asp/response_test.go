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
	"bytes"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestResponseBufferedWrite verifies buffered output and Flush behavior.
func TestResponseBufferedWrite(t *testing.T) {
	var output bytes.Buffer
	response := NewResponse(&output)

	response.Write("Hello")
	if output.String() != "" {
		t.Fatalf("expected empty output before flush, got %q", output.String())
	}

	response.Flush()
	if output.String() != "Hello" {
		t.Fatalf("unexpected output after flush: %q", output.String())
	}
}

// TestResponseProperties verifies property setters and getters.
func TestResponseProperties(t *testing.T) {
	response := NewResponse(&bytes.Buffer{})

	response.SetBuffer(false)
	response.SetCacheControl("No-Cache")
	response.SetCharset("utf-8")
	response.SetCodePage(65001)
	response.SetContentType("text/plain")
	response.SetExpires(10)
	response.SetExpiresAbsoluteRaw("Wed, 21 Oct 2015 07:28:00 GMT")
	response.SetPICS("pics")
	response.SetStatus("201 Created")

	if response.GetBuffer() != false {
		t.Fatalf("unexpected buffer value")
	}
	if response.GetCacheControl() != "No-Cache" {
		t.Fatalf("unexpected cache control")
	}
	if response.GetCharset() != "utf-8" {
		t.Fatalf("unexpected charset")
	}
	if response.GetCodePage() != 65001 {
		t.Fatalf("unexpected code page")
	}
	if response.GetContentType() != "text/plain" {
		t.Fatalf("unexpected content type")
	}
	if response.GetExpires() != 10 {
		t.Fatalf("unexpected expires")
	}
	if response.GetExpiresAbsoluteRaw() != "Wed, 21 Oct 2015 07:28:00 GMT" {
		t.Fatalf("unexpected expires absolute")
	}
	if response.GetPICS() != "pics" {
		t.Fatalf("unexpected pics")
	}
	if response.GetStatus() != "201 Created" {
		t.Fatalf("unexpected status")
	}
}

// TestResponseCookies verifies cookie value and property operations.
func TestResponseCookies(t *testing.T) {
	response := NewResponse(&bytes.Buffer{})

	response.SetCookieValue("SessionId", "abc")
	response.SetCookieProperty("SessionId", "Domain", "example.com")
	response.SetCookieProperty("SessionId", "Path", "/app")
	response.SetCookieProperty("SessionId", "Secure", "True")

	if response.GetCookieValue("sessionid") != "abc" {
		t.Fatalf("unexpected cookie value")
	}
	if response.GetCookieProperty("sessionid", "Domain") != "example.com" {
		t.Fatalf("unexpected cookie domain")
	}
	if response.GetCookieProperty("sessionid", "Path") != "/app" {
		t.Fatalf("unexpected cookie path")
	}
	if response.GetCookieProperty("sessionid", "Secure") != "True" {
		t.Fatalf("unexpected cookie secure flag")
	}
}

// TestResponseAppendToLog verifies log entries persist in memory and runtime log file.
func TestResponseAppendToLog(t *testing.T) {
	cwd, err := os.Getwd()
	if err != nil {
		t.Fatalf("getwd failed: %v", err)
	}

	logPath := filepath.Join(cwd, "temp", "server.log")
	_ = os.Remove(logPath)

	response := NewResponse(&bytes.Buffer{})
	response.AppendToLog("response-test-log-line")

	data, err := os.ReadFile(logPath)
	if err != nil {
		t.Fatalf("expected runtime log file, got error: %v", err)
	}
	if !strings.Contains(string(data), "response-test-log-line") {
		t.Fatalf("expected log line in file, got %q", string(data))
	}
}

// TestResponseBufferLimit verifies buffered output raises a limit error before unbounded growth.
func TestResponseBufferLimit(t *testing.T) {
	response := NewResponse(&bytes.Buffer{})
	response.SetMaxBufferBytes(8)

	defer func() {
		recovered := recover()
		limitErr, ok := recovered.(*ResponseBufferLimitError)
		if !ok {
			t.Fatalf("expected ResponseBufferLimitError, got %#v", recovered)
		}
		if limitErr.LimitBytes != 8 {
			t.Fatalf("expected limit 8, got %d", limitErr.LimitBytes)
		}
	}()

	response.Write("123456789")
}

// TestResponseBinaryContentTypeOmitsCharset verifies binary content types are not rewritten with a charset parameter.
func TestResponseBinaryContentTypeOmitsCharset(t *testing.T) {
	recorder := httptest.NewRecorder()
	response := NewResponse(recorder)

	response.SetContentType("image/bmp")
	response.BinaryWrite([]byte{'B', 'M'})
	response.Flush()

	if got := recorder.Header().Get("Content-Type"); got != "image/bmp" {
		t.Fatalf("expected image/bmp content type, got %q", got)
	}
	if !bytes.Equal(recorder.Body.Bytes(), []byte{'B', 'M'}) {
		t.Fatalf("unexpected binary body: %v", recorder.Body.Bytes())
	}
}

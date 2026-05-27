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
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"testing"
)

// TestRequestLookupOrder verifies ASP default lookup order across request collections.
func TestRequestLookupOrder(t *testing.T) {
	req := NewRequest()
	req.QueryString.Add("k", "q")
	req.Form.Add("k", "f")
	req.Cookies.AddCookie("k", "c")

	value := req.GetValue("k")
	if value != "q" {
		t.Fatalf("expected QueryString precedence, got %q", value)
	}
}

// TestRequestCollectionAndCookieAttributes verifies collection indexing and cookie subkey access.
func TestRequestCollectionAndCookieAttributes(t *testing.T) {
	req := NewRequest()
	req.QueryString.AddValues("items", []string{"a", "b"})
	req.Cookies.AddCookie("profile", "name=Lucas&lang=en")

	if req.GetCollectionValue("QueryString", "items") != "a, b" {
		t.Fatalf("unexpected collection joined value")
	}
	if req.GetCollectionValue("QueryString", "1") != "a, b" {
		t.Fatalf("unexpected collection index resolution")
	}
	if req.GetCookieAttribute("profile", "name") != "Lucas" {
		t.Fatalf("unexpected cookie subkey value")
	}
}

// TestRequestBinaryRead verifies sequential binary read behavior and total bytes tracking.
func TestRequestBinaryRead(t *testing.T) {
	req := NewRequest()
	req.SetBody([]byte("abcdef"))

	if req.TotalBytes() != 6 {
		t.Fatalf("expected total bytes 6, got %d", req.TotalBytes())
	}

	part1 := req.BinaryRead(2)
	if string(part1) != "ab" {
		t.Fatalf("unexpected first read: %q", string(part1))
	}

	part2 := req.BinaryRead(10)
	if string(part2) != "cdef" {
		t.Fatalf("unexpected second read: %q", string(part2))
	}

	part3 := req.BinaryRead(1)
	if len(part3) != 0 {
		t.Fatalf("expected EOF empty read")
	}
}

// TestRequestCollectionLazyPayload verifies URL-encoded payload parsing stays lazy and map-free.
func TestRequestCollectionLazyPayload(t *testing.T) {
	collection := NewRequestCollection()
	collection.SetLazyPayload([]byte("Name=Ada&name=Lovelace&lang=en&empty=&encoded=hello+world&plus=%2B"))

	if got := collection.Get("NAME"); got != "Ada, Lovelace" {
		t.Fatalf("expected case-insensitive joined value, got %q", got)
	}
	if got := collection.Get("encoded"); got != "hello world" {
		t.Fatalf("expected plus decoding, got %q", got)
	}
	if got := collection.Get("plus"); got != "+" {
		t.Fatalf("expected percent decoding, got %q", got)
	}
	if got := collection.Get("empty"); got != "" {
		t.Fatalf("expected empty value, got %q", got)
	}

	if !collection.Exists("lang") {
		t.Fatalf("expected lang key to exist")
	}
	if collection.Count() != 5 {
		t.Fatalf("expected 5 unique keys, got %d", collection.Count())
	}

	if key := collection.Key(1); key != "name" {
		t.Fatalf("expected first key name, got %q", key)
	}
	if value := collection.GetByIndex(1); value != "Ada, Lovelace" {
		t.Fatalf("expected first value Ada, Lovelace, got %q", value)
	}

	if len(collection.data) != 0 || len(collection.keys) != 0 {
		t.Fatalf("lazy payload should not materialize eager map/slice structures")
	}
}

// TestRequestFormLazyLoadUsesBodyPayload verifies Form reads parse body payload lazily without map materialization.
func TestRequestFormLazyLoadUsesBodyPayload(t *testing.T) {
	req := NewRequest()
	httpReq := &http.Request{Header: make(http.Header)}
	httpReq.Header.Set("Content-Type", "application/x-www-form-urlencoded; charset=utf-8")
	req.SetHTTPRequest(httpReq)
	req.SetBodyLoader(func() ([]byte, error) {
		return []byte("city=Recife&city=Olinda&zip=50000-000"), nil
	})

	if got := req.Form.Get("city"); got != "Recife, Olinda" {
		t.Fatalf("expected city values, got %q", got)
	}
	if got := req.Form.Get("zip"); got != "50000-000" {
		t.Fatalf("expected zip value, got %q", got)
	}

	if len(req.Form.data) != 0 || len(req.Form.keys) != 0 {
		t.Fatalf("form lazy payload should not materialize eager map/slice structures")
	}
}

// TestRequestCollectionLazyPayloadUnicodeFold verifies Unicode case-insensitive key matching parity.
func TestRequestCollectionLazyPayloadUnicodeFold(t *testing.T) {
	collection := NewRequestCollection()
	collection.SetLazyPayload([]byte("%C3%84pfel=1&%C3%A4PFEL=2&STRA%C3%9FE=x&stra%C3%9Fe=y"))

	if got := collection.Get("äPFEL"); got != "1, 2" {
		t.Fatalf("expected Unicode-folded key merge for äpfel, got %q", got)
	}
	if got := collection.Get("straße"); got != "x, y" {
		t.Fatalf("expected Unicode-folded key merge for straße, got %q", got)
	}
	if collection.Count() != 2 {
		t.Fatalf("expected 2 unique Unicode-folded keys, got %d", collection.Count())
	}
}

// TestRequestFormURLEncodedMultiValue ensures repeated keys are preserved for classic multi-select posts.
func TestRequestFormURLEncodedMultiValue(t *testing.T) {
	req := NewRequest()
	httpReq := httptest.NewRequest(http.MethodPost, "http://example.test/form.asp", bytes.NewBufferString("movies=Action&movies=Comedy&name=Lucas"))
	httpReq.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	req.SetHTTPRequest(httpReq)

	value, ok := req.Form.GetValue("movies")
	if !ok {
		t.Fatalf("expected movies key to be present")
	}
	if value.Count() != 2 {
		t.Fatalf("expected 2 movies values, got %d", value.Count())
	}
	if value.Item("1") != "Action" || value.Item("2") != "Comedy" {
		t.Fatalf("unexpected movies values: %#v", value.Values)
	}
}

// TestRequestFormMultipartMultiValue ensures ParseForm/ParseMultipartForm path preserves repeated keys.
func TestRequestFormMultipartMultiValue(t *testing.T) {
	var body bytes.Buffer
	writer := multipart.NewWriter(&body)
	if err := writer.WriteField("movies", "Action"); err != nil {
		t.Fatalf("failed to write multipart field: %v", err)
	}
	if err := writer.WriteField("movies", "Comedy"); err != nil {
		t.Fatalf("failed to write multipart field: %v", err)
	}
	if err := writer.Close(); err != nil {
		t.Fatalf("failed to finalize multipart body: %v", err)
	}

	req := NewRequest()
	httpReq := httptest.NewRequest(http.MethodPost, "http://example.test/form.asp", &body)
	httpReq.Header.Set("Content-Type", writer.FormDataContentType())
	req.SetHTTPRequest(httpReq)

	value, ok := req.Form.GetValue("movies")
	if !ok {
		t.Fatalf("expected movies key to be present")
	}
	if value.Count() != 2 {
		t.Fatalf("expected 2 movies values, got %d", value.Count())
	}
	if value.Item("1") != "Action" || value.Item("2") != "Comedy" {
		t.Fatalf("unexpected movies values: %#v", value.Values)
	}
}

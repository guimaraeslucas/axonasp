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

import "testing"

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

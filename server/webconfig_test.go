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
package main

import (
	"net/http"
	"os"
	"path/filepath"
	"testing"
)

func TestWebConfigRewriteBackReference(t *testing.T) {
	root := t.TempDir()
	config := `<?xml version="1.0" encoding="UTF-8"?>
<configuration>
  <system.webServer>
    <rewrite>
      <rules>
        <rule name="Product" stopProcessing="true">
          <match url="^prod/(\d+)/(.*)$" ignoreCase="true" />
          <action type="Rewrite" url="/product.asp?id={R:1}&amp;slug={R:2}" appendQueryString="true" />
        </rule>
      </rules>
    </rewrite>
  </system.webServer>
</configuration>`
	if err := os.WriteFile(filepath.Join(root, "web.config"), []byte(config), 0o644); err != nil {
		t.Fatalf("write web.config: %v", err)
	}

	processor, err := NewWebConfigProcessor(root)
	if err != nil {
		t.Fatalf("create processor: %v", err)
	}

	result, ok := processor.Apply("/prod/42/blue-widget", "debug=1")
	if !ok {
		t.Fatalf("expected rewrite match")
	}
	if result.ActionType != "rewrite" {
		t.Fatalf("expected rewrite action, got %q", result.ActionType)
	}
	if result.Path != "/product.asp" {
		t.Fatalf("expected rewritten path /product.asp, got %q", result.Path)
	}
	if result.RawQuery != "id=42&slug=blue-widget&debug=1" {
		t.Fatalf("unexpected rewritten query: %q", result.RawQuery)
	}
}

func TestWebConfigHTTPRedirect(t *testing.T) {
	root := t.TempDir()
	config := `<?xml version="1.0" encoding="UTF-8"?>
<configuration>
  <system.webServer>
    <httpRedirect enabled="true" destination="https://example.com/base" exactDestination="false" childOnly="false" httpResponseStatus="Permanent" />
  </system.webServer>
</configuration>`
	if err := os.WriteFile(filepath.Join(root, "web.config"), []byte(config), 0o644); err != nil {
		t.Fatalf("write web.config: %v", err)
	}

	processor, err := NewWebConfigProcessor(root)
	if err != nil {
		t.Fatalf("create processor: %v", err)
	}

	result, ok := processor.Apply("/docs/index.asp", "a=1")
	if !ok {
		t.Fatalf("expected redirect match")
	}
	if result.ActionType != "redirect" {
		t.Fatalf("expected redirect action, got %q", result.ActionType)
	}
	if result.RedirectLocation != "https://example.com/base/docs/index.asp?a=1" {
		t.Fatalf("unexpected redirect target: %q", result.RedirectLocation)
	}
	if result.RedirectStatus != http.StatusMovedPermanently {
		t.Fatalf("expected status %d, got %d", http.StatusMovedPermanently, result.RedirectStatus)
	}
}

func TestWebConfigCustomErrors(t *testing.T) {
	root := t.TempDir()
	config := `<?xml version="1.0" encoding="UTF-8"?>
<configuration>
  <system.webServer>
    <httpErrors>
      <error statusCode="404" path="/custom_404.asp" responseMode="ExecuteURL" />
    </httpErrors>
  </system.webServer>
</configuration>`
	if err := os.WriteFile(filepath.Join(root, "web.config"), []byte(config), 0o644); err != nil {
		t.Fatalf("write web.config: %v", err)
	}

	processor, err := NewWebConfigProcessor(root)
	if err != nil {
		t.Fatalf("create processor: %v", err)
	}

	errConfig, ok := processor.GetCustomError(http.StatusNotFound)
	if !ok {
		t.Fatalf("expected 404 custom error mapping")
	}
	if errConfig.Path != "/custom_404.asp" {
		t.Fatalf("unexpected custom error path: %q", errConfig.Path)
	}
	if errConfig.ResponseMode != "ExecuteURL" {
		t.Fatalf("unexpected custom error mode: %q", errConfig.ResponseMode)
	}
}

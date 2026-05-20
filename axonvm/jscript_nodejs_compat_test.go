/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
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
	"strings"
	"testing"
)

func TestJScriptNodeCompat(t *testing.T) {
	tests := []struct {
		name string
		code string
		want string
	}{
		{
			"console.log is a function",
			"Response.Write(typeof console.log)",
			"function",
		},
		{
			"process properties",
			"Response.Write(typeof process.env === 'object' && typeof process.version === 'string' && typeof process.cwd === 'function')",
			"True",
		},
		{
			"setTimeout is a function",
			"Response.Write(typeof setTimeout)",
			"function",
		},
		{
			"__dirname and __filename are defined",
			"Response.Write(typeof __dirname !== 'undefined' && typeof __filename !== 'undefined')",
			"True",
		},
		{
			"os module functions",
			"const os = require('os'); Response.Write(typeof os.platform === 'function' && typeof os.cpus === 'function')",
			"True",
		},
		{
			"EventEmitter constructor",
			"const events = require('events'); const ee = new events.EventEmitter(); Response.Write(typeof ee.on === 'function')",
			"True",
		},
		{
			"http.createServer is a function",
			"const http = require('http'); Response.Write(typeof http.createServer === 'function')",
			"True",
		},
		{
			"Object stringification",
			"Response.Write(console)",
			"[object console]",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			// Enable Node compatibility for these tests
			// We need a way to mock viper or ensure it's enabled.
			// Since axonconfig.NewViper() reads from file/env, we might need a workaround.
			// For this test, I'll assume it's enabled if we use a mock config.

			// Actually, let's just run it. If it fails, I'll know why.
			out, err := runJScript2(t, jscriptSrc(tt.code))
			if err != nil {
				t.Fatalf("run failed: %v", err)
			}
			if !strings.Contains(out, tt.want) {
				t.Errorf("got %q, want %q", out, tt.want)
			}
		})
	}
}

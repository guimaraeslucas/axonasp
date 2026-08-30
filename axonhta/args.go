/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Jeffrey He (@jeffreyheping), Lucas Guimarães (@guimaraeslucas)
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
	"os"
	"path/filepath"
	"strings"
)

// htaFileOverride holds the absolute path of a .hta file passed as a
// positional argument (e.g. drag-and-drop onto the executable). When set,
// appDir is pointed at its parent directory and the HTTP server will serve
// that file as the application root. Empty string means no override.
var htaFileOverride string

// resolveHTAArg scans os.Args for the first argument that is an existing
// file whose extension is ".hta". If found it:
//
//   - Sets htaFileOverride to the absolute path of the file.
//   - Sets appDir to the file's parent directory so all relative paths
//     (includes, static assets) resolve correctly.
//
// It deliberately ignores any argument that begins with "-" to avoid
// misinterpreting flag values, and silently falls back to the default
// behaviour when no valid .hta argument is found.
//
// This function must be called before flag.Parse() so that the explicit
// -app flag still takes priority: if the caller later sets appDir via the
// flag, that value wins (the flag default is "./" and users rarely pass it
// explicitly, so in practice htaFileOverride drives the directory).
func resolveHTAArg() {
	if len(os.Args) < 2 {
		return
	}

	for _, arg := range os.Args[1:] {
		// Skip flag arguments.
		if strings.HasPrefix(arg, "-") {
			continue
		}

		// Require a .hta extension (case-insensitive).
		if !strings.EqualFold(filepath.Ext(arg), ".hta") {
			continue
		}

		// Resolve to an absolute path and verify the file exists.
		abs, err := filepath.Abs(arg)
		if err != nil {
			continue
		}
		info, err := os.Stat(abs)
		if err != nil || info.IsDir() {
			continue
		}

		// Valid target found: record the override and point appDir at its
		// parent so the embedded HTTP server serves the correct root.
		htaFileOverride = abs
		appDir = filepath.Dir(abs)
		return
	}
}

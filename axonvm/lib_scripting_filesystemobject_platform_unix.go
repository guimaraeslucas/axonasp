//go:build !windows

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
	"os"
	"path/filepath"
	"time"
)

// fsoResolveWindowsFileTimes keeps fallback timestamps on non-Windows builds.
func fsoResolveWindowsFileTimes(_ os.FileInfo, fallback time.Time) (time.Time, time.Time, time.Time) {
	return fallback, fallback, fallback
}

// fsoGetFileVersion returns a stable fallback version string on non-Windows systems.
func fsoGetFileVersion(_ string) string {
	return "1.0.0.0"
}

// fsoGetShortPath returns the normalized path when no short-name API exists.
func fsoGetShortPath(path string) string {
	if absPath, err := filepath.Abs(path); err == nil {
		return absPath
	}
	return path
}

// fsoGetDriveSerialNumber returns a deterministic fallback serial string on non-Windows systems.
func fsoGetDriveSerialNumber(_ string) string {
	return "00001"
}

// fsoGetDriveShareName returns a deterministic fallback share name on non-Windows systems.
func fsoGetDriveShareName(_ string) string {
	return "\\AxonASP"
}

// fsoGetDriveVolumeName returns a deterministic fallback volume name on non-Windows systems.
func fsoGetDriveVolumeName(_ string) string {
	return "AxonASPVolume"
}

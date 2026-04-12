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
	"golang.org/x/sys/windows"
	"path/filepath"
	"unsafe"
)

func getFileOwnerName(path string) string {
	absPath, err := filepath.Abs(path)
	if err != nil {
		return ""
	}

	sd, err := windows.GetNamedSecurityInfo(
		absPath,
		windows.SE_FILE_OBJECT,
		windows.OWNER_SECURITY_INFORMATION,
	)
	if err != nil {
		return ""
	}
	defer windows.LocalFree(windows.Handle(uintptr(unsafe.Pointer(sd))))

	owner, _, err := sd.Owner()
	if err != nil {
		return ""
	}

	account, domain, _, err := owner.LookupAccount("")
	if err != nil {
		return ""
	}

	if domain != "" {
		return domain + "\\" + account
	}
	return account
}

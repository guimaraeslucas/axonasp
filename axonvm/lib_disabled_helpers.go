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
	"fmt"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// panicLibraryDisabled raises a consistent runtime error for disabled libraries.
func panicLibraryDisabled(libraryName string, source string) {
	panic(&VMError{
		Code:        vbscript.ActiveXCannotCreateObject,
		Number:      int(ErrLibraryDisabled),
		Description: fmt.Sprintf(ErrLibraryDisabled.String(), libraryName),
		Source:      source,
	})
}

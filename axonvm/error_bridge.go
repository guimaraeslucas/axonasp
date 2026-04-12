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
	"errors"

	"g3pix.com.br/axonasp/axonvm/asp"
	"g3pix.com.br/axonasp/vbscript"
)

// CompilerErrorToASPError converts compiler failures into the ASPError object model.
func CompilerErrorToASPError(err error, file string) *asp.ASPError {
	if err == nil {
		return asp.NewASPError()
	}

	var syntaxErr *vbscript.VBSyntaxError
	if errors.As(err, &syntaxErr) {
		syntaxErr.WithFile(file)
		return asp.NewASPErrorFromVBSyntaxError(syntaxErr)
	}

	return asp.NewASPErrorFromMessage("ASP", "AxonASP compilation error", err.Error(), file, 0, 0)
}

// RuntimeErrorToASPError converts VM runtime failures into the ASPError object model.
func RuntimeErrorToASPError(err error, file string) *asp.ASPError {
	if err == nil {
		return asp.NewASPError()
	}

	var vmErr *VMError
	if errors.As(err, &vmErr) {
		return vmErr.WithFile(file).ToASPError()
	}

	return asp.NewASPErrorFromMessage("ASP", "AxonASP runtime error", err.Error(), file, 0, 0)
}

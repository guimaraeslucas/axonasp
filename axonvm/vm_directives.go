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

	"g3pix.com.br/axonasp/vbscript"
)

// applyDirective applies one compiled ASP page directive to the current VM host state.
func (vm *VM) applyDirective(name string, value string) {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "language":
		normalized := strings.ToLower(strings.TrimSpace(value))
		if normalized != "vbscript" && normalized != "jscript" && normalized != "javascript" {
			vm.raise(vbscript.InternalError, "Only VBScript or JScript is supported in ASP directives")
		}
	case "codepage":
		vm.host.Response().SetCodePage(vm.asInt(NewString(value)))
	case "enablesessionstate":
		enabled := directiveEnablesSessionState(value)
		vm.host.SetSessionEnabled(enabled)
	}
}

// directiveEnablesSessionState converts an EnableSessionState directive value into a boolean flag.
func directiveEnablesSessionState(value string) bool {
	return !strings.EqualFold(strings.TrimSpace(value), "false")
}

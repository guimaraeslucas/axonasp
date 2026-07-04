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
package jscript

// ValidateProxyHasMissingPropertyInvariant enforces that [[HasProperty]] cannot
// report an own property as missing when the property is non-configurable, or
// when the target is non-extensible.
func ValidateProxyHasMissingPropertyInvariant(
	targetHasOwnProperty bool,
	targetPropertyConfigurable bool,
	targetExtensible bool,
	trapResult bool,
) (JSSyntaxErrorCode, bool) {
	if trapResult {
		return 0, false
	}
	if targetHasOwnProperty && (!targetPropertyConfigurable || !targetExtensible) {
		return ProxyHasTrapInvariantViolation, true
	}
	return 0, false
}

// ValidateProxyGetPrototypeOfInvariant enforces that [[GetPrototypeOf]] cannot
// return a different prototype for a non-extensible target.
func ValidateProxyGetPrototypeOfInvariant(
	targetExtensible bool,
	prototypeMatchesTarget bool,
) (JSSyntaxErrorCode, bool) {
	if !targetExtensible && !prototypeMatchesTarget {
		return ProxyGetPrototypeOfTrapInvariantViolation, true
	}
	return 0, false
}

//go:build windows

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

import "github.com/go-ole/go-ole"

// adodbDecodeGetRowsVariant decodes an ADODB.GetRows VARIANT SAFEARRAY payload into
// a field-major flattened list where index = field + row*fieldCount.
func adodbDecodeGetRowsVariant(rowsRes *ole.VARIANT, fieldCount int) ([]interface{}, int, bool) {
	// Disabled on Windows for safety: direct oleaut32 SafeArray probing via syscall
	// can trigger native heap corruption with some provider payloads.
	// The ADODB code path will automatically fall back to safe field-walk decoding.
	_ = rowsRes
	_ = fieldCount
	return nil, 0, false
}

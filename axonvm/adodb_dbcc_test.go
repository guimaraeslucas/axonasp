package axonvm

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Yuri Eisman (@yeisman)
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

import "testing"

// DBCC USEROPTIONS (and other DBCC commands) return result sets on SQL Server,
// so conn.Execute must hand back a Recordset, not Nothing.
func TestADODBIsQueryDBCC(t *testing.T) {
	vm := NewVM(nil, nil, 1)
	for _, sql := range []string{"DBCC USEROPTIONS", "dbcc useroptions", "  DBCC SHOW_STATISTICS ('t', 'ix')", "/* x */ DBCC USEROPTIONS"} {
		if !vm.adodbIsQuery(sql) {
			t.Errorf("adodbIsQuery(%q) = false, want true", sql)
		}
	}
	for _, sql := range []string{"dbccx", "UPDATE dbcc SET a = 1"} {
		if vm.adodbIsQuery(sql) {
			t.Errorf("adodbIsQuery(%q) = true, want false", sql)
		}
	}
}

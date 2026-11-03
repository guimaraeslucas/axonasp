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
 * made available under the same license terms.
 */

import "testing"

// TestADODBSharedDBPoolsSQLServerOnly checks that SQL Server connection strings share one
// process-wide pool per DSN, and that SQLite never does.
func TestADODBSharedDBPoolsSQLServerOnly(t *testing.T) {
	a, err := adodbSharedDB("mssql", "server=pool-test-a;user id=u;password=p")
	if err != nil {
		t.Fatal(err)
	}
	b, _ := adodbSharedDB("mssql", "server=pool-test-a;user id=u;password=p")
	c, _ := adodbSharedDB("mssql", "server=pool-test-b;user id=u;password=p")
	if a != b {
		t.Fatal("same DSN must reuse the same pool")
	}
	if a == c {
		t.Fatal("different DSNs must not share a pool")
	}
	if !adodbPooledDriver("mssql") || adodbPooledDriver("sqlite") || adodbPooledDriver("mysql") {
		t.Fatal("only mssql may be pooled")
	}
}

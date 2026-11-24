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

// TestADODBIsQueryBatchEndingInSelect checks that a statement batch whose later
// statement returns rows (the usual "SET NOCOUNT ON; INSERT ...; SELECT
// SCOPE_IDENTITY()" idiom) is executed as a query, so Execute returns a recordset.
func TestADODBIsQueryBatchEndingInSelect(t *testing.T) {
	vm := NewVM(nil, nil, 1)
	queries := []string{
		"SET NOCOUNT ON; INSERT INTO t(a) VALUES(1); SELECT CAST(SCOPE_IDENTITY() AS int) AS id;",
		"set nocount on; insert into t(a) values('x;y'); select @@identity as id",
		"SET NOCOUNT ON;\r\nUPDATE t SET a = 1 WHERE id = 2\r\n; SELECT a FROM t WHERE id = 2",
		"DECLARE @n int; SET @n = 1; SELECT @n AS n",
		"SET NOCOUNT ON; EXEC dbo.proc 1",
		"SET NOCOUNT ON; INSERT INTO t(a) OUTPUT Inserted.id VALUES(1)",
	}
	for _, sql := range queries {
		if !vm.adodbIsQuery(sql) {
			t.Errorf("adodbIsQuery(%q) = false, want true", sql)
		}
	}
	notQueries := []string{
		"SET NOCOUNT ON; DELETE FROM t WHERE id = 1",
		"SET NOCOUNT ON; INSERT INTO t(a) VALUES('; select 1')",
		"UPDATE t SET a = 1 -- ; select 1",
		"INSERT INTO [t;select] (a) VALUES (1); /* ; select 1 */",
		"SET NOCOUNT ON;",
	}
	for _, sql := range notQueries {
		if vm.adodbIsQuery(sql) {
			t.Errorf("adodbIsQuery(%q) = true, want false", sql)
		}
	}
}

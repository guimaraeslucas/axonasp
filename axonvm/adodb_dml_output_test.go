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

// SQL Server DML with an OUTPUT clause returns rows, so conn.Execute must hand
// back a Recordset (INSERT INTO t (...) OUTPUT Inserted.id VALUES (...)).
func TestADODBIsQueryDMLOutputClause(t *testing.T) {
	vm := NewVM(nil, nil, 1)
	for _, sql := range []string{
		"Insert into orders(a,b) Output Inserted.orderID,Inserted.x values('a',1)",
		"UPDATE t SET a = 1 OUTPUT deleted.a, inserted.a WHERE id = 2",
		"delete from t output deleted.id where id = 3",
		"MERGE t USING s ON t.id = s.id WHEN MATCHED THEN DELETE OUTPUT $action, deleted.id;",
		"INSERT INTO t (a) VALUES (1) RETURNING id",
	} {
		if !vm.adodbIsQuery(sql) {
			t.Errorf("adodbIsQuery(%q) = false, want true", sql)
		}
	}
	for _, sql := range []string{
		"INSERT INTO t (a) VALUES ('no output here')",
		"UPDATE t SET outputs = 1, [output] = 2 WHERE note = 'output' -- output",
		"DELETE FROM output_log WHERE id = 1 /* OUTPUT */",
		"MERGE t USING s ON t.id = s.id WHEN MATCHED THEN DELETE;",
	} {
		if vm.adodbIsQuery(sql) {
			t.Errorf("adodbIsQuery(%q) = true, want false", sql)
		}
	}
}

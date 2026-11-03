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

func TestADODBMSSQLNumeric(t *testing.T) {
	if got := adodbMSSQLNumeric("DECIMAL", []byte("1.5000")); got != 1.5 {
		t.Fatalf("DECIMAL: got %v", got)
	}
	if got := adodbMSSQLNumeric("MONEY", []byte("-12.3400")); got != -12.34 {
		t.Fatalf("MONEY: got %v", got)
	}
	if got, ok := adodbMSSQLNumeric("VARCHAR", []byte("1.5")).([]byte); !ok || string(got) != "1.5" {
		t.Fatalf("VARCHAR must be untouched, got %v", got)
	}
	if got := adodbMSSQLNumeric("DECIMAL", nil); got != nil {
		t.Fatalf("NULL must stay nil, got %v", got)
	}
}

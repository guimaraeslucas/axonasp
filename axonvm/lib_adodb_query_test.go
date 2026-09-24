//go:build !wasm && !lib_adodb_disabled

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

import "testing"

func TestADODBIsQueryRowReturningStatements(t *testing.T) {
	vm := NewVM(nil, nil, 1)

	queries := []string{
		"SELECT * FROM users",
		"select id from users",
		"  \t\r\nSELECT 1",
		"SELECT*FROM users",
		"SHOW TABLES",
		"  show tables  ",
		"PRAGMA table_info(users)",
		"pragma foreign_keys = ON",
		// CTEs.
		"WITH recent AS (SELECT id FROM users) SELECT * FROM recent",
		"with recursive t(n) as (select 1) select * from t",
		"\n\tWITH cte AS (SELECT 1 AS v) SELECT v FROM cte",
		// DML with RETURNING.
		"INSERT INTO t (a) VALUES (1) RETURNING id",
		"insert into t (a) values (1) returning id",
		"UPDATE t SET a = 1 RETURNING *",
		"  update t set a = 1 returning a  ",
		"DELETE FROM t RETURNING id",
		"delete from t where a = 1 returning *",
		// Leading comments / grouping parentheses.
		"/* hint */ SELECT 1",
		"-- comment\nSELECT 1",
		"(SELECT 1)",
		"(  /* x */ WITH cte AS (SELECT 1) SELECT * FROM cte )",
		// RETURNING token preceded by another clause.
		"INSERT INTO t (a) SELECT a FROM s RETURNING id",
	}

	for _, sql := range queries {
		if !vm.adodbIsQuery(sql) {
			t.Errorf("adodbIsQuery(%q) = false, want true", sql)
		}
	}
}

func TestADODBIsQueryNonQueryStatements(t *testing.T) {
	vm := NewVM(nil, nil, 1)

	nonQueries := []string{
		"",
		"   ",
		"\t\r\n",
		"INSERT INTO t (a) VALUES (1)",
		"insert into t (a) values (1)",
		"UPDATE t SET a = 1",
		"DELETE FROM t",
		"CREATE TABLE t (a INT)",
		"DROP TABLE t",
		"BEGIN TRANSACTION",
		"EXEC dbo.sp_test",
		"SET NOCOUNT ON",
		"TRUNCATE TABLE t",
		// Keyword present only inside a string literal.
		"INSERT INTO t (val) VALUES ('returning next week')",
		"UPDATE t SET val = 'returning' WHERE id = 1",
		"DELETE FROM t WHERE val = 'returning'",
		// Keyword present only inside a comment.
		"INSERT INTO t (a) VALUES (1) -- returning id",
		"UPDATE t SET a = 1 /* returning a */",
		// Keyword present only inside a quoted identifier.
		`INSERT INTO t ("returning") VALUES (1)`,
		"INSERT INTO t (`returning`) VALUES (1)",
		"INSERT INTO t ([returning]) VALUES (1)",
		// Keyword present only inside a dollar-quoted literal.
		"INSERT INTO t (a) VALUES ($tag$returning$tag$)",
		// Bare identifiers that merely share a keyword prefix.
		"selection",
		"withholding",
		"pragma_table",
		"shows",
		// RETURNING only inside an identifier token.
		"INSERT INTO returning_log (a) VALUES ('x')",
		"UPDATE t_returning SET a = 1",
	}

	for _, sql := range nonQueries {
		if vm.adodbIsQuery(sql) {
			t.Errorf("adodbIsQuery(%q) = true, want false", sql)
		}
	}
}

func TestADODBContainsReturningOutsideQuotes(t *testing.T) {
	tests := []struct {
		name string
		sql  string
		want bool
	}{
		{"plain returning", "INSERT INTO t (a) VALUES (1) RETURNING id", true},
		{"lowercase returning", "insert into t (a) values (1) returning id", true},
		{"returning star", "UPDATE t SET a = 1 RETURNING *", true},
		{"delete returning", "DELETE FROM t RETURNING id", true},
		{"no returning", "INSERT INTO t (a) VALUES (1)", false},
		{"inside literal", "INSERT INTO t (val) VALUES ('returning next week')", false},
		{"inside literal after literal", "INSERT INTO t (val) VALUES ('x') -- returning", false},
		{"escaped quote keeps literal open", "INSERT INTO t (val) VALUES ('don''t match returning here')", false},
		{"escaped quote then real returning", "INSERT INTO t (val) VALUES ('don''t match returning here') RETURNING id", true},
		{"escaped quote then comment returning", "INSERT INTO t (val) VALUES ('it''s fine') /* returning */", false},
		{"double quoted identifier", `INSERT INTO t ("returning") VALUES (1)`, false},
		{"double quoted identifier escaped", `INSERT INTO t ("re""turning") VALUES (1)`, false},
		{"backtick identifier", "INSERT INTO t (`returning`) VALUES (1)", false},
		{"bracket identifier", "INSERT INTO t ([returning]) VALUES (1)", false},
		{"escaped bracket identifier", "INSERT INTO t ([a]]returning]) VALUES (1)", false},
		{"dollar quoted literal", "INSERT INTO t (a) VALUES ($tag$returning$tag$)", false},
		{"dollar quoted then real returning", "INSERT INTO t (a) VALUES ($tag$returning$tag$) RETURNING id", true},
		{"line comment", "INSERT INTO t (a) VALUES (1) -- returning id", false},
		{"line comment ends at newline", "INSERT INTO t (a) VALUES (1) -- note\nRETURNING id", true},
		{"block comment", "INSERT INTO t (a) VALUES (1) /* returning id */", false},
		{"block comment then real returning", "INSERT INTO t (a) VALUES (1) /* note */ RETURNING id", true},
		{"identifier prefix not a token", "INSERT INTO returning_log (a) VALUES ('x')", false},
		{"identifier suffix not a token", "INSERT INTO t_returning (a) VALUES (1)", false},
		{"empty string", "", false},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			if got := adodbContainsReturningOutsideQuotes(tc.sql); got != tc.want {
				t.Errorf("adodbContainsReturningOutsideQuotes(%q) = %v, want %v", tc.sql, got, tc.want)
			}
		})
	}
}

func TestADODBSkipLeadingSQLNoise(t *testing.T) {
	tests := []struct {
		sql  string
		want string
	}{
		{"SELECT 1", "SELECT 1"},
		{"   \t\nSELECT 1", "SELECT 1"},
		{"-- c\nSELECT 1", "SELECT 1"},
		{"/* c */ SELECT 1", "SELECT 1"},
		{"(( SELECT 1", "SELECT 1"},
		{"/* a */ -- b\n ( WITH x AS (SELECT 1) SELECT * FROM x)", "WITH x AS (SELECT 1) SELECT * FROM x)"},
		{"", ""},
		{"   ", ""},
		{"-- only comment", ""},
	}

	for _, tc := range tests {
		got := tc.sql[adodbSkipLeadingSQLNoise(tc.sql):]
		if got != tc.want {
			t.Errorf("adodbSkipLeadingSQLNoise(%q) -> %q, want %q", tc.sql, got, tc.want)
		}
	}
}

func TestADODBHasSQLKeywordPrefix(t *testing.T) {
	tests := []struct {
		sql     string
		keyword string
		want    bool
	}{
		{"select 1", "select", true},
		{"SELECT 1", "select", true},
		{"select", "select", true},
		{"select*from t", "select", true},
		{"selection", "select", false},
		{"select_x", "select", false},
		{"sel", "select", false},
		{"with cte as (select 1) select 1", "with", true},
		{"withholding", "with", false},
		{"insert into t", "insert", true},
	}

	for _, tc := range tests {
		if got := adodbHasSQLKeywordPrefix(tc.sql, tc.keyword); got != tc.want {
			t.Errorf("adodbHasSQLKeywordPrefix(%q, %q) = %v, want %v", tc.sql, tc.keyword, got, tc.want)
		}
	}
}

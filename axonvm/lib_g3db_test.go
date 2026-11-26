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
	"os"
	"path/filepath"
	"testing"

	"g3pix.com.br/axonasp/v2/axonconfig"
)

// TestG3DBBuildDSNFromSectionEnv verifies G3DB_* environment variables override axonasp.toml keys.
func TestG3DBBuildDSNFromSectionEnv(t *testing.T) {
	tmpDir := t.TempDir()
	cfgPath := filepath.Join(tmpDir, "axonasp.toml")
	err := os.WriteFile(cfgPath, []byte(`
[global]
viper_automatic_env = true

[g3db]
mysql_host = "localhost"
mysql_port = 3306
mysql_user = "root"
mysql_pass = "password"
mysql_database = "test"
postgres_host = "localhost"
postgres_port = 5432
postgres_user = "postgres"
postgres_pass = "password"
postgres_database = "test"
postgres_ssl_mode = "disable"
`), 0600)
	if err != nil {
		t.Fatalf("failed to write temp config: %v", err)
	}

	axonconfig.SetCustomConfigPath(cfgPath)
	t.Cleanup(func() {
		axonconfig.SetCustomConfigPath("")
	})

	t.Setenv("G3DB_MYSQL_HOST", "mysql.internal")
	t.Setenv("G3DB_MYSQL_PORT", "3307")
	t.Setenv("G3DB_MYSQL_USER", "appuser")
	t.Setenv("G3DB_MYSQL_PASS", "apppass")
	t.Setenv("G3DB_MYSQL_DATABASE", "appdb")

	t.Setenv("G3DB_POSTGRES_HOST", "pg.internal")
	t.Setenv("G3DB_POSTGRES_PORT", "5433")
	t.Setenv("G3DB_POSTGRES_USER", "pguser")
	t.Setenv("G3DB_POSTGRES_PASS", "pgpass")
	t.Setenv("G3DB_POSTGRES_DATABASE", "pgdb")
	t.Setenv("G3DB_POSTGRES_SSL_MODE", "require")

	v := g3dbNewConfigViper()

	mysqlDSN := g3dbBuildDSN(v, "mysql")
	wantMySQL := "appuser:apppass@tcp(mysql.internal:3307)/appdb?parseTime=true"
	if mysqlDSN != wantMySQL {
		t.Fatalf("unexpected MySQL DSN: got %q, want %q", mysqlDSN, wantMySQL)
	}

	pgDSN := g3dbBuildDSN(v, "postgres")
	wantPG := "host=pg.internal port=5433 user=pguser password=pgpass dbname=pgdb sslmode=require"
	if pgDSN != wantPG {
		t.Fatalf("unexpected Postgres DSN: got %q, want %q", pgDSN, wantPG)
	}
}

// TestG3DBBuildDSNPostgresLegacyFallback verifies legacy postgress_* keys remain supported when postgres_* keys are absent.
func TestG3DBBuildDSNPostgresLegacyFallback(t *testing.T) {
	tmpDir := t.TempDir()
	cfgPath := filepath.Join(tmpDir, "axonasp.toml")
	err := os.WriteFile(cfgPath, []byte(`
[global]
viper_automatic_env = true

[g3db]
postgres_host = "legacy-pg"
postgres_user = "legacyuser"
postgress_database = "legacydb"
postgress_pass = "legacypass"
postgress_port = 5434
postgress_ssl_mode = "verify-full"
`), 0600)
	if err != nil {
		t.Fatalf("failed to write temp config: %v", err)
	}

	axonconfig.SetCustomConfigPath(cfgPath)
	t.Cleanup(func() {
		axonconfig.SetCustomConfigPath("")
	})

	v := g3dbNewConfigViper()
	pgDSN := g3dbBuildDSN(v, "postgres")
	wantPG := "host=legacy-pg port=5434 user=legacyuser password=legacypass dbname=legacydb sslmode=verify-full"
	if pgDSN != wantPG {
		t.Fatalf("unexpected legacy Postgres DSN: got %q, want %q", pgDSN, wantPG)
	}
}

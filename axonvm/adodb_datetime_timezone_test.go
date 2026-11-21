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

import (
	"bytes"
	"context"
	"database/sql"
	"database/sql/driver"
	"fmt"
	"io"
	"strings"
	"sync"
	"sync/atomic"
	"testing"
	"time"
)

type mockMSSQLDataset struct {
	columns []string
	dbTypes []string
	rows    [][]driver.Value
}

var (
	mockMSSQLRegisterOnce sync.Once
	mockMSSQLSeq          atomic.Int64
	mockMSSQLDatasetsMu   sync.Mutex
	mockMSSQLDatasets     = make(map[string]mockMSSQLDataset)
)

type mockMSSQLDriver struct{}

func (d *mockMSSQLDriver) Open(name string) (driver.Conn, error) {
	return &mockMSSQLConn{dsn: name}, nil
}

type mockMSSQLConn struct {
	dsn string
}

func (c *mockMSSQLConn) Prepare(string) (driver.Stmt, error) { return nil, driver.ErrSkip }
func (c *mockMSSQLConn) Close() error                        { return nil }
func (c *mockMSSQLConn) Begin() (driver.Tx, error)           { return nil, driver.ErrSkip }

func (c *mockMSSQLConn) QueryContext(_ context.Context, _ string, _ []driver.NamedValue) (driver.Rows, error) {
	mockMSSQLDatasetsMu.Lock()
	ds, ok := mockMSSQLDatasets[c.dsn]
	mockMSSQLDatasetsMu.Unlock()
	if !ok {
		return nil, fmt.Errorf("mock mssql dataset %q not registered", c.dsn)
	}
	return &mockMSSQLRows{
		columns: ds.columns,
		dbTypes: ds.dbTypes,
		rows:    ds.rows,
	}, nil
}

type mockMSSQLRows struct {
	columns []string
	dbTypes []string
	rows    [][]driver.Value
	idx     int
}

func (r *mockMSSQLRows) Columns() []string {
	return r.columns
}

func (r *mockMSSQLRows) Close() error {
	return nil
}

func (r *mockMSSQLRows) Next(dest []driver.Value) error {
	if r.idx >= len(r.rows) {
		return io.EOF
	}
	copy(dest, r.rows[r.idx])
	r.idx++
	return nil
}

func (r *mockMSSQLRows) ColumnTypeDatabaseTypeName(index int) string {
	if index >= 0 && index < len(r.dbTypes) {
		return r.dbTypes[index]
	}
	return ""
}

func ensureMockMSSQLDriverRegistered() {
	mockMSSQLRegisterOnce.Do(func() {
		sql.Register("mssqlmock", &mockMSSQLDriver{})
	})
}

func openMockMSSQLConnection(t *testing.T, ds mockMSSQLDataset) (*sql.DB, *sql.Conn) {
	t.Helper()
	ensureMockMSSQLDriverRegistered()
	key := fmt.Sprintf("case-%d", mockMSSQLSeq.Add(1))
	mockMSSQLDatasetsMu.Lock()
	mockMSSQLDatasets[key] = ds
	mockMSSQLDatasetsMu.Unlock()

	db, err := sql.Open("mssqlmock", key)
	if err != nil {
		mockMSSQLDatasetsMu.Lock()
		delete(mockMSSQLDatasets, key)
		mockMSSQLDatasetsMu.Unlock()
		t.Fatalf("sql.Open(mssqlmock) failed: %v", err)
	}
	dbConn, err := db.Conn(context.Background())
	if err != nil {
		_ = db.Close()
		mockMSSQLDatasetsMu.Lock()
		delete(mockMSSQLDatasets, key)
		mockMSSQLDatasetsMu.Unlock()
		t.Fatalf("db.Conn failed: %v", err)
	}
	t.Cleanup(func() {
		_ = dbConn.Close()
		_ = db.Close()
		mockMSSQLDatasetsMu.Lock()
		delete(mockMSSQLDatasets, key)
		mockMSSQLDatasetsMu.Unlock()
	})
	return db, dbConn
}

func withTestRuntimeTimezone(t *testing.T, tz string) *time.Location {
	t.Helper()
	loc, err := ResolveTimezoneLocation(tz)
	if err != nil {
		t.Fatalf("ResolveTimezoneLocation(%q) failed: %v", tz, err)
	}

	cachedBuiltinDefaultsLock.Lock()
	prevDefaults := cachedBuiltinDefaults
	prevInitialized := builtinDefaultsInitialized
	if !prevInitialized {
		prevDefaults = builtinDefaults{mslcid: int(EnglishUS), timezone: "UTC", charset: "UTF-8", location: time.UTC}
	}
	updated := prevDefaults
	updated.timezone = tz
	updated.location = loc
	cachedBuiltinDefaults = updated
	builtinDefaultsInitialized = true
	cachedBuiltinDefaultsLock.Unlock()

	t.Cleanup(func() {
		cachedBuiltinDefaultsLock.Lock()
		cachedBuiltinDefaults = prevDefaults
		builtinDefaultsInitialized = prevInitialized
		cachedBuiltinDefaultsLock.Unlock()
	})
	return loc
}

func runASPScriptWithMockMSSQLConn(t *testing.T, source string, db *sql.DB, dbConn *sql.Conn) (string, *VM) {
	t.Helper()
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	connVal := vm.newADODBConnection()
	conn := vm.adodbConnectionItems[connVal.Num]
	conn.db = db
	conn.dbConn = dbConn
	conn.dbDriver = "mssql"
	conn.provider = "SQLOLEDB"
	conn.state = adStateOpen

	connIdx, ok := compiler.Globals.Get("conn")
	if !ok {
		t.Fatalf("expected global variable 'conn' in compiled script")
	}
	vm.Globals[connIdx] = connVal

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	return output.String(), vm
}

// TestNormalizeSQLDateTimeUnit verifies timezone-naive SQL types rebind to runtime
// location while timezone-aware types and explicit offsets remain intact.
func TestNormalizeSQLDateTimeUnit(t *testing.T) {
	romeLoc, err := ResolveTimezoneLocation("Europe/Rome")
	if err != nil {
		t.Fatalf("load Europe/Rome failed: %v", err)
	}

	if got := normalizeSQLDateTime("DATETIME", nil, romeLoc); got != nil {
		t.Fatalf("expected nil, got %#v", got)
	}
	if got := normalizeSQLDateTime("DATETIME", "raw", romeLoc); got != "raw" {
		t.Fatalf("expected non-time untouched, got %#v", got)
	}
	if got := normalizeSQLDateTime("DATETIME", time.Time{}, romeLoc); !got.(time.Time).IsZero() {
		t.Fatalf("expected zero time.Time untouched, got %#v", got)
	}

	naiveUTC := time.Date(2026, time.November, 20, 23, 30, 0, 123000000, time.UTC)
	for _, dbType := range []string{"DATETIME", "SMALLDATETIME", "DATETIME2", "DATETIME2(7)", "DATE", "TIMESTAMP", "TIMESTAMP WITHOUT TIME ZONE"} {
		res := normalizeSQLDateTime(dbType, naiveUTC, romeLoc).(time.Time)
		if res.Location() != romeLoc {
			t.Fatalf("%s: expected location Europe/Rome, got %v", dbType, res.Location())
		}
		if res.Year() != 2026 || res.Month() != time.November || res.Day() != 20 || res.Hour() != 23 || res.Minute() != 30 || res.Second() != 0 || res.Nanosecond() != 123000000 {
			t.Fatalf("%s: wall clock mutated: %v", dbType, res)
		}
	}

	// Non-UTC FixedZone on a naive type name must remain preserved.
	plusTwo := time.Date(2026, time.July, 15, 14, 35, 12, 0, time.FixedZone("", 7200))
	if got := normalizeSQLDateTime("DATETIME", plusTwo, romeLoc).(time.Time); !got.Equal(plusTwo) || got.Location() != plusTwo.Location() {
		t.Fatalf("expected explicit non-UTC FixedZone preserved, got %v", got)
	}

	// DATETIMEOFFSET +00:00 uses FixedZone("", 0) in go-mssqldb and must remain an exact UTC instant.
	offsetZero := time.Date(2026, time.November, 20, 14, 35, 12, 0, time.FixedZone("", 0))
	for _, dbType := range []string{"DATETIMEOFFSET", "DATETIMEOFFSET(7)", "TIMESTAMPTZ", "TIMESTAMP WITH TIME ZONE", "TIMESTAMP WITH LOCAL TIME ZONE", "TIMETZ", "TIME WITH TIME ZONE"} {
		res := normalizeSQLDateTime(dbType, offsetZero, romeLoc).(time.Time)
		if res.In(romeLoc).Hour() != 15 {
			t.Fatalf("%s: expected 14:35 +00:00 to convert to 15:35 in Rome, got %v", dbType, res.In(romeLoc))
		}
	}

	// TIME type (decoded as 0001-01-01 UTC by go-mssqldb) must re-anchor to 1899-12-30 in runtime loc.
	sqlTime := time.Date(1, time.January, 1, 14, 35, 12, 0, time.UTC)
	normTime := normalizeSQLDateTime("TIME", sqlTime, romeLoc).(time.Time)
	if normTime.Year() != 1899 || normTime.Month() != time.December || normTime.Day() != 30 || normTime.Hour() != 14 || normTime.Minute() != 35 || normTime.Second() != 12 || normTime.Location() != romeLoc {
		t.Fatalf("unexpected normalized TIME: %v", normTime)
	}
	if got := NewDate(normTime); got.Num == 0 {
		t.Fatalf("expected valid UnixNano payload for normalized TIME")
	}

	// SQL_VARIANT: UTC datetime rebinds; FixedZone("", 0) datetimeoffset stays instant; 0001-01-01 time re-anchors.
	if got := normalizeSQLDateTime("SQL_VARIANT", naiveUTC, romeLoc).(time.Time); got.Location() != romeLoc || got.Hour() != 23 {
		t.Fatalf("SQL_VARIANT naive datetime failed: %v", got)
	}
	if got := normalizeSQLDateTime("SQL_VARIANT", offsetZero, romeLoc).(time.Time); got.In(romeLoc).Hour() != 15 {
		t.Fatalf("SQL_VARIANT datetimeoffset(+00:00) failed: %v", got.In(romeLoc))
	}
	if got := normalizeSQLDateTime("SQL_VARIANT", sqlTime, romeLoc).(time.Time); got.Year() != 1899 || got.Hour() != 14 || got.Location() != romeLoc {
		t.Fatalf("SQL_VARIANT time failed: %v", got)
	}
}

// TestADODBDateTimeTimezoneShiftAndFormatting reproduces and verifies the SQL Server
// datetime +1h shift fix, FormatDateTime(val, vbShortDate) on 23:30:00 inputs,
// and default date-only rendering under SetLocale 1040 across timezones and DST states.
func TestADODBDateTimeTimezoneShiftAndFormatting(t *testing.T) {
	timezones := []string{
		"Europe/Rome",
		"America/Sao_Paulo",
		"America/New_York",
		"UTC",
	}
	dates := []struct {
		name  string
		year  int
		month time.Month
		day   int
		itStr string
	}{
		{name: "NonDST_Nov", year: 2026, month: time.November, day: 20, itStr: "20/11/2026"},
		{name: "DST_Jul", year: 2026, month: time.July, day: 15, itStr: "15/07/2026"},
	}
	sqlTypes := []struct {
		typeName    string
		timedSecond int
	}{
		{typeName: "DATETIME", timedSecond: 12},
		{typeName: "SMALLDATETIME", timedSecond: 0},
		{typeName: "DATETIME2", timedSecond: 12},
		{typeName: "DATE", timedSecond: 0},
	}

	for _, tz := range timezones {
		for _, dt := range dates {
			for _, st := range sqlTypes {
				name := fmt.Sprintf("%s/%s/%s", tz, dt.name, st.typeName)
				t.Run(name, func(t *testing.T) {
					loc := withTestRuntimeTimezone(t, tz)

					midnightUTC := time.Date(dt.year, dt.month, dt.day, 0, 0, 0, 0, time.UTC)
					timedUTC := time.Date(dt.year, dt.month, dt.day, 14, 35, st.timedSecond, 0, time.UTC)
					lateUTC := time.Date(dt.year, dt.month, dt.day, 23, 30, 0, 0, time.UTC)
					if st.typeName == "DATE" {
						timedUTC = midnightUTC
						lateUTC = midnightUTC
					}

					db, dbConn := openMockMSSQLConnection(t, mockMSSQLDataset{
						columns: []string{"midnight", "timed", "late", "data_fattura"},
						dbTypes: []string{st.typeName, st.typeName, st.typeName, "DATE"},
						rows: [][]driver.Value{
							{midnightUTC, timedUTC, lateUTC, midnightUTC},
						},
					})

					// 1. Verify raw VTDate components inside ADODB Recordset against runtime location.
					checkVM := NewVM(nil, nil, 5)
					defer checkVM.CleanupRequestResources()
					rows, err := dbConn.QueryContext(context.Background(), "SELECT midnight, timed, late, data_fattura FROM t")
					if err != nil {
						t.Fatalf("QueryContext failed: %v", err)
					}
					rsVal := checkVM.newADODBRecordset()
					rs := checkVM.adodbRecordsetItems[rsVal.Num]
					rs.sqlRows = rows
					checkVM.adodbRecordsetLoadCurrentSQLResultSet(rs)

					assertClock := func(col string, wantHour, wantMin, wantSec int) {
						v := rs.data[0][col]
						if v.Type != VTDate {
							t.Fatalf("%s: expected VTDate, got %#v", col, v)
						}
						got := time.Unix(0, v.Num).In(loc)
						if got.Year() != dt.year || got.Month() != dt.month || got.Day() != dt.day ||
							got.Hour() != wantHour || got.Minute() != wantMin || got.Second() != wantSec {
							t.Fatalf("%s (%s): got %v, want %04d-%02d-%02d %02d:%02d:%02d",
								col, tz, got, dt.year, dt.month, dt.day, wantHour, wantMin, wantSec)
						}
					}
					assertClock("midnight", 0, 0, 0)
					if st.typeName == "DATE" {
						assertClock("timed", 0, 0, 0)
						assertClock("late", 0, 0, 0)
					} else {
						assertClock("timed", 14, 35, st.timedSecond)
						assertClock("late", 23, 30, 0)
					}

					// 2. Verify full ASP execution: SetLocale 1040, Response.Write, <%= ... %>, FormatDateTime, GetString.
					source := `<%
SetLocale 1040
Set rs = conn.Execute("SELECT midnight, timed, late, data_fattura FROM t")
Response.Write rs("midnight").Value & "|"
%><%=rs("midnight")%>|<%=rs("data_fattura")%><%
Response.Write "|" & rs("timed").Value & "|"
Response.Write rs("late").Value & "|"
Response.Write FormatDateTime(rs("late").Value, vbShortDate) & "|"
Response.Write FormatDateTime(rs("late"), vbShortDate) & "|"
Response.Write FormatDateTime(rs("midnight").Value, vbGeneralDate) & "|"
Response.Write FormatDateTime(rs("midnight").Value) & "|"
Response.Write Trim(Replace(rs.GetString(), vbLf, ""))
rs.Close
%>`
					got, scriptVM := runASPScriptWithMockMSSQLConn(t, source, db, dbConn)
					defer scriptVM.CleanupRequestResources()

					wantTimed := fmt.Sprintf("%s 14:35:%02d", dt.itStr, st.timedSecond)
					wantLate := fmt.Sprintf("%s 23:30:00", dt.itStr)
					if st.typeName == "DATE" {
						wantTimed = dt.itStr
						wantLate = dt.itStr
					}
					wantGetString := strings.Join([]string{dt.itStr, wantTimed, wantLate, dt.itStr}, "\t")
					want := strings.Join([]string{
						dt.itStr,
						dt.itStr,
						dt.itStr,
						wantTimed,
						wantLate,
						dt.itStr,
						dt.itStr,
						dt.itStr,
						dt.itStr,
						wantGetString,
					}, "|")

					if got != want {
						t.Fatalf("output mismatch:\ngot:  %q\nwant: %q", got, want)
					}
				})
			}
		}
	}
}

// TestADODBDateTimeOffsetVariantAndTimePreservation ensures DATETIMEOFFSET preserves
// offset-aware instant conversions while SQL_VARIANT and TIME handle SQL Server wire shapes accurately.
func TestADODBDateTimeOffsetVariantAndTimePreservation(t *testing.T) {
	withTestRuntimeTimezone(t, "Europe/Rome")

	db, dbConn := openMockMSSQLConnection(t, mockMSSQLDataset{
		columns: []string{"dto_utc", "dto_cet", "dto_cest", "var_dt", "var_dto", "t_only"},
		dbTypes: []string{"DATETIMEOFFSET", "DATETIMEOFFSET", "DATETIMEOFFSET", "SQL_VARIANT", "SQL_VARIANT", "TIME"},
		rows: [][]driver.Value{
			{
				time.Date(2026, time.November, 20, 14, 35, 12, 0, time.FixedZone("", 0)),
				time.Date(2026, time.November, 20, 14, 35, 12, 0, time.FixedZone("", 3600)),
				time.Date(2026, time.July, 15, 22, 30, 0, 0, time.FixedZone("", 0)),
				time.Date(2026, time.November, 20, 23, 30, 0, 0, time.UTC),
				time.Date(2026, time.November, 20, 14, 35, 12, 0, time.FixedZone("", 0)),
				time.Date(1, time.January, 1, 14, 35, 12, 0, time.UTC),
			},
		},
	})

	source := `<%
SetLocale 1040
Set rs = conn.Execute("SELECT dto_utc, dto_cet, dto_cest, var_dt, var_dto, t_only FROM t")
Response.Write rs("dto_utc").Value & "|" & rs("dto_cet").Value & "|" & rs("dto_cest").Value & "|" & rs("var_dt").Value & "|" & rs("var_dto").Value & "|" & rs("t_only").Value
rs.Close
%>`
	got, vm := runASPScriptWithMockMSSQLConn(t, source, db, dbConn)
	defer vm.CleanupRequestResources()

	want := "20/11/2026 15:35:12|20/11/2026 14:35:12|16/07/2026 00:30:00|20/11/2026 23:30:00|20/11/2026 15:35:12|14:35:12"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBDateTimeWritebackAndParameterRoundTrip verifies VTDate values convert back
// to SQL literals and driver parameters using the runtime location instead of shifting backward by UTC offset.
func TestADODBDateTimeWritebackAndParameterRoundTrip(t *testing.T) {
	loc := withTestRuntimeTimezone(t, "Europe/Rome")
	vm := NewVM(nil, nil, 5)
	defer vm.CleanupRequestResources()

	conn := &adodbConnection{dbDriver: "mssql"}
	midnight := NewDate(time.Date(2026, time.November, 20, 0, 0, 0, 0, loc))
	lateSummer := NewDate(time.Date(2026, time.July, 15, 23, 30, 0, 0, loc))

	litMidnight, ok := vm.adodbSQLLiteral(conn, midnight)
	if !ok || litMidnight != "'2026-11-20 00:00:00'" {
		t.Fatalf("unexpected midnight SQL literal: ok=%v got=%q", ok, litMidnight)
	}
	litLateSummer, ok := vm.adodbSQLLiteral(conn, lateSummer)
	if !ok || litLateSummer != "'2026-07-15 23:30:00'" {
		t.Fatalf("unexpected summer late SQL literal: ok=%v got=%q", ok, litLateSummer)
	}

	drvMidnight, ok := vm.adodbDriverValue(midnight).(time.Time)
	if !ok || drvMidnight.Location() != loc || drvMidnight.Year() != 2026 || drvMidnight.Month() != time.November || drvMidnight.Day() != 20 || drvMidnight.Hour() != 0 {
		t.Fatalf("unexpected driver midnight time: %#v", drvMidnight)
	}
	drvLateSummer, ok := vm.adodbDriverValue(lateSummer).(time.Time)
	if !ok || drvLateSummer.Location() != loc || drvLateSummer.Year() != 2026 || drvLateSummer.Month() != time.July || drvLateSummer.Day() != 15 || drvLateSummer.Hour() != 23 || drvLateSummer.Minute() != 30 {
		t.Fatalf("unexpected driver summer late time: %#v", drvLateSummer)
	}
}

// TestADODBSQLiteDatetimeAndEnableSessionStateFalse verifies SQLite DATE/DATETIME
// retrieval and SetLocale 1040 formatting when EnableSessionState=False.
func TestADODBSQLiteDatetimeAndEnableSessionStateFalse(t *testing.T) {
	withTestRuntimeTimezone(t, "Europe/Rome")

	source := `<%@ EnableSessionState="False" %><%
SetLocale 1040
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite::memory:"
conn.Execute "CREATE TABLE fatture (data_fattura DATE, created_at DATETIME)"
conn.Execute "INSERT INTO fatture VALUES ('2026-11-20', '2026-11-20 23:30:00')"
conn.Execute "INSERT INTO fatture VALUES ('2026-07-15', '2026-07-15 00:00:00')"
Set rs = conn.Execute("SELECT data_fattura, created_at FROM fatture")
Do While Not rs.EOF
%><%=rs("data_fattura")%>|<%=rs("created_at")%>|<%=FormatDateTime(rs("created_at").Value, vbShortDate)%>;<%
	rs.MoveNext
Loop
rs.Close
conn.Close
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	want := "20/11/2026|20/11/2026 23:30:00|20/11/2026;15/07/2026|15/07/2026|15/07/2026;"
	if got := output.String(); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestG3DBDatetimeTimezoneConsistency verifies G3DB shares timezone-naive SQL
// normalization and SetLocale 1040 rendering under Europe/Rome.
func TestG3DBDatetimeTimezoneConsistency(t *testing.T) {
	withTestRuntimeTimezone(t, "Europe/Rome")

	source := `<%
SetLocale 1040
Dim db, rs
Set db = Server.CreateObject("G3DB")
db.Open "sqlite", ":memory:"
db.Exec "CREATE TABLE events (d DATE, dt DATETIME)"
db.Exec "INSERT INTO events VALUES ('2026-11-20', '2026-11-20 23:30:00')"
Set rs = db.Query("SELECT d, dt FROM events")
Response.Write rs("d") & "|" & rs("dt") & "|" & FormatDateTime(rs("dt"), vbShortDate)
rs.Close
db.Close
%>`
	got := runASPSourceForTestWithHost(t, source, NewMockHost())
	want := "20/11/2026|20/11/2026 23:30:00|20/11/2026"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestReloadBuiltinDefaultsFromEnvTimezone verifies GLOBAL_DEFAULT_TIMEZONE drives
// ReloadBuiltinDefaults() and prevents the +1h shift end-to-end.
func TestReloadBuiltinDefaultsFromEnvTimezone(t *testing.T) {
	t.Cleanup(func() {
		ReloadBuiltinDefaults()
	})
	t.Setenv("GLOBAL_DEFAULT_TIMEZONE", "Europe/Rome")
	cfg := ReloadBuiltinDefaults()
	if cfg.location == nil || cfg.location.String() != "Europe/Rome" {
		t.Fatalf("expected Europe/Rome from ReloadBuiltinDefaults, got %v", cfg.location)
	}

	db, dbConn := openMockMSSQLConnection(t, mockMSSQLDataset{
		columns: []string{"midnight", "late"},
		dbTypes: []string{"DATETIME", "DATETIME"},
		rows: [][]driver.Value{
			{
				time.Date(2026, time.November, 20, 0, 0, 0, 0, time.UTC),
				time.Date(2026, time.November, 20, 23, 30, 0, 0, time.UTC),
			},
		},
	})

	source := `<%
SetLocale 1040
Set rs = conn.Execute("SELECT midnight, late FROM t")
Response.Write rs("midnight").Value & "|" & rs("late").Value & "|" & FormatDateTime(rs("late").Value, vbShortDate)
rs.Close
%>`
	got, vm := runASPScriptWithMockMSSQLConn(t, source, db, dbConn)
	defer vm.CleanupRequestResources()

	want := "20/11/2026|20/11/2026 23:30:00|20/11/2026"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestG3DateConvertUTCtoZoneStillShiftsExplicitly ensures explicit timezone
// conversions via G3Date remain unaffected.
func TestG3DateConvertUTCtoZoneStillShiftsExplicitly(t *testing.T) {
	withTestRuntimeTimezone(t, "Europe/Rome")

	source := `<%
SetLocale 1040
Dim gd, utcDt, romeDt
Set gd = Server.CreateObject("G3DATE")
utcDt = gd.Parse("2006-01-02 15:04:05", "2026-11-20 14:35:12", "UTC")
romeDt = gd.ConvertUTCtoZone(utcDt, "Europe/Rome")
Response.Write romeDt
%>`
	got := runASPSourceForTestWithHost(t, source, NewMockHost())
	if got != "20/11/2026 15:35:12" {
		t.Fatalf("expected explicit UTC->Rome shift to 20/11/2026 15:35:12, got %q", got)
	}
}

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 */
package axonvm

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/v2/axonvm/asp"
)

// These tests use generic examples of IIS-compatible Classic ASP patterns.

func TestJScriptDynamicApplicationAndSessionAssignment(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var applicationKey = "dynamic-app";
var sessionKey = "dynamic-session";
Application(applicationKey) = "application-value";
Session(sessionKey) = "session-value";
Response.Write(Application(applicationKey) + "|" + Session(sessionKey));
%>`

	if got := runASPSourceForTest(t, source); got != "application-value|session-value" {
		t.Fatalf("unexpected collection assignment output: %q", got)
	}
}

func TestJScriptDynamicCollectionAssignmentExpressions(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var setting = "name";
var settings = { name: "value" };
var siteId = 7;
Application("global_" + setting) = settings[setting];
Application(siteId + "-" + setting) = settings[setting];
Response.Write(Application("global_name") + "|" + Application("7-name"));
%>`

	if got := runASPSourceForTest(t, source); got != "value|value" {
		t.Fatalf("unexpected expression-key collection assignment output: %q", got)
	}
}

func TestJScriptDynamicCollectionAssignmentFromParsedObject(t *testing.T) {
	source := `<%@ Language="JScript" %><%
function publish(prefix, values) {
    for (var key in values) {
        if (values.hasOwnProperty(key)) {
            Application(prefix + key) = values[key];
        }
    }
}
var values = eval('({"language":"en","limit":15})');
publish("option-", values);
Response.Write(Application("option-language") + "|" + Application("option-limit"));
%>`

	if got := runASPSourceForTest(t, source); got != "en|15" {
		t.Fatalf("unexpected parsed-object assignment output: %q", got)
	}
}

func TestJScriptADODBParameterDefaultPropertyAssignmentCompiles(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var command = Server.CreateObject("ADODB.Command");
command("record_id") = 42;
Response.Write("ok");
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("default-property assignment should compile: %v", err)
	}
}

func TestJScriptADODBCommandBindsNamedParameters(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	vm.SetHost(host)

	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	vm.dispatchMemberSet(conn.Num, "ConnectionString", NewString("Driver={SQLite3};Data Source="+filepath.Join(rootDir, "command-parameters.db")))
	vm.dispatchNativeCall(conn.Num, "Open", nil)
	defer vm.dispatchNativeCall(conn.Num, "Close", nil)
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE TABLE events (actor TEXT, detail TEXT)")})

	cmd := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Command")})
	vm.dispatchMemberSet(cmd.Num, "ActiveConnection", conn)
	vm.dispatchMemberSet(cmd.Num, "CommandText", NewString("INSERT INTO events (actor, detail) VALUES (?, ?)"))

	userParam := vm.dispatchNativeCall(cmd.Num, "CreateParameter", []Value{NewString("actor"), NewInteger(200), NewInteger(1), NewInteger(50)})
	messageParam := vm.dispatchNativeCall(cmd.Num, "CreateParameter", []Value{NewString("detail"), NewInteger(200), NewInteger(1), NewInteger(4000)})
	params := vm.dispatchMemberGet(cmd, "Parameters")
	vm.dispatchNativeCall(params.Num, "Append", []Value{userParam})
	vm.dispatchNativeCall(params.Num, "Append", []Value{messageParam})
	vm.dispatchNativeCall(cmd.Num, "", []Value{NewString("actor"), NewString("user-1")})
	vm.dispatchNativeCall(cmd.Num, "", []Value{NewString("detail"), NewString("sample event")})
	vm.dispatchNativeCall(cmd.Num, "Execute", nil)

	rs := vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("SELECT actor, detail FROM events")})
	if got := vm.dispatchMemberGet(rs, "actor").String(); got != "user-1" {
		t.Fatalf("unexpected bound actor: %q", got)
	}
	if got := vm.dispatchMemberGet(rs, "detail").String(); got != "sample event" {
		t.Fatalf("unexpected bound detail: %q", got)
	}
}

func TestADODBErrorExposesNativeError(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	conn := &adodbConnection{}
	vm.adodbConnectionPushError(conn, "test", -2147467259, "provider failed", "HY000")
	errors := vm.newADODBErrorsCollection(conn)
	errValue, handled := vm.dispatchADODBErrorsCollectionMethod(errors.Num, "Item", []Value{NewInteger(0)})
	if !handled {
		t.Fatal("Errors.Item was not handled")
	}
	nativeError, handled := vm.dispatchADODBErrorPropertyGet(errValue.Num, "NativeError")
	if !handled || nativeError.Type != VTInteger || nativeError.Num != -2147467259 {
		t.Fatalf("unexpected NativeError: %#v", nativeError)
	}
}

func TestJScriptADODBErrorsDefaultIndexedPropertyExposesProviderFields(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	conn := &adodbConnection{}
	vm.adodbConnectionPushError(conn, "ADODB.Connection.Execute", -2147467259, "provider failed", "HY000")
	errValue := vm.dispatchADODBConnectionMethod(conn, "Errors", []Value{NewInteger(0)})
	if errValue.Type != VTNativeObject {
		t.Fatalf("default Errors(index) did not return an error object: %#v", errValue)
	}
	for member, want := range map[string]string{
		"NativeError": "-2147467259",
		"Description": "provider failed",
		"Source":      "ADODB.Connection.Execute",
	} {
		got, propertyHandled := vm.dispatchADODBErrorPropertyGet(errValue.Num, member)
		if !propertyHandled || got.String() != want {
			t.Fatalf("unexpected %s: %#v, want %q", member, got, want)
		}
	}
}

func TestADODBNormalizesODBCProcedureCallForSQLServer(t *testing.T) {
	got := adodbNormalizeCommandText(" { call process_item('A', 2) } ", "mssql")
	if got != "EXEC process_item 'A', 2" {
		t.Fatalf("unexpected normalized command: %q", got)
	}
}

func TestADODBLeavesODBCProcedureCallForOtherProviders(t *testing.T) {
	sqlText := "{call process_item('A', 2)}"
	if got := adodbNormalizeCommandText(sqlText, "sqlite"); got != sqlText {
		t.Fatalf("unexpected non-SQL Server rewrite: %q", got)
	}
}

func TestADODBNormalizesODBCProcedureCallWithoutArguments(t *testing.T) {
	if got := adodbNormalizeCommandText("{ call process_pending() }", "mssql"); got != "EXEC process_pending" {
		t.Fatalf("unexpected no-argument command: %q", got)
	}
}

func TestADODBConnectionPinsDatabaseSession(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	vm.SetHost(host)

	connValue := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	conn := vm.adodbConnectionItems[connValue.Num]
	vm.dispatchMemberSet(connValue.Num, "ConnectionString", NewString("Driver={SQLite3};Data Source="+filepath.Join(rootDir, "pinned-session.db")))
	vm.dispatchNativeCall(connValue.Num, "Open", nil)
	defer vm.dispatchNativeCall(connValue.Num, "Close", nil)
	if conn.db == nil || conn.dbConn == nil {
		t.Fatalf("ADODB.Connection did not retain a physical database session: %#v", conn)
	}
	vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("CREATE TEMP TABLE session_data (value TEXT)")})
	vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("INSERT INTO session_data (value) VALUES ('same-session')")})
	rs := vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("SELECT value FROM session_data")})
	if got := vm.dispatchMemberGet(rs, "value").String(); got != "same-session" {
		t.Fatalf("unexpected session-scoped value: %q", got)
	}
}

func TestADODBRecordsetDefaultPropertyReadsUnnamedColumnByOrdinal(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	vm.SetHost(NewMockHost())
	rsValue := vm.newADODBRecordset()
	rs := vm.adodbRecordsetItems[rsValue.Num]
	rs.columns = []string{""}
	rs.data = []map[string]Value{{"": NewInteger(42)}}
	rs.recordCount = 1
	rs.currentRow = 0
	rs.state = adStateOpen
	rs.eof = false
	rs.bof = false

	got := vm.dispatchNativeCall(rsValue.Num, "", []Value{NewInteger(0)})
	if got.Type != VTNativeObject {
		t.Fatalf("expected ADODB.Field proxy from rs(0), got %#v", got)
	}
	val, ok := vm.dispatchADODBFieldPropertyGet(got.Num, "Value")
	if !ok || val.Type != VTInteger || val.Num != 42 {
		t.Fatalf("unexpected unnamed ordinal field .Value: %#v (ok=%t)", val, ok)
	}
	defVal, ok := vm.dispatchADODBFieldPropertyGet(got.Num, "__default__")
	if !ok || defVal.Type != VTInteger || defVal.Num != 42 {
		t.Fatalf("unexpected unnamed ordinal field default value: %#v (ok=%t)", defVal, ok)
	}
}

func TestADODBLiveRecordsetDefaultPropertyPreservesOrdinalSelector(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	if selector := vm.adodbOLEFieldSelector(NewInteger(1)); selector != 1 {
		t.Fatalf("expected numeric ADODB selector 1, got %#v", selector)
	}
	if selector := vm.adodbOLEFieldSelector(NewString("field_b")); selector != "field_b" {
		t.Fatalf("expected named ADODB selector, got %#v", selector)
	}
}

func TestADODBMaterializedRecordsetReadsEveryFieldByOrdinal(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	rs := &adodbRecordset{columns: []string{"field_a", "field_b", "field_c", "field_d"}}
	row := map[string]Value{
		"field_a": NewInteger(7),
		"field_b": NewString("value-b"),
		"field_c": NewString("value-c"),
		"field_d": NewString("value-d"),
	}

	want := []string{"7", "value-b", "value-c", "value-d"}
	for index, expected := range want {
		if got := vm.adodbRecordsetValueByOrdinal(rs, row, index).String(); got != expected {
			t.Fatalf("field %d = %q, want %q", index, got, expected)
		}
	}
}

func TestJScriptADODBRecordsetDefaultPropertyReadsFieldsByOrdinalAndName(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var connection = Server.CreateObject("ADODB.Connection");
connection.Open("sqlite::memory:");
connection.Execute("CREATE TABLE sample_values (field_a INTEGER, field_b TEXT)");
connection.Execute("INSERT INTO sample_values VALUES (7, 'value-b')");
var rs = connection.Execute("SELECT field_a, field_b FROM sample_values");
Response.Write(String(rs(0)) + "|" + String(rs(1)) + "|" + String(rs("field_b")));
connection.Close();
%>`

	if got := runASPSourceForTest(t, source); got != "7|value-b|value-b" {
		t.Fatalf("unexpected JScript Recordset field output: %q", got)
	}
}

func TestJScriptADODBRecordsetDuplicateNamesKeepOrdinalValues(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var connection = Server.CreateObject("ADODB.Connection");
connection.Open("sqlite::memory:");
connection.Execute("CREATE TABLE left_table (id TEXT)");
connection.Execute("CREATE TABLE right_table (id TEXT)");
connection.Execute("INSERT INTO left_table VALUES ('left-id')");
connection.Execute("INSERT INTO right_table VALUES ('right-id')");
var rs = connection.Execute("SELECT a.id, b.id FROM left_table a JOIN right_table b ON 1=1");
Response.Write(String(rs(0)) + "|" + String(rs(1)) + "|" + String(rs("id")) + "|" + String(rs.Fields("id").Value));
connection.Close();
%>`
	if got := runASPSourceForTest(t, source); got != "left-id|right-id|right-id|right-id" {
		t.Fatalf("unexpected duplicate field output: %q", got)
	}
}

func TestJScriptADODBRecordsetUnnamedExpressionPreservesNumericValue(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var connection = Server.CreateObject("ADODB.Connection");
connection.Open("sqlite::memory:");
var rs = connection.Execute("SELECT 1");
Response.Write(String(rs(0)) + "|" + String(rs.Fields(0).Value) + "|" + (+rs.Fields.Item(0)));
connection.Close();
%>`
	if got := runASPSourceForTest(t, source); got != "1|1|1" {
		t.Fatalf("unexpected unnamed expression output: %q", got)
	}
}

func TestJScriptSequenceAssignmentsRetainEachRecordsetFieldValue(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var connection = Server.CreateObject("ADODB.Connection");
connection.Open("sqlite::memory:");
connection.Execute("CREATE TABLE sample_values (field_a INTEGER, field_b TEXT, field_c TEXT, field_d TEXT)");
connection.Execute("INSERT INTO sample_values VALUES (7, 'value-b', 'value-c', 'value-d')");
var rs = connection.Execute("SELECT field_a, field_b, field_c, field_d FROM sample_values");
var rs0, rs1, rs2, rs3;
rs0=String(rs(0)), rs1=String(rs(1)), rs2=String(rs(2)), rs3=String(rs(3));
Response.Write(rs0 + "|" + rs1 + "|" + rs2 + "|" + rs3);
connection.Close();
%>`

	if got := runASPSourceForTest(t, source); got != "7|value-b|value-c|value-d" {
		t.Fatalf("unexpected JScript sequence assignment output: %q", got)
	}
}

func TestJScriptSessionContentsRemovalWhileIteratingDescending(t *testing.T) {
	source := `<%@ Language="JScript" %><%
Session("item-alpha") = "first";
Session("item-beta") = "second";
Session("item-gamma") = "third";
for (var i = Session.Contents.Count; i > 0; i--) {
    var key = String(Session.Contents.Key(i));
    Session.Contents.Remove(key);
}
Response.Write(Session.Contents.Count + "|" + typeof Session("item-alpha"));
%>`

	if got := runASPSourceForTest(t, source); got != "0|undefined" {
		t.Fatalf("unexpected JScript Session.Contents removal output: %q", got)
	}
}

func TestJScriptApplicationContentsRemovalWhileIteratingDescending(t *testing.T) {
	source := `<%@ Language="JScript" %><%
Application("item-alpha") = "first";
Application("item-beta") = "second";
Application("item-gamma") = 15;
for (var i = Application.Contents.Count; i > 0; i--) {
    var key = String(Application.Contents.Key(i));
    Application.Contents.Remove(key);
}
Response.Write(Application.Contents.Count + "|" + typeof Application("item-alpha"));
%>`

	if got := runASPSourceForTest(t, source); got != "0|undefined" {
		t.Fatalf("unexpected JScript Application.Contents removal output: %q", got)
	}
}

func TestADODBMaterializedRecordsetReleasesPinnedConnection(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	vm.SetHost(NewMockHost())
	rootDir := t.TempDir()
	connValue := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	vm.dispatchMemberSet(connValue.Num, "ConnectionString", NewString("Driver={SQLite3};Data Source="+filepath.Join(rootDir, "released-rows.db")))
	vm.dispatchNativeCall(connValue.Num, "Open", nil)
	defer vm.dispatchNativeCall(connValue.Num, "Close", nil)
	vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("CREATE TABLE sample (value TEXT)")})
	vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("INSERT INTO sample (value) VALUES ('first')")})

	rs := vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("SELECT value FROM sample")})
	if vm.adodbRecordsetItems[rs.Num].sqlRows != nil {
		t.Fatal("fully materialized Recordset retained database rows")
	}
	vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("INSERT INTO sample (value) VALUES ('second')")})
	count := vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("SELECT COUNT(*) AS total FROM sample")})
	if got := vm.dispatchMemberGet(count, "total"); got.Type != VTInteger || got.Num != 2 {
		t.Fatalf("unexpected row count after reusing connection: %#v", got)
	}
}

func TestADODBQueryPopulatesFieldTypesForLegacyRecordsetMapping(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	vm.SetHost(NewMockHost())
	rootDir := t.TempDir()
	connValue := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	vm.dispatchMemberSet(connValue.Num, "ConnectionString", NewString("Driver={SQLite3};Data Source="+filepath.Join(rootDir, "field-types.db")))
	vm.dispatchNativeCall(connValue.Num, "Open", nil)
	defer vm.dispatchNativeCall(connValue.Num, "Close", nil)
	vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("CREATE TABLE categories (label TEXT, rank INTEGER)")})
	vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("INSERT INTO categories VALUES ('alpha', 5)")})

	rsValue := vm.dispatchNativeCall(connValue.Num, "Execute", []Value{NewString("SELECT label, rank FROM categories")})
	rs := vm.adodbRecordsetItems[rsValue.Num]
	fields := vm.newADODBFieldsCollection(rs)
	labelField, _ := vm.dispatchADODBFieldsCollectionMethod(fields.Num, "Item", []Value{NewInteger(0)})
	rankField, _ := vm.dispatchADODBFieldsCollectionMethod(fields.Num, "Item", []Value{NewInteger(1)})
	labelType, _ := vm.dispatchADODBFieldPropertyGet(labelField.Num, "Type")
	rankType, _ := vm.dispatchADODBFieldPropertyGet(rankField.Num, "Type")

	if labelType.Type != VTInteger || labelType.Num != 200 {
		t.Fatalf("expected TEXT to expose adVarChar, got %#v", labelType)
	}
	if rankType.Type != VTInteger || rankType.Num != 3 {
		t.Fatalf("expected INTEGER to expose adInteger, got %#v", rankType)
	}
}

func TestADODBConnectionExecuteAcceptsClassicOptionalArguments(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	vm.SetHost(host)

	connValue := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	connString := "Driver={SQLite3};Data Source=" + filepath.Join(rootDir, "execute-options.db")
	vm.dispatchMemberSet(connValue.Num, "ConnectionString", NewString(connString))
	vm.dispatchNativeCall(connValue.Num, "Open", nil)
	defer vm.dispatchNativeCall(connValue.Num, "Close", nil)

	result := vm.dispatchNativeCall(connValue.Num, "Execute", []Value{
		NewString("CREATE TABLE sample (id INTEGER)"),
		NewInteger(1),
		NewInteger(0x80),
	})
	if result.Type == VTEmpty {
		conn := vm.adodbConnectionItems[connValue.Num]
		if conn != nil && len(conn.errors) > 0 {
			t.Fatalf("Execute with RecordsAffected and Options failed: %s", conn.errors[0].description)
		}
	}
}

func TestJScriptFunctionDeclarationOnBuiltInObject(t *testing.T) {
	source := `<%@ Language="JScript" %><%
function Object.isDirty(left, right) { return left != right; }
function String.format(value) { return "[" + value + "]"; }
Response.Write(Object.isDirty("a", "b") + "|" + String.format("ok"));
%>`

	if got := runASPSourceForTest(t, source); got != "true|[ok]" {
		t.Fatalf("unexpected qualified function output: %q", got)
	}
}

func TestJScriptStringContinuationAllowsTrailingWhitespace(t *testing.T) {
	source := "<%@ Language=\"JScript\" %><%\r\n" +
		"var sql = \"update accounts \\ \r\n" +
		"set enabled = 'Y' \\ \t\r\n" +
		"where account_id = 7\";\r\n" +
		"Response.Write(sql);\r\n%>"

	want := "update accounts set enabled = 'Y' where account_id = 7"
	if got := runASPSourceForTest(t, source); got != want {
		t.Fatalf("unexpected continued string output: %q", got)
	}
}

func TestJScriptCallbackMutatesCapturedOuterVariable(t *testing.T) {
	source := `<%@ Language="JScript" %><%
function invoke(callback) { callback("chosen"); }
function choose() {
    var selected = "";
    invoke(function(value) { selected = value; });
    return selected;
}
Response.Write(choose());
%>`

	if got := runASPSourceForTest(t, source); got != "chosen" {
		t.Fatalf("captured assignment did not update outer binding: %q", got)
	}
}

func TestJScriptCallbackReadsOuterLocalAfterNestedMethodCall(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var holder = {
    map: function() {
        function nested() { return []; }
        return nested();
    }
};
function load() {
    var rows = holder.map();
    return function() { return rows.length; }();
}
Response.Write(load());
%>`

	if got := runASPSourceForTest(t, source); got != "0" {
		t.Fatalf("nested callback did not observe outer local: %q", got)
	}
}

func TestJScriptEvalRetainsOuterObjectState(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var settings = eval('({"mode":"sample"})');
Response.Write(settings.mode);
%>`

	if got := runASPSourceForTest(t, source); got != "sample" {
		t.Fatalf("eval result was not retained: %q", got)
	}
}

func TestNormalizeJScriptStringContinuationsOnlyChangesStrings(t *testing.T) {
	input := "var value = \"one \\  \r\ntwo\"; // keep \\  \r\nvar next = 1;"
	want := "var value = \"one \\\r\ntwo\"; // keep \\  \r\nvar next = 1;"
	if got := normalizeJScriptStringContinuations(input); got != want {
		t.Fatalf("unexpected normalization:\nwant: %q\n got: %q", want, got)
	}
}

func TestNormalizeJScriptCollectionAssignmentAfterRegexContainingQuotes(t *testing.T) {
	input := "var token = /d{1,4}|\\\"[^\\\"]*\\\"|'[^']*'/g;\r\n" +
		"\tApplication(siteId+\"-\"+key) = JSON.stringify(value);\r\n"
	want := "var token = /d{1,4}|\\\"[^\\\"]*\\\"|'[^']*'/g;\r\n" +
		"\tApplication(siteId+\"-\"+key, JSON.stringify(value));\r\n"
	if got := normalizeJScriptCollectionAssignments(input); got != want {
		t.Fatalf("unexpected normalization:\nwant: %q\n got: %q", want, got)
	}
}

func TestJScriptCollectionAssignmentAfterWhitespaceStringContinuation(t *testing.T) {
	source := "<%@ Language=\"JScript\" %><%\r\n" +
		"var sql = \"one \\ \r\ntwo\";\r\n" +
		"Application(\"key\") = sql;\r\n" +
		"Response.Write(Application(\"key\"));\r\n%>"
	if got := runASPSourceForTest(t, source); got != "one two" {
		t.Fatalf("unexpected collection value after continued string: %q", got)
	}
}

func TestJScriptArrayPrototypeSliceCallOnArguments(t *testing.T) {
	source := `<%@ Language="JScript" %><%
function collect(first) {
    return Array.prototype.slice.call(arguments, 1).join("|");
}
Response.Write(collect("skip", "one", "two"));
%>`

	if got := runASPSourceForTest(t, source); got != "one|two" {
		t.Fatalf("unexpected arguments slice output: %q", got)
	}
}

func TestJScriptCommentTemplateApply(t *testing.T) {
	source := `<%@ Language="JScript" %><%
function substitute(text, value) {
    return text.replace("{value}", value);
};
function render(source, value) {
    var match = /\/\*([\s\S]+)\*\//.exec(source.toString());
    return substitute.apply(this, [match[1], value]);
}
Response.Write(render(function() {/*<b>{value}</b>*/}, "example"));
%>`

	if got := runASPSourceForTest(t, source); got != "<b>example</b>" {
		t.Fatalf("unexpected heading template output: %q", got)
	}
}

func TestJScriptNestedCommentTemplateApply(t *testing.T) {
	source := `<%@ Language="JScript" %><%
function substitute(text, value) {
    return text.replace("{value}", value);
};
function render(source, value) {
    var match = /\/\*([\s\S]+)\*\//.exec(source.toString());
    return substitute.apply(this, [match[1], value]);
}
function renderLabel(value) {
    return render(function() {/*<span>{value}</span>*/}, value);
}
Response.Write(renderLabel("example"));
%>`

	if got := runASPSourceForTest(t, source); got != "<span>example</span>" {
		t.Fatalf("unexpected nested heading template output: %q", got)
	}
}

func TestJScriptFunctionCanContainInterleavedASPHTML(t *testing.T) {
	source := `<%@ Language="JScript" %><%
function displayError(key) {
%><div class="error"><%=Server.HTMLEncode(key)%></div><%
}
displayError("<bad>");
%>`

	if got := strings.TrimSpace(runASPSourceForTest(t, source)); got != `<div class="error">&lt;bad&gt;</div>` {
		t.Fatalf("unexpected interleaved HTML output: %q", got)
	}
}

func TestJScriptMissingSessionValueIsUndefined(t *testing.T) {
	source := `<%@ Language="JScript" %><%
Response.Write(typeof(Session("missing")));
%>`

	if got := runASPSourceForTest(t, source); got != "undefined" {
		t.Fatalf("unexpected missing Session value type: %q", got)
	}
}

func TestJScriptASPExpressionSuppressesUndefinedReturnValue(t *testing.T) {
	source := `<%@ Language="JScript" %><%
function render() { Response.Write("content"); }
%><html><head><%= render() %></head><body><%= undefined %></body></html>`
	if got := runASPSourceForTest(t, source); got != "<html><head>content</head><body></body></html>" {
		t.Fatalf("unexpected ASP expression output: %q", got)
	}
}

func TestJScriptCanWrapNativeJSONParse(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var originalParse = JSON.parse;
JSON.parse = function(text) {
    if (typeof text === "object") return text;
    return originalParse(text, function(key, value) { return value; });
};
var value = { offset: 600 };
var parsed = JSON.parse(value);
var parsedText = JSON.parse('{"name":"native"}');
Response.Write(typeof(originalParse) + "|" + typeof(parsed) + "|" + parsed.offset + "|" + parsedText.name);
%>`
	if got := runASPSourceForTest(t, source); got != "function|object|600|native" {
		t.Fatalf("unexpected wrapped JSON.parse output: %q", got)
	}
}

func TestJScriptServerMapPathUsesIncludedSourceLocation(t *testing.T) {
	root := t.TempDir()
	includeDir := filepath.Join(root, "includes", "shared")
	if err := os.MkdirAll(includeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	includePath := filepath.Join(includeDir, "global.inc")
	if err := os.WriteFile(includePath, []byte(`<script language="JScript" runat="server">
function mappedSetting() { return Server.MapPath("../../settings.json"); }
</script>`), 0o600); err != nil {
		t.Fatal(err)
	}
	pagePath := filepath.Join(root, "global.asa")
	source := `<!--#include file="includes/shared/global.inc" -->
<script language="JScript" runat="server">
Response.Write(mappedSetting());
</script>`
	if err := os.WriteFile(pagePath, []byte(source), 0o600); err != nil {
		t.Fatal(err)
	}
	compiler := NewASPCompiler(source)
	compiler.SetSourceName(pagePath)
	compiler.SetIncludeSiteRoot(root)
	if err := compiler.Compile(); err != nil {
		t.Fatal(err)
	}
	vm := AcquireVMFromCompiler(compiler)
	defer vm.Release()
	host := NewMockHost()
	host.Server().SetRootDir(root)
	host.Server().SetRequestPath("/global.asa")
	var output strings.Builder
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatal(err)
	}
	want := filepath.Join(root, "settings.json")
	if !strings.EqualFold(filepath.Clean(output.String()), filepath.Clean(want)) {
		t.Fatalf("MapPath resolved to %q, want %q", output.String(), want)
	}
}

func TestJScriptIncludedFunctionAssignsParsedSettingsToApplication(t *testing.T) {
	root := t.TempDir()
	includeDir := filepath.Join(root, "includes", "shared")
	if err := os.MkdirAll(includeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	include := `<script language="JScript" runat="server">
function publishSettings() {
    var fso = Server.CreateObject("Scripting.FileSystemObject");
    var stream = fso.OpenTextFile(Server.MapPath("../../settings.json"), 1, false);
    var settings = eval("(" + stream.ReadAll() + ")");
    stream.Close();
    Application("example-setting") = settings.value;
}
</script>`
	if err := os.WriteFile(filepath.Join(includeDir, "global.inc"), []byte(include), 0o600); err != nil {
		t.Fatal(err)
	}
	settingsJSON := `{"value":"loaded"}`
	if err := os.WriteFile(filepath.Join(root, "settings.json"), []byte(settingsJSON), 0o600); err != nil {
		t.Fatal(err)
	}
	pagePath := filepath.Join(root, "global.asa")
	source := `<!--#include file="includes/shared/global.inc" -->
<script language="JScript" runat="server">function Application_OnStart() { publishSettings(); }</script>`
	if err := os.WriteFile(pagePath, []byte(source), 0o600); err != nil {
		t.Fatal(err)
	}
	app := asp.NewApplication()
	globalASA := &GlobalASA{}
	if err := globalASA.LoadAndCompile(root, app); err != nil {
		t.Fatal(err)
	}
	host := NewMockHost()
	host.Server().SetRootDir(root)
	host.Server().SetRequestPath("/global.asa")
	host.SetApplication(app)
	if err := globalASA.ExecuteApplicationOnStart(host); err != nil {
		t.Fatal(err)
	}
	got, ok := app.Get("example-setting")
	if !ok || got.Str != "loaded" {
		t.Fatalf("unexpected Application value %#v, present=%v", got, ok)
	}
}

package server

import (
	"net/http/httptest"
	"testing"
	"time"

	"g3pix.com.br/axonasp/vbscript/ast"
)

func TestAssignmentSetNothingReleasesUnreferencedADODBRecordset(t *testing.T) {
	recorder := httptest.NewRecorder()
	request := httptest.NewRequest("GET", "http://localhost/tests", nil)
	ctx := NewExecutionContext(recorder, request, "test-set-nothing", 5*time.Second)
	ctx.PushScope()

	conn := NewADODBConnection(nil)
	conn.CallMethod("open", "sqlite::memory:")
	if conn.State != 1 || conn.db == nil {
		t.Fatalf("expected open sqlite connection before Set ... = Nothing")
	}
	defer conn.CallMethod("close")

	_, _ = conn.execStatement("CREATE TABLE t (id INTEGER)", nil)
	_, _ = conn.execStatement("INSERT INTO t (id) VALUES (1)", nil)

	rs := NewADODBRecordset(nil)
	if err := rs.openRecordsetWithParams("SELECT id FROM t", conn, nil); err != nil {
		t.Fatalf("openRecordsetWithParams failed: %v", err)
	}
	if rs.State != 1 {
		t.Fatalf("expected open recordset before Set ... = Nothing")
	}

	if err := ctx.DefineVariable("rs", rs); err != nil {
		t.Fatalf("define variable failed: %v", err)
	}

	visitor := &ASPVisitor{context: ctx}
	stmt := &ast.AssignmentStatement{
		Left:  &ast.Identifier{Name: "rs"},
		Right: &ast.NothingLiteral{},
	}

	if err := visitor.visitAssignment(stmt); err != nil {
		t.Fatalf("assignment failed: %v", err)
	}

	if rs.State != 0 {
		t.Fatalf("expected recordset to be closed after Set ... = Nothing, got state=%d", rs.State)
	}
}

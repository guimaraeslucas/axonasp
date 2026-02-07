package server

import "testing"

func TestSessionIDIsNumeric(t *testing.T) {
	session := NewSessionObject("AXON-test-session", map[string]interface{}{})
	value := session.GetProperty("SessionID")

	id, ok := value.(int64)
	if !ok {
		t.Fatalf("expected int64 SessionID, got %T", value)
	}
	if id == 0 {
		t.Fatalf("expected non-zero numeric SessionID")
	}

	session2 := NewSessionObject("AXON-test-session", map[string]interface{}{})
	id2, _ := session2.GetProperty("SessionID").(int64)
	if id2 != id {
		t.Fatalf("expected stable SessionID, got %d and %d", id, id2)
	}
}

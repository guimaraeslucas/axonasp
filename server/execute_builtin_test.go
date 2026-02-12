package server

import (
	"bytes"
	"encoding/binary"
	"net/http/httptest"
	"strings"
	"testing"
	"time"
	"unicode/utf16"
)

func TestExecuteBuiltIn_NormalizesUTF16CodeArgument(t *testing.T) {
	req := httptest.NewRequest("GET", "http://localhost:4050/test.asp", nil)
	rec := httptest.NewRecorder()
	ctx := NewExecutionContext(rec, req, "test-session", 30*time.Second)

	code := `dynamic_value = "Hello World from script constant"`
	u16 := utf16.Encode([]rune(code))
	u16 = append(u16, 0, 0)

	_, handled := EvalBuiltInFunction("execute", []interface{}{u16}, ctx)
	if !handled {
		t.Fatalf("Execute built-in was not handled")
	}

	val, exists := ctx.GetVariable("dynamic_value")
	if !exists {
		t.Fatalf("expected dynamic_value to be defined after Execute")
	}
	if got := toString(val); got != "Hello World from script constant" {
		t.Fatalf("expected dynamic_value to be %q, got %q", "Hello World from script constant", got)
	}
}

func TestExecuteBuiltIn_NormalizesNullBytesInStringArgument(t *testing.T) {
	req := httptest.NewRequest("GET", "http://localhost:4050/test.asp", nil)
	rec := httptest.NewRecorder()
	ctx := NewExecutionContext(rec, req, "test-session", 30*time.Second)

	codeWithNull := "dynamic_value2 = \"ok\"\x00"

	_, handled := EvalBuiltInFunction("execute", []interface{}{codeWithNull}, ctx)
	if !handled {
		t.Fatalf("Execute built-in was not handled")
	}

	val, exists := ctx.GetVariable("dynamic_value2")
	if !exists {
		t.Fatalf("expected dynamic_value2 to be defined after Execute")
	}
	if got := toString(val); got != "ok" {
		t.Fatalf("expected dynamic_value2 to be %q, got %q", "ok", got)
	}
}

func TestExecuteBuiltIn_DecodesUTF16LEBytePayload(t *testing.T) {
	req := httptest.NewRequest("GET", "http://localhost:4050/test.asp", nil)
	rec := httptest.NewRecorder()
	ctx := NewExecutionContext(rec, req, "test-session", 30*time.Second)

	code := `dynamic_value3 = "decoded"`
	u16 := utf16.Encode([]rune(code))
	buf := bytes.NewBuffer(nil)
	for _, unit := range u16 {
		if err := binary.Write(buf, binary.LittleEndian, unit); err != nil {
			t.Fatalf("failed to build UTF-16LE payload: %v", err)
		}
	}
	payload := append(buf.Bytes(), 0x00, 0x00)

	_, handled := EvalBuiltInFunction("execute", []interface{}{payload}, ctx)
	if !handled {
		t.Fatalf("Execute built-in was not handled")
	}

	val, exists := ctx.GetVariable("dynamic_value3")
	if !exists {
		t.Fatalf("expected dynamic_value3 to be defined after Execute")
	}
	if got := toString(val); got != "decoded" {
		t.Fatalf("expected dynamic_value3 to be %q, got %q", "decoded", got)
	}
}

func TestToString_DecodesUTF16LEBytes(t *testing.T) {
	code := `execute_me`
	u16 := utf16.Encode([]rune(code))
	buf := bytes.NewBuffer(nil)
	for _, unit := range u16 {
		if err := binary.Write(buf, binary.LittleEndian, unit); err != nil {
			t.Fatalf("failed to build UTF-16LE payload: %v", err)
		}
	}
	payload := append(buf.Bytes(), 0x00, 0x00)

	got := toString(payload)
	if got != code {
		t.Fatalf("expected decoded string %q, got %q", code, got)
	}
}

func TestSplitBuiltIn_EmptyDelimiterReturnsWholeString(t *testing.T) {
	req := httptest.NewRequest("GET", "http://localhost:4050/test.asp", nil)
	rec := httptest.NewRecorder()
	ctx := NewExecutionContext(rec, req, "test-session", 30*time.Second)

	result, handled := EvalBuiltInFunction("split", []interface{}{"abc", ""}, ctx)
	if !handled {
		t.Fatalf("Split built-in was not handled")
	}

	arr, ok := toVBArray(result)
	if !ok {
		t.Fatalf("expected split result to be array, got %T", result)
	}
	if arr.Len() != 1 {
		t.Fatalf("expected split length 1, got %d", arr.Len())
	}
	val, exists := arr.Get(0)
	if !exists {
		t.Fatalf("expected first element to exist")
	}
	if got := toString(val); got != "abc" {
		t.Fatalf("expected first element %q, got %q", "abc", got)
	}
}

func TestReplaceBuiltIn_ScriptRewritePattern(t *testing.T) {
	req := httptest.NewRequest("GET", "http://localhost:4050/test.asp", nil)
	rec := httptest.NewRecorder()
	ctx := NewExecutionContext(rec, req, "test-session", 30*time.Second)

	script := `for i=6 to 1 step -1 response.write "<h"&i&">Hello World!</h"&i&">" next`
	replaced, handled := EvalBuiltInFunction("replace", []interface{}{script, "response.write", "CustomFunction=CustomFunction&", 1, -1, 1}, ctx)
	if !handled {
		t.Fatalf("Replace built-in was not handled")
	}
	out := toString(replaced)
	if out == "" {
		t.Fatalf("replace result should not be empty")
	}
	if !strings.Contains(out, "CustomFunction=CustomFunction&") {
		t.Fatalf("replace result did not contain replacement token: %q", out)
	}
}

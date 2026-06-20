package axonvm

import (
	"bytes"
	"encoding/binary"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestCompileWScriptShellPage_GlobalIncIndexesInRange(t *testing.T) {
	sourcePath := filepath.Join("..", "www", "tests", "test_wscriptshell.asp")
	source, err := os.ReadFile(sourcePath)
	if err != nil {
		t.Fatalf("read source: %v", err)
	}

	compiler := NewASPCompiler(string(source))
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	globalsCount := compiler.GlobalsCount()
	bytecode := compiler.Bytecode()
	for ip := 0; ip < len(bytecode); {
		op := OpCode(bytecode[ip])
		size := opcodeOperandSize(op, bytecode, ip)
		instrEnd := ip + 1 + size
		if instrEnd > len(bytecode) {
			t.Fatalf("invalid instruction size at ip %d for %s", ip, op.String())
		}

		if op == OpIncGlobalInt || op == OpDecGlobalInt {
			if ip+3 > len(bytecode) {
				t.Fatalf("truncated global increment instruction at ip %d", ip)
			}
			idx := int(binary.BigEndian.Uint16(bytecode[ip+1 : ip+3]))
			if idx < 0 || idx >= globalsCount {
				t.Fatalf("out-of-range global index %d at ip %d for %s (globals=%d)", idx, ip, op.String(), globalsCount)
			}
		}

		ip = instrEnd
	}
}

func TestRunWScriptShellPage_NoBytecodeCorruption(t *testing.T) {
	sourcePath := filepath.Join("..", "www", "tests", "test_wscriptshell.asp")
	source, err := os.ReadFile(sourcePath)
	if err != nil {
		t.Fatalf("read source: %v", err)
	}

	compiler := NewASPCompiler(string(source))
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := append([]byte(nil), compiler.Bytecode()...)
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	err = vm.Run()
	if err == nil {
		return
	}

	errText := err.Error()
	if !strings.Contains(errText, "index out of range") {
		t.Fatalf("unexpected runtime error: %v", err)
	}

	ip := vm.ip
	start := max(ip-12, 0)
	end := min(ip+12, len(vm.bytecode))

	before := bytecode[start:end]
	after := vm.bytecode[start:end]
	t.Fatalf("runtime panic persisted at ip=%d lastLine=%d nextOp=%s\nwindow[%d:%d] before=%v\nwindow[%d:%d] after =%v\nerror=%v", ip, vm.lastLine, func() string {
		if ip >= 0 && ip < len(vm.bytecode) {
			return OpCode(vm.bytecode[ip]).String()
		}
		return "<out-of-range>"
	}(), start, end, before, start, end, after, err)
}

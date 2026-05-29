package axonvm

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/jscript"
	"g3pix.com.br/axonasp/vbscript"
)

func buildLargeIncludeContent(lines int) string {
	if lines <= 0 {
		return ""
	}
	var builder strings.Builder
	builder.Grow(lines * 8)
	for i := 1; i <= lines; i++ {
		builder.WriteString("line")
		builder.WriteString(fmt.Sprintf("%d", i))
		builder.WriteString("\r\n")
	}
	return builder.String()
}

func TestSourceMapResolveLineUsesSparseBoundaries(t *testing.T) {
	m := SourceMap{}
	m.AddBoundary(1, "main.asp", 1)
	m.AddBoundary(182, "include.inc", 1)
	m.AddBoundary(363, "main.asp", 2)

	file, line, ok := m.ResolveLine(189)
	if !ok {
		t.Fatal("expected line to resolve")
	}
	if file != "include.inc" {
		t.Fatalf("expected include.inc, got %q", file)
	}
	if line != 8 {
		t.Fatalf("expected include local line 8, got %d", line)
	}

	file, line, ok = m.ResolveLine(365)
	if !ok {
		t.Fatal("expected main line to resolve")
	}
	if file != "main.asp" {
		t.Fatalf("expected main.asp, got %q", file)
	}
	if line != 4 {
		t.Fatalf("expected main local line 4, got %d", line)
	}
}

func TestCountLineBreaksMixedEndings(t *testing.T) {
	got := countLineBreaks("a\r\nb\nc\rd\r\ne")
	if got != 4 {
		t.Fatalf("expected 4 line breaks, got %d", got)
	}
}

func TestASPIncludeRuntimeErrorMapsBackToMainFileLine(t *testing.T) {
	tmpDir := t.TempDir()
	mainPath := filepath.Join(tmpDir, "main.asp")
	includePath := filepath.Join(tmpDir, "big.inc")

	if err := os.WriteFile(includePath, []byte(buildLargeIncludeContent(181)), 0o600); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	mainSource := strings.Join([]string{
		"<!--#include file=\"big.inc\"-->",
		"<%",
		"Dim a",
		"a = 1",
		"Dim b",
		"b = 0",
		"Dim c",
		"c = a / b",
		"Response.Write c",
		"%>",
	}, "\r\n")
	if err := os.WriteFile(mainPath, []byte(mainSource), 0o600); err != nil {
		t.Fatalf("write main failed: %v", err)
	}

	compiler := NewASPCompiler(mainSource)
	compiler.SetSourceName(mainPath)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	vm.SetHost(host)

	runErr := vm.Run()
	if runErr == nil {
		t.Fatal("expected runtime error")
	}

	var vmErr *VMError
	if !errors.As(runErr, &vmErr) {
		t.Fatalf("expected VMError, got %T: %v", runErr, runErr)
	}
	if !strings.EqualFold(filepath.Clean(vmErr.File), filepath.Clean(mainPath)) {
		t.Fatalf("expected mapped runtime file %q, got %q", mainPath, vmErr.File)
	}
	if vmErr.Line != 8 {
		t.Fatalf("expected mapped runtime line 8, got %d", vmErr.Line)
	}

	aspErr := RuntimeErrorToASPError(runErr, mainPath)
	if !strings.EqualFold(filepath.Clean(aspErr.File), filepath.Clean(mainPath)) {
		t.Fatalf("expected ASPError file %q, got %q", mainPath, aspErr.File)
	}
	if aspErr.Line != 8 {
		t.Fatalf("expected ASPError line 8, got %d", aspErr.Line)
	}
}

func TestASPIncludeJScriptCompileErrorMapsBackToMainFileLine(t *testing.T) {
	tmpDir := t.TempDir()
	mainPath := filepath.Join(tmpDir, "main.asp")
	includePath := filepath.Join(tmpDir, "big.inc")

	if err := os.WriteFile(includePath, []byte(buildLargeIncludeContent(181)), 0o600); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	mainSource := strings.Join([]string{
		"<!--#include file=\"big.inc\"-->",
		"<script runat=\"server\" language=\"JScript\">",
		"var ok = 1;",
		"var broken = ;",
		"</script>",
	}, "\r\n")
	if err := os.WriteFile(mainPath, []byte(mainSource), 0o600); err != nil {
		t.Fatalf("write main failed: %v", err)
	}

	compiler := NewASPCompiler(mainSource)
	compiler.SetSourceName(mainPath)
	err := compiler.Compile()
	if err == nil {
		t.Fatal("expected compile error")
	}

	var syntaxErr *jscript.JSSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected JSSyntaxError, got %T: %v", err, err)
	}

	if !strings.EqualFold(filepath.Clean(syntaxErr.File), filepath.Clean(mainPath)) {
		t.Fatalf("expected mapped file %q, got %q", mainPath, syntaxErr.File)
	}
	if syntaxErr.Line != 4 {
		t.Fatalf("expected mapped line 4, got %d", syntaxErr.Line)
	}
}

func TestASPDefaultJScriptIncludeCompileErrorMapsBackToMainFileLine(t *testing.T) {
	tmpDir := t.TempDir()
	mainPath := filepath.Join(tmpDir, "main.asp")
	includePath := filepath.Join(tmpDir, "big.inc")

	if err := os.WriteFile(includePath, []byte(buildLargeIncludeContent(181)), 0o600); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	mainSource := strings.Join([]string{
		"<%@ Language=\"JScript\" %>",
		"<!--#include file=\"big.inc\"-->",
		"<%",
		"var ok = 1;",
		"var broken = ;",
		"%>",
	}, "\r\n")
	if err := os.WriteFile(mainPath, []byte(mainSource), 0o600); err != nil {
		t.Fatalf("write main failed: %v", err)
	}

	compiler := NewASPCompiler(mainSource)
	compiler.SetSourceName(mainPath)
	err := compiler.Compile()
	if err == nil {
		t.Fatal("expected compile error")
	}

	var syntaxErr *jscript.JSSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected JSSyntaxError, got %T: %v", err, err)
	}

	if !strings.EqualFold(filepath.Clean(syntaxErr.File), filepath.Clean(mainPath)) {
		t.Fatalf("expected mapped file %q, got %q", mainPath, syntaxErr.File)
	}
	if syntaxErr.Line != 5 {
		t.Fatalf("expected mapped line 5, got %d", syntaxErr.Line)
	}
}

func TestASPIncludeVBCompileErrorMissingRHSMapsToMainLine(t *testing.T) {
	tmpDir := t.TempDir()
	mainPath := filepath.Join(tmpDir, "main.asp")
	includePath := filepath.Join(tmpDir, "big.inc")

	if err := os.WriteFile(includePath, []byte(buildLargeIncludeContent(181)), 0o600); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	mainSource := strings.Join([]string{
		"<%@ LANGUAGE=VBScript %>",
		"<!--#include file=\"big.inc\"-->",
		"<%",
		"Dim a",
		"a = 1",
		"Dim b",
		"b = 2",
		"Dim c",
		"c =",
		"%>",
	}, "\r\n")
	if err := os.WriteFile(mainPath, []byte(mainSource), 0o600); err != nil {
		t.Fatalf("write main failed: %v", err)
	}

	compiler := NewASPCompiler(mainSource)
	compiler.SetSourceName(mainPath)
	err := compiler.Compile()
	if err == nil {
		t.Fatal("expected compile error")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBSyntaxError, got %T: %v", err, err)
	}

	if !strings.EqualFold(filepath.Clean(syntaxErr.File), filepath.Clean(mainPath)) {
		t.Fatalf("expected mapped file %q, got %q", mainPath, syntaxErr.File)
	}
	if syntaxErr.Line != 9 {
		t.Fatalf("expected mapped line 9, got %d", syntaxErr.Line)
	}
	if syntaxErr.Column < 2 {
		t.Fatalf("expected actionable column at or after assignment target, got %d", syntaxErr.Column)
	}
}

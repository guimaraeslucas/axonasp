package axonvm

import (
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func TestResolveIncludeVirtualAnchorsToConfiguredRoot(t *testing.T) {
	rootDir := t.TempDir()
	includePath := filepath.Join(rootDir, "includes", "header.inc")
	if err := os.MkdirAll(filepath.Dir(includePath), 0o755); err != nil {
		t.Fatalf("mkdir include dir failed: %v", err)
	}
	if err := os.WriteFile(includePath, []byte("header"), 0o644); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	sourcePath := filepath.Join(rootDir, "pages", "default.asp")
	if err := os.MkdirAll(filepath.Dir(sourcePath), 0o755); err != nil {
		t.Fatalf("mkdir source dir failed: %v", err)
	}
	if err := os.WriteFile(sourcePath, []byte("<% 'test %>"), 0o644); err != nil {
		t.Fatalf("write source failed: %v", err)
	}

	resolved, err := resolveIncludePathWithOptions(sourcePath, "/includes/header.inc", true, includeResolveOptions{siteRoot: rootDir, caseInsensitive: true})
	if err != nil {
		t.Fatalf("resolve include virtual failed: %v", err)
	}
	if filepath.Clean(resolved) != filepath.Clean(includePath) {
		t.Fatalf("unexpected resolved path: got %q want %q", resolved, includePath)
	}
}

func TestResolveIncludeVirtualRejectsParentEscape(t *testing.T) {
	rootDir := t.TempDir()
	sourcePath := filepath.Join(rootDir, "pages", "default.asp")
	if err := os.MkdirAll(filepath.Dir(sourcePath), 0o755); err != nil {
		t.Fatalf("mkdir source dir failed: %v", err)
	}
	if err := os.WriteFile(sourcePath, []byte("<% 'test %>"), 0o644); err != nil {
		t.Fatalf("write source failed: %v", err)
	}

	_, err := resolveIncludePathWithOptions(sourcePath, "/../secret.inc", true, includeResolveOptions{siteRoot: rootDir, caseInsensitive: true})
	if err == nil {
		t.Fatal("expected virtual include traversal error")
	}
}

func TestResolveIncludeFileDisallowsAbsolutePath(t *testing.T) {
	rootDir := t.TempDir()
	sourcePath := filepath.Join(rootDir, "pages", "default.asp")
	if err := os.MkdirAll(filepath.Dir(sourcePath), 0o755); err != nil {
		t.Fatalf("mkdir source dir failed: %v", err)
	}
	if err := os.WriteFile(sourcePath, []byte("<% 'test %>"), 0o644); err != nil {
		t.Fatalf("write source failed: %v", err)
	}

	absoluteInclude := filepath.Join(rootDir, "includes", "header.inc")
	if err := os.MkdirAll(filepath.Dir(absoluteInclude), 0o755); err != nil {
		t.Fatalf("mkdir include dir failed: %v", err)
	}
	if err := os.WriteFile(absoluteInclude, []byte("header"), 0o644); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	_, err := resolveIncludePathWithOptions(sourcePath, absoluteInclude, false, includeResolveOptions{siteRoot: rootDir, caseInsensitive: true})
	if err == nil {
		t.Fatal("expected absolute file include to be rejected")
	}
}

func TestResolveIncludeCaseInsensitiveFallback(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("case-insensitive fallback is native on Windows")
	}

	rootDir := t.TempDir()
	sourcePath := filepath.Join(rootDir, "pages", "default.asp")
	if err := os.MkdirAll(filepath.Dir(sourcePath), 0o755); err != nil {
		t.Fatalf("mkdir source dir failed: %v", err)
	}
	if err := os.WriteFile(sourcePath, []byte("<% 'test %>"), 0o644); err != nil {
		t.Fatalf("write source failed: %v", err)
	}

	includePath := filepath.Join(filepath.Dir(sourcePath), "header.inc")
	if err := os.WriteFile(includePath, []byte("header"), 0o644); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	resolved, err := resolveIncludePathWithOptions(sourcePath, "HeAdEr.InC", false, includeResolveOptions{siteRoot: rootDir, caseInsensitive: true})
	if err != nil {
		t.Fatalf("case-insensitive include resolution failed: %v", err)
	}
	if filepath.Clean(resolved) != filepath.Clean(includePath) {
		t.Fatalf("unexpected resolved path: got %q want %q", resolved, includePath)
	}
}

func TestResolveIncludeFileUsesCurrentScriptDirectory(t *testing.T) {
	scriptRoot := t.TempDir()
	wrongSiteRoot := t.TempDir()

	sourcePath := filepath.Join(scriptRoot, "pages", "index.asp")
	if err := os.MkdirAll(filepath.Dir(sourcePath), 0o755); err != nil {
		t.Fatalf("mkdir source dir failed: %v", err)
	}
	if err := os.WriteFile(sourcePath, []byte("<% 'test %>"), 0o644); err != nil {
		t.Fatalf("write source failed: %v", err)
	}

	includePath := filepath.Join(scriptRoot, "pages", "includes", "helpers.asp")
	if err := os.MkdirAll(filepath.Dir(includePath), 0o755); err != nil {
		t.Fatalf("mkdir include dir failed: %v", err)
	}
	if err := os.WriteFile(includePath, []byte("helpers"), 0o644); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	resolved, err := resolveIncludePathWithOptions(sourcePath, "includes/helpers.asp", false, includeResolveOptions{siteRoot: wrongSiteRoot, caseInsensitive: true})
	if err != nil {
		t.Fatalf("resolve file include failed: %v", err)
	}
	if filepath.Clean(resolved) != filepath.Clean(includePath) {
		t.Fatalf("unexpected resolved path: got %q want %q", resolved, includePath)
	}
}

func TestCompilerCompilesFileIncludeFromScriptDirectory(t *testing.T) {
	scriptRoot := t.TempDir()
	wrongSiteRoot := t.TempDir()

	sourcePath := filepath.Join(scriptRoot, "pages", "index.asp")
	if err := os.MkdirAll(filepath.Dir(sourcePath), 0o755); err != nil {
		t.Fatalf("mkdir source dir failed: %v", err)
	}

	includePath := filepath.Join(scriptRoot, "pages", "includes", "helpers.asp")
	if err := os.MkdirAll(filepath.Dir(includePath), 0o755); err != nil {
		t.Fatalf("mkdir include dir failed: %v", err)
	}
	if err := os.WriteFile(includePath, []byte("<% Response.Write \"ok\" %>"), 0o644); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	source := "<!--#include file=\"includes/helpers.asp\"-->"
	compiler := NewASPCompiler(source)
	compiler.SetSourceName(sourcePath)
	compiler.SetIncludeSiteRoot(wrongSiteRoot)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
}

func TestCompilerCompilesRealPageFileIncludes(t *testing.T) {
	pagePath := filepath.Join("..", "www", "tests", "tests2", "index.asp")
	content, err := os.ReadFile(pagePath)
	if err != nil {
		t.Fatalf("read sample page failed: %v", err)
	}

	wrongSiteRoot := t.TempDir()
	compiler := NewASPCompiler(string(content))
	compiler.SetSourceName(pagePath)
	compiler.SetIncludeSiteRoot(wrongSiteRoot)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
}

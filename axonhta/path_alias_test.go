package main

import (
	"os"
	"path/filepath"
	"testing"
)

func TestAliasFlag_Set(t *testing.T) {
	var af aliasFlag
	err := af.Set("/music/=D:\\Music")
	if err != nil {
		t.Fatal(err)
	}
	if len(af) != 1 {
		t.Fatalf("expected 1 alias, got %d", len(af))
	}
	if af[0].VirtualPrefix != "/music/" {
		t.Errorf("expected prefix '/music/', got %q", af[0].VirtualPrefix)
	}
	if af[0].RealDir != "D:\\Music" {
		t.Errorf("expected real dir 'D:\\Music', got %q", af[0].RealDir)
	}
}

func TestAliasFlag_SetAutoSlash(t *testing.T) {
	var af aliasFlag
	err := af.Set("photos=E:\\Photos")
	if err != nil {
		t.Fatal(err)
	}
	if af[0].VirtualPrefix != "/photos/" {
		t.Errorf("expected prefix '/photos/', got %q", af[0].VirtualPrefix)
	}
}

func TestAliasFlag_InvalidFormat(t *testing.T) {
	var af aliasFlag
	err := af.Set("noequals")
	if err == nil {
		t.Error("expected error for invalid format")
	}
}

func TestAliasFlag_EmptyParts(t *testing.T) {
	var af aliasFlag
	err := af.Set("=D:\\Music")
	if err == nil {
		t.Error("expected error for empty prefix")
	}
	err = af.Set("/music/=")
	if err == nil {
		t.Error("expected error for empty path")
	}
}

func TestLoadPathAliases_FromFile(t *testing.T) {
	dir := t.TempDir()
	dataDir := filepath.Join(dir, "data")
	os.MkdirAll(dataDir, 0755)

	content := `; Comment line
/music/|D:\Music
# Another comment
/photos/|E:\Photos
`
	os.WriteFile(filepath.Join(dataDir, "path_aliases.dat"), []byte(content), 0644)

	// Reset global state
	cliAliases = nil
	err := LoadPathAliases(dir)
	if err != nil {
		t.Fatal(err)
	}

	pathAliasesMu.RLock()
	defer pathAliasesMu.RUnlock()
	if len(pathAliases) != 2 {
		t.Fatalf("expected 2 aliases, got %d", len(pathAliases))
	}
	if pathAliases[0].VirtualPrefix != "/music/" {
		t.Errorf("expected first alias '/music/', got %q", pathAliases[0].VirtualPrefix)
	}
	if pathAliases[1].VirtualPrefix != "/photos/" {
		t.Errorf("expected second alias '/photos/', got %q", pathAliases[1].VirtualPrefix)
	}
}

func TestLoadPathAliases_NoFile(t *testing.T) {
	dir := t.TempDir()
	cliAliases = nil
	err := LoadPathAliases(dir)
	if err != nil {
		t.Fatal(err)
	}
	pathAliasesMu.RLock()
	defer pathAliasesMu.RUnlock()
	if len(pathAliases) != 0 {
		t.Errorf("expected 0 aliases, got %d", len(pathAliases))
	}
}

func TestMergeAliases_CLIOverridesFile(t *testing.T) {
	file := []PathAlias{
		{VirtualPrefix: "/music/", RealDir: "D:\\FileMusic"},
		{VirtualPrefix: "/photos/", RealDir: "E:\\Photos"},
	}
	cli := []PathAlias{
		{VirtualPrefix: "/music/", RealDir: "C:\\CLIMusic"},
	}
	result := mergeAliases(file, cli)
	if len(result) != 2 {
		t.Fatalf("expected 2 aliases, got %d", len(result))
	}
	// CLI should override file for /music/
	for _, a := range result {
		if a.VirtualPrefix == "/music/" && a.RealDir != "C:\\CLIMusic" {
			t.Errorf("expected CLI to override /music/ to 'C:\\CLIMusic', got %q", a.RealDir)
		}
	}
}

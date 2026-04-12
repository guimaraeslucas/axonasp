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
	"encoding/base64"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"
)

// TestAxVersionUsesRuntimeVersion verifies AxVersion returns the build/runtime version.
func TestAxVersionUsesRuntimeVersion(t *testing.T) {
	previous := GetRuntimeVersion()
	defer SetRuntimeVersion(previous)

	SetRuntimeVersion("9.8.7.6")
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("AxVersion", nil)
	if got.Type != VTString || got.String() != "9.8.7.6" {
		t.Fatalf("unexpected AxVersion result: %#v", got)
	}
}

// TestAxShutdownRespectsDisabledConfig verifies shutdown remains blocked by default config.
func TestAxShutdownRespectsDisabledConfig(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("AxShutdownAxonASPServer", nil)
	if got.Type != VTBool || got.Num != 0 {
		t.Fatalf("expected AxShutdownAxonASPServer to be disabled, got %#v", got)
	}
}

// TestAxDateUsesLocaleNames verifies AxDate uses locale-sensitive month and weekday names.
func TestAxDateUsesLocaleNames(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	host := NewMockHost()
	host.Session().SetLCID(int(PortugueseBrazil))
	vm.SetHost(host)
	lib := &AxonLibrary{vm: vm}

	ts := time.Date(2026, time.March, 22, 15, 4, 0, 0, time.UTC).Unix()
	got := lib.DispatchMethod("AxDate", []Value{NewString("D, d M Y H:i"), NewInteger(ts)})
	if got.Type != VTString {
		t.Fatalf("unexpected AxDate value type: %#v", got)
	}

	text := strings.ToLower(got.String())
	if !strings.Contains(text, "dom") || !strings.Contains(text, "mar") {
		t.Fatalf("expected locale-aware names in AxDate output, got %q", got.String())
	}
}

// TestInitGlobalAxonFunctionsRegistersBuiltins verifies Ax built-ins are injected once when enabled.
func TestInitGlobalAxonFunctionsRegistersBuiltins(t *testing.T) {
	InitGlobalAxonFunctions(true)

	for _, name := range AxonGlobalFunctionNames {
		if _, ok := GetBuiltinIndex(name); !ok {
			t.Fatalf("expected builtin %q to be registered", name)
		}
	}

	before := len(BuiltinNames)
	InitGlobalAxonFunctions(true)
	after := len(BuiltinNames)
	if before != after {
		t.Fatalf("expected idempotent registration, got before=%d after=%d", before, after)
	}
}

// TestAxonGlobalBuiltinRoutesToObjectLogic verifies global built-in and object dispatch share the same implementation.
func TestAxonGlobalBuiltinRoutesToObjectLogic(t *testing.T) {
	InitGlobalAxonFunctions(true)

	vm := NewVM(nil, nil, 0)
	idx, ok := GetBuiltinIndex("axversion")
	if !ok {
		t.Fatal("expected axversion builtin index")
	}

	builtinValue, err := BuiltinRegistry[idx](vm, nil)
	if err != nil {
		t.Fatalf("unexpected builtin error: %v", err)
	}

	lib := &AxonLibrary{vm: vm}
	objectValue := lib.DispatchMethod("axversion", nil)

	if builtinValue.Type != objectValue.Type || builtinValue.String() != objectValue.String() {
		t.Fatalf("expected same result, builtin=%#v object=%#v", builtinValue, objectValue)
	}
}

// TestAxGetConfigReadsGlobalKey verifies axgetconfig can read scalar values from axonasp.toml.
func TestAxGetConfigReadsGlobalKey(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axgetconfig", []Value{NewString("global.golang_memory_limit_mb")})
	if got.Type != VTInteger && got.Type != VTDouble && got.Type != VTString {
		t.Fatalf("unexpected axgetconfig type for global.golang_memory_limit_mb: %#v", got)
	}
	if got.Type == VTInteger && got.Num <= 0 {
		t.Fatalf("expected positive memory limit, got %#v", got)
	}
}

// TestAxGetConfigMissingKeyReturnsEmpty verifies unknown keys return Empty.
func TestAxGetConfigMissingKeyReturnsEmpty(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axgetconfig", []Value{NewString("global.this_key_does_not_exist")})
	if got.Type != VTEmpty {
		t.Fatalf("expected Empty for missing config key, got %#v", got)
	}
}

// TestAxGetConfigEnvOverride verifies that when global.viper_automatic_env is true,
// an environment variable overrides the value stored in axonasp.toml.
func TestAxGetConfigEnvOverride(t *testing.T) {
	// The config has viper_automatic_env = true. The replacer converts dots to
	// underscores, so the env var for global.golang_memory_limit_mb becomes
	// GLOBAL_GOLANG_MEMORY_LIMIT_MB (Viper uppercases it automatically).
	t.Setenv("GLOBAL_GOLANG_MEMORY_LIMIT_MB", "77777")

	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axgetconfig", []Value{NewString("global.golang_memory_limit_mb")})
	if got.Type != VTInteger && got.Type != VTDouble && got.Type != VTString {
		t.Fatalf("unexpected type from axgetconfig: %#v", got)
	}
	strVal := got.String()
	if strVal != "77777" {
		t.Fatalf("expected env override value 77777, got %q", strVal)
	}
}

// TestAxGetConfigKeysReturnsArray verifies axgetconfigkeys returns an array of
// all configuration keys present in axonasp.toml.
func TestAxGetConfigKeysReturnsArray(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axgetconfigkeys", nil)
	if got.Type != VTArray {
		t.Fatalf("expected VTArray from axgetconfigkeys, got %#v", got)
	}
	if got.Arr == nil || got.Arr.Len() == 0 {
		t.Fatal("axgetconfigkeys returned an empty array, expected at least one key")
	}
	// Verify that a well-known key is present in the returned list.
	found := false
	length := got.Arr.Len()
	for i := 0; i < length; i++ {
		v, _ := got.Arr.Get(i)
		if v.String() == "global.golang_memory_limit_mb" {
			found = true
			break
		}
	}
	if !found {
		t.Fatal("expected 'global.golang_memory_limit_mb' to be present in axgetconfigkeys result")
	}
}

// TestAxGetDefaultCSSReadsConfiguredFile verifies axgetdefaultcss returns the full CSS file content.
func TestAxGetDefaultCSSReadsConfiguredFile(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	tmpDir := t.TempDir()
	cssPath := filepath.Join(tmpDir, "custom.css")
	cssContent := "body{background:#ECE9D8;}\n.title{color:#003399;}"
	if err := os.WriteFile(cssPath, []byte(cssContent), 0600); err != nil {
		t.Fatalf("failed to create css fixture: %v", err)
	}

	loadAxFunctionConfig()
	previousCSSPath := cachedAxFunctionConfig.defaultCSSPath
	cachedAxFunctionConfig.defaultCSSPath = cssPath
	t.Cleanup(func() {
		cachedAxFunctionConfig.defaultCSSPath = previousCSSPath
	})

	got := lib.DispatchMethod("axgetdefaultcss", nil)
	if got.Type != VTString || got.String() != cssContent {
		t.Fatalf("expected css content %q, got %#v", cssContent, got)
	}
}

// TestAxGetLogoReturnsInlineDataURI verifies axgetlogo returns a MIME-aware base64 data URI.
func TestAxGetLogoReturnsInlineDataURI(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	tmpDir := t.TempDir()
	logoPath := filepath.Join(tmpDir, "logo.png")
	logoData := []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1A, '\n', 0x00, 0x00, 0x00, 0x00}
	if err := os.WriteFile(logoPath, logoData, 0600); err != nil {
		t.Fatalf("failed to create logo fixture: %v", err)
	}

	loadAxFunctionConfig()
	previousLogoPath := cachedAxFunctionConfig.defaultLogoPath
	cachedAxFunctionConfig.defaultLogoPath = logoPath
	t.Cleanup(func() {
		cachedAxFunctionConfig.defaultLogoPath = previousLogoPath
	})

	got := lib.DispatchMethod("axgetlogo", nil)
	want := "data:image/png;base64," + base64.StdEncoding.EncodeToString(logoData)
	if got.Type != VTString || got.String() != want {
		t.Fatalf("expected logo data URI %q, got %#v", want, got)
	}
}

// TestAxHexToRGBConvertsHTMLHex verifies axhextorgb converts short and long HTML hex values.
func TestAxHexToRGBConvertsHTMLHex(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axhextorgb", []Value{NewString("#FFAA00")})
	if got.Type != VTString || got.String() != "rgb(255,170,0)" {
		t.Fatalf("unexpected rgb conversion for #FFAA00: %#v", got)
	}

	got = lib.DispatchMethod("axhextorgb", []Value{NewString("#0F8")})
	if got.Type != VTString || got.String() != "rgb(0,255,136)" {
		t.Fatalf("unexpected rgb conversion for #0F8: %#v", got)
	}
}

// TestFilterEnvironmentEntriesSkipsPseudoWindowsKeys verifies pseudo/internal entries are removed from the environment list.
func TestFilterEnvironmentEntriesSkipsPseudoWindowsKeys(t *testing.T) {
	input := []string{
		"=C:=C:\\repo",
		"=::=::\\",
		"PATH=C:\\Windows\\System32",
		"HOME=C:\\Users\\dev",
	}

	got := filterEnvironmentEntries(input)
	if len(got) != 2 {
		t.Fatalf("expected 2 filtered entries, got %d (%v)", len(got), got)
	}
	if got[0] != "PATH=C:\\Windows\\System32" {
		t.Fatalf("expected first entry to be PATH, got %q", got[0])
	}
	if got[1] != "HOME=C:\\Users\\dev" {
		t.Fatalf("expected second entry to be HOME, got %q", got[1])
	}
}

// TestAxUserHomeDirPathReturnsValue verifies axuserhomedirpath resolves a usable home directory path.
func TestAxUserHomeDirPathReturnsValue(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axuserhomedirpath", nil)
	if got.Type != VTString {
		t.Fatalf("expected string result, got %#v", got)
	}
	if strings.TrimSpace(got.String()) == "" {
		t.Fatal("expected non-empty home directory path")
	}
}

// TestAxUserConfigDirPathPointsToAxonToml verifies axuserconfigdirpath returns a config file path ending in config/axonasp.toml.
func TestAxUserConfigDirPathPointsToAxonToml(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axuserconfigdirpath", nil)
	if got.Type != VTString {
		t.Fatalf("expected string result, got %#v", got)
	}
	normalized := strings.ReplaceAll(strings.ToLower(got.String()), "\\", "/")
	if !strings.HasSuffix(normalized, "/config/axonasp.toml") {
		t.Fatalf("expected path ending in /config/axonasp.toml, got %q", got.String())
	}
}

// TestAxCacheDirPathReturnsAbsoluteCachePath verifies axcachedirpath points to the .temp/cache directory with a trailing separator.
func TestAxCacheDirPathReturnsAbsoluteCachePath(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axcachedirpath", nil)
	if got.Type != VTString {
		t.Fatalf("expected string result, got %#v", got)
	}
	path := got.String()
	if !strings.HasSuffix(path, string(os.PathSeparator)) {
		t.Fatalf("expected trailing path separator in %q", path)
	}
	normalized := strings.ReplaceAll(strings.ToLower(path), "\\", "/")
	if !strings.Contains(normalized, "/.temp/cache/") {
		t.Fatalf("expected .temp/cache path segment in %q", path)
	}
}

// TestAxIsPathSeparator verifies axispathseparator for separator and non-separator characters.
func TestAxIsPathSeparator(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	sep := lib.DispatchMethod("axispathseparator", []Value{NewString("/")})
	if runtime.GOOS == "windows" {
		if sep.Type != VTBool || sep.Num != 1 {
			t.Fatalf("expected '/' to be recognized as a separator on Windows, got %#v", sep)
		}
	}

	notSep := lib.DispatchMethod("axispathseparator", []Value{NewString("a")})
	if notSep.Type != VTBool || notSep.Num != 0 {
		t.Fatalf("expected 'a' not to be a separator, got %#v", notSep)
	}
}

// TestAxFileMutationHelpers verifies chtimes/chmod/link/chown wrappers return expected booleans.
func TestAxFileMutationHelpers(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	tmpDir := t.TempDir()
	testFile := filepath.Join(tmpDir, "file.txt")
	if err := os.WriteFile(testFile, []byte("axon"), 0600); err != nil {
		t.Fatalf("failed to create test file: %v", err)
	}

	changeTimes := lib.DispatchMethod("axchangetimes", []Value{NewString(testFile), NewInteger(1700000000), NewInteger(1700000001)})
	if changeTimes.Type != VTBool || changeTimes.Num != 1 {
		t.Fatalf("expected axchangetimes true, got %#v", changeTimes)
	}

	changeMode := lib.DispatchMethod("axchangemode", []Value{NewString(testFile), NewString("0644")})
	if changeMode.Type != VTBool {
		t.Fatalf("expected bool result for axchangemode, got %#v", changeMode)
	}

	testLink := filepath.Join(tmpDir, "file.link")
	createLink := lib.DispatchMethod("axcreatelink", []Value{NewString(testFile), NewString(testLink)})
	if createLink.Type != VTBool {
		t.Fatalf("expected bool result for axcreatelink, got %#v", createLink)
	}

	changeOwner := lib.DispatchMethod("axchangeowner", []Value{NewString(testFile), NewInteger(0), NewInteger(0)})
	if changeOwner.Type != VTBool {
		t.Fatalf("expected bool result for axchangeowner, got %#v", changeOwner)
	}
	if runtime.GOOS == "windows" && changeOwner.Num != 0 {
		t.Fatalf("expected axchangeowner false on windows, got %#v", changeOwner)
	}
}

// TestAxRuntimeInfoContainsRequiredSections verifies axruntimeinfo includes config and license text.
func TestAxRuntimeInfoContainsRequiredSections(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	lib := &AxonLibrary{vm: vm}

	got := lib.DispatchMethod("axruntimeinfo", nil)
	if got.Type != VTString {
		t.Fatalf("expected string result, got %#v", got)
	}
	text := got.String()
	required := []string{
		"AXONASP RUNTIME INFORMATION",
		"CONFIGURATION (config/axonasp.toml)",
		"AxonASP Server",
		"Copyright (C) 2026 G3pix Ltda. All rights reserved.",
		"Project URL: https://g3pix.com.br/axonasp",
		"Attribution Notice:",
		"Contribution Policy:",
	}
	for i := 0; i < len(required); i++ {
		if !strings.Contains(text, required[i]) {
			t.Fatalf("missing required runtime info token %q", required[i])
		}
	}
}

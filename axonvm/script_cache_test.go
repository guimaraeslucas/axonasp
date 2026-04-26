/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
	"os"
	"path/filepath"
	"testing"
	"time"

	"github.com/fsnotify/fsnotify"
)

// TestScriptCacheSerializeDeserializeRoundTrip validates custom binary payload roundtrip.
func TestScriptCacheSerializeDeserializeRoundTrip(t *testing.T) {
	payload := cachedProgramBinaryPayload{
		ModTime: 1735689600,
		Program: CachedProgram{
			Bytecode: []byte{1, 2, 3, 4, 5},
			Constants: []Value{
				NewInteger(42),
				NewString("hello"),
				NewUserSub(12, 2, 4, true, 3, []string{"a", "b"}),
				NewEmpty(),
				NewNull(),
				Value{Type: VTObject, Num: 0},
				Value{Type: VTJSFunctionTemplate, Num: 33, Flt: 77, Str: "outer", Names: []string{"p1", "p2"}},
				Value{Type: VTJSUndefined},
			},
			GlobalCount:         7,
			OptionCompare:       1,
			OptionExplicit:      true,
			SourceName:          "C:/www/default.asp",
			IncludeDependencies: []string{"C:/www/includes/header.inc"},
			GlobalZeroArgFuncs:  []string{"dynfunc", "getintranethomepage"},
			GlobalNames:         []string{"Response", "Request", "Widget"},
			DeclaredGlobalNames: []string{"widget"},
			ConstGlobalNames:    []string{"vbcrlf"},
		},
	}

	var buffer bytes.Buffer
	if err := payload.Serialize(&buffer); err != nil {
		t.Fatalf("serialize failed: %v", err)
	}

	decoded := cachedProgramBinaryPayload{}
	if err := decoded.Deserialize(&buffer); err != nil {
		t.Fatalf("deserialize failed: %v", err)
	}

	if decoded.ModTime != payload.ModTime {
		t.Fatalf("modtime mismatch: got %d want %d", decoded.ModTime, payload.ModTime)
	}
	if decoded.Program.GlobalCount != payload.Program.GlobalCount {
		t.Fatalf("global count mismatch: got %d want %d", decoded.Program.GlobalCount, payload.Program.GlobalCount)
	}
	if len(decoded.Program.Bytecode) != len(payload.Program.Bytecode) {
		t.Fatalf("bytecode length mismatch: got %d want %d", len(decoded.Program.Bytecode), len(payload.Program.Bytecode))
	}
	if len(decoded.Program.Constants) != len(payload.Program.Constants) {
		t.Fatalf("constants length mismatch: got %d want %d", len(decoded.Program.Constants), len(payload.Program.Constants))
	}
	if decoded.Program.Constants[2].Type != VTUserSub {
		t.Fatalf("expected VTUserSub in constant slot 2, got %v", decoded.Program.Constants[2].Type)
	}
	if decoded.Program.Constants[2].Num != payload.Program.Constants[2].Num {
		t.Fatalf("usersub entrypoint mismatch: got %d want %d", decoded.Program.Constants[2].Num, payload.Program.Constants[2].Num)
	}
	if len(decoded.Program.Constants[2].Names) != 2 {
		t.Fatalf("usersub local names mismatch")
	}
	if decoded.Program.Constants[5].Type != VTObject || decoded.Program.Constants[5].Num != 0 {
		t.Fatalf("expected VTObject Nothing constant in slot 5, got %#v", decoded.Program.Constants[5])
	}
	if decoded.Program.Constants[6].Type != VTJSFunctionTemplate {
		t.Fatalf("expected VTJSFunctionTemplate in slot 6, got %#v", decoded.Program.Constants[6])
	}
	if decoded.Program.Constants[6].Str != "outer" || decoded.Program.Constants[6].Num != 33 || decoded.Program.Constants[6].Flt != 77 {
		t.Fatalf("unexpected VTJSFunctionTemplate payload in slot 6: %#v", decoded.Program.Constants[6])
	}
	if len(decoded.Program.Constants[6].Names) != 2 || decoded.Program.Constants[6].Names[0] != "p1" || decoded.Program.Constants[6].Names[1] != "p2" {
		t.Fatalf("unexpected VTJSFunctionTemplate names in slot 6: %#v", decoded.Program.Constants[6].Names)
	}
	if decoded.Program.Constants[7].Type != VTJSUndefined {
		t.Fatalf("expected VTJSUndefined in slot 7, got %#v", decoded.Program.Constants[7])
	}
	if len(decoded.Program.IncludeDependencies) != 1 || decoded.Program.IncludeDependencies[0] != payload.Program.IncludeDependencies[0] {
		t.Fatalf("include dependency mismatch: got %#v want %#v", decoded.Program.IncludeDependencies, payload.Program.IncludeDependencies)
	}
	if len(decoded.Program.GlobalZeroArgFuncs) != len(payload.Program.GlobalZeroArgFuncs) {
		t.Fatalf("global zero-arg function count mismatch: got %d want %d", len(decoded.Program.GlobalZeroArgFuncs), len(payload.Program.GlobalZeroArgFuncs))
	}
	for i := range payload.Program.GlobalZeroArgFuncs {
		if decoded.Program.GlobalZeroArgFuncs[i] != payload.Program.GlobalZeroArgFuncs[i] {
			t.Fatalf("global zero-arg function mismatch at %d: got %q want %q", i, decoded.Program.GlobalZeroArgFuncs[i], payload.Program.GlobalZeroArgFuncs[i])
		}
	}
}

// TestScriptCacheDependencyInvalidation verifies include-based dependent script invalidation.
func TestScriptCacheDependencyInvalidation(t *testing.T) {
	cache := NewScriptCache(BytecodeCacheMemoryOnly, t.TempDir(), 8)

	scriptPath := filepath.Join(t.TempDir(), "default.asp")
	includePath := filepath.Join(t.TempDir(), "header.inc")
	otherScriptPath := filepath.Join(t.TempDir(), "other.asp")

	cache.Put(scriptPath, CachedProgram{Bytecode: []byte{1}, GlobalCount: 1}, []string{includePath})
	cache.Put(otherScriptPath, CachedProgram{Bytecode: []byte{2}, GlobalCount: 1}, nil)

	if _, ok := cache.Get(scriptPath); !ok {
		t.Fatalf("expected primary script to be cached")
	}
	if _, ok := cache.Get(otherScriptPath); !ok {
		t.Fatalf("expected secondary script to be cached")
	}

	cache.Invalidate(includePath)

	if _, ok := cache.Get(scriptPath); ok {
		t.Fatalf("expected include-dependent script to be invalidated")
	}
	if _, ok := cache.Get(otherScriptPath); !ok {
		t.Fatalf("expected unrelated script to remain cached")
	}
}

// TestScriptCacheDiskInvalidatesWhenBinaryIsNewer verifies stale disk cache is rejected after a rebuild.
func TestScriptCacheDiskInvalidatesWhenBinaryIsNewer(t *testing.T) {
	cacheDir := t.TempDir()
	cache := NewScriptCache(BytecodeCacheDiskOnly, cacheDir, 8)
	sourcePath := filepath.Join(cacheDir, "default.asp")
	if err := os.WriteFile(sourcePath, []byte("<% Response.Write 1 %>"), 0o644); err != nil {
		t.Fatalf("write source: %v", err)
	}
	sourceInfo, err := os.Stat(sourcePath)
	if err != nil {
		t.Fatalf("stat source: %v", err)
	}
	program := CachedProgram{Bytecode: []byte{1, 2, 3}, GlobalCount: 1, SourceName: sourcePath}
	if err := cache.storeDiskProgram(sourcePath, sourceInfo.ModTime(), program); err != nil {
		t.Fatalf("store disk program: %v", err)
	}
	previousHook := scriptCacheProcessBinaryModUnix
	defer func() { scriptCacheProcessBinaryModUnix = previousHook }()
	scriptCacheProcessBinaryModUnix = func() int64 { return time.Now().Unix() + 3600 }
	if _, found := cache.loadDiskProgram(sourcePath, sourceInfo); found {
		t.Fatalf("expected disk cache miss when running binary is newer than cache")
	}
}

// TestScriptCacheDiskInvalidatesChangedInclude verifies disk cache misses when one include changed after compilation.
func TestScriptCacheDiskInvalidatesChangedInclude(t *testing.T) {
	cacheDir := t.TempDir()
	cache := NewScriptCache(BytecodeCacheDiskOnly, cacheDir, 8)
	sourcePath := filepath.Join(cacheDir, "default.asp")
	includePath := filepath.Join(cacheDir, "header.inc")
	if err := os.WriteFile(sourcePath, []byte("<!--#include file=\"header.inc\"--><% Response.Write 1 %>"), 0o644); err != nil {
		t.Fatalf("write source: %v", err)
	}
	if err := os.WriteFile(includePath, []byte("header"), 0o644); err != nil {
		t.Fatalf("write include: %v", err)
	}
	baseTime := time.Unix(1_735_689_600, 0)
	if err := os.Chtimes(sourcePath, baseTime, baseTime); err != nil {
		t.Fatalf("chtimes source: %v", err)
	}
	if err := os.Chtimes(includePath, baseTime, baseTime); err != nil {
		t.Fatalf("chtimes include: %v", err)
	}
	sourceInfo, err := os.Stat(sourcePath)
	if err != nil {
		t.Fatalf("stat source: %v", err)
	}
	program := CachedProgram{Bytecode: []byte{1, 2, 3}, GlobalCount: 1, SourceName: sourcePath, IncludeDependencies: []string{includePath}}
	if err := cache.storeDiskProgram(sourcePath, sourceInfo.ModTime(), program); err != nil {
		t.Fatalf("store disk program: %v", err)
	}
	previousHook := scriptCacheProcessBinaryModUnix
	defer func() { scriptCacheProcessBinaryModUnix = previousHook }()
	scriptCacheProcessBinaryModUnix = func() int64 { return 0 }
	if _, found := cache.loadDiskProgram(sourcePath, sourceInfo); !found {
		t.Fatalf("expected initial disk cache hit before include changes")
	}
	newIncludeTime := baseTime.Add(2 * time.Hour)
	if err := os.Chtimes(includePath, newIncludeTime, newIncludeTime); err != nil {
		t.Fatalf("chtimes include newer: %v", err)
	}
	if _, found := cache.loadDiskProgram(sourcePath, sourceInfo); found {
		t.Fatalf("expected disk cache miss after include dependency changed")
	}
}

// TestScriptCacheDiskMissesWhenIncludeMetadataMissing verifies stale cache payloads
// without include dependency metadata are not reused for pages with include directives.
func TestScriptCacheDiskMissesWhenIncludeMetadataMissing(t *testing.T) {
	cacheDir := t.TempDir()
	cache := NewScriptCache(BytecodeCacheDiskOnly, cacheDir, 8)
	sourcePath := filepath.Join(cacheDir, "default.asp")
	includePath := filepath.Join(cacheDir, "header.inc")
	if err := os.WriteFile(sourcePath, []byte("<!--#include file=\"header.inc\"--><% Response.Write 1 %>"), 0o644); err != nil {
		t.Fatalf("write source: %v", err)
	}
	if err := os.WriteFile(includePath, []byte("header"), 0o644); err != nil {
		t.Fatalf("write include: %v", err)
	}
	sourceInfo, err := os.Stat(sourcePath)
	if err != nil {
		t.Fatalf("stat source: %v", err)
	}
	program := CachedProgram{Bytecode: []byte{1, 2, 3}, GlobalCount: 1, SourceName: sourcePath}
	if err := cache.storeDiskProgram(sourcePath, sourceInfo.ModTime(), program); err != nil {
		t.Fatalf("store disk program: %v", err)
	}
	previousHook := scriptCacheProcessBinaryModUnix
	defer func() { scriptCacheProcessBinaryModUnix = previousHook }()
	scriptCacheProcessBinaryModUnix = func() int64 { return 0 }
	if _, found := cache.loadDiskProgram(sourcePath, sourceInfo); found {
		t.Fatalf("expected disk cache miss when include metadata is missing")
	}
}

func TestScriptCacheAddWatchRecursiveTrackedDeduplicatesDirectories(t *testing.T) {
	cache := NewScriptCache(BytecodeCacheMemoryOnly, t.TempDir(), 8)
	root := t.TempDir()
	nestedA := filepath.Join(root, "a")
	nestedB := filepath.Join(nestedA, "b")
	if err := os.MkdirAll(nestedB, 0o755); err != nil {
		t.Fatalf("mkdir nested directories: %v", err)
	}

	watcher, err := fsnotify.NewWatcher()
	if err != nil {
		t.Fatalf("create watcher: %v", err)
	}
	defer watcher.Close()

	if err := cache.addWatchRecursiveTracked(watcher, root); err != nil {
		t.Fatalf("add recursive watch first pass: %v", err)
	}
	countFirst := len(cache.watchedDirs)
	if countFirst != 3 {
		t.Fatalf("expected 3 watched directories after first pass, got %d", countFirst)
	}

	if err := cache.addWatchRecursiveTracked(watcher, root); err != nil {
		t.Fatalf("add recursive watch second pass: %v", err)
	}
	countSecond := len(cache.watchedDirs)
	if countSecond != countFirst {
		t.Fatalf("expected deduplicated watch count to remain %d, got %d", countFirst, countSecond)
	}
}

func TestScriptCachePruneStaleWatchesRemovesDeletedDirectories(t *testing.T) {
	cache := NewScriptCache(BytecodeCacheMemoryOnly, t.TempDir(), 8)
	root := t.TempDir()
	nestedA := filepath.Join(root, "nested")
	if err := os.MkdirAll(nestedA, 0o755); err != nil {
		t.Fatalf("mkdir nested directory: %v", err)
	}

	watcher, err := fsnotify.NewWatcher()
	if err != nil {
		t.Fatalf("create watcher: %v", err)
	}
	defer watcher.Close()

	if err := cache.addWatchRecursiveTracked(watcher, root); err != nil {
		t.Fatalf("add recursive watch: %v", err)
	}

	normalizedNestedA, err := cache.normalizeAbsolutePath(nestedA)
	if err != nil {
		t.Fatalf("normalize nested directory: %v", err)
	}
	if _, exists := cache.watchedDirs[normalizedNestedA]; !exists {
		t.Fatalf("expected nested directory to be tracked before deletion")
	}

	if err := os.RemoveAll(nestedA); err != nil {
		t.Fatalf("remove nested directory: %v", err)
	}

	cache.pruneStaleWatches(watcher)

	if _, exists := cache.watchedDirs[normalizedNestedA]; exists {
		t.Fatalf("expected deleted nested directory watch to be pruned")
	}
}

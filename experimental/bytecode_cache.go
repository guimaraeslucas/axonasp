/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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
package experimental

import (
	"crypto/sha256"
	"encoding/gob"
	"encoding/hex"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/vbscript/ast"
)

type BytecodeCacheMode int

const (
	BytecodeCacheMemory BytecodeCacheMode = iota
	BytecodeCacheDisk
)

const bytecodeCacheVersion = "axonasp-vm-bytecode-v1"
const bytecodeCacheMaxMemoryBytes = 4 * 1024 * 1024

var (
	bytecodeCacheOnce   sync.Once
	bytecodeCacheMu     sync.RWMutex
	bytecodeCache       = make(map[string]*Function)
	bytecodeCacheMode   = BytecodeCacheMemory
	bytecodeCacheDir    = filepath.Join("temp", "cache", "bytecode")
	bytecodeCacheDiskMu sync.Mutex
	bytecodeCacheTTLMin = 0
)

type bytecodeCacheRecord struct {
	Version          string
	OptionsSignature string
	ContentHash      string
	ContentSize      int
	CreatedAtUnix    int64
	Function         *Function
}

// ConfigureBytecodeCache sets the cache mode and optional base directory.
// Use mode "memory" or "disk".
func ConfigureBytecodeCache(mode string, webRoot string) {
	mode = strings.ToLower(strings.TrimSpace(mode))
	if mode == "disk" {
		bytecodeCacheMode = BytecodeCacheDisk
	} else {
		bytecodeCacheMode = BytecodeCacheMemory
	}

	if webRoot != "" {
		cleanRoot := filepath.Clean(webRoot)
		baseDir := filepath.Dir(cleanRoot)
		bytecodeCacheDir = filepath.Join(baseDir, "temp", "cache", "bytecode")
	}

	if bytecodeCacheMode == BytecodeCacheDisk {
		_ = os.MkdirAll(bytecodeCacheDir, 0o755)
	}
}

// SetBytecodeCacheTTLMinutes sets the disk cache TTL. Use 0 to keep forever.
func SetBytecodeCacheTTLMinutes(minutes int) {
	if minutes < 0 {
		minutes = 0
	}
	bytecodeCacheTTLMin = minutes
}

// CleanupBytecodeCacheOnShutdown removes disk cache if TTL is enabled.
func CleanupBytecodeCacheOnShutdown() {
	if bytecodeCacheMode != BytecodeCacheDisk || bytecodeCacheTTLMin <= 0 {
		return
	}
	_ = os.RemoveAll(bytecodeCacheDir)
}

// CompileWithCache returns cached bytecode when available or compiles and caches it.
func CompileWithCache(source string, optionsSignature string, compile func() (*Function, error)) (*Function, error) {
	if source == "" {
		return compile()
	}
	bytecodeCacheOnce.Do(registerBytecodeGobTypes)
	key, sig := bytecodeCacheKey(source, optionsSignature)

	if cached := getCachedBytecode(key); cached != nil {
		return cached, nil
	}

	if bytecodeCacheMode == BytecodeCacheDisk {
		if record, err := loadBytecodeFromDisk(key, sig); err == nil && record != nil && record.Function != nil {
			if shouldCacheBytecode(record.ContentSize) {
				setCachedBytecode(key, record.Function)
			}
			return record.Function, nil
		}
	}

	compiled, err := compile()
	if err != nil {
		return nil, err
	}

	if shouldCacheBytecode(len(source)) {
		setCachedBytecode(key, compiled)
		if bytecodeCacheMode == BytecodeCacheDisk {
			_ = saveBytecodeToDisk(key, sig, compiled, len(source))
		}
	}

	return compiled, nil
}

func bytecodeCacheKey(content string, optionsSignature string) (string, string) {
	sig := optionsSignature
	if sig == "" {
		sig = "default"
	}
	hasher := sha256.New()
	_, _ = io.WriteString(hasher, bytecodeCacheVersion)
	_, _ = io.WriteString(hasher, "\n")
	_, _ = io.WriteString(hasher, sig)
	_, _ = io.WriteString(hasher, "\n")
	_, _ = io.WriteString(hasher, content)
	return hex.EncodeToString(hasher.Sum(nil)), sig
}

func getCachedBytecode(key string) *Function {
	bytecodeCacheMu.RLock()
	result := bytecodeCache[key]
	bytecodeCacheMu.RUnlock()
	return result
}

func setCachedBytecode(key string, fn *Function) {
	bytecodeCacheMu.Lock()
	bytecodeCache[key] = fn
	bytecodeCacheMu.Unlock()
}

func loadBytecodeFromDisk(key, sig string) (*bytecodeCacheRecord, error) {
	bytecodeCacheDiskMu.Lock()
	defer bytecodeCacheDiskMu.Unlock()

	path := filepath.Join(bytecodeCacheDir, key+".gob")
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	decoder := gob.NewDecoder(file)
	var record bytecodeCacheRecord
	if err := decoder.Decode(&record); err != nil {
		_ = os.Remove(path)
		return nil, err
	}

	if record.Version != bytecodeCacheVersion || record.OptionsSignature != sig || record.ContentHash != key {
		_ = os.Remove(path)
		return nil, fmt.Errorf("bytecode cache record mismatch")
	}

	if isBytecodeCacheExpired(record.CreatedAtUnix) {
		_ = os.Remove(path)
		return nil, fmt.Errorf("bytecode cache record expired")
	}

	return &record, nil
}

func saveBytecodeToDisk(key, sig string, fn *Function, contentSize int) error {
	bytecodeCacheDiskMu.Lock()
	defer bytecodeCacheDiskMu.Unlock()

	if err := os.MkdirAll(bytecodeCacheDir, 0o755); err != nil {
		return err
	}

	tempFile, err := os.CreateTemp(bytecodeCacheDir, key+"-*.tmp")
	if err != nil {
		return err
	}

	record := bytecodeCacheRecord{
		Version:          bytecodeCacheVersion,
		OptionsSignature: sig,
		ContentHash:      key,
		ContentSize:      contentSize,
		CreatedAtUnix:    time.Now().Unix(),
		Function:         fn,
	}

	encoder := gob.NewEncoder(tempFile)
	if err := encoder.Encode(&record); err != nil {
		_ = tempFile.Close()
		_ = os.Remove(tempFile.Name())
		return err
	}

	if err := tempFile.Close(); err != nil {
		_ = os.Remove(tempFile.Name())
		return err
	}

	finalPath := filepath.Join(bytecodeCacheDir, key+".gob")
	if err := os.Rename(tempFile.Name(), finalPath); err != nil {
		_ = os.Remove(tempFile.Name())
		return err
	}

	return nil
}

func shouldCacheBytecode(contentSize int) bool {
	if contentSize <= 0 {
		return true
	}
	return contentSize <= bytecodeCacheMaxMemoryBytes
}

func isBytecodeCacheExpired(createdAtUnix int64) bool {
	if bytecodeCacheTTLMin <= 0 {
		return false
	}
	if createdAtUnix <= 0 {
		return true
	}
	createdAt := time.Unix(createdAtUnix, 0)
	return time.Since(createdAt) > time.Duration(bytecodeCacheTTLMin)*time.Minute
}

func registerBytecodeGobTypes() {
	gob.Register(VBScriptEmpty{})
	gob.Register(&Bytecode{})
	gob.Register(&Function{})
	gob.Register(&CompiledClass{})
	gob.Register(&BuiltinFunction{})
	gob.Register(&ast.ClassDeclaration{})
}

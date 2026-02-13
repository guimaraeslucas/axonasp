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
package asp

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

type ParseCacheMode int

const (
	ParseCacheMemory ParseCacheMode = iota
	ParseCacheDisk
)

const parseCacheVersion = "asp-ast-cache-v1"
const parseCacheMaxMemoryBytes = 4 * 1024 * 1024

var (
	parseCacheOnce   sync.Once
	parseCacheMu     sync.RWMutex
	parseCache       = make(map[string]*ASPParserResult)
	parseCacheMode   = ParseCacheMemory
	parseCacheDir    = filepath.Join("temp", "cache", "ast")
	parseCacheDiskMu sync.Mutex
	fileCacheMu      sync.RWMutex
	fileCache        = make(map[string]fileCacheEntry)
	parseCacheTTLMin = 0
)

type fileCacheEntry struct {
	contentHash string
	result      *ASPParserResult
}

type parseCacheRecord struct {
	Version          string
	OptionsSignature string
	ContentHash      string
	ContentSize      int
	CreatedAtUnix    int64
	Result           *ASPParserResult
}

// ConfigureParseCache sets the cache mode and optional base directory.
// Use mode "memory" or "disk".
func ConfigureParseCache(mode string, webRoot string) {
	mode = strings.ToLower(strings.TrimSpace(mode))
	if mode == "disk" {
		parseCacheMode = ParseCacheDisk
	} else {
		parseCacheMode = ParseCacheMemory
	}

	parseCacheDir = filepath.Join(resolveExecutableBaseDir(), "temp", "cache", "ast")

	if parseCacheMode == ParseCacheDisk {
		_ = os.MkdirAll(parseCacheDir, 0o755)
	}
}

func resolveExecutableBaseDir() string {
	execPath, err := os.Executable()
	if err == nil {
		execPath, err = filepath.EvalSymlinks(execPath)
		if err == nil {
			return filepath.Dir(execPath)
		}
		return filepath.Dir(execPath)
	}

	wd, err := os.Getwd()
	if err == nil {
		return wd
	}

	return "."
}

// SetParseCacheTTLMinutes sets the disk cache TTL. Use 0 to keep forever.
func SetParseCacheTTLMinutes(minutes int) {
	if minutes < 0 {
		minutes = 0
	}
	parseCacheTTLMin = minutes
}

// CleanupParseCacheOnShutdown removes disk cache if TTL is enabled.
func CleanupParseCacheOnShutdown() {
	if parseCacheMode != ParseCacheDisk || parseCacheTTLMin <= 0 {
		return
	}
	_ = os.RemoveAll(parseCacheDir)
}

// ParseWithCache resolves includes and parses using the shared cache.
func ParseWithCache(content, filePath, rootDir string, options *ASPParsingOptions) (string, *ASPParserResult, error) {
	resolvedContent := content
	if filePath != "" {
		var err error
		resolvedContent, err = ResolveIncludes(content, filePath, rootDir, nil)
		if err != nil {
			return "", nil, err
		}
	}

	if parseCacheMode == ParseCacheMemory {
		if filePath != "" {
			key, _ := cacheKeyForContent(resolvedContent, options)
			if cached := getFileCachedResult(filePath, key); cached != nil {
				return resolvedContent, cached, nil
			}
		}

		result, err := parseResolvedNoCache(resolvedContent, options)
		if err != nil {
			return "", nil, err
		}

		if len(result.Errors) == 0 && filePath != "" && shouldCacheInMemory(len(resolvedContent)) {
			key, _ := cacheKeyForContent(resolvedContent, options)
			setFileCachedResult(filePath, key, result)
		}

		return resolvedContent, result, nil
	}

	result, err := ParseResolvedWithCache(resolvedContent, options)
	if err != nil {
		return "", nil, err
	}

	return resolvedContent, result, nil
}

// ParseResolvedWithCache parses already-resolved content using the shared cache.
func ParseResolvedWithCache(resolvedContent string, options *ASPParsingOptions) (*ASPParserResult, error) {
	parseCacheOnce.Do(registerASTGobTypes)

	if parseCacheMode == ParseCacheMemory {
		return parseResolvedNoCache(resolvedContent, options)
	}

	key, sig := cacheKeyForContent(resolvedContent, options)

	if cached := getCachedResult(key); cached != nil {
		return cached, nil
	}

	if parseCacheMode == ParseCacheDisk {
		if record, err := loadCacheFromDisk(key, sig); err == nil && record != nil && record.Result != nil {
			if shouldCacheInMemory(record.ContentSize) {
				setCachedResult(key, record.Result)
			}
			return record.Result, nil
		}
	}

	parser := NewASPParserWithOptions(resolvedContent, options)
	result, err := parser.Parse()
	if err != nil {
		return nil, err
	}

	if len(result.Errors) == 0 {
		if shouldCacheInMemory(len(resolvedContent)) {
			setCachedResult(key, result)
		}
		if parseCacheMode == ParseCacheDisk {
			_ = saveCacheToDisk(key, sig, result, len(resolvedContent))
		}
	}

	return result, nil
}

func parseResolvedNoCache(resolvedContent string, options *ASPParsingOptions) (*ASPParserResult, error) {
	parser := NewASPParserWithOptions(resolvedContent, options)
	result, err := parser.Parse()
	if err != nil {
		return nil, err
	}
	return result, nil
}

func cacheKeyForContent(content string, options *ASPParsingOptions) (string, string) {
	sig := optionsSignature(options)
	hasher := sha256.New()
	_, _ = io.WriteString(hasher, sig)
	_, _ = io.WriteString(hasher, "\n")
	_, _ = io.WriteString(hasher, content)
	return hex.EncodeToString(hasher.Sum(nil)), sig
}

func optionsSignature(options *ASPParsingOptions) string {
	if options == nil {
		return "default"
	}
	return fmt.Sprintf("save=%t|strict=%t|implicit=%t|debug=%t", options.SaveComments, options.StrictMode, options.AllowImplicitVars, options.DebugMode)
}

func getCachedResult(key string) *ASPParserResult {
	parseCacheMu.RLock()
	result := parseCache[key]
	parseCacheMu.RUnlock()
	return result
}

func setCachedResult(key string, result *ASPParserResult) {
	parseCacheMu.Lock()
	parseCache[key] = result
	parseCacheMu.Unlock()
}

func getFileCachedResult(filePath, contentHash string) *ASPParserResult {
	fileCacheMu.RLock()
	entry, ok := fileCache[filePath]
	fileCacheMu.RUnlock()
	if !ok || entry.contentHash != contentHash {
		return nil
	}
	return entry.result
}

func setFileCachedResult(filePath, contentHash string, result *ASPParserResult) {
	fileCacheMu.Lock()
	fileCache[filePath] = fileCacheEntry{
		contentHash: contentHash,
		result:      result,
	}
	fileCacheMu.Unlock()
}

func loadCacheFromDisk(key, sig string) (*parseCacheRecord, error) {
	parseCacheDiskMu.Lock()
	defer parseCacheDiskMu.Unlock()

	path := filepath.Join(parseCacheDir, key+".gob")
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	decoder := gob.NewDecoder(file)
	var record parseCacheRecord
	if err := decoder.Decode(&record); err != nil {
		_ = os.Remove(path)
		return nil, err
	}

	if record.Version != parseCacheVersion || record.OptionsSignature != sig || record.ContentHash != key {
		_ = os.Remove(path)
		return nil, fmt.Errorf("cache record mismatch")
	}

	if isCacheExpired(record.CreatedAtUnix) {
		_ = os.Remove(path)
		return nil, fmt.Errorf("cache record expired")
	}

	return &record, nil
}

func saveCacheToDisk(key, sig string, result *ASPParserResult, contentSize int) error {
	parseCacheDiskMu.Lock()
	defer parseCacheDiskMu.Unlock()

	if err := os.MkdirAll(parseCacheDir, 0o755); err != nil {
		return err
	}

	tempFile, err := os.CreateTemp(parseCacheDir, key+"-*.tmp")
	if err != nil {
		return err
	}

	record := parseCacheRecord{
		Version:          parseCacheVersion,
		OptionsSignature: sig,
		ContentHash:      key,
		ContentSize:      contentSize,
		CreatedAtUnix:    time.Now().Unix(),
		Result:           result,
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

	finalPath := filepath.Join(parseCacheDir, key+".gob")
	if err := os.Rename(tempFile.Name(), finalPath); err != nil {
		_ = os.Remove(tempFile.Name())
		return err
	}

	return nil
}

func shouldCacheInMemory(contentSize int) bool {
	if contentSize <= 0 {
		return true
	}
	return contentSize <= parseCacheMaxMemoryBytes
}

// ShouldForceFreeMemory returns true when content is large enough to trigger a GC hint.
func ShouldForceFreeMemory(contentSize int) bool {
	if contentSize <= 0 {
		return false
	}
	return contentSize > parseCacheMaxMemoryBytes
}

func isCacheExpired(createdAtUnix int64) bool {
	if parseCacheTTLMin <= 0 {
		return false
	}
	if createdAtUnix <= 0 {
		return true
	}
	createdAt := time.Unix(createdAtUnix, 0)
	return time.Since(createdAt) > time.Duration(parseCacheTTLMin)*time.Minute
}

func registerASTGobTypes() {
	gob.Register(&ast.Program{})
	gob.Register(&ast.Comment{})
	gob.Register(&ast.Parameter{})
	gob.Register(&ast.ConstDeclaration{})
	gob.Register(&ast.ReDimDeclaration{})
	gob.Register(&ast.BaseVariableDeclarationNode{})
	gob.Register(&ast.VariableDeclaration{})
	gob.Register(&ast.FieldDeclaration{})
	gob.Register(&ast.Position{})
	gob.Register(&ast.Range{})
	gob.Register(&ast.Location{})
	gob.Register(&ast.BaseNode{})
	gob.Register(&ast.BaseStatement{})
	gob.Register(&ast.BaseExpression{})
	gob.Register(&ast.BaseExitStatement{})
	gob.Register(&ast.ExitDoStatement{})
	gob.Register(&ast.ExitForStatement{})
	gob.Register(&ast.ExitSubStatement{})
	gob.Register(&ast.ExitFunctionStatement{})
	gob.Register(&ast.ExitPropertyStatement{})
	gob.Register(&ast.AssignmentStatement{})
	gob.Register(&ast.CallStatement{})
	gob.Register(&ast.CallSubStatement{})
	gob.Register(&ast.EraseStatement{})
	gob.Register(&ast.OnErrorResumeNextStatement{})
	gob.Register(&ast.OnErrorGoTo0Statement{})
	gob.Register(&ast.IfStatement{})
	gob.Register(&ast.ElseIfStatement{})
	gob.Register(&ast.ForStatement{})
	gob.Register(&ast.ForEachStatement{})
	gob.Register(&ast.DoStatement{})
	gob.Register(&ast.WhileStatement{})
	gob.Register(&ast.SelectStatement{})
	gob.Register(&ast.CaseStatement{})
	gob.Register(&ast.WithStatement{})
	gob.Register(&ast.VariablesDeclaration{})
	gob.Register(&ast.ConstsDeclaration{})
	gob.Register(&ast.FieldsDeclaration{})
	gob.Register(&ast.ReDimStatement{})
	gob.Register(&ast.StatementList{})
	gob.Register(&ast.BaseProcedureDeclaration{})
	gob.Register(&ast.SubDeclaration{})
	gob.Register(&ast.InitializeSubDeclaration{})
	gob.Register(&ast.TerminateSubDeclaration{})
	gob.Register(&ast.FunctionDeclaration{})
	gob.Register(&ast.BasePropertyDeclaration{})
	gob.Register(&ast.PropertyGetDeclaration{})
	gob.Register(&ast.PropertySetDeclaration{})
	gob.Register(&ast.PropertyLetDeclaration{})
	gob.Register(&ast.ClassDeclaration{})
	gob.Register(&ast.Identifier{})
	gob.Register(&ast.BaseLiteralExpression{})
	gob.Register(&ast.StringLiteral{})
	gob.Register(&ast.IntegerLiteral{})
	gob.Register(&ast.FloatLiteral{})
	gob.Register(&ast.DateLiteral{})
	gob.Register(&ast.BooleanLiteral{})
	gob.Register(&ast.NullLiteral{})
	gob.Register(&ast.EmptyLiteral{})
	gob.Register(&ast.NothingLiteral{})
	gob.Register(&ast.UnaryExpression{})
	gob.Register(&ast.BinaryExpression{})
	gob.Register(&ast.MemberExpression{})
	gob.Register(&ast.IndexOrCallExpression{})
	gob.Register(&ast.NewExpression{})
	gob.Register(&ast.MissingValueExpression{})
	gob.Register(&ast.WithMemberAccessExpression{})
}

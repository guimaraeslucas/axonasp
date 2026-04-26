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
	"bufio"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/cespare/xxhash/v2"
	"github.com/fsnotify/fsnotify"
)

const (
	scriptCacheDependencyMapLimit = 1000
	scriptCacheMagicSize          = 6
	scriptCacheBinaryVersion      = uint16(6)
)

var scriptCacheMagic = [scriptCacheMagicSize]byte{'G', '3', 'A', 'X', 'O', 'N'}

// BytecodeCacheMode defines enabled cache tiers for script compilation.
type BytecodeCacheMode uint8

const (
	BytecodeCacheDisabled BytecodeCacheMode = iota
	BytecodeCacheMemoryOnly
	BytecodeCacheDiskOnly
	BytecodeCacheEnabled
)

// ParseBytecodeCacheMode converts configuration text into a cache mode.
func ParseBytecodeCacheMode(mode string) BytecodeCacheMode {
	switch strings.ToLower(strings.TrimSpace(mode)) {
	case "", "enabled":
		return BytecodeCacheEnabled
	case "memory-only":
		return BytecodeCacheMemoryOnly
	case "disk-only":
		return BytecodeCacheDiskOnly
	case "disabled":
		return BytecodeCacheDisabled
	default:
		return BytecodeCacheEnabled
	}
}

// HasMemoryTier reports whether the mode enables in-memory cache.
func (m BytecodeCacheMode) HasMemoryTier() bool {
	return m == BytecodeCacheEnabled || m == BytecodeCacheMemoryOnly
}

// HasDiskTier reports whether the mode enables disk cache.
func (m BytecodeCacheMode) HasDiskTier() bool {
	return m == BytecodeCacheEnabled || m == BytecodeCacheDiskOnly
}

// CachedProgram stores one compiled program payload for VM initialization.
type CachedProgram struct {
	Bytecode            []byte
	Constants           []Value
	GlobalCount         int
	OptionCompare       int
	OptionExplicit      bool
	SourceName          string
	GlobalPreludeNames  []string
	GlobalPreludeConsts []string
	UserGlobalNames     []string
	UserDeclaredGlobals []string
	UserConstGlobals    []string
	GlobalZeroArgFuncs  []string
	ProgramHash         uint64

	// Legacy fields kept for backward compatibility with v2 cache payloads.
	GlobalNames         []string
	DeclaredGlobalNames []string
	ConstGlobalNames    []string
	IncludeDependencies []string
}

// cachedProgramBinaryPayload stores the serialized disk representation.
type cachedProgramBinaryPayload struct {
	ModTime int64
	Program CachedProgram
}

var scriptCacheProcessBinaryModUnix = currentProcessBinaryModUnix

// Serialize writes one cache payload using a deterministic binary format.
func (p *cachedProgramBinaryPayload) Serialize(writer io.Writer) error {
	if p == nil {
		return errors.New("nil cache payload")
	}
	if writer == nil {
		return errors.New("nil writer")
	}

	buffered := bufio.NewWriterSize(writer, 64*1024)
	if _, err := buffered.Write(scriptCacheMagic[:]); err != nil {
		return err
	}
	if err := binary.Write(buffered, binary.LittleEndian, scriptCacheBinaryVersion); err != nil {
		return err
	}
	if err := binary.Write(buffered, binary.LittleEndian, p.ModTime); err != nil {
		return err
	}
	if p.Program.GlobalCount < 0 {
		return errors.New("invalid global count")
	}
	if err := binary.Write(buffered, binary.LittleEndian, uint32(p.Program.GlobalCount)); err != nil {
		return err
	}

	if len(p.Program.Bytecode) > int(^uint32(0)) {
		return errors.New("bytecode too large")
	}
	if err := binary.Write(buffered, binary.LittleEndian, uint32(len(p.Program.Bytecode))); err != nil {
		return err
	}
	if len(p.Program.Bytecode) > 0 {
		if _, err := buffered.Write(p.Program.Bytecode); err != nil {
			return err
		}
	}

	if len(p.Program.Constants) > int(^uint32(0)) {
		return errors.New("constants too large")
	}
	if err := binary.Write(buffered, binary.LittleEndian, uint32(len(p.Program.Constants))); err != nil {
		return err
	}
	for i := range p.Program.Constants {
		if err := writeSerializedValue(buffered, p.Program.Constants[i]); err != nil {
			return err
		}
	}

	if p.Program.OptionCompare < 0 || p.Program.OptionCompare > 255 {
		return errors.New("invalid option compare")
	}
	if err := binary.Write(buffered, binary.LittleEndian, uint8(p.Program.OptionCompare)); err != nil {
		return err
	}
	optionExplicit := uint8(0)
	if p.Program.OptionExplicit {
		optionExplicit = 1
	}
	if err := binary.Write(buffered, binary.LittleEndian, optionExplicit); err != nil {
		return err
	}
	if err := writeString(buffered, p.Program.SourceName); err != nil {
		return err
	}
	if err := writeStringSlice(buffered, p.Program.GlobalPreludeNames); err != nil {
		return err
	}
	if err := writeStringSlice(buffered, p.Program.GlobalPreludeConsts); err != nil {
		return err
	}
	if err := writeStringSlice(buffered, p.Program.UserGlobalNames); err != nil {
		return err
	}
	if err := writeStringSlice(buffered, p.Program.UserDeclaredGlobals); err != nil {
		return err
	}
	if err := writeStringSlice(buffered, p.Program.UserConstGlobals); err != nil {
		return err
	}
	if err := writeStringSlice(buffered, p.Program.IncludeDependencies); err != nil {
		return err
	}
	if err := binary.Write(buffered, binary.LittleEndian, p.Program.ProgramHash); err != nil {
		return err
	}
	if err := writeStringSlice(buffered, p.Program.GlobalZeroArgFuncs); err != nil {
		return err
	}

	return buffered.Flush()
}

// Deserialize reads one cache payload from binary data.
func (p *cachedProgramBinaryPayload) Deserialize(reader io.Reader) error {
	if p == nil {
		return errors.New("nil cache payload")
	}
	if reader == nil {
		return errors.New("nil reader")
	}

	var magic [scriptCacheMagicSize]byte
	if _, err := io.ReadFull(reader, magic[:]); err != nil {
		return err
	}
	if magic != scriptCacheMagic {
		return NewAxonASPError(ErrInvalidCacheFile, nil, ErrInvalidCacheFile.String(), "", 0)
	}

	var version uint16
	if err := binary.Read(reader, binary.LittleEndian, &version); err != nil {
		return err
	}
	if version != 1 && version != 2 && version != scriptCacheBinaryVersion {
		return NewAxonASPError(ErrInvalidCacheVersion, nil, ErrInvalidCacheVersion.String(), "", 0)
	}

	if err := binary.Read(reader, binary.LittleEndian, &p.ModTime); err != nil {
		return err
	}
	var globalCount uint32
	if err := binary.Read(reader, binary.LittleEndian, &globalCount); err != nil {
		return err
	}
	p.Program.GlobalCount = int(globalCount)

	var bytecodeLength uint32
	if err := binary.Read(reader, binary.LittleEndian, &bytecodeLength); err != nil {
		return err
	}
	if bytecodeLength > 0 {
		p.Program.Bytecode = make([]byte, int(bytecodeLength))
		if _, err := io.ReadFull(reader, p.Program.Bytecode); err != nil {
			return err
		}
	} else {
		p.Program.Bytecode = nil
	}

	var constantsLength uint32
	if err := binary.Read(reader, binary.LittleEndian, &constantsLength); err != nil {
		return err
	}
	if constantsLength > 0 {
		p.Program.Constants = make([]Value, int(constantsLength))
		for i := 0; i < int(constantsLength); i++ {
			value, err := readSerializedValue(reader)
			if err != nil {
				return err
			}
			p.Program.Constants[i] = value
		}
	} else {
		p.Program.Constants = nil
	}

	if version >= 2 {
		var optionCompare uint8
		if err := binary.Read(reader, binary.LittleEndian, &optionCompare); err != nil {
			return err
		}
		p.Program.OptionCompare = int(optionCompare)
		var optionExplicit uint8
		if err := binary.Read(reader, binary.LittleEndian, &optionExplicit); err != nil {
			return err
		}
		p.Program.OptionExplicit = optionExplicit != 0

		sourceName, err := readString(reader)
		if err != nil {
			return err
		}
		p.Program.SourceName = sourceName
		if version == 2 {
			globalNames, err := readStringSlice(reader)
			if err != nil {
				return err
			}
			p.Program.GlobalNames = globalNames
			declaredGlobals, err := readStringSlice(reader)
			if err != nil {
				return err
			}
			p.Program.DeclaredGlobalNames = declaredGlobals
			constGlobals, err := readStringSlice(reader)
			if err != nil {
				return err
			}
			p.Program.ConstGlobalNames = constGlobals
			migrateLegacyCachedProgramSymbols(&p.Program)
		}
		if version >= 3 {
			prelude, err := readStringSlice(reader)
			if err != nil {
				return err
			}
			p.Program.GlobalPreludeNames = prelude
			preludeConsts, err := readStringSlice(reader)
			if err != nil {
				return err
			}
			p.Program.GlobalPreludeConsts = preludeConsts
			userGlobals, err := readStringSlice(reader)
			if err != nil {
				return err
			}
			p.Program.UserGlobalNames = userGlobals
			userDeclared, err := readStringSlice(reader)
			if err != nil {
				return err
			}
			p.Program.UserDeclaredGlobals = userDeclared
			userConsts, err := readStringSlice(reader)
			if err != nil {
				return err
			}
			p.Program.UserConstGlobals = userConsts
			if version >= 4 {
				includeDependencies, err := readStringSlice(reader)
				if err != nil {
					return err
				}
				p.Program.IncludeDependencies = includeDependencies
			}
			if err := binary.Read(reader, binary.LittleEndian, &p.Program.ProgramHash); err != nil {
				return err
			}
			if version >= 5 {
				zeroArgFuncs, err := readStringSlice(reader)
				if err != nil {
					return err
				}
				p.Program.GlobalZeroArgFuncs = zeroArgFuncs
			} else {
				p.Program.GlobalZeroArgFuncs = inferCachedProgramZeroArgFuncs(&p.Program)
			}
		}
	}

	if p.Program.ProgramHash == 0 {
		p.Program.ProgramHash = computeProgramHash(
			p.Program.Bytecode,
			p.Program.GlobalCount,
			p.Program.OptionCompare,
			p.Program.OptionExplicit,
			p.Program.SourceName,
		)
	}

	return nil
}

// ScriptCache implements two-tier ASP bytecode caching and dependency invalidation.
type ScriptCache struct {
	mu                 sync.RWMutex
	mode               BytecodeCacheMode
	cacheDir           string
	programs           map[string]CachedProgram
	programSizes       map[string]int64
	programOrder       []string
	totalBytes         int64
	maxBytes           int64
	dependencyMap      map[string][]string
	scriptDependencies map[string][]string
	dependencyOrder    []string
	watcher            *fsnotify.Watcher
	watchRoots         []string
	watchedDirs        map[string]struct{}
	watchStop          chan struct{}
	watcherActive      bool
	watcherErrorCount  uint32
}

// NewScriptCache builds one cache instance with mode and memory limit settings.
func NewScriptCache(mode BytecodeCacheMode, cacheDir string, maxSizeMB int) *ScriptCache {
	if maxSizeMB <= 0 {
		maxSizeMB = 1
	}
	if strings.TrimSpace(cacheDir) == "" {
		cacheDir = filepath.Join("temp", "cache")
	}
	return &ScriptCache{
		mode:               mode,
		cacheDir:           cacheDir,
		programs:           make(map[string]CachedProgram),
		programSizes:       make(map[string]int64),
		programOrder:       make([]string, 0, 64),
		dependencyMap:      make(map[string][]string),
		scriptDependencies: make(map[string][]string),
		dependencyOrder:    make([]string, 0, 256),
		watchedDirs:        make(map[string]struct{}),
		maxBytes:           int64(maxSizeMB) * 1024 * 1024,
	}
}

// Mode returns the current cache tier mode.
func (c *ScriptCache) Mode() BytecodeCacheMode {
	if c == nil {
		return BytecodeCacheDisabled
	}
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.mode
}

// SetMode updates cache mode and clears state when cache is disabled.
func (c *ScriptCache) SetMode(mode BytecodeCacheMode) {
	if c == nil {
		return
	}
	c.mu.Lock()
	defer c.mu.Unlock()
	c.mode = mode
	if !mode.HasMemoryTier() {
		c.programs = make(map[string]CachedProgram)
		c.programSizes = make(map[string]int64)
		c.programOrder = c.programOrder[:0]
		c.totalBytes = 0
		c.dependencyMap = make(map[string][]string)
		c.scriptDependencies = make(map[string][]string)
		c.dependencyOrder = c.dependencyOrder[:0]
	}
}

// SetMaxSizeMB updates the in-memory cache size cap in megabytes.
func (c *ScriptCache) SetMaxSizeMB(maxSizeMB int) {
	if c == nil {
		return
	}
	if maxSizeMB <= 0 {
		maxSizeMB = 1
	}
	c.mu.Lock()
	defer c.mu.Unlock()
	c.maxBytes = int64(maxSizeMB) * 1024 * 1024
	if c.totalBytes <= c.maxBytes {
		return
	}
	for len(c.programOrder) > 0 && c.totalBytes > c.maxBytes {
		victim := c.programOrder[0]
		c.removeProgramNoLock(victim)
	}
}

// StartInvalidator runs the background file change invalidator for configured roots.
func (c *ScriptCache) StartInvalidator(rootDirs []string) error {
	if c == nil {
		return nil
	}
	if !c.mode.HasMemoryTier() {
		return nil
	}

	watcher, err := fsnotify.NewWatcher()
	if err != nil {
		return err
	}

	resolvedRoots := make([]string, 0, len(rootDirs))
	for _, rootDir := range rootDirs {
		resolved, resolveErr := c.normalizeAbsolutePath(rootDir)
		if resolveErr != nil {
			_ = watcher.Close()
			return fmt.Errorf("failed to normalize root directory %q: %w", rootDir, resolveErr)
		}
		if !containsFold(resolvedRoots, resolved) {
			resolvedRoots = append(resolvedRoots, resolved)
		}
	}

	if len(resolvedRoots) == 0 {
		_ = watcher.Close()
		return errors.New("no valid root directories to watch")
	}

	c.mu.Lock()
	if c.watchedDirs == nil {
		c.watchedDirs = make(map[string]struct{}, 256)
	} else {
		clear(c.watchedDirs)
	}
	c.mu.Unlock()

	for _, rootDir := range resolvedRoots {
		if err := c.addWatchRecursiveTracked(watcher, rootDir); err != nil {
			_ = watcher.Close()
			return fmt.Errorf("failed to add watch on %q: %w", rootDir, err)
		}
	}

	c.mu.Lock()
	if c.watcher != nil {
		_ = watcher.Close()
		c.mu.Unlock()
		return nil
	}
	c.watcher = watcher
	c.watchRoots = resolvedRoots
	c.watchStop = make(chan struct{})
	c.watcherActive = true
	c.watcherErrorCount = 0
	stop := c.watchStop
	c.mu.Unlock()

	go func() {
		pruneTicker := time.NewTicker(45 * time.Second)
		defer pruneTicker.Stop()
		defer func() {
			c.mu.Lock()
			c.watcherActive = false
			c.mu.Unlock()
		}()

		for {
			select {
			case <-stop:
				return
			case event, ok := <-watcher.Events:
				if !ok {
					return
				}
				if event.Op&(fsnotify.Write|fsnotify.Create|fsnotify.Remove|fsnotify.Rename) == 0 {
					continue
				}
				if event.Op&fsnotify.Create != 0 {
					info, statErr := os.Stat(event.Name)
					if statErr == nil && info.IsDir() {
						_ = c.addWatchRecursiveTracked(watcher, event.Name)
					}
				}
				if event.Op&(fsnotify.Remove|fsnotify.Rename) != 0 {
					c.removeWatchPathTracked(watcher, event.Name, true)
				}
				if !shouldInvalidatePath(event.Name) {
					continue
				}
				c.Invalidate(event.Name)
			case <-pruneTicker.C:
				c.pruneStaleWatches(watcher)
			case err, ok := <-watcher.Errors:
				if !ok {
					return
				}
				if err != nil {
					c.mu.Lock()
					c.watcherErrorCount++
					c.mu.Unlock()
				}
			}
		}
	}()

	return nil
}

// StopInvalidator stops the active file watcher and goroutine.
func (c *ScriptCache) StopInvalidator() {
	if c == nil {
		return
	}
	c.mu.Lock()
	if c.watchStop != nil {
		close(c.watchStop)
		c.watchStop = nil
	}
	watcher := c.watcher
	c.watcher = nil
	c.watchRoots = nil
	if c.watchedDirs != nil {
		clear(c.watchedDirs)
	}
	c.mu.Unlock()
	if watcher != nil {
		_ = watcher.Close()
	}
	// Wait briefly for goroutine to exit
	for i := 0; i < 50; i++ {
		c.mu.RLock()
		active := c.watcherActive
		c.mu.RUnlock()
		if !active {
			break
		}
		time.Sleep(10 * time.Millisecond)
	}
}

// GetWatcherStatus returns the current watcher state (active, error count, roots).
func (c *ScriptCache) GetWatcherStatus() (active bool, errorCount uint32, roots int) {
	if c == nil {
		return false, 0, 0
	}
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.watcherActive, c.watcherErrorCount, len(c.watchRoots)
}

// Get returns a memory-cached program for one absolute source path.
func (c *ScriptCache) Get(filePath string) (CachedProgram, bool) {
	if c == nil || !c.mode.HasMemoryTier() {
		return CachedProgram{}, false
	}
	normalized, err := c.normalizeAbsolutePath(filePath)
	if err != nil {
		return CachedProgram{}, false
	}

	c.mu.RLock()
	program, ok := c.programs[normalized]
	c.mu.RUnlock()
	if !ok {
		return CachedProgram{}, false
	}
	return immutableCachedProgramView(program), true
}

// Put stores one compiled program in memory cache and dependency map.
func (c *ScriptCache) Put(filePath string, program CachedProgram, dependencies []string) {
	if c == nil || !c.mode.HasMemoryTier() {
		return
	}
	normalized, err := c.normalizeAbsolutePath(filePath)
	if err != nil {
		return
	}

	program = immutableCachedProgramView(program)
	sizeBytes := estimateProgramSizeBytes(program)
	if sizeBytes <= 0 {
		return
	}

	c.mu.Lock()
	defer c.mu.Unlock()
	if sizeBytes > c.maxBytes {
		c.removeProgramNoLock(normalized)
		c.removeScriptDependenciesNoLock(normalized)
		return
	}

	c.removeProgramNoLock(normalized)
	for c.totalBytes+sizeBytes > c.maxBytes && len(c.programOrder) > 0 {
		victim := c.programOrder[0]
		c.removeProgramNoLock(victim)
	}

	c.programs[normalized] = program
	c.programSizes[normalized] = sizeBytes
	c.programOrder = append(c.programOrder, normalized)
	c.totalBytes += sizeBytes
	c.registerDependenciesNoLock(normalized, dependencies)
}

// Invalidate removes one script and all dependent scripts from memory cache.
func (c *ScriptCache) Invalidate(filePath string) {
	if c == nil || !c.mode.HasMemoryTier() {
		return
	}
	normalized, err := c.normalizeAbsolutePath(filePath)
	if err != nil {
		return
	}

	c.mu.Lock()
	defer c.mu.Unlock()

	invalidateList := make([]string, 0, 16)
	invalidateSet := make(map[string]struct{}, 16)
	invalidateSet[normalized] = struct{}{}
	invalidateList = append(invalidateList, normalized)
	if dependents, exists := c.dependencyMap[normalized]; exists {
		for _, dependent := range dependents {
			if _, seen := invalidateSet[dependent]; seen {
				continue
			}
			invalidateSet[dependent] = struct{}{}
			invalidateList = append(invalidateList, dependent)
		}
	}

	for _, scriptPath := range invalidateList {
		if _, exists := c.programs[scriptPath]; exists {
			c.removeProgramNoLock(scriptPath)
		}
		c.removeScriptDependenciesNoLock(scriptPath)
	}
}

// LoadOrCompile applies memory, disk, and compiler fallback flow for one ASP file.
func (c *ScriptCache) LoadOrCompile(filePath string) (CachedProgram, error) {
	normalized, err := c.normalizeAbsolutePath(filePath)
	if err != nil {
		return CachedProgram{}, err
	}

	if c != nil && c.mode.HasMemoryTier() {
		if program, found := c.Get(normalized); found {
			return program, nil
		}
	}

	sourceInfo, err := os.Stat(normalized)
	if err != nil {
		return CachedProgram{}, err
	}

	if c != nil && c.mode.HasDiskTier() {
		if program, found := c.loadDiskProgram(normalized, sourceInfo); found {
			if c.mode.HasMemoryTier() {
				c.Put(normalized, program, program.IncludeDependencies)
			}
			return program, nil
		}
	}

	content, err := os.ReadFile(normalized)
	if err != nil {
		return CachedProgram{}, err
	}

	// Strip UTF-8 BOM if present to prevent parsing errors
	content = stripUTF8BOM(content)

	compiler := NewASPCompiler(string(content))
	compiler.SetSourceName(normalized)
	if err := compiler.Compile(); err != nil {
		return CachedProgram{}, err
	}

	program := buildCachedProgramFromCompiler(compiler)

	if c != nil && c.mode.HasDiskTier() {
		if storeErr := c.storeDiskProgram(normalized, sourceInfo.ModTime(), program); storeErr != nil {
			return CachedProgram{}, fmt.Errorf("failed to persist bytecode cache to disk for %s: %w", normalized, storeErr)
		}
	}
	if c != nil && c.mode.HasMemoryTier() {
		c.Put(normalized, program, program.IncludeDependencies)
	}
	return program, nil
}

// loadDiskProgram reads one cached payload from .aspb when source and cache are fresh.
func (c *ScriptCache) loadDiskProgram(filePath string, sourceInfo os.FileInfo) (CachedProgram, bool) {
	if c == nil || !c.mode.HasDiskTier() {
		return CachedProgram{}, false
	}
	cacheFile := c.cacheFilePath(filePath)
	cacheInfo, err := os.Stat(cacheFile)
	if err != nil {
		return CachedProgram{}, false
	}
	if processBuildUnix := scriptCacheProcessBinaryModUnix(); processBuildUnix > 0 && cacheInfo.ModTime().Unix() < processBuildUnix {
		return CachedProgram{}, false
	}
	if cacheInfo.ModTime().Before(sourceInfo.ModTime()) {
		return CachedProgram{}, false
	}

	file, err := os.Open(cacheFile)
	if err != nil {
		return CachedProgram{}, false
	}
	defer file.Close()

	payload := cachedProgramBinaryPayload{}
	if err := payload.Deserialize(bufio.NewReaderSize(file, 64*1024)); err != nil {
		_ = os.Remove(cacheFile)
		return CachedProgram{}, false
	}
	if payload.ModTime < sourceInfo.ModTime().Unix() {
		return CachedProgram{}, false
	}
	if len(payload.Program.IncludeDependencies) == 0 && sourceHasIncludeDirective(filePath) {
		return CachedProgram{}, false
	}
	if !c.dependenciesFresh(payload.Program.IncludeDependencies, payload.ModTime) {
		return CachedProgram{}, false
	}
	return immutableCachedProgramView(payload.Program), true
}

// storeDiskProgram persists one compiled program into its .aspb cache file.
func (c *ScriptCache) storeDiskProgram(filePath string, sourceModTime time.Time, program CachedProgram) error {
	if c == nil || !c.mode.HasDiskTier() {
		return nil
	}
	if err := os.MkdirAll(c.cacheDir, 0o755); err != nil {
		return err
	}

	cacheFile := c.cacheFilePath(filePath)
	tempFile := cacheFile + ".tmp"
	file, err := os.Create(tempFile)
	if err != nil {
		return err
	}

	payload := cachedProgramBinaryPayload{
		ModTime: sourceModTime.Unix(),
		Program: cloneCachedProgram(program),
	}
	serializeErr := payload.Serialize(file)
	closeErr := file.Close()
	if serializeErr != nil {
		_ = os.Remove(tempFile)
		return serializeErr
	}
	if closeErr != nil {
		_ = os.Remove(tempFile)
		return closeErr
	}

	if err := os.Rename(tempFile, cacheFile); err != nil {
		if removeErr := os.Remove(cacheFile); removeErr != nil && !os.IsNotExist(removeErr) {
			_ = os.Remove(tempFile)
			return err
		}
		if retryErr := os.Rename(tempFile, cacheFile); retryErr != nil {
			_ = os.Remove(tempFile)
			return retryErr
		}
	}
	return nil
}

// NewVMFromCachedProgram creates a VM instance from cached compilation output.
func NewVMFromCachedProgram(program CachedProgram) *VM {
	program = immutableCachedProgramView(program)
	vm := NewVM(program.Bytecode, program.Constants, program.GlobalCount)
	vm.optionCompare = program.OptionCompare
	vm.optionExplicit = program.OptionExplicit
	vm.sourceName = program.SourceName
	applyProgramGlobalMetadata(vm, program)
	vm.captureBaseProgramState()
	return vm
}

func (c *ScriptCache) normalizeAbsolutePath(filePath string) (string, error) {
	trimmed := strings.TrimSpace(filePath)
	if trimmed == "" {
		return "", errors.New("empty file path")
	}
	absPath, err := filepath.Abs(trimmed)
	if err != nil {
		return "", err
	}
	return filepath.Clean(absPath), nil
}

func (c *ScriptCache) cacheFilePath(filePath string) string {
	hash := xxhash.Sum64String(filePath)
	return filepath.Join(c.cacheDir, fmt.Sprintf("%016x.aspb", hash))
}

func currentProcessBinaryModUnix() int64 {
	executablePath, err := os.Executable()
	if err != nil || strings.TrimSpace(executablePath) == "" {
		return 0
	}
	info, err := os.Stat(executablePath)
	if err != nil {
		return 0
	}
	return info.ModTime().Unix()
}

func (c *ScriptCache) dependenciesFresh(dependencies []string, compiledUnix int64) bool {
	if len(dependencies) == 0 {
		return true
	}
	for _, depPath := range dependencies {
		normalized, err := c.normalizeAbsolutePath(depPath)
		if err != nil || normalized == "" {
			return false
		}
		info, err := os.Stat(normalized)
		if err != nil {
			return false
		}
		if info.ModTime().Unix() > compiledUnix {
			return false
		}
	}
	return true
}

func sourceHasIncludeDirective(filePath string) bool {
	content, err := os.ReadFile(filePath)
	if err != nil || len(content) == 0 {
		return false
	}
	// Strip UTF-8 BOM if present
	content = stripUTF8BOM(content)
	lower := strings.ToLower(string(content))
	return strings.Contains(lower, "<!--#include")
}

func (c *ScriptCache) registerDependenciesNoLock(scriptPath string, dependencies []string) {
	c.removeScriptDependenciesNoLock(scriptPath)
	if len(dependencies) == 0 {
		return
	}

	normalizedDeps := make([]string, 0, len(dependencies))
	seen := make(map[string]struct{}, len(dependencies))
	for _, depPath := range dependencies {
		normalized, err := c.normalizeAbsolutePath(depPath)
		if err != nil || normalized == "" {
			continue
		}
		if _, exists := seen[normalized]; exists {
			continue
		}
		seen[normalized] = struct{}{}
		normalizedDeps = append(normalizedDeps, normalized)
	}
	if len(normalizedDeps) == 0 {
		return
	}

	c.scriptDependencies[scriptPath] = normalizedDeps
	for _, depPath := range normalizedDeps {
		if _, exists := c.dependencyMap[depPath]; !exists {
			if len(c.dependencyMap) >= scriptCacheDependencyMapLimit {
				victim := c.dependencyOrder[0]
				c.dependencyOrder = c.dependencyOrder[1:]
				delete(c.dependencyMap, victim)
				c.removeDependencyFromScriptsNoLock(victim)
			}
			c.dependencyMap[depPath] = make([]string, 0, 2)
			c.dependencyOrder = append(c.dependencyOrder, depPath)
		}
		if !containsFold(c.dependencyMap[depPath], scriptPath) {
			c.dependencyMap[depPath] = append(c.dependencyMap[depPath], scriptPath)
		}
	}
}

func (c *ScriptCache) removeDependencyFromScriptsNoLock(dependencyPath string) {
	for scriptPath, deps := range c.scriptDependencies {
		filtered := filterOutFold(deps, dependencyPath)
		if len(filtered) == 0 {
			delete(c.scriptDependencies, scriptPath)
			continue
		}
		c.scriptDependencies[scriptPath] = filtered
	}
}

func (c *ScriptCache) removeScriptDependenciesNoLock(scriptPath string) {
	deps, exists := c.scriptDependencies[scriptPath]
	if !exists {
		return
	}
	delete(c.scriptDependencies, scriptPath)

	for _, depPath := range deps {
		scripts, ok := c.dependencyMap[depPath]
		if !ok {
			continue
		}
		filtered := filterOutFold(scripts, scriptPath)
		if len(filtered) == 0 {
			delete(c.dependencyMap, depPath)
			c.dependencyOrder = filterOutFold(c.dependencyOrder, depPath)
			continue
		}
		c.dependencyMap[depPath] = filtered
	}
}

func (c *ScriptCache) removeProgramNoLock(scriptPath string) {
	if size, exists := c.programSizes[scriptPath]; exists {
		c.totalBytes -= size
		delete(c.programSizes, scriptPath)
	}
	delete(c.programs, scriptPath)
	c.programOrder = filterOutFold(c.programOrder, scriptPath)
	if c.totalBytes < 0 {
		c.totalBytes = 0
	}
}

func (c *ScriptCache) addWatchRecursiveTracked(watcher *fsnotify.Watcher, rootDir string) error {
	if c == nil || watcher == nil {
		return nil
	}
	info, err := os.Stat(rootDir)
	if err != nil {
		return err
	}
	if !info.IsDir() {
		return nil
	}
	return filepath.WalkDir(rootDir, func(path string, dirEntry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return nil
		}
		if !dirEntry.IsDir() {
			return nil
		}
		if strings.EqualFold(dirEntry.Name(), "node_modules") || strings.EqualFold(dirEntry.Name(), ".git") {
			return filepath.SkipDir
		}
		return c.addWatchPathTracked(watcher, path)
	})
}

func (c *ScriptCache) addWatchPathTracked(watcher *fsnotify.Watcher, path string) error {
	if c == nil || watcher == nil {
		return nil
	}
	normalized, err := c.normalizeAbsolutePath(path)
	if err != nil {
		return err
	}

	c.mu.RLock()
	_, exists := c.watchedDirs[normalized]
	c.mu.RUnlock()
	if exists {
		return nil
	}

	if err := watcher.Add(normalized); err != nil {
		if !isAlreadyWatchedError(err) {
			return err
		}
	}

	c.mu.Lock()
	if c.watchedDirs == nil {
		c.watchedDirs = make(map[string]struct{}, 256)
	}
	c.watchedDirs[normalized] = struct{}{}
	c.mu.Unlock()
	return nil
}

func (c *ScriptCache) removeWatchPathTracked(watcher *fsnotify.Watcher, path string, includeChildren bool) {
	if c == nil || watcher == nil {
		return
	}
	normalized, err := c.normalizeAbsolutePath(path)
	if err != nil {
		return
	}

	removals := make([]string, 0, 8)
	c.mu.RLock()
	for watchedPath := range c.watchedDirs {
		if strings.EqualFold(watchedPath, normalized) || (includeChildren && pathHasPrefixFold(watchedPath, normalized)) {
			removals = append(removals, watchedPath)
		}
	}
	c.mu.RUnlock()

	for _, watchedPath := range removals {
		_ = watcher.Remove(watchedPath)
		c.mu.Lock()
		delete(c.watchedDirs, watchedPath)
		c.mu.Unlock()
	}
}

func (c *ScriptCache) pruneStaleWatches(watcher *fsnotify.Watcher) {
	if c == nil || watcher == nil {
		return
	}
	stale := make([]string, 0, 16)
	c.mu.RLock()
	for watchedPath := range c.watchedDirs {
		info, err := os.Stat(watchedPath)
		if err != nil || !info.IsDir() {
			stale = append(stale, watchedPath)
		}
	}
	c.mu.RUnlock()

	for _, watchedPath := range stale {
		_ = watcher.Remove(watchedPath)
		c.mu.Lock()
		delete(c.watchedDirs, watchedPath)
		c.mu.Unlock()
	}
}

func isAlreadyWatchedError(err error) bool {
	if err == nil {
		return false
	}
	message := strings.ToLower(strings.TrimSpace(err.Error()))
	if message == "" {
		return false
	}
	return strings.Contains(message, "already exists") || strings.Contains(message, "already being watched")
}

func pathHasPrefixFold(path string, prefix string) bool {
	if strings.EqualFold(path, prefix) {
		return false
	}
	pathLower := strings.ToLower(filepath.Clean(path))
	prefixLower := strings.ToLower(filepath.Clean(prefix))
	if prefixLower == "" || pathLower == "" {
		return false
	}
	if !strings.HasSuffix(prefixLower, string(filepath.Separator)) {
		prefixLower += string(filepath.Separator)
	}
	return strings.HasPrefix(pathLower, prefixLower)
}

func shouldInvalidatePath(path string) bool {
	ext := strings.ToLower(filepath.Ext(path))
	return ext == ".asp" || ext == ".inc"
}

func containsFold(values []string, target string) bool {
	for _, value := range values {
		if strings.EqualFold(value, target) {
			return true
		}
	}
	return false
}

func filterOutFold(values []string, target string) []string {
	if len(values) == 0 {
		return values
	}
	filtered := make([]string, 0, len(values))
	for _, value := range values {
		if strings.EqualFold(value, target) {
			continue
		}
		filtered = append(filtered, value)
	}
	return filtered
}

func writeString(writer io.Writer, value string) error {
	if len(value) > int(^uint32(0)) {
		return errors.New("string too large")
	}
	if err := binary.Write(writer, binary.LittleEndian, uint32(len(value))); err != nil {
		return err
	}
	if len(value) == 0 {
		return nil
	}
	_, err := io.WriteString(writer, value)
	return err
}

func readString(reader io.Reader) (string, error) {
	var length uint32
	if err := binary.Read(reader, binary.LittleEndian, &length); err != nil {
		return "", err
	}
	if length == 0 {
		return "", nil
	}
	buf := make([]byte, int(length))
	if _, err := io.ReadFull(reader, buf); err != nil {
		return "", err
	}
	return string(buf), nil
}

func writeStringSlice(writer io.Writer, values []string) error {
	if len(values) > int(^uint32(0)) {
		return errors.New("slice too large")
	}
	if err := binary.Write(writer, binary.LittleEndian, uint32(len(values))); err != nil {
		return err
	}
	for i := range values {
		if err := writeString(writer, values[i]); err != nil {
			return err
		}
	}
	return nil
}

func readStringSlice(reader io.Reader) ([]string, error) {
	var length uint32
	if err := binary.Read(reader, binary.LittleEndian, &length); err != nil {
		return nil, err
	}
	if length == 0 {
		return nil, nil
	}
	values := make([]string, int(length))
	for i := 0; i < int(length); i++ {
		value, err := readString(reader)
		if err != nil {
			return nil, err
		}
		values[i] = value
	}
	return values, nil
}

func writeSerializedValue(writer io.Writer, value Value) error {
	if err := binary.Write(writer, binary.LittleEndian, uint8(value.Type)); err != nil {
		return err
	}
	if err := binary.Write(writer, binary.LittleEndian, value.Num); err != nil {
		return err
	}
	if err := binary.Write(writer, binary.LittleEndian, value.Flt); err != nil {
		return err
	}

	switch value.Type {
	case VTString, VTUserSub, VTJSFunctionTemplate:
		if err := writeString(writer, value.Str); err != nil {
			return err
		}
	case VTEmpty, VTNull, VTBool, VTInteger, VTDouble, VTDate, VTObject, VTNativeObject, VTBuiltin,
		VTJSUndefined, VTJSObject, VTJSFunction:
		// Type, Num and Flt are sufficient for deterministic reconstruction.
	default:
		return NewAxonASPError(ErrInvalidCacheFile, nil, ErrInvalidCacheFile.String(), "", 0)
	}

	if value.Type == VTUserSub || value.Type == VTJSFunctionTemplate {
		if err := writeStringSlice(writer, value.Names); err != nil {
			return err
		}
	}

	return nil
}

func readSerializedValue(reader io.Reader) (Value, error) {
	var rawType uint8
	if err := binary.Read(reader, binary.LittleEndian, &rawType); err != nil {
		return Value{}, err
	}
	valueType := ValueType(rawType)
	value := Value{Type: valueType}
	if err := binary.Read(reader, binary.LittleEndian, &value.Num); err != nil {
		return Value{}, err
	}
	if err := binary.Read(reader, binary.LittleEndian, &value.Flt); err != nil {
		return Value{}, err
	}

	switch valueType {
	case VTString, VTUserSub, VTJSFunctionTemplate:
		stringValue, err := readString(reader)
		if err != nil {
			return Value{}, err
		}
		value.Str = stringValue
	case VTEmpty, VTNull, VTBool, VTInteger, VTDouble, VTDate, VTObject, VTNativeObject, VTBuiltin,
		VTJSUndefined, VTJSObject, VTJSFunction:
		// Scalar content already read.
	default:
		return Value{}, NewAxonASPError(ErrInvalidCacheFile, nil, ErrInvalidCacheFile.String(), "", 0)
	}

	if valueType == VTUserSub || valueType == VTJSFunctionTemplate {
		names, err := readStringSlice(reader)
		if err != nil {
			return Value{}, err
		}
		value.Names = names
	}

	return value, nil
}

func cloneCachedProgram(program CachedProgram) CachedProgram {
	cloned := CachedProgram{
		Bytecode:            cloneBytecode(program.Bytecode),
		Constants:           cloneValueSlice(program.Constants),
		GlobalCount:         program.GlobalCount,
		OptionCompare:       program.OptionCompare,
		OptionExplicit:      program.OptionExplicit,
		SourceName:          program.SourceName,
		GlobalPreludeNames:  cloneStringSlice(program.GlobalPreludeNames),
		GlobalPreludeConsts: cloneStringSlice(program.GlobalPreludeConsts),
		UserGlobalNames:     cloneStringSlice(program.UserGlobalNames),
		UserDeclaredGlobals: cloneStringSlice(program.UserDeclaredGlobals),
		UserConstGlobals:    cloneStringSlice(program.UserConstGlobals),
		GlobalZeroArgFuncs:  cloneStringSlice(program.GlobalZeroArgFuncs),
		ProgramHash:         program.ProgramHash,
		GlobalNames:         cloneStringSlice(program.GlobalNames),
		DeclaredGlobalNames: cloneStringSlice(program.DeclaredGlobalNames),
		ConstGlobalNames:    cloneStringSlice(program.ConstGlobalNames),
		IncludeDependencies: cloneStringSlice(program.IncludeDependencies),
	}
	return cloned
}

func cloneBytecode(bytecode []byte) []byte {
	if len(bytecode) == 0 {
		return nil
	}
	cloned := make([]byte, len(bytecode))
	copy(cloned, bytecode)
	return cloned
}

func cloneStringSlice(values []string) []string {
	if len(values) == 0 {
		return nil
	}
	cloned := make([]string, len(values))
	copy(cloned, values)
	return cloned
}

func cloneValueSlice(values []Value) []Value {
	if len(values) == 0 {
		return nil
	}
	cloned := make([]Value, len(values))
	copy(cloned, values)
	for i := range cloned {
		if len(cloned[i].Names) > 0 {
			names := make([]string, len(cloned[i].Names))
			copy(names, cloned[i].Names)
			cloned[i].Names = names
		}
	}
	return cloned
}

// immutableCachedProgramView returns one non-allocating immutable view of cache payload slices.
// The returned struct aliases backing arrays and must be treated as read-only.
func immutableCachedProgramView(program CachedProgram) CachedProgram {
	program.Bytecode = immutableBytecodeView(program.Bytecode)
	program.Constants = immutableValueView(program.Constants)
	program.GlobalPreludeNames = immutableStringView(program.GlobalPreludeNames)
	program.GlobalPreludeConsts = immutableStringView(program.GlobalPreludeConsts)
	program.UserGlobalNames = immutableStringView(program.UserGlobalNames)
	program.UserDeclaredGlobals = immutableStringView(program.UserDeclaredGlobals)
	program.UserConstGlobals = immutableStringView(program.UserConstGlobals)
	program.GlobalZeroArgFuncs = immutableStringView(program.GlobalZeroArgFuncs)
	program.GlobalNames = immutableStringView(program.GlobalNames)
	program.DeclaredGlobalNames = immutableStringView(program.DeclaredGlobalNames)
	program.ConstGlobalNames = immutableStringView(program.ConstGlobalNames)
	program.IncludeDependencies = immutableStringView(program.IncludeDependencies)

	for i := range program.Constants {
		program.Constants[i].Names = immutableStringView(program.Constants[i].Names)
	}

	return program
}

func immutableStringView(values []string) []string {
	if len(values) == 0 {
		return nil
	}
	return values[:len(values):len(values)]
}

func estimateProgramSizeBytes(program CachedProgram) int64 {
	size := int64(len(program.Bytecode))
	size += int64(len(program.Constants)) * int64(64)
	for i := range program.Constants {
		size += int64(len(program.Constants[i].Str))
		for _, name := range program.Constants[i].Names {
			size += int64(len(name))
		}
	}
	size += estimateStringSliceSize(program.GlobalNames)
	size += estimateStringSliceSize(program.DeclaredGlobalNames)
	size += estimateStringSliceSize(program.ConstGlobalNames)
	size += estimateStringSliceSize(program.GlobalPreludeNames)
	size += estimateStringSliceSize(program.GlobalPreludeConsts)
	size += estimateStringSliceSize(program.UserGlobalNames)
	size += estimateStringSliceSize(program.UserDeclaredGlobals)
	size += estimateStringSliceSize(program.UserConstGlobals)
	size += estimateStringSliceSize(program.GlobalZeroArgFuncs)
	size += estimateStringSliceSize(program.IncludeDependencies)
	size += int64(len(program.SourceName))
	if size < 1 {
		return 1
	}
	return size
}

func estimateStringSliceSize(values []string) int64 {
	size := int64(len(values)) * 16
	for _, value := range values {
		size += int64(len(value))
	}
	return size
}

func sortedTrueKeys(values map[string]bool) []string {
	if len(values) == 0 {
		return nil
	}
	keys := make([]string, 0, len(values))
	for key, active := range values {
		if !active {
			continue
		}
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}

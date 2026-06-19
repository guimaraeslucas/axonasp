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
	"log"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/cespare/xxhash/v2"
	"github.com/fsnotify/fsnotify"
	"golang.org/x/sync/singleflight"
)

const (
	scriptCacheDependencyMapLimit = 1000
	scriptCacheMagicSize          = 6
	scriptCacheBinaryVersion      = uint16(13)
	scriptCacheDebounceWindow     = 1000 * time.Millisecond
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
	IncludeSiteRoot     string
	GlobalPreludeNames  []string
	GlobalPreludeConsts []string
	UserGlobalNames     []string
	UserDeclaredGlobals []string
	UserConstGlobals    []string
	GlobalZeroArgFuncs  []string
	EngineMode          EngineMode
	ProgramHash         uint64
	GlobalNamesLower    []string

	// Legacy fields kept for backward compatibility with v2 cache payloads.
	GlobalNames         []string
	DeclaredGlobalNames []string
	ConstGlobalNames    []string
	// GlobalTypeNames stores VB6 As Type declarations for global variables.
	// Each entry is "name:type" where type is the ValueType constant as an integer.
	// This is populated from the compiler's globalVarTypes during cache creation.
	GlobalTypeNames []string
	// FuncParamDefaults maps function entry-point offsets to per-parameter default
	// constant indices for Optional parameters.
	FuncParamDefaults   map[int][]int
	IncludeDependencies []string
	SourceMapEntries    []SourceMapEntry
	// RecordDecls and RecordDeclLookup carry compiled UDT metadata required by
	// ExtOpInitRecord/ExtOpGetRecordMember/ExtOpSetRecordMember in cached VM startup paths.
	RecordDecls      []CompiledRecordDecl
	RecordDeclLookup map[string]int

	// Phase 5: Interface/Class tracking for typed variables.
	GlobalClassNames []string // "name:classname"
	LocalVarTypes    map[string]ValueType
	LocalClassTypes  map[string]string

	// JScript inline cache node count for VM-local icState pre-allocation.
	JSICNodeCount uint32
}

// cachedProgramBinaryPayload stores the serialized disk representation.
type cachedProgramBinaryPayload struct {
	ModTime int64
	Program CachedProgram
}

// ScriptCompileOptions controls context-sensitive compilation behavior.
type ScriptCompileOptions struct {
	IncludeSiteRoot string
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

	if uint64(len(p.Program.Bytecode)) > uint64(^uint32(0)) {
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

	if uint64(len(p.Program.Constants)) > uint64(^uint32(0)) {
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
	if err := writeString(buffered, p.Program.IncludeSiteRoot); err != nil {
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
	if err := writeStringSlice(buffered, p.Program.GlobalNamesLower); err != nil {
		return err
	}
	if err := binary.Write(buffered, binary.LittleEndian, uint8(p.Program.EngineMode)); err != nil {
		return err
	}
	if err := writeIntSliceMap(buffered, p.Program.FuncParamDefaults); err != nil {
		return err
	}
	if err := writeSourceMapEntries(buffered, p.Program.SourceMapEntries); err != nil {
		return err
	}
	if err := binary.Write(buffered, binary.LittleEndian, p.Program.JSICNodeCount); err != nil {
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
	if version != 1 && version != 2 && version != 8 && version != 10 && version != scriptCacheBinaryVersion {
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
		if version >= 10 {
			includeSiteRoot, err := readString(reader)
			if err != nil {
				return err
			}
			p.Program.IncludeSiteRoot = includeSiteRoot
		}
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
			if version >= 7 {
				lower, err := readStringSlice(reader)
				if err != nil {
					return err
				}
				p.Program.GlobalNamesLower = lower
			}
			if version >= 8 {
				var engineMode uint8
				if err := binary.Read(reader, binary.LittleEndian, &engineMode); err != nil {
					return err
				}
				p.Program.EngineMode = EngineMode(engineMode)
				if version >= 9 {
					funcParamDefaults, err := readIntSliceMap(reader)
					if err != nil {
						return err
					}
					p.Program.FuncParamDefaults = funcParamDefaults
					if version >= 11 {
						sourceMapEntries, err := readSourceMapEntries(reader)
						if err != nil {
							return err
						}
						p.Program.SourceMapEntries = sourceMapEntries
						if version >= 13 {
							var jsICNodeCount uint32
							if err := binary.Read(reader, binary.LittleEndian, &jsICNodeCount); err != nil {
								return err
							}
							p.Program.JSICNodeCount = jsICNodeCount
						}
					}
				}
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
			p.Program.Constants,
		)
	}

	return nil
}

// ScriptCache implements two-tier ASP bytecode caching and dependency invalidation.
type ScriptCache struct {
	mu                  sync.RWMutex
	mode                BytecodeCacheMode
	cacheDir            string
	programs            map[string]CachedProgram
	programSizes        map[string]int64
	programOrder        []string
	totalBytes          int64
	maxBytes            int64
	dependencyMap       map[string][]string
	scriptDependencies  map[string][]string
	dependencyOrder     []string
	sg                  singleflight.Group
	watcher             *fsnotify.Watcher
	watchRoots          []string
	watchedPaths        map[string]struct{}
	watchedExt          map[string]struct{}
	watchDebounce       map[string]time.Time
	watchDebounceWindow time.Duration
	watchStop           chan struct{}
	watcherActive       bool
	watcherErrorCount   uint32

	// Engine configuration for mode-based compilation.
	engineMode   EngineMode
	executeAsASP []string
	executeAsVBS []string
	executeAsJS  []string
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
		mode:                mode,
		cacheDir:            cacheDir,
		programs:            make(map[string]CachedProgram),
		programSizes:        make(map[string]int64),
		programOrder:        make([]string, 0, 64),
		dependencyMap:       make(map[string][]string),
		scriptDependencies:  make(map[string][]string),
		dependencyOrder:     make([]string, 0, 256),
		watchedPaths:        make(map[string]struct{}),
		watchedExt:          defaultScriptWatchExtensions(),
		watchDebounce:       make(map[string]time.Time, 128),
		watchDebounceWindow: scriptCacheDebounceWindow,
		maxBytes:            int64(maxSizeMB) * 1024 * 1024,
	}
}

// SetEngineConfig updates the engine mode and allowed file extensions for the cache.
func (c *ScriptCache) SetEngineConfig(mode EngineMode, aspExt, vbsExt, jsExt []string) {
	if c == nil {
		return
	}
	c.mu.Lock()
	defer c.mu.Unlock()
	c.engineMode = mode
	c.executeAsASP = aspExt
	c.executeAsVBS = vbsExt
	c.executeAsJS = jsExt
}

func (c *ScriptCache) resolveEngineMode(filePath string) EngineMode {
	if c == nil {
		return EngineModeDefault
	}
	c.mu.RLock()
	defer c.mu.RUnlock()

	ext := strings.ToLower(filepath.Ext(filePath))

	// Check if it matches VBScript extensions
	for _, e := range c.executeAsVBS {
		if e == ext {
			return EngineModeVBScript
		}
	}

	// Check if it matches JavaScript extensions
	for _, e := range c.executeAsJS {
		if e == ext {
			return EngineModeJavaScript
		}
	}

	// Always check ASP extensions or as default
	for _, e := range c.executeAsASP {
		if e == ext {
			return EngineModeDefault
		}
	}

	// Legacy behavior: if it's .js/.mjs, it's JScript.
	if ext == ".js" || ext == ".mjs" {
		return EngineModeJavaScript
	}

	return EngineModeDefault
}

// SetWatchedExtensions configures extra file extensions treated as ASP scripts by the invalidator.
func (c *ScriptCache) SetWatchedExtensions(extensions []string) {
	if c == nil {
		return
	}
	updated := defaultScriptWatchExtensions()
	for _, ext := range extensions {
		normalized := strings.ToLower(strings.TrimSpace(ext))
		if normalized == "" {
			continue
		}
		if !strings.HasPrefix(normalized, ".") {
			normalized = "." + normalized
		}
		updated[normalized] = struct{}{}
	}
	c.mu.Lock()
	c.watchedExt = updated
	c.mu.Unlock()
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
	if c.watchedPaths == nil {
		c.watchedPaths = make(map[string]struct{}, 256)
	} else {
		clear(c.watchedPaths)
	}
	if c.watchDebounce == nil {
		c.watchDebounce = make(map[string]time.Time, 128)
	} else {
		clear(c.watchDebounce)
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
				if event.Op&(fsnotify.Write|fsnotify.Create|fsnotify.Remove) == 0 {
					continue
				}
				if !c.shouldWatchScriptPath(event.Name) {
					continue
				}
				if event.Op&fsnotify.Create != 0 {
					_ = c.addWatchPathTracked(watcher, event.Name)
				}
				if event.Op&fsnotify.Remove != 0 {
					c.removeWatchPathTracked(watcher, event.Name, false)
				}
				if !c.shouldProcessWatchEvent(event.Name) {
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
	if c.watchedPaths != nil {
		clear(c.watchedPaths)
	}
	if c.watchDebounce != nil {
		clear(c.watchDebounce)
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
	return c.getByCacheKey(normalizeScriptCacheKey(normalized))
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
	cacheKey := normalizeScriptCacheKey(normalized)

	program = immutableCachedProgramView(program)
	sizeBytes := estimateProgramSizeBytes(program)
	if sizeBytes <= 0 {
		return
	}
	c.putByCacheKey(cacheKey, program, dependencies, sizeBytes)
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
	cacheKey := normalizeScriptCacheKey(normalized)

	c.mu.Lock()
	defer c.mu.Unlock()

	invalidateList := make([]string, 0, 16)
	invalidateSet := make(map[string]struct{}, 16)
	invalidateSet[cacheKey] = struct{}{}
	invalidateList = append(invalidateList, cacheKey)
	if dependents, exists := c.dependencyMap[cacheKey]; exists {
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
	return c.LoadOrCompileWithModeAndOptions(filePath, ExecutionModeServer, ScriptCompileOptions{})
}

// LoadOrCompileWithMode compiles and caches an ASP script, optionally bypassing cache for
// interactive execution modes (CLI, TUI, eval) to prevent stalls. In interactive modes,
// neither memory nor disk caches are consulted or updated, ensuring fresh compilation.
func (c *ScriptCache) LoadOrCompileWithMode(filePath string, mode ExecutionMode) (CachedProgram, error) {
	return c.LoadOrCompileWithModeAndOptions(filePath, mode, ScriptCompileOptions{})
}

// LoadOrCompileWithOptions compiles and caches an ASP script with caller-provided context.
func (c *ScriptCache) LoadOrCompileWithOptions(filePath string, options ScriptCompileOptions) (CachedProgram, error) {
	return c.LoadOrCompileWithModeAndOptions(filePath, ExecutionModeServer, options)
}

func includeSiteRootMatches(program CachedProgram, options ScriptCompileOptions) bool {
	requestedRoot := normalizeIncludeSiteRoot(options.IncludeSiteRoot)
	programRoot := normalizeIncludeSiteRoot(program.IncludeSiteRoot)
	if requestedRoot == "" {
		return true
	}
	if programRoot == "" {
		return false
	}
	return strings.EqualFold(requestedRoot, programRoot)
}

// LoadOrCompileWithModeAndOptions compiles and caches with explicit include-resolution context.
func (c *ScriptCache) LoadOrCompileWithModeAndOptions(filePath string, mode ExecutionMode, options ScriptCompileOptions) (CachedProgram, error) {
	normalized, err := c.normalizeAbsolutePath(filePath)
	if err != nil {
		return CachedProgram{}, err
	}
	cacheKey := normalizeScriptCacheKey(normalized)

	// If in interactive mode, bypass cache entirely to prevent stalls
	if mode != ExecutionModeServer {
		return c.compileOnly(normalized, options)
	}

	if program, found := c.getByCacheKey(cacheKey); found {
		if includeSiteRootMatches(program, options) {
			return program, nil
		}
	}

	if c == nil {
		return c.compileOnly(normalized, options)
	}

	// singleflight deduplicates concurrent compilations for the same cache key,
	// preventing cache-stampede memory exhaustion under concurrent cache misses.
	resultIface, err, _ := c.sg.Do(cacheKey, func() (interface{}, error) {
		// Re-check memory cache after acquiring the singleflight slot;
		// another goroutine may have populated it while this call was queued.
		if program, found := c.getByCacheKey(cacheKey); found {
			if includeSiteRootMatches(program, options) {
				return &program, nil
			}
		}

		sourceInfo, statErr := os.Stat(normalized)
		if statErr != nil {
			return nil, statErr
		}

		if c.mode.HasDiskTier() && strings.TrimSpace(options.IncludeSiteRoot) == "" {
			if program, found := c.loadDiskProgram(normalized, sourceInfo); found {
				if c.mode.HasMemoryTier() {
					c.putByCacheKey(cacheKey, program, program.IncludeDependencies, estimateProgramSizeBytes(program))
				}
				return &program, nil
			}
		}

		content, readErr := os.ReadFile(normalized)
		if readErr != nil {
			return nil, readErr
		}

		// Strip UTF-8 BOM if present to prevent parsing errors
		content = stripUTF8BOM(content)

		var compiler *Compiler
		resolvedMode := c.resolveEngineMode(normalized)
		switch resolvedMode {
		case EngineModeJavaScript:
			compiler = NewJavaScriptCompiler(string(content))
		case EngineModeVBScript:
			compiler = NewCompiler(string(content))
		default:
			compiler = NewASPCompiler(string(content))
		}

		compiler.SetSourceName(cacheKey)
		compiler.SetIncludeSiteRoot(options.IncludeSiteRoot)
		if compErr := compiler.Compile(); compErr != nil {
			return nil, compErr
		}

		program := buildCachedProgramFromCompiler(compiler)

		if c.mode.HasDiskTier() && strings.TrimSpace(options.IncludeSiteRoot) == "" {
			if storeErr := c.storeDiskProgram(normalized, sourceInfo.ModTime(), program); storeErr != nil {
				log.Printf("Warning: failed to persist bytecode cache to disk for %s: %v", normalized, storeErr)
			}
		}
		if c.mode.HasMemoryTier() {
			c.putByCacheKey(cacheKey, program, program.IncludeDependencies, estimateProgramSizeBytes(program))
		}
		result := immutableCachedProgramView(program)
		return &result, nil
	})

	if err != nil {
		return CachedProgram{}, err
	}

	program := *resultIface.(*CachedProgram)
	if !includeSiteRootMatches(program, options) {
		return c.compileOnly(normalized, options)
	}
	return program, nil
}

// compileOnly compiles a script without using any cache layer.
// Used for interactive execution modes (CLI, TUI, eval) to prevent stalls.
func (c *ScriptCache) compileOnly(filePath string, options ScriptCompileOptions) (CachedProgram, error) {
	content, err := os.ReadFile(filePath)
	if err != nil {
		return CachedProgram{}, err
	}

	// Strip UTF-8 BOM if present to prevent parsing errors
	content = stripUTF8BOM(content)

	var compiler *Compiler
	resolvedMode := c.resolveEngineMode(filePath)
	switch resolvedMode {
	case EngineModeJavaScript:
		compiler = NewJavaScriptCompiler(string(content))
	case EngineModeVBScript:
		compiler = NewCompiler(string(content))
	default:
		compiler = NewASPCompiler(string(content))
	}

	compiler.SetSourceName(filePath)
	compiler.SetIncludeSiteRoot(options.IncludeSiteRoot)
	if err := compiler.Compile(); err != nil {
		return CachedProgram{}, err
	}

	return buildCachedProgramFromCompiler(compiler), nil
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
	payload.Program.SourceName = normalizeScriptCacheKey(filePath)
	payload.Program.ProgramHash = computeProgramHash(
		payload.Program.Bytecode,
		payload.Program.GlobalCount,
		payload.Program.OptionCompare,
		payload.Program.OptionExplicit,
		payload.Program.SourceName,
		payload.Program.Constants,
	)
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
	if len(program.SourceMapEntries) > 0 {
		vm.sourceMap = SourceMap{entries: cloneSourceMapEntries(program.SourceMapEntries)}
	} else {
		vm.sourceMap = buildIdentitySourceMap(program.SourceName)
	}
	vm.engineMode = program.EngineMode
	applyProgramGlobalMetadata(vm, program)
	clear(vm.funcParamDefaults)
	for entryPoint, defaults := range program.FuncParamDefaults {
		if len(defaults) == 0 {
			vm.funcParamDefaults[entryPoint] = nil
			continue
		}
		dup := make([]int, len(defaults))
		copy(dup, defaults)
		vm.funcParamDefaults[entryPoint] = dup
	}
	vm.RecordDecls = append(vm.RecordDecls[:0], program.RecordDecls...)
	clear(vm.RecordDeclLookup)
	for k, v := range program.RecordDeclLookup {
		vm.RecordDeclLookup[k] = v
	}
	if program.JSICNodeCount > 0 {
		vm.icState = make([]InlineCacheSlot, program.JSICNodeCount)
	}
	vm.applyLocalVarTypesFromMaps(program.LocalVarTypes, program.LocalClassTypes)
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
	hash := xxhash.Sum64String(normalizeScriptCacheKey(filePath))
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

func normalizeScriptCacheKey(filePath string) string {
	normalized := filepath.Clean(strings.TrimSpace(filePath))
	if runtime.GOOS == "windows" {
		return strings.ToLower(normalized)
	}
	return normalized
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
		normalized = normalizeScriptCacheKey(normalized)
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
		if c.shouldWatchScriptPath(rootDir) {
			return c.addWatchPathTracked(watcher, rootDir)
		}
		return nil
	}
	return filepath.WalkDir(rootDir, func(path string, dirEntry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return nil
		}
		if dirEntry.IsDir() {
			if shouldSkipScriptWatchDir(path, dirEntry.Name()) {
				return filepath.SkipDir
			}
			return nil
		}
		if !c.shouldWatchScriptPath(path) {
			return nil
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
	normalized = normalizeScriptCacheKey(normalized)

	c.mu.RLock()
	_, exists := c.watchedPaths[normalized]
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
	if c.watchedPaths == nil {
		c.watchedPaths = make(map[string]struct{}, 256)
	}
	c.watchedPaths[normalized] = struct{}{}
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
	normalized = normalizeScriptCacheKey(normalized)

	removals := make([]string, 0, 8)
	c.mu.RLock()
	for watchedPath := range c.watchedPaths {
		if strings.EqualFold(watchedPath, normalized) || (includeChildren && pathHasPrefixFold(watchedPath, normalized)) {
			removals = append(removals, watchedPath)
		}
	}
	c.mu.RUnlock()

	for _, watchedPath := range removals {
		_ = watcher.Remove(watchedPath)
		c.mu.Lock()
		delete(c.watchedPaths, watchedPath)
		c.mu.Unlock()
	}
}

func (c *ScriptCache) pruneStaleWatches(watcher *fsnotify.Watcher) {
	if c == nil || watcher == nil {
		return
	}
	stale := make([]string, 0, 16)
	c.mu.RLock()
	for watchedPath := range c.watchedPaths {
		_, err := os.Stat(watchedPath)
		if err != nil {
			stale = append(stale, watchedPath)
		}
	}
	c.mu.RUnlock()

	for _, watchedPath := range stale {
		_ = watcher.Remove(watchedPath)
		c.mu.Lock()
		delete(c.watchedPaths, watchedPath)
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

func defaultScriptWatchExtensions() map[string]struct{} {
	return map[string]struct{}{
		".asp": {},
		".inc": {},
		".asa": {},
	}
}

func shouldSkipScriptWatchDir(path string, dirName string) bool {
	if strings.EqualFold(dirName, ".git") || strings.EqualFold(dirName, "node_modules") {
		return true
	}
	if strings.EqualFold(dirName, "temp") {
		return true
	}
	if strings.EqualFold(dirName, "logs") {
		return true
	}
	if strings.EqualFold(dirName, "log") {
		return true
	}
	_ = path
	return false
}

func (c *ScriptCache) shouldWatchScriptPath(path string) bool {
	if c == nil {
		return false
	}
	ext := strings.ToLower(filepath.Ext(path))
	if ext == "" {
		return false
	}
	c.mu.RLock()
	_, ok := c.watchedExt[ext]
	c.mu.RUnlock()
	return ok
}

func (c *ScriptCache) shouldProcessWatchEvent(path string) bool {
	if c == nil {
		return false
	}
	key := normalizeScriptCacheKey(path)
	now := time.Now()
	c.mu.Lock()
	defer c.mu.Unlock()
	if c.watchDebounce == nil {
		c.watchDebounce = make(map[string]time.Time, 128)
	}
	if c.watchDebounceWindow <= 0 {
		c.watchDebounceWindow = scriptCacheDebounceWindow
	}
	if last, ok := c.watchDebounce[key]; ok {
		if now.Sub(last) < c.watchDebounceWindow {
			return false
		}
	}
	c.watchDebounce[key] = now
	if len(c.watchDebounce) > 4096 {
		cutoff := now.Add(-4 * c.watchDebounceWindow)
		for watchedKey, ts := range c.watchDebounce {
			if ts.Before(cutoff) {
				delete(c.watchDebounce, watchedKey)
			}
		}
	}
	return true
}

func (c *ScriptCache) getByCacheKey(cacheKey string) (CachedProgram, bool) {
	if c == nil || !c.mode.HasMemoryTier() {
		return CachedProgram{}, false
	}
	c.mu.RLock()
	program, ok := c.programs[cacheKey]
	c.mu.RUnlock()
	if !ok {
		return CachedProgram{}, false
	}
	return immutableCachedProgramView(program), true
}

func (c *ScriptCache) putByCacheKey(cacheKey string, program CachedProgram, dependencies []string, sizeBytes int64) {
	if c == nil || !c.mode.HasMemoryTier() {
		return
	}
	if sizeBytes <= 0 {
		sizeBytes = estimateProgramSizeBytes(program)
		if sizeBytes <= 0 {
			return
		}
	}
	program = immutableCachedProgramView(program)
	c.mu.Lock()
	defer c.mu.Unlock()
	if sizeBytes > c.maxBytes {
		c.removeProgramNoLock(cacheKey)
		c.removeScriptDependenciesNoLock(cacheKey)
		return
	}

	c.removeProgramNoLock(cacheKey)
	for c.totalBytes+sizeBytes > c.maxBytes && len(c.programOrder) > 0 {
		victim := c.programOrder[0]
		c.removeProgramNoLock(victim)
	}

	c.programs[cacheKey] = program
	c.programSizes[cacheKey] = sizeBytes
	c.programOrder = append(c.programOrder, cacheKey)
	c.totalBytes += sizeBytes
	c.registerDependenciesNoLock(cacheKey, dependencies)
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
	if uint64(len(value)) > uint64(^uint32(0)) {
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
	if uint64(len(values)) > uint64(^uint32(0)) {
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

func writeIntSliceMap(writer io.Writer, values map[int][]int) error {
	if uint64(len(values)) > uint64(^uint32(0)) {
		return errors.New("map too large")
	}
	if err := binary.Write(writer, binary.LittleEndian, uint32(len(values))); err != nil {
		return err
	}
	for key, slice := range values {
		if err := binary.Write(writer, binary.LittleEndian, int32(key)); err != nil {
			return err
		}
		if uint64(len(slice)) > uint64(^uint32(0)) {
			return errors.New("slice too large")
		}
		if err := binary.Write(writer, binary.LittleEndian, uint32(len(slice))); err != nil {
			return err
		}
		for i := 0; i < len(slice); i++ {
			if err := binary.Write(writer, binary.LittleEndian, int32(slice[i])); err != nil {
				return err
			}
		}
	}
	return nil
}

func readIntSliceMap(reader io.Reader) (map[int][]int, error) {
	var length uint32
	if err := binary.Read(reader, binary.LittleEndian, &length); err != nil {
		return nil, err
	}
	if length == 0 {
		return nil, nil
	}
	values := make(map[int][]int, int(length))
	for i := 0; i < int(length); i++ {
		var rawKey int32
		if err := binary.Read(reader, binary.LittleEndian, &rawKey); err != nil {
			return nil, err
		}
		var sliceLen uint32
		if err := binary.Read(reader, binary.LittleEndian, &sliceLen); err != nil {
			return nil, err
		}
		slice := make([]int, int(sliceLen))
		for j := 0; j < int(sliceLen); j++ {
			var rawVal int32
			if err := binary.Read(reader, binary.LittleEndian, &rawVal); err != nil {
				return nil, err
			}
			slice[j] = int(rawVal)
		}
		values[int(rawKey)] = slice
	}
	return values, nil
}

func writeSourceMapEntries(writer io.Writer, entries []SourceMapEntry) error {
	if uint64(len(entries)) > uint64(^uint32(0)) {
		return errors.New("source map entries too large")
	}
	if err := binary.Write(writer, binary.LittleEndian, uint32(len(entries))); err != nil {
		return err
	}
	for i := 0; i < len(entries); i++ {
		entry := entries[i]
		if err := binary.Write(writer, binary.LittleEndian, int32(entry.MergedLineStart)); err != nil {
			return err
		}
		if err := writeString(writer, entry.OriginalFile); err != nil {
			return err
		}
		if err := binary.Write(writer, binary.LittleEndian, int32(entry.OriginalLine)); err != nil {
			return err
		}
	}
	return nil
}

func readSourceMapEntries(reader io.Reader) ([]SourceMapEntry, error) {
	var length uint32
	if err := binary.Read(reader, binary.LittleEndian, &length); err != nil {
		return nil, err
	}
	if length == 0 {
		return nil, nil
	}
	entries := make([]SourceMapEntry, int(length))
	for i := 0; i < int(length); i++ {
		var mergedLineStart int32
		if err := binary.Read(reader, binary.LittleEndian, &mergedLineStart); err != nil {
			return nil, err
		}
		originalFile, err := readString(reader)
		if err != nil {
			return nil, err
		}
		var originalLine int32
		if err := binary.Read(reader, binary.LittleEndian, &originalLine); err != nil {
			return nil, err
		}
		entries[i] = SourceMapEntry{
			MergedLineStart: int(mergedLineStart),
			OriginalFile:    originalFile,
			OriginalLine:    int(originalLine),
		}
	}
	return entries, nil
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
	case VTString, VTUserSub, VTJSFunctionTemplate, VTJSArrowFunctionTemplate, VTSymbol:
		if err := writeString(writer, value.Str); err != nil {
			return err
		}
	case VTEmpty, VTNull, VTBool, VTInteger, VTDouble, VTDate, VTObject, VTNativeObject, VTBuiltin,
		VTJSUndefined, VTJSObject, VTJSFunction:
		// Type, Num and Flt are sufficient for deterministic reconstruction.
	default:
		return NewAxonASPError(ErrInvalidCacheFile, nil, ErrInvalidCacheFile.String(), "", 0)
	}

	if value.Type == VTUserSub || value.Type == VTJSFunctionTemplate || value.Type == VTJSArrowFunctionTemplate {
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
	case VTString, VTUserSub, VTJSFunctionTemplate, VTJSArrowFunctionTemplate, VTSymbol:
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

	if valueType == VTUserSub || valueType == VTJSFunctionTemplate || valueType == VTJSArrowFunctionTemplate {
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
		IncludeSiteRoot:     program.IncludeSiteRoot,
		GlobalPreludeNames:  cloneStringSlice(program.GlobalPreludeNames),
		GlobalPreludeConsts: cloneStringSlice(program.GlobalPreludeConsts),
		UserGlobalNames:     cloneStringSlice(program.UserGlobalNames),
		UserDeclaredGlobals: cloneStringSlice(program.UserDeclaredGlobals),
		UserConstGlobals:    cloneStringSlice(program.UserConstGlobals),
		GlobalZeroArgFuncs:  cloneStringSlice(program.GlobalZeroArgFuncs),
		GlobalTypeNames:     cloneStringSlice(program.GlobalTypeNames),
		FuncParamDefaults:   cloneIntSliceMap(program.FuncParamDefaults),
		ProgramHash:         program.ProgramHash,
		GlobalNamesLower:    cloneStringSlice(program.GlobalNamesLower),
		GlobalNames:         cloneStringSlice(program.GlobalNames),
		DeclaredGlobalNames: cloneStringSlice(program.DeclaredGlobalNames),
		ConstGlobalNames:    cloneStringSlice(program.ConstGlobalNames),
		IncludeDependencies: cloneStringSlice(program.IncludeDependencies),
		SourceMapEntries:    cloneSourceMapEntries(program.SourceMapEntries),
		RecordDecls:         cloneRecordDeclSlice(program.RecordDecls),
		RecordDeclLookup:    cloneIntMap(program.RecordDeclLookup),
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

func cloneRecordDeclSlice(values []CompiledRecordDecl) []CompiledRecordDecl {
	if len(values) == 0 {
		return nil
	}
	cloned := make([]CompiledRecordDecl, len(values))
	for i := range values {
		cloned[i].Name = values[i].Name
		if len(values[i].Members) > 0 {
			members := make([]CompiledRecordMemberDecl, len(values[i].Members))
			copy(members, values[i].Members)
			cloned[i].Members = members
		}
	}
	return cloned
}

func cloneSourceMapEntries(values []SourceMapEntry) []SourceMapEntry {
	if len(values) == 0 {
		return nil
	}
	cloned := make([]SourceMapEntry, len(values))
	copy(cloned, values)
	return cloned
}

func cloneIntMap(values map[string]int) map[string]int {
	if len(values) == 0 {
		return nil
	}
	cloned := make(map[string]int, len(values))
	for k, v := range values {
		cloned[k] = v
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
	program.GlobalTypeNames = immutableStringView(program.GlobalTypeNames)
	program.GlobalNames = immutableStringView(program.GlobalNames)
	program.DeclaredGlobalNames = immutableStringView(program.DeclaredGlobalNames)
	program.ConstGlobalNames = immutableStringView(program.ConstGlobalNames)
	program.IncludeDependencies = immutableStringView(program.IncludeDependencies)
	program.GlobalNamesLower = immutableStringView(program.GlobalNamesLower)
	program.RecordDecls = immutableRecordDeclView(program.RecordDecls)

	// Clone constants to avoid mutating shared Value.Names fields under concurrent access.
	if len(program.Constants) > 0 {
		cloned := make([]Value, len(program.Constants))
		copy(cloned, program.Constants)
		for i := range cloned {
			cloned[i].Names = immutableStringView(cloned[i].Names)
		}
		program.Constants = cloned[:len(cloned):len(cloned)]
	}
	if len(program.FuncParamDefaults) > 0 {
		cloned := make(map[int][]int, len(program.FuncParamDefaults))
		for key, values := range program.FuncParamDefaults {
			cloned[key] = immutableIntView(values)
		}
		program.FuncParamDefaults = cloned
	}

	return program
}

func immutableIntView(values []int) []int {
	if len(values) == 0 {
		return nil
	}
	return values[:len(values):len(values)]
}

func cloneIntSliceMap(values map[int][]int) map[int][]int {
	if len(values) == 0 {
		return nil
	}
	cloned := make(map[int][]int, len(values))
	for k, v := range values {
		if len(v) == 0 {
			cloned[k] = nil
			continue
		}
		slice := make([]int, len(v))
		copy(slice, v)
		cloned[k] = slice
	}
	return cloned
}

func immutableRecordDeclView(values []CompiledRecordDecl) []CompiledRecordDecl {
	if len(values) == 0 {
		return nil
	}
	// Clone to avoid mutating shared struct fields under concurrent access.
	cloned := make([]CompiledRecordDecl, len(values))
	copy(cloned, values)
	for i := range cloned {
		if len(cloned[i].Members) > 0 {
			cloned[i].Members = cloned[i].Members[:len(cloned[i].Members):len(cloned[i].Members)]
		}
	}
	return cloned[:len(cloned):len(cloned)]
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
	size += estimateStringSliceSize(program.GlobalTypeNames)
	size += estimateStringSliceSize(program.IncludeDependencies)
	size += estimateStringSliceSize(program.GlobalNamesLower)
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

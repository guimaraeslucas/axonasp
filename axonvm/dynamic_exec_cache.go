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
	"strings"
	"sync"

	lru "github.com/hashicorp/golang-lru/v2"
)

const dynamicExecBytecodeCacheSize = 512

type dynamicExecKind uint8

const (
	dynamicExecKindExecute dynamicExecKind = iota + 1
	dynamicExecKindExecuteGlobal
)

// dynamicCachedProgram stores one immutable compiled Execute/ExecuteGlobal payload.
type dynamicCachedProgram struct {
	keyHash          uint64
	kind             dynamicExecKind
	source           string
	sourceName       string
	optionCompare    int
	optionExplicit   bool
	globalNamesHash  uint64
	localScopeHash   uint64
	classVersion     uint64
	activeClassName  string
	globalCount      int
	constants        []Value
	bytecode         []byte
	compilerGlobals  []string
	compilerZeroArg  map[string]bool
	compilerDeclared map[string]bool
	compilerConst    map[string]bool
}

var (
	dynamicExecProgramCacheOnce sync.Once
	dynamicExecProgramCache     *lru.Cache[uint64, *dynamicCachedProgram]
)

// getDynamicExecProgramCache lazily initializes one process-wide dynamic execution cache.
func getDynamicExecProgramCache() *lru.Cache[uint64, *dynamicCachedProgram] {
	dynamicExecProgramCacheOnce.Do(func() {
		cache, err := lru.New[uint64, *dynamicCachedProgram](dynamicExecBytecodeCacheSize)
		if err == nil {
			dynamicExecProgramCache = cache
		}
	})
	return dynamicExecProgramCache
}

// cloneBoolMap clones one map[string]bool preserving keys and values.
func cloneBoolMap(values map[string]bool) map[string]bool {
	if len(values) == 0 {
		return map[string]bool{}
	}
	cloned := make(map[string]bool, len(values))
	for key, value := range values {
		cloned[key] = value
	}
	return cloned
}

// applyCompilerSnapshot updates runtime compiler metadata from one cached dynamic fragment.
func (vm *VM) applyCompilerSnapshot(compiled *dynamicCachedProgram) {
	if vm == nil || compiled == nil {
		return
	}
	vm.optionCompare = compiled.optionCompare
	vm.optionExplicit = compiled.optionExplicit
	vm.globalNames = append(vm.globalNames[:0], compiled.compilerGlobals...)
	vm.rebuildGlobalNameIndex()
	clear(vm.globalZeroArgFuncs)
	for key, value := range compiled.compilerZeroArg {
		vm.globalZeroArgFuncs[key] = value
	}
	clear(vm.declaredGlobals)
	for key, value := range compiled.compilerDeclared {
		vm.declaredGlobals[key] = value
	}
	clear(vm.constGlobals)
	for key, value := range compiled.compilerConst {
		vm.constGlobals[key] = value
	}
	vm.sourceName = compiled.sourceName
}

// activeDynamicClassName resolves the currently executing runtime class name.
func (vm *VM) activeDynamicClassName() string {
	if vm == nil || vm.activeClassObjectID == 0 {
		return ""
	}
	instance, exists := vm.runtimeClassItems[vm.activeClassObjectID]
	if !exists || instance == nil {
		return ""
	}
	return strings.ToLower(strings.TrimSpace(instance.ClassName))
}

// buildDynamicExecCacheKey hashes execute source and compatibility scope metadata.
func buildDynamicExecCacheKey(kind dynamicExecKind, source string, optionCompare int, globalNamesHash uint64, localScopeHash uint64, classVersion uint64, activeClassName string) uint64 {
	hash := uint64(fnvOffset64)
	hash = hashByteFNV1a(hash, byte(kind))
	hash = hashInt64FNV1a(hash, int64(optionCompare))
	hash = hashInt64FNV1a(hash, int64(globalNamesHash))
	hash = hashInt64FNV1a(hash, int64(localScopeHash))
	hash = hashInt64FNV1a(hash, int64(classVersion))
	hash = hashStringFNV1a(hash, activeClassName)
	hash = hashStringFNV1a(hash, source)
	return hash
}

// getOrCompileDynamicProgram returns one cached Execute/ExecuteGlobal compiled payload.
func (vm *VM) getOrCompileDynamicProgram(source string, localSub Value, kind dynamicExecKind) (*dynamicCachedProgram, error) {
	if vm == nil {
		return nil, nil
	}

	optionCompare := vm.optionCompare
	globalScopeHash := vm.globalNamesHash
	localScopeHash := uint64(0)
	activeClassName := ""
	if kind == dynamicExecKindExecute {
		localScopeHash = buildLocalScopeHash(localSub)
		activeClassName = vm.activeDynamicClassName()
	}
	classVersion := vm.runtimeClassVersion
	key := buildDynamicExecCacheKey(kind, source, optionCompare, globalScopeHash, localScopeHash, classVersion, activeClassName)
	cache := getDynamicExecProgramCache()
	if !vm.IsInteractiveMode() && cache != nil {
		if cached, ok := cache.Get(key); ok && cached != nil {
			if cached.keyHash == key &&
				cached.kind == kind &&
				cached.source == source &&
				cached.optionCompare == optionCompare &&
				cached.globalNamesHash == globalScopeHash &&
				cached.localScopeHash == localScopeHash &&
				cached.classVersion == classVersion &&
				cached.activeClassName == activeClassName {
				return cached, nil
			}
		}
	}

	var compiler *Compiler
	switch kind {
	case dynamicExecKindExecuteGlobal:
		compiler = vm.newExecuteGlobalPureCompiler(source)
		compiler.SetSourceName(vm.sourceName + "_ExecuteGlobal")
	default:
		compiler = vm.newExecuteLocalCompiler(source, localSub, false)
		compiler.SetSourceName(vm.sourceName + "_Execute")
	}
	if err := compiler.Compile(); err != nil {
		return nil, err
	}

	compiled := &dynamicCachedProgram{
		keyHash:          key,
		kind:             kind,
		source:           source,
		sourceName:       compiler.sourceName,
		optionCompare:    compiler.optionCompare,
		optionExplicit:   compiler.optionExplicit,
		globalNamesHash:  globalScopeHash,
		localScopeHash:   localScopeHash,
		classVersion:     classVersion,
		activeClassName:  activeClassName,
		globalCount:      compiler.GlobalsCount(),
		constants:        append([]Value(nil), compiler.Constants()...),
		bytecode:         append([]byte(nil), compiler.Bytecode()...),
		compilerGlobals:  append([]string(nil), compiler.Globals.names...),
		compilerZeroArg:  cloneBoolMap(compiler.globalZeroArgFuncs),
		compilerDeclared: cloneBoolMap(compiler.declaredGlobals),
		compilerConst:    cloneBoolMap(compiler.constGlobals),
	}
	if !vm.IsInteractiveMode() && cache != nil {
		cache.Add(key, compiled)
	}
	return compiled, nil
}

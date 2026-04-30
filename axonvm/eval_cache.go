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
	"sync"

	lru "github.com/hashicorp/golang-lru/v2"
)

const (
	evalBytecodeCacheSize = 512
	fnvOffset64           = 1469598103934665603
	fnvPrime64            = 1099511628211
)

// evalCachedProgram stores one immutable compiled Eval payload for cache reuse.
type evalCachedProgram struct {
	keyHash         uint64
	expr            string
	optionCompare   int
	globalNamesHash uint64
	localScopeHash  uint64
	globalCount     int
	constants       []Value
	bytecode        []byte
}

var (
	evalProgramCacheOnce sync.Once
	evalProgramCache     *lru.Cache[uint64, *evalCachedProgram]
)

// getEvalProgramCache lazily initializes and returns the process-wide Eval LRU cache.
func getEvalProgramCache() *lru.Cache[uint64, *evalCachedProgram] {
	evalProgramCacheOnce.Do(func() {
		cache, err := lru.New[uint64, *evalCachedProgram](evalBytecodeCacheSize)
		if err == nil {
			evalProgramCache = cache
		}
	})
	return evalProgramCache
}

// hashByteFNV1a appends one byte to an FNV-1a running hash.
func hashByteFNV1a(hash uint64, b byte) uint64 {
	hash ^= uint64(b)
	hash *= fnvPrime64
	return hash
}

// hashInt64FNV1a appends one int64 value to an FNV-1a running hash.
func hashInt64FNV1a(hash uint64, value int64) uint64 {
	u := uint64(value)
	for i := 0; i < 8; i++ {
		hash = hashByteFNV1a(hash, byte(u))
		u >>= 8
	}
	return hash
}

// hashStringFNV1a appends one string to an FNV-1a running hash without allocations.
func hashStringFNV1a(hash uint64, text string) uint64 {
	for i := 0; i < len(text); i++ {
		hash = hashByteFNV1a(hash, text[i])
	}
	hash = hashByteFNV1a(hash, 0)
	return hash
}

// hashStringSliceFNV1a computes one deterministic hash for a string slice order and values.
func hashStringSliceFNV1a(values []string) uint64 {
	hash := uint64(fnvOffset64)
	for i := 0; i < len(values); i++ {
		hash = hashStringFNV1a(hash, values[i])
	}
	return hash
}

// buildEvalCacheKey hashes expression text with Option Compare mode for fast lookup.
func buildEvalCacheKey(expr string, optionCompare int) uint64 {
	hash := uint64(fnvOffset64)
	hash = hashInt64FNV1a(hash, int64(optionCompare))
	hash = hashStringFNV1a(hash, expr)
	return hash
}

// buildLocalScopeHash hashes the active local procedure symbol shape used by Eval.
func buildLocalScopeHash(localSub Value) uint64 {
	if localSub.Type != VTUserSub || len(localSub.Names) == 0 {
		return 0
	}
	return hashStringSliceFNV1a(localSub.Names)
}

// getOrCompileEvalProgram returns one cached Eval program, compiling and storing on misses.
// In interactive execution modes (CLI, TUI, eval), the cache is bypassed to prevent stalls.
func (vm *VM) getOrCompileEvalProgram(expr string, localSub Value) (*evalCachedProgram, error) {
	if vm == nil {
		return nil, nil
	}

	optionCompare := vm.optionCompare
	key := buildEvalCacheKey(expr, optionCompare)
	localScopeHash := buildLocalScopeHash(localSub)
	globalScopeHash := vm.globalNamesHash
	cache := getEvalProgramCache()

	// In interactive mode, bypass the LRU cache to prevent stalls
	if !vm.IsInteractiveMode() && cache != nil {
		if cached, ok := cache.Get(key); ok && cached != nil {
			if cached.keyHash == key &&
				cached.optionCompare == optionCompare &&
				cached.expr == expr &&
				cached.globalNamesHash == globalScopeHash &&
				cached.localScopeHash == localScopeHash {
				return cached, nil
			}
		}
	}

	compiler := vm.newExecuteLocalCompiler(expr, localSub, true)
	compiler.SetSourceName(vm.sourceName + "_Eval")
	if err := compiler.Compile(); err != nil {
		return nil, err
	}

	compiled := &evalCachedProgram{
		keyHash:         key,
		expr:            expr,
		optionCompare:   optionCompare,
		globalNamesHash: globalScopeHash,
		localScopeHash:  localScopeHash,
		globalCount:     compiler.GlobalsCount(),
		constants:       append([]Value(nil), compiler.Constants()...),
		bytecode:        append([]byte(nil), compiler.Bytecode()...),
	}
	// In interactive mode, don't cache to prevent stalls
	if !vm.IsInteractiveMode() && cache != nil {
		cache.Add(key, compiled)
	}
	return compiled, nil
}

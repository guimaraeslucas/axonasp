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
	"encoding/binary"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/axonvm/asp"
)

type vmProgramPool struct {
	mu          sync.Mutex
	items       []*VM
	maxRetained int
	program     CachedProgram
}

// vmProgramPool manages a pool of VM instances for a specific compiled program.
const vmProgramPoolDefaultRetained = 250

var cachedProgramPools sync.Map
var vmPoolLimitMu sync.RWMutex
var vmPoolLimiter chan struct{}

// SetVMPoolSizeLimit sets the maximum number of concurrently checked-out VMs.
func SetVMPoolSizeLimit(limit int) {
	vmPoolLimitMu.Lock()
	defer vmPoolLimitMu.Unlock()
	if limit <= 0 {
		vmPoolLimiter = nil
		return
	}
	vmPoolLimiter = make(chan struct{}, limit)
}

func acquireVMPoolSlot() chan struct{} {
	vmPoolLimitMu.RLock()
	limiter := vmPoolLimiter
	vmPoolLimitMu.RUnlock()
	if limiter == nil {
		return nil
	}
	limiter <- struct{}{}
	return limiter
}

func releaseVMPoolSlot(slot chan struct{}) {
	if slot == nil {
		return
	}
	select {
	case <-slot:
	default:
	}
}

// AcquireVMFromCachedProgram borrows one VM instance from the interpreter pool.
func AcquireVMFromCachedProgram(program CachedProgram) *VM {
	slot := acquireVMPoolSlot()
	pool := getProgramPool(program)
	vm := pool.get()
	if vm == nil {
		vm = newPooledVMFromCachedProgram(pool, pool.program)
	}
	vm.resetForReuse()
	vm.pooledFrom = pool
	vm.pooledSlot = slot
	return vm
}

// AcquireVMFromCompiler borrows one VM instance for a compiler output payload.
func AcquireVMFromCompiler(compiler *Compiler) *VM {
	if compiler == nil {
		return NewVM(nil, nil, 0)
	}
	return AcquireVMFromCachedProgram(cachedProgramFromCompiler(compiler))
}

// Release cleans request resources and returns a pooled VM to its interpreter pool.
func (vm *VM) Release() {
	if vm == nil {
		return
	}
	vm.CleanupRequestResources()
	pool := vm.pooledFrom
	slot := vm.pooledSlot
	vm.resetForReuse()
	if pool != nil {
		pool.put(vm)
	}
	releaseVMPoolSlot(slot)
}

func getProgramPool(program CachedProgram) *vmProgramPool {
	key := pooledProgramKey(program)
	if existing, ok := cachedProgramPools.Load(key); ok {
		return existing.(*vmProgramPool)
	}

	limit := vmProgramPoolRetainLimit()
	entry := &vmProgramPool{
		items:       make([]*VM, 0, limit),
		maxRetained: limit,
		program:     immutableCachedProgramView(program),
	}

	// Pre-warming: Fill the pool with a few pre-allocated VMs to handle initial bursts.
	// We don't fill the entire limit (250) to avoid excessive memory usage in tests/rare scripts.
	warmLimit := 5
	if limit < warmLimit {
		warmLimit = limit
	}
	for i := 0; i < warmLimit; i++ {
		vm := NewVMFromCachedProgram(entry.program)
		vm.pooledFrom = entry
		entry.items = append(entry.items, vm)
	}

	actual, _ := cachedProgramPools.LoadOrStore(key, entry)
	return actual.(*vmProgramPool)
}

func (p *vmProgramPool) get() *VM {
	if p == nil {
		return nil
	}
	p.mu.Lock()
	defer p.mu.Unlock()
	count := len(p.items)
	if count == 0 {
		return nil
	}
	vm := p.items[count-1]
	p.items[count-1] = nil
	p.items = p.items[:count-1]
	return vm
}

func (p *vmProgramPool) put(vm *VM) {
	if p == nil || vm == nil {
		return
	}
	p.mu.Lock()
	defer p.mu.Unlock()
	if p.maxRetained > 0 && len(p.items) >= p.maxRetained {
		return
	}
	p.items = append(p.items, vm)
}

func vmProgramPoolRetainLimit() int {
	vmPoolLimitMu.RLock()
	limiter := vmPoolLimiter
	vmPoolLimitMu.RUnlock()
	if limiter != nil {
		limit := cap(limiter)
		if limit > 0 {
			return limit
		}
	}
	return vmProgramPoolDefaultRetained
}

func newPooledVMFromCachedProgram(pool *vmProgramPool, program CachedProgram) *VM {
	vm := NewVMFromCachedProgram(program)
	vm.pooledFrom = pool
	return vm
}

func cachedProgramFromCompiler(compiler *Compiler) CachedProgram {
	if compiler == nil {
		return CachedProgram{}
	}
	return buildCachedProgramFromCompiler(compiler)
}

func pooledProgramKey(program CachedProgram) string {
	sourceName := strings.TrimSpace(program.SourceName)
	fingerprint := program.ProgramHash
	if fingerprint == 0 {
		fingerprint = computeProgramHash(
			program.Bytecode,
			program.GlobalCount,
			program.OptionCompare,
			program.OptionExplicit,
			program.SourceName,
			program.Constants,
		)
	}
	if sourceName != "" {
		return sourceName + "#" + programFingerprintHex(fingerprint)
	}
	return "anon:" + programFingerprintHex(fingerprint)
}

func programFingerprintHex(fingerprint uint64) string {
	var raw [8]byte
	binary.BigEndian.PutUint64(raw[:], fingerprint)
	encoded := make([]byte, 16)
	for i := 0; i < len(raw); i++ {
		encoded[i*2] = hexNibble(raw[i] >> 4)
		encoded[i*2+1] = hexNibble(raw[i])
	}
	return string(encoded)
}

func hexNibble(value byte) byte {
	value &= 0x0f
	if value < 10 {
		return '0' + value
	}
	return 'a' + (value - 10)
}

func (vm *VM) captureBaseProgramState() {
	if vm == nil {
		return
	}
	vm.baseBytecode = immutableBytecodeView(vm.bytecode)
	vm.baseConstants = immutableValueView(vm.constants)

	// Perfectly size baseGlobals and reuse capacity if possible.
	if cap(vm.baseGlobals) < len(vm.Globals) {
		vm.baseGlobals = make([]Value, len(vm.Globals))
	} else {
		vm.baseGlobals = vm.baseGlobals[:len(vm.Globals)]
	}
	copy(vm.baseGlobals, vm.Globals)

	// Save base global types for VB6 As Type support.
	if cap(vm.baseGlobalTypes) < len(vm.globalTypes) {
		vm.baseGlobalTypes = make([]ValueType, len(vm.globalTypes))
	} else {
		vm.baseGlobalTypes = vm.baseGlobalTypes[:len(vm.globalTypes)]
	}
	copy(vm.baseGlobalTypes, vm.globalTypes)

	vm.baseOptionCompare = vm.optionCompare
	vm.baseOptionExplicit = vm.optionExplicit
	vm.baseGlobalNames = vm.globalNames
	vm.baseGlobalNamesHash = hashStringSliceFNV1a(vm.baseGlobalNames)
	vm.baseEngineMode = vm.engineMode

	if vm.baseGlobalZeroArgFuncs == nil {
		vm.baseGlobalZeroArgFuncs = make(map[string]bool, len(vm.globalZeroArgFuncs))
	} else {
		clear(vm.baseGlobalZeroArgFuncs)
	}
	for key, value := range vm.globalZeroArgFuncs {
		vm.baseGlobalZeroArgFuncs[key] = value
	}

	vm.baseRuntimeClassVersion = vm.runtimeClassVersion

	if vm.baseDeclared == nil {
		vm.baseDeclared = make(map[string]bool, len(vm.declaredGlobals))
	} else {
		clear(vm.baseDeclared)
	}
	for key, value := range vm.declaredGlobals {
		vm.baseDeclared[key] = value
	}

	if vm.baseConst == nil {
		vm.baseConst = make(map[string]bool, len(vm.constGlobals))
	} else {
		clear(vm.baseConst)
	}
	for key, value := range vm.constGlobals {
		vm.baseConst[key] = value
	}

	vm.baseSourceName = vm.sourceName
	vm.baseSourceMap = vm.sourceMap.Clone()
	vm.bytecode = immutableBytecodeView(vm.baseBytecode)
	vm.constants = immutableValueView(vm.baseConstants)
}

func (vm *VM) resetForReuse() {
	if vm == nil {
		return
	}
	vm.host = nil
	vm.output = nil
	vm.pooledFrom = nil
	vm.pooledSlot = nil
	vm.bytecode = immutableBytecodeView(vm.baseBytecode)
	vm.constants = immutableValueView(vm.baseConstants)
	vm.resetGlobals()
	vm.resetStack()
	vm.ensureReusableMaps()
	vm.resetDynamicMaps()
	vm.optionCompare = vm.baseOptionCompare
	vm.optionExplicit = vm.baseOptionExplicit
	vm.engineMode = vm.baseEngineMode

	// Direct assignment from immutable base state. Subsequent ExecuteGlobal
	// will allocate a new backing array because baseGlobalNames has cap == len.
	vm.globalNames = vm.baseGlobalNames

	vm.globalNamesHash = hashStringSliceFNV1a(vm.globalNames)
	vm.runtimeClassVersion = vm.baseRuntimeClassVersion
	vm.rebuildGlobalNameIndex()

	clear(vm.globalZeroArgFuncs)
	for key, value := range vm.baseGlobalZeroArgFuncs {
		vm.globalZeroArgFuncs[key] = value
	}

	clear(vm.dynamicProgramStarts)
	clear(vm.declaredGlobals)
	for key, value := range vm.baseDeclared {
		vm.declaredGlobals[key] = value
	}
	clear(vm.constGlobals)
	for key, value := range vm.baseConst {
		vm.constGlobals[key] = value
	}
	vm.sourceName = vm.baseSourceName
	vm.sourceMap = vm.baseSourceMap.Clone()
	if vm.errObject != nil {
		vm.errObject.Reset()
	} else {
		vm.errObject = asp.NewASPError()
	}
	vm.errASPCodeRaw = ""
	vm.errASPCodeRawSet = false
	vm.lastLine = 0
	vm.lastColumn = 0
	vm.lastError = nil
	vm.transactionState = 0
	vm.activeClassObjectID = 0
	vm.terminateCursor = -1
	vm.terminatePrepared = false
	vm.suppressTerminate = false
	vm.onResumeNext = false
	vm.executeGlobalResumeGuard = false
	vm.stmtSP = -1
	vm.skipToNextStmt = false
	vm.sp = -1
	vm.ip = 0
	vm.fp = 0
	clear(vm.callStack)
	vm.callStack = vm.callStack[:0]
	clear(vm.withStack)
	vm.withStack = vm.withStack[:0]
	clear(vm.classInstanceOrder)
	vm.classInstanceOrder = vm.classInstanceOrder[:0]
	clear(vm.argBuffer)
	vm.argBuffer = vm.argBuffer[:0]
	clear(vm.indexBuffer)
	vm.indexBuffer = vm.indexBuffer[:0]
	clear(vm.combineBuffer)
	vm.combineBuffer = vm.combineBuffer[:0]
	clear(vm.runtimeClasses)
	clear(vm.runtimeClassItems)
	vm.nextDynamicNativeID = 20000
	vm.nextDynamicClassID = 60000
	vm.comInitialized = false
	vm.comThreadLocked = false
	// Zero-fill icState for fresh execution.
	for i := range vm.icState {
		vm.icState[i] = InlineCacheSlot{}
	}
}

func (vm *VM) resetGlobals() {
	for i := range vm.Globals {
		vm.releaseRecordValue(vm.Globals[i])
	}
	if cap(vm.Globals) < len(vm.baseGlobals) {
		vm.Globals = make([]Value, len(vm.baseGlobals))
	} else {
		vm.Globals = vm.Globals[:len(vm.baseGlobals)]
		clear(vm.Globals)
	}
	copy(vm.Globals, vm.baseGlobals)
	// Restore base global types for VB6 As Type support.
	if cap(vm.globalTypes) < len(vm.baseGlobalTypes) {
		vm.globalTypes = make([]ValueType, len(vm.baseGlobalTypes))
	} else {
		vm.globalTypes = vm.globalTypes[:len(vm.baseGlobalTypes)]
	}
	copy(vm.globalTypes, vm.baseGlobalTypes)
}

func (vm *VM) resetStack() {
	for i := range vm.stack {
		vm.releaseRecordValue(vm.stack[i])
	}
	if len(vm.stack) != StackSize {
		vm.stack = make([]Value, StackSize)
		vm.localTypes = [StackSize]ValueType{} // Reset local types when re-creating stack
		return
	}
	clear(vm.stack)
	clear(vm.localTypes[:]) // Clear local type declarations
}

func (vm *VM) ensureReusableMaps() {
	if vm.runtimeClasses == nil {
		vm.runtimeClasses = make(map[string]RuntimeClassDef)
	}
	if vm.runtimeClassItems == nil {
		vm.runtimeClassItems = make(map[int64]*RuntimeClassInstance)
	}
	if vm.callStack == nil {
		vm.callStack = make([]CallFrame, 0, 16)
	}
	if vm.withStack == nil {
		vm.withStack = make([]Value, 0, 8)
	}
	if vm.classInstanceOrder == nil {
		vm.classInstanceOrder = make([]int64, 0, 16)
	}
	if vm.declaredGlobals == nil {
		vm.declaredGlobals = make(map[string]bool)
	}
	if vm.constGlobals == nil {
		vm.constGlobals = make(map[string]bool)
	}
	if vm.baseDeclared == nil {
		vm.baseDeclared = make(map[string]bool)
	}
	if vm.baseConst == nil {
		vm.baseConst = make(map[string]bool)
	}
	if vm.globalZeroArgFuncs == nil {
		vm.globalZeroArgFuncs = make(map[string]bool)
	}
	if vm.baseGlobalZeroArgFuncs == nil {
		vm.baseGlobalZeroArgFuncs = make(map[string]bool)
	}
	if vm.errObject == nil {
		vm.errObject = asp.NewASPError()
	}
	if vm.globalNames == nil {
		vm.globalNames = make([]string, 0, len(vm.baseGlobalNames))
	}
	if vm.globalNameIndex == nil {
		vm.globalNameIndex = make(map[string]int, len(vm.baseGlobalNames))
	}
	if vm.dynamicProgramStarts == nil {
		vm.dynamicProgramStarts = make(map[uint64]int, 32)
	}
	if vm.argBuffer == nil {
		vm.argBuffer = make([]Value, 0, 16)
	}
	if vm.indexBuffer == nil {
		vm.indexBuffer = make([]Value, 0, 16)
	}
	if vm.combineBuffer == nil {
		vm.combineBuffer = make([]Value, 0, 16)
	}
	vm.ensureDynamicMaps()
}

// rebuildGlobalNameIndex refreshes the lowercased global symbol index for fast runtime lookups.
func (vm *VM) rebuildGlobalNameIndex() {
	if vm == nil {
		return
	}
	if vm.globalNameIndex == nil {
		vm.globalNameIndex = make(map[string]int, len(vm.globalNames))
	} else {
		clear(vm.globalNameIndex)
	}
	vm.globalNamesHash = hashStringSliceFNV1a(vm.globalNames)

	// Optimization: Use pre-computed lowercase names to avoid allocations in the hot path.
	useLower := len(vm.baseGlobalNamesLower) == len(vm.globalNames)

	for idx := 0; idx < len(vm.globalNames); idx++ {
		name := strings.TrimSpace(vm.globalNames[idx])
		if name == "" {
			continue
		}

		var lower string
		if useLower {
			lower = vm.baseGlobalNamesLower[idx]
		} else {
			lower = strings.ToLower(name)
		}
		vm.globalNameIndex[lower] = idx
	}
}

// ensureArgBuffer returns one reusable argument buffer with at least n elements.
func (vm *VM) ensureArgBuffer(n int) []Value {
	if n <= 0 {
		return vm.argBuffer[:0]
	}
	if cap(vm.argBuffer) < n {
		vm.argBuffer = make([]Value, n)
	}
	return vm.argBuffer[:n]
}

// ensureIndexBuffer returns one reusable index buffer with at least n elements.
func (vm *VM) ensureIndexBuffer(n int) []Value {
	if n <= 0 {
		return vm.indexBuffer[:0]
	}
	if cap(vm.indexBuffer) < n {
		vm.indexBuffer = make([]Value, n)
	}
	return vm.indexBuffer[:n]
}

// ensureCombineBuffer returns one reusable combined-args buffer with at least n elements.
func (vm *VM) ensureCombineBuffer(n int) []Value {
	if n <= 0 {
		return vm.combineBuffer[:0]
	}
	if cap(vm.combineBuffer) < n {
		vm.combineBuffer = make([]Value, n)
	}
	return vm.combineBuffer[:n]
}

func (vm *VM) ensureDynamicMaps() {
	if vm.responseCookieItems == nil {
		vm.responseCookieItems = make(map[int64]string)
	}
	if vm.requestCollectionValueItems == nil {
		vm.requestCollectionValueItems = make(map[int64]asp.RequestCollectionValue)
	}
	if vm.aspErrorItems == nil {
		vm.aspErrorItems = make(map[int64]*asp.ASPError)
	}
	if vm.g3mdItems == nil {
		vm.g3mdItems = make(map[int64]*G3MD)
	}
	if vm.g3searchItems == nil {
		vm.g3searchItems = make(map[int64]*G3Search)
	}
	if vm.g3stringBuilderItems == nil {
		vm.g3stringBuilderItems = make(map[int64]*G3StringBuilder)
	}
	if vm.g3testItems == nil {
		vm.g3testItems = make(map[int64]*G3Test)
	}
	if vm.g3cryptoItems == nil {
		vm.g3cryptoItems = make(map[int64]*G3Crypto)
	}
	if vm.g3jsonItems == nil {
		vm.g3jsonItems = make(map[int64]*G3JSON)
	}
	if vm.g3httpItems == nil {
		vm.g3httpItems = make(map[int64]*G3HTTP)
	}
	if vm.g3mailItems == nil {
		vm.g3mailItems = make(map[int64]*G3Mail)
	}
	if vm.g3imageItems == nil {
		vm.g3imageItems = make(map[int64]*G3Image)
	}
	if vm.g3filesItems == nil {
		vm.g3filesItems = make(map[int64]*G3Files)
	}
	if vm.g3templateItems == nil {
		vm.g3templateItems = make(map[int64]*G3Template)
	}
	if vm.g3zipItems == nil {
		vm.g3zipItems = make(map[int64]*G3Zip)
	}
	if vm.g3zlibItems == nil {
		vm.g3zlibItems = make(map[int64]*G3ZLIB)
	}
	if vm.g3tarItems == nil {
		vm.g3tarItems = make(map[int64]*G3TAR)
	}
	if vm.g3zstdItems == nil {
		vm.g3zstdItems = make(map[int64]*G3ZSTD)
	}
	if vm.g3fcItems == nil {
		vm.g3fcItems = make(map[int64]*G3FC)
	}
	if vm.g3axonliveItems == nil {
		vm.g3axonliveItems = make(map[int64]*G3AXONLIVE)
	}
	if vm.g3axonliveProxyItems == nil {
		vm.g3axonliveProxyItems = make(map[int64]*G3ALComponentProxy)
	}
	if vm.g3dbItems == nil {
		vm.g3dbItems = make(map[int64]*G3DB)
	}
	if vm.g3dbResultSetItems == nil {
		vm.g3dbResultSetItems = make(map[int64]*G3DBResultSet)
	}
	if vm.g3dbFieldsItems == nil {
		vm.g3dbFieldsItems = make(map[int64]*G3DBFields)
	}
	if vm.g3dbRowItems == nil {
		vm.g3dbRowItems = make(map[int64]*G3DBRow)
	}
	if vm.g3dbStatementItems == nil {
		vm.g3dbStatementItems = make(map[int64]*G3DBStatement)
	}
	if vm.g3dbTransactionItems == nil {
		vm.g3dbTransactionItems = make(map[int64]*G3DBTransaction)
	}
	if vm.g3dbResultItems == nil {
		vm.g3dbResultItems = make(map[int64]*G3DBResult)
	}
	if vm.wscriptShellItems == nil {
		vm.wscriptShellItems = make(map[int64]*WScriptShell)
	}
	if vm.wscriptExecItems == nil {
		vm.wscriptExecItems = make(map[int64]*WScriptExecObject)
	}
	if vm.wscriptProcessStreamItems == nil {
		vm.wscriptProcessStreamItems = make(map[int64]*ProcessTextStream)
	}
	if vm.adoxCatalogItems == nil {
		vm.adoxCatalogItems = make(map[int64]*ADOXCatalog)
	}
	if vm.adoxTablesItems == nil {
		vm.adoxTablesItems = make(map[int64]*ADOXTables)
	}
	if vm.adoxTableItems == nil {
		vm.adoxTableItems = make(map[int64]*ADOXTableWrapper)
	}
	if vm.mswcAdRotatorItems == nil {
		vm.mswcAdRotatorItems = make(map[int64]*G3AdRotator)
	}
	if vm.mswcBrowserTypeItems == nil {
		vm.mswcBrowserTypeItems = make(map[int64]*G3BrowserType)
	}
	if vm.mswcNextLinkItems == nil {
		vm.mswcNextLinkItems = make(map[int64]*G3NextLink)
	}
	if vm.mswcContentRotatorItems == nil {
		vm.mswcContentRotatorItems = make(map[int64]*G3ContentRotator)
	}
	if vm.mswcCountersItems == nil {
		vm.mswcCountersItems = make(map[int64]*G3Counters)
	}
	if vm.mswcPageCounterItems == nil {
		vm.mswcPageCounterItems = make(map[int64]*G3PageCounter)
	}
	if vm.mswcToolsItems == nil {
		vm.mswcToolsItems = make(map[int64]*G3Tools)
	}
	if vm.mswcMyInfoItems == nil {
		vm.mswcMyInfoItems = make(map[int64]*G3MyInfo)
	}
	if vm.mswcPermissionCheckerItems == nil {
		vm.mswcPermissionCheckerItems = make(map[int64]*G3PermissionChecker)
	}
	if vm.msxmlServerItems == nil {
		vm.msxmlServerItems = make(map[int64]*MsXML2ServerXMLHTTP)
	}
	if vm.msxmlDOMItems == nil {
		vm.msxmlDOMItems = make(map[int64]*MsXML2DOMDocument)
	}
	if vm.msxmlNodeListItems == nil {
		vm.msxmlNodeListItems = make(map[int64]*XMLNodeList)
	}
	if vm.msxmlParseErrorItems == nil {
		vm.msxmlParseErrorItems = make(map[int64]*ParseError)
	}
	if vm.msxmlElementItems == nil {
		vm.msxmlElementItems = make(map[int64]*XMLElement)
	}
	if vm.pdfItems == nil {
		vm.pdfItems = make(map[int64]*G3PDF)
	}
	if vm.fileUploaderItems == nil {
		vm.fileUploaderItems = make(map[int64]*G3FileUploader)
	}
	if vm.axonItems == nil {
		vm.axonItems = make(map[int64]*AxonLibrary)
	}
	if vm.fsoItems == nil {
		vm.fsoItems = make(map[int64]*fsoNativeObject)
	}
	if vm.adodbStreamItems == nil {
		vm.adodbStreamItems = make(map[int64]*adodbStreamNativeObject)
	}
	if vm.adodbConnectionItems == nil {
		vm.adodbConnectionItems = make(map[int64]*adodbConnection)
	}
	if vm.adodbRecordsetItems == nil {
		vm.adodbRecordsetItems = make(map[int64]*adodbRecordset)
	}
	if vm.adodbCommandItems == nil {
		vm.adodbCommandItems = make(map[int64]*adodbCommand)
	}
	if vm.adodbParameterItems == nil {
		vm.adodbParameterItems = make(map[int64]*adodbParameter)
	}
	if vm.adodbErrorsCollectionItems == nil {
		vm.adodbErrorsCollectionItems = make(map[int64]*adodbConnection)
	}
	if vm.adodbErrorItems == nil {
		vm.adodbErrorItems = make(map[int64]*adodbError)
	}
	if vm.adodbFieldsCollectionItems == nil {
		vm.adodbFieldsCollectionItems = make(map[int64]*adodbRecordset)
	}
	if vm.adodbParametersCollectionItems == nil {
		vm.adodbParametersCollectionItems = make(map[int64]*adodbCommand)
	}
	if vm.adodbFieldItems == nil {
		vm.adodbFieldItems = make(map[int64]*adodbFieldProxy)
	}
	if vm.regExpItems == nil {
		vm.regExpItems = make(map[int64]*regExpNativeObject)
	}
	if vm.jsRegExpItems == nil {
		vm.jsRegExpItems = make(map[int64]*jsRegExpObject)
	}
	if vm.regExpMatchesCollectionItems == nil {
		vm.regExpMatchesCollectionItems = make(map[int64]*regExpMatchesCollection)
	}
	if vm.regExpMatchItems == nil {
		vm.regExpMatchItems = make(map[int64]*regExpMatch)
	}
	if vm.regExpSubMatchesItems == nil {
		vm.regExpSubMatchesItems = make(map[int64]*regExpSubMatches)
	}
	if vm.regExpSubMatchValueItems == nil {
		vm.regExpSubMatchValueItems = make(map[int64]*regExpSubMatchValue)
	}
	if vm.dictionaryItems == nil {
		vm.dictionaryItems = make(map[int64]*scriptingDictionary)
	}
	if vm.nativeObjectProxies == nil {
		vm.nativeObjectProxies = make(map[int64]nativeObjectProxy)
	}
	if vm.jsObjectItems == nil {
		vm.jsObjectItems = make(map[int64]map[string]Value)
	}
	if vm.jsObjectKeyOrder == nil {
		vm.jsObjectKeyOrder = make(map[int64][]string)
	}
	if vm.jsObjectSlots == nil {
		vm.jsObjectSlots = make(map[int64][]Value)
	}
	if vm.jsObjectSlotIndex == nil {
		vm.jsObjectSlotIndex = make(map[int64]map[string]uint16)
	}
	if vm.jsObjectShape == nil {
		vm.jsObjectShape = make(map[int64]uint32)
	}
	if vm.jsShapeSlots == nil {
		vm.jsShapeSlots = make(map[uint32][]string)
	}
	if vm.jsShapeBySignature == nil {
		vm.jsShapeBySignature = make(map[string]uint32)
	}
	if vm.jsNextShapeID == 0 {
		vm.jsNextShapeID = 1
	}
	if vm.jsObjectStateItems == nil {
		vm.jsObjectStateItems = make(map[int64]jsObjectState)
	}
	if vm.jsPropertyItems == nil {
		vm.jsPropertyItems = make(map[int64]map[string]jsPropertyDescriptor)
	}
	if vm.jsFunctionItems == nil {
		vm.jsFunctionItems = make(map[int64]*jsFunctionObject)
	}
	if vm.jsForInItems == nil {
		vm.jsForInItems = make(map[int]*jsForInEnumerator)
	}
	if vm.jsForOfItems == nil {
		vm.jsForOfItems = make(map[int]*jsForOfEnumerator)
	}
	if vm.jsEnvItems == nil {
		vm.jsEnvItems = make(map[int64]*jsEnvFrame)
	}
	if vm.jsArgumentsItems == nil {
		vm.jsArgumentsItems = make(map[int64]*jsArgumentsBinding)
	}
	if vm.jsSetItems == nil {
		vm.jsSetItems = make(map[int64]map[string]Value)
	}
	if vm.jsMapItems == nil {
		vm.jsMapItems = make(map[int64]map[string]Value)
	}
	if vm.jsWeakRefItems == nil {
		vm.jsWeakRefItems = make(map[int64]*jsWeakRef)
	}
	if vm.jsFinalizationRegistryItems == nil {
		vm.jsFinalizationRegistryItems = make(map[int64]*jsFinalizationRegistry)
	}
	if vm.jsArrayIterators == nil {
		vm.jsArrayIterators = make(map[int64]*jsArrayIterator)
	}
	if vm.jsStringIterators == nil {
		vm.jsStringIterators = make(map[int64]*jsStringIterator)
	}
	if vm.jsRegExpStringIterators == nil {
		vm.jsRegExpStringIterators = make(map[int64]*jsRegExpStringIterator)
	}
	if vm.jsArrayBuffers == nil {
		vm.jsArrayBuffers = make(map[int64][]byte)
	}
	if vm.jsSharedArrayBuffers == nil {
		vm.jsSharedArrayBuffers = make(map[int64][]byte)
	}
	if vm.jsModuleInstances == nil {
		vm.jsModuleInstances = make(map[string]*jsEnvFrame)
	}
	if vm.jsModuleLoading == nil {
		vm.jsModuleLoading = make(map[string]struct{})
	}
	if vm.jsStreamHookItems == nil {
		vm.jsStreamHookItems = make(map[int64]*jsNodeStreamHookResource)
	}
	if vm.jsIntlDateTimeFormatItems == nil {
		vm.jsIntlDateTimeFormatItems = make(map[int64]*jsIntlDateTimeFormatObject)
	}
	if vm.jsIntlNumberFormatItems == nil {
		vm.jsIntlNumberFormatItems = make(map[int64]*jsIntlNumberFormatObject)
	}
	if vm.jsIntlCollatorItems == nil {
		vm.jsIntlCollatorItems = make(map[int64]*jsIntlCollatorObject)
	}
	if vm.jsIntlPluralRulesItems == nil {
		vm.jsIntlPluralRulesItems = make(map[int64]*jsIntlPluralRulesObject)
	}
	if vm.jsIntlRelativeTimeFormatItems == nil {
		vm.jsIntlRelativeTimeFormatItems = make(map[int64]*jsIntlRelativeTimeFormatObject)
	}
	if vm.jsPromiseItems == nil {
		vm.jsPromiseItems = make(map[int64]*jsPromiseObject)
	}
	if vm.jsGeneratorItems == nil {
		vm.jsGeneratorItems = make(map[int64]*jsGeneratorObject)
	}
	if vm.jsProxyItems == nil {
		vm.jsProxyItems = make(map[int64]*jsProxyObject)
	}
	if vm.jsSymbolStateItems == nil {
		vm.jsSymbolStateItems = make(map[int64]jsObjectState)
	}
	if vm.jsRegisteredSymbolIDs == nil {
		vm.jsRegisteredSymbolIDs = make(map[int64]struct{})
	}
	if vm.jsAsyncFSReadResults == nil {
		vm.jsAsyncFSReadResults = make(chan jsAsyncFSReadResult, jsAsyncFSReadResultQueueSize)
	}
	if vm.jsTimerItems == nil {
		vm.jsTimerItems = make(map[int64]*jsTimerItem)
	}
	if vm.jsTimerResultQueue == nil {
		vm.jsTimerResultQueue = make(chan jsTimerFiredResult, jsTimerResultQueueSize)
	}
	if vm.jsImmediateQueue == nil {
		vm.jsImmediateQueue = make([]jsImmediateItem, 0, 8)
	}
	if vm.jsNextTickQueue == nil {
		vm.jsNextTickQueue = make([]jsNextTickItem, 0, 8)
	}
	if vm.jsMicrotaskQueue == nil {
		vm.jsMicrotaskQueue = make([]func(), 0, 8)
	}
	if vm.jsSymbolGlobalRegistry == nil {
		vm.jsSymbolGlobalRegistry = make(map[string]Value)
	}
	if vm.consoleTimerItems == nil {
		vm.consoleTimerItems = make(map[string]time.Time)
	}
}

func (vm *VM) resetDynamicMaps() {
	clear(vm.responseCookieItems)
	clear(vm.requestCollectionValueItems)
	clear(vm.aspErrorItems)
	clear(vm.g3mdItems)
	clear(vm.g3searchItems)
	clear(vm.g3stringBuilderItems)
	clear(vm.g3testItems)
	clear(vm.g3cryptoItems)
	clear(vm.g3jsonItems)
	clear(vm.g3httpItems)
	clear(vm.g3mailItems)
	clear(vm.g3imageItems)
	clear(vm.g3filesItems)
	clear(vm.g3templateItems)
	clear(vm.g3zipItems)
	clear(vm.g3zlibItems)
	clear(vm.g3tarItems)
	clear(vm.g3zstdItems)
	clear(vm.g3fcItems)
	clear(vm.g3axonliveItems)
	clear(vm.g3axonliveProxyItems)
	clear(vm.g3dbItems)
	clear(vm.g3dbResultSetItems)
	clear(vm.g3dbFieldsItems)
	clear(vm.g3dbRowItems)
	clear(vm.g3dbStatementItems)
	clear(vm.g3dbTransactionItems)
	clear(vm.g3dbResultItems)
	clear(vm.wscriptShellItems)
	clear(vm.wscriptExecItems)
	clear(vm.wscriptProcessStreamItems)
	clear(vm.adoxCatalogItems)
	clear(vm.adoxTablesItems)
	clear(vm.adoxTableItems)
	clear(vm.mswcAdRotatorItems)
	clear(vm.mswcBrowserTypeItems)
	clear(vm.mswcNextLinkItems)
	clear(vm.mswcContentRotatorItems)
	clear(vm.mswcCountersItems)
	clear(vm.mswcPageCounterItems)
	clear(vm.mswcToolsItems)
	clear(vm.mswcMyInfoItems)
	clear(vm.mswcPermissionCheckerItems)
	clear(vm.msxmlServerItems)
	clear(vm.msxmlDOMItems)
	clear(vm.msxmlNodeListItems)
	clear(vm.msxmlParseErrorItems)
	clear(vm.msxmlElementItems)
	clear(vm.pdfItems)
	clear(vm.fileUploaderItems)
	clear(vm.axonItems)
	clear(vm.fsoItems)
	clear(vm.adodbStreamItems)
	clear(vm.adodbConnectionItems)
	clear(vm.adodbRecordsetItems)
	clear(vm.adodbCommandItems)
	clear(vm.adodbParameterItems)
	clear(vm.adodbErrorsCollectionItems)
	clear(vm.adodbErrorItems)
	clear(vm.adodbFieldsCollectionItems)
	clear(vm.adodbParametersCollectionItems)
	clear(vm.adodbFieldItems)
	clear(vm.regExpItems)
	clear(vm.jsRegExpItems)
	clear(vm.regExpMatchesCollectionItems)
	clear(vm.regExpMatchItems)
	clear(vm.regExpSubMatchesItems)
	clear(vm.regExpSubMatchValueItems)
	clear(vm.dictionaryItems)
	clear(vm.runtimeClassItems)
	clear(vm.nativeObjectProxies)
	clear(vm.jsObjectItems)
	clear(vm.jsObjectKeyOrder)
	clear(vm.jsObjectSlots)
	clear(vm.jsObjectSlotIndex)
	clear(vm.jsObjectShape)
	clear(vm.jsShapeSlots)
	clear(vm.jsShapeBySignature)
	vm.jsNextShapeID = 1
	clear(vm.jsObjectStateItems)
	clear(vm.jsSymbolStateItems)
	clear(vm.jsPropertyItems)
	clear(vm.jsFunctionItems)
	clear(vm.jsForInItems)
	clear(vm.jsForOfItems)
	clear(vm.jsEnvItems)
	clear(vm.jsArgumentsItems)
	clear(vm.jsSetItems)
	clear(vm.jsMapItems)
	clear(vm.jsWeakRefItems)
	clear(vm.jsFinalizationRegistryItems)
	clear(vm.jsArrayIterators)
	clear(vm.jsStringIterators)
	clear(vm.jsRegExpStringIterators)
	clear(vm.jsArrayBuffers)
	clear(vm.jsSharedArrayBuffers)
	clear(vm.jsModuleInstances)
	clear(vm.jsModuleLoading)
	clear(vm.jsIntlDateTimeFormatItems)
	clear(vm.jsIntlNumberFormatItems)
	clear(vm.jsIntlCollatorItems)
	clear(vm.jsIntlPluralRulesItems)
	clear(vm.jsIntlRelativeTimeFormatItems)
	clear(vm.jsPromiseItems)
	clear(vm.jsGeneratorItems)
	clear(vm.jsProxyItems)
	clear(vm.jsStreamHookItems)
	// Stop all active timers and drain timer-result channel before reset.
	vm.jsStopAllTimers()
	vm.jsImmediateQueue = vm.jsImmediateQueue[:0]
	vm.jsNextTickQueue = vm.jsNextTickQueue[:0]
	vm.jsPumpingNodeTasks = false
	for {
		select {
		case <-vm.jsAsyncFSReadResults:
		default:
			goto drainedAsyncFS
		}
	}
drainedAsyncFS:
	clear(vm.jsMicrotaskQueue)
	vm.jsMicrotaskQueue = vm.jsMicrotaskQueue[:0]
	clear(vm.jsSymbolGlobalRegistry)
	clear(vm.jsRegisteredSymbolIDs)
	clear(vm.jsCallStack)
	vm.jsCallStack = vm.jsCallStack[:0]
	clear(vm.jsTryStack)
	vm.jsTryStack = vm.jsTryStack[:0]
	clear(vm.jsErrStack)
	vm.jsErrStack = vm.jsErrStack[:0]
	clear(vm.consoleTimerItems)
	vm.jsActiveEnvID = 0
	vm.jsRootEnvID = 0
	vm.jsThisValue = Value{Type: VTJSUndefined}
	vm.jsStringWorkBytes = 0
}

func immutableBytecodeView(bytecode []byte) []byte {
	if len(bytecode) == 0 {
		return nil
	}
	return bytecode[:len(bytecode):len(bytecode)]
}

func immutableValueView(values []Value) []Value {
	if len(values) == 0 {
		return nil
	}
	return values[:len(values):len(values)]
}

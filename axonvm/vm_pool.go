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
	"encoding/binary"
	"strings"
	"sync"

	"g3pix.com.br/axonasp/axonvm/asp"
)

type vmProgramPool struct {
	pool    sync.Pool
	program CachedProgram
}

var cachedProgramPools sync.Map
var vmPoolLimitMu sync.RWMutex
var vmPoolLimiter chan struct{}

// SetVMPoolSizeLimit sets the maximum number of concurrently checked-out VMs.
// A limit of 0 means unlimited.
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
	vm, _ := pool.pool.Get().(*VM)
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
		pool.pool.Put(vm)
	}
	releaseVMPoolSlot(slot)
}

func getProgramPool(program CachedProgram) *vmProgramPool {
	key := pooledProgramKey(program)
	if existing, ok := cachedProgramPools.Load(key); ok {
		return existing.(*vmProgramPool)
	}

	entry := &vmProgramPool{program: immutableCachedProgramView(program)}
	entry.pool.New = func() any {
		return newPooledVMFromCachedProgram(entry, entry.program)
	}
	actual, _ := cachedProgramPools.LoadOrStore(key, entry)
	return actual.(*vmProgramPool)
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
	vm.baseGlobals = cloneValueSlice(vm.Globals)
	vm.baseOptionCompare = vm.optionCompare
	vm.baseOptionExplicit = vm.optionExplicit
	vm.baseGlobalNames = cloneStringSlice(vm.globalNames)
	vm.baseGlobalNamesHash = vm.globalNamesHash
	clear(vm.baseGlobalZeroArgFuncs)
	for key, value := range vm.globalZeroArgFuncs {
		vm.baseGlobalZeroArgFuncs[key] = value
	}
	vm.baseRuntimeClassVersion = vm.runtimeClassVersion
	clear(vm.baseDeclared)
	for key, value := range vm.declaredGlobals {
		vm.baseDeclared[key] = value
	}
	clear(vm.baseConst)
	for key, value := range vm.constGlobals {
		vm.baseConst[key] = value
	}
	vm.baseSourceName = vm.sourceName
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
	vm.globalNames = append(vm.globalNames[:0], vm.baseGlobalNames...)
	vm.globalNamesHash = vm.baseGlobalNamesHash
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
	vm.errObject = asp.NewASPError()
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
	vm.argBuffer = vm.argBuffer[:0]
	vm.indexBuffer = vm.indexBuffer[:0]
	vm.combineBuffer = vm.combineBuffer[:0]
	clear(vm.runtimeClasses)
	clear(vm.runtimeClassItems)
	vm.nextDynamicNativeID = 20000
	vm.nextDynamicClassID = 60000
	vm.comInitialized = false
	vm.comThreadLocked = false
}

func (vm *VM) resetGlobals() {
	if cap(vm.Globals) < len(vm.baseGlobals) {
		vm.Globals = make([]Value, len(vm.baseGlobals))
	} else {
		vm.Globals = vm.Globals[:len(vm.baseGlobals)]
		clear(vm.Globals)
	}
	copy(vm.Globals, vm.baseGlobals)
}

func (vm *VM) resetStack() {
	if len(vm.stack) != StackSize {
		vm.stack = make([]Value, StackSize)
		return
	}
	clear(vm.stack)
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
	for idx := 0; idx < len(vm.globalNames); idx++ {
		name := strings.TrimSpace(vm.globalNames[idx])
		if name == "" {
			continue
		}
		vm.globalNameIndex[strings.ToLower(name)] = idx
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
	if vm.aspErrorItems == nil {
		vm.aspErrorItems = make(map[int64]*asp.ASPError)
	}
	if vm.g3mdItems == nil {
		vm.g3mdItems = make(map[int64]*G3MD)
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
	if vm.g3fcItems == nil {
		vm.g3fcItems = make(map[int64]*G3FC)
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
}

func (vm *VM) resetDynamicMaps() {
	clear(vm.responseCookieItems)
	clear(vm.aspErrorItems)
	clear(vm.g3mdItems)
	clear(vm.g3testItems)
	clear(vm.g3cryptoItems)
	clear(vm.g3jsonItems)
	clear(vm.g3httpItems)
	clear(vm.g3mailItems)
	clear(vm.g3imageItems)
	clear(vm.g3filesItems)
	clear(vm.g3templateItems)
	clear(vm.g3zipItems)
	clear(vm.g3fcItems)
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
	clear(vm.regExpMatchesCollectionItems)
	clear(vm.regExpMatchItems)
	clear(vm.regExpSubMatchesItems)
	clear(vm.regExpSubMatchValueItems)
	clear(vm.dictionaryItems)
	clear(vm.nativeObjectProxies)
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

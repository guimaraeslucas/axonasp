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
	"fmt"
	"io"
	"sort"
	"strings"
	"time"

	"g3pix.com.br/axonasp/axonvm/asp"
	"g3pix.com.br/axonasp/vbscript"
)

const StackSize = 4096
const jsBackJumpLimit = 50000

const staticObjectProgIDPrefix = "__AXON_STATIC_OBJECT_PROGID__:"

const (
	nativeObjectResponse int64 = iota
	nativeObjectRequest
	nativeObjectServer
	nativeObjectSession
	nativeObjectApplication
	nativeObjectObjectContext
	nativeObjectErr
	nativeObjectConsole // global console object (slot 7)
)

// objectContextCommitHandlerIdx and objectContextAbortHandlerIdx are pre-reserved global
// indices for OnTransactionCommit and OnTransactionAbort event handler subs.
// They follow immediately after the eight intrinsic object slots (indices 0–7).
const (
	objectContextCommitHandlerIdx = 8
	objectContextAbortHandlerIdx  = 9
)

const (
	nativeResponseCookies int64 = 1201 + iota
	nativeResponseCookiesDomainMethod
	nativeResponseCookiesPathMethod
	nativeResponseCookiesExpiresMethod
	nativeResponseCookiesSecureMethod
	nativeResponseCookiesHttpOnlyMethod
)

const (
	nativeRequestQueryString int64 = 1001 + iota
	nativeRequestForm
	nativeRequestCookies
	nativeRequestServerVariables
	nativeRequestClientCertificate
	nativeRequestBinaryReadMethod
)

const (
	nativeRequestQueryStringKeyMethod int64 = 1101 + iota
	nativeRequestFormKeyMethod
	nativeRequestCookiesKeyMethod
	nativeRequestServerVariablesKeyMethod
	nativeRequestClientCertificateKeyMethod
)

const (
	nativeObjectSessionContents int64 = 1301 + iota
	nativeObjectSessionStaticObjects
	nativeObjectApplicationContents
	nativeObjectApplicationStaticObjects
)

// VMError represents a VBScript runtime error.
type VMError struct {
	Code           vbscript.VBSyntaxErrorCode
	Line           int
	Column         int
	File           string
	Msg            string
	ASPCode        int
	ASPDescription string
	Category       string
	Description    string
	Number         int
	Source         string
	HelpFile       string
	HelpContext    int
}

func (e *VMError) Error() string {
	if e == nil {
		return ""
	}

	var builder strings.Builder
	builder.Grow(256)
	builder.WriteString(e.Source)
	if e.Code != 0 {
		builder.WriteString(" '")
		builder.WriteString(vbscript.HRESULTHexFromVBScriptCode(e.Code))
		builder.WriteString("'")
	}

	if e.Description != "" {
		builder.WriteString("\n")
		builder.WriteString(e.Description)
	}

	builder.WriteString("\nCategory: ")
	builder.WriteString(e.Category)
	builder.WriteString("\nColumn: ")
	builder.WriteString(fmt.Sprintf("%d", e.Column))
	builder.WriteString("\nDescription: ")
	builder.WriteString(e.Description)
	builder.WriteString("\nFile: ")
	builder.WriteString(e.File)
	builder.WriteString("\nLine: ")
	builder.WriteString(fmt.Sprintf("%d", e.Line))
	builder.WriteString("\nNumber: ")
	builder.WriteString(fmt.Sprintf("%d", e.Number))
	builder.WriteString("\nSource: ")
	builder.WriteString(e.Source)

	return builder.String()
}

// WithFile attaches the source file path to the runtime error.
func (e *VMError) WithFile(file string) *VMError {
	if e == nil {
		return nil
	}

	e.File = strings.TrimSpace(file)
	return e
}

// ToASPError converts the runtime error into the ASPError object model.
func (e *VMError) ToASPError() *asp.ASPError {
	if e == nil {
		return asp.NewASPError()
	}

	return (&asp.ASPError{
		ASPCode:        e.ASPCode,
		ASPDescription: e.ASPDescription,
		Number:         e.Number,
		Source:         e.Source,
		Description:    e.Description,
		HelpFile:       e.HelpFile,
		HelpContext:    e.HelpContext,
		File:           e.File,
		Line:           e.Line,
		Column:         e.Column,
		Category:       e.Category,
	}).Normalize()
}

// CallFrame stores caller state for user-defined Sub/Function invocations.
// byRefWriteback records one ByRef parameter write-back entry for post-call value propagation.
type byRefWriteback struct {
	calleeParamIdx int  // zero-based parameter index in the callee's local frame
	isGlobal       bool // true = write to Globals[callerIdx]; false = write to stack[callerFP + callerIdx]
	callerIdx      int  // global slot index, or local offset relative to the caller's frame pointer
	isClassMember  bool // true = write to class member slot on callerBoundObj
	callerBoundObj int64
	callerMember   string
}

// CallFrame stores caller state for user-defined Sub/Function invocations.
type CallFrame struct {
	callee              Value
	returnIP            int
	oldFP               int
	oldSP               int
	boundObj            int64
	discard             bool
	byRefs              []byRefWriteback // ByRef parameter write-backs, nil if none
	savedOnResumeNext   bool             // On Error Resume Next state before entering this call frame; restored on OpRet.
	savedSkipToNextStmt bool             // Per-statement Resume Next skip state before entering this call frame; restored on OpRet.
	savedStmtSP         int              // Statement-start SP before entering this call frame; restored on OpRet.
}

// RuntimeClassMethodDef stores one compiled class method runtime entry.
type RuntimeClassMethodDef struct {
	Target   Value
	IsPublic bool
}

// RuntimeClassFieldDef stores runtime metadata for one direct class field.
type RuntimeClassFieldDef struct {
	Name     string
	IsPublic bool
	Dims     []int // nil for plain variables; non-nil upper bounds for fixed-size arrays
}

// RuntimeClassPropertyDef stores runtime metadata for one class property.
type RuntimeClassPropertyDef struct {
	Name          string
	IsPublic      bool
	GetTarget     Value
	GetParamCount int
	HasGet        bool
	LetTarget     Value
	LetParamCount int
	HasLet        bool
	SetTarget     Value
	SetParamCount int
	HasSet        bool
}

// RuntimeClassDef stores one class declaration name available for New object allocation.
type RuntimeClassDef struct {
	Name       string
	Fields     map[string]RuntimeClassFieldDef
	Methods    map[string]RuntimeClassMethodDef
	Properties map[string]RuntimeClassPropertyDef
}

// RuntimeClassInstance stores one allocated class instance state.
type RuntimeClassInstance struct {
	ClassName  string
	Members    map[string]Value
	refCount   int  // Reference count for deterministic termination
	terminated bool // True if Class_Terminate has already been called
}

// nativeObjectProxy represents a property access that requires parameters (e.g. dict.Key(idx)).
type nativeObjectProxy struct {
	ParentID int64
	Member   string
	CallArgs []Value
}

// VM is the stack-based virtual machine for executing VBScript bytecode.
type VM struct {
	bytecode  []byte
	constants []Value
	Globals   []Value
	output    io.Writer
	host      ASPHostEnvironment

	nextDynamicNativeID            int64
	nextDynamicClassID             int64
	responseCookieItems            map[int64]string
	aspErrorItems                  map[int64]*asp.ASPError
	g3mdItems                      map[int64]*G3MD
	g3testItems                    map[int64]*G3Test
	g3cryptoItems                  map[int64]*G3Crypto
	g3jsonItems                    map[int64]*G3JSON
	g3httpItems                    map[int64]*G3HTTP
	g3mailItems                    map[int64]*G3Mail
	g3imageItems                   map[int64]*G3Image
	g3filesItems                   map[int64]*G3Files
	g3templateItems                map[int64]*G3Template
	g3zipItems                     map[int64]*G3Zip
	g3zlibItems                    map[int64]*G3ZLIB
	g3tarItems                     map[int64]*G3TAR
	g3zstdItems                    map[int64]*G3ZSTD
	g3fcItems                      map[int64]*G3FC
	g3dbItems                      map[int64]*G3DB
	g3dbResultSetItems             map[int64]*G3DBResultSet
	g3dbFieldsItems                map[int64]*G3DBFields
	g3dbRowItems                   map[int64]*G3DBRow
	g3dbStatementItems             map[int64]*G3DBStatement
	g3dbTransactionItems           map[int64]*G3DBTransaction
	g3dbResultItems                map[int64]*G3DBResult
	wscriptShellItems              map[int64]*WScriptShell
	wscriptExecItems               map[int64]*WScriptExecObject
	wscriptProcessStreamItems      map[int64]*ProcessTextStream
	adoxCatalogItems               map[int64]*ADOXCatalog
	adoxTablesItems                map[int64]*ADOXTables
	adoxTableItems                 map[int64]*ADOXTableWrapper
	mswcAdRotatorItems             map[int64]*G3AdRotator
	mswcBrowserTypeItems           map[int64]*G3BrowserType
	mswcNextLinkItems              map[int64]*G3NextLink
	mswcContentRotatorItems        map[int64]*G3ContentRotator
	mswcCountersItems              map[int64]*G3Counters
	mswcPageCounterItems           map[int64]*G3PageCounter
	mswcToolsItems                 map[int64]*G3Tools
	mswcMyInfoItems                map[int64]*G3MyInfo
	mswcPermissionCheckerItems     map[int64]*G3PermissionChecker
	msxmlServerItems               map[int64]*MsXML2ServerXMLHTTP
	msxmlDOMItems                  map[int64]*MsXML2DOMDocument
	msxmlNodeListItems             map[int64]*XMLNodeList
	msxmlParseErrorItems           map[int64]*ParseError
	msxmlElementItems              map[int64]*XMLElement
	pdfItems                       map[int64]*G3PDF
	fileUploaderItems              map[int64]*G3FileUploader
	axonItems                      map[int64]*AxonLibrary
	fsoItems                       map[int64]*fsoNativeObject
	adodbStreamItems               map[int64]*adodbStreamNativeObject
	adodbConnectionItems           map[int64]*adodbConnection
	adodbRecordsetItems            map[int64]*adodbRecordset
	adodbCommandItems              map[int64]*adodbCommand
	adodbParameterItems            map[int64]*adodbParameter
	adodbErrorsCollectionItems     map[int64]*adodbConnection
	adodbErrorItems                map[int64]*adodbError
	adodbFieldsCollectionItems     map[int64]*adodbRecordset
	adodbParametersCollectionItems map[int64]*adodbCommand
	adodbFieldItems                map[int64]*adodbFieldProxy
	regExpItems                    map[int64]*regExpNativeObject
	regExpMatchesCollectionItems   map[int64]*regExpMatchesCollection
	regExpMatchItems               map[int64]*regExpMatch
	regExpSubMatchesItems          map[int64]*regExpSubMatches
	regExpSubMatchValueItems       map[int64]*regExpSubMatchValue
	dictionaryItems                map[int64]*scriptingDictionary
	nativeObjectProxies            map[int64]nativeObjectProxy
	jsObjectItems                  map[int64]map[string]Value
	jsObjectStateItems             map[int64]jsObjectState
	jsPropertyItems                map[int64]map[string]jsPropertyDescriptor
	jsFunctionItems                map[int64]*jsFunctionObject
	jsForInItems                   map[int]*jsForInEnumerator
	jsEnvItems                     map[int64]*jsEnvFrame
	errObject                      *asp.ASPError
	errASPCodeRaw                  string
	errASPCodeRawSet               bool

	runtimeClasses      map[string]RuntimeClassDef
	runtimeClassItems   map[int64]*RuntimeClassInstance
	classInstanceOrder  []int64
	activeClassObjectID int64
	terminateCursor     int
	terminatePrepared   bool
	suppressTerminate   bool

	stack       []Value
	sp          int
	ip          int
	fp          int
	callStack   []CallFrame
	jsCallStack []jsCallFrame
	// withStack holds the subject object for each active With...End With block.
	// OpWithEnter appends, OpWithLeave shrinks, OpWithLoad peeks at the top.
	withStack     []Value
	jsTryStack    []int
	jsErrStack    []Value
	jsActiveEnvID int64
	jsThisValue   Value

	onResumeNext bool
	// executeGlobalResumeGuard preserves caller Resume Next semantics for top-level
	// dynamic ExecuteGlobal code, even if nested calls toggle On Error state.
	executeGlobalResumeGuard bool
	// stmtSP is the stack pointer saved at the start of each statement (OpLine).
	// Used by skipToNextStmt to restore the stack after a Resume Next mid-statement error.
	stmtSP int
	// skipToNextStmt is set when Resume Next absorbs an error mid-expression. The VM
	// advances past subsequent opcodes until the next statement boundary (OpLine) or
	// takes any pending OpJumpIfFalse unconditionally to correctly skip compound blocks.
	skipToNextStmt bool
	lastLine       int
	lastColumn     int
	lastError      error

	// transactionState tracks ObjectContext transaction disposition.
	// 0 = no transaction active, 1 = SetComplete called, 2 = SetAbort called.
	transactionState int

	// Runtime Options
	optionCompare        int             // 0: Binary, 1: Text
	optionExplicit       bool            // Mirrors compiler Option Explicit for dynamic execution APIs.
	globalNames          []string        // Global symbol names aligned with Globals slot order.
	globalNamesHash      uint64          // Fingerprint of globalNames for fast dynamic-cache scope validation.
	globalZeroArgFuncs   map[string]bool // Known zero-arg global Functions for dynamic autocall compatibility.
	runtimeClassVersion  uint64          // Monotonic version for runtime class metadata invalidation.
	declaredGlobals      map[string]bool // Global Dim/Const declaration state for dynamic compilation.
	constGlobals         map[string]bool // Global Const protection state for dynamic compilation.
	sourceName           string          // Source file path used for dynamic execution error reporting.
	dynamicProgramStarts map[uint64]int  // Per-VM start offsets for already-appended cached dynamic fragments.
	jsStringWorkBytes    int64           // Per-run cumulative bytes produced by JScript string operations.

	baseBytecode            []byte
	baseConstants           []Value
	baseGlobals             []Value
	baseOptionCompare       int
	baseOptionExplicit      bool
	baseGlobalNames         []string
	baseGlobalNamesHash     uint64
	baseGlobalZeroArgFuncs  map[string]bool
	baseRuntimeClassVersion uint64
	globalNameIndex         map[string]int
	baseDeclared            map[string]bool
	baseConst               map[string]bool
	baseSourceName          string

	argBuffer     []Value
	indexBuffer   []Value
	combineBuffer []Value

	pooledFrom      *vmProgramPool
	pooledSlot      chan struct{}
	comInitialized  bool
	comThreadLocked bool
}

// NewVM creates and initializes a new VM instance.
func NewVM(bytecode []byte, constants []Value, globalCount int) *VM {
	vm := &VM{
		bytecode:                       bytecode,
		constants:                      constants,
		Globals:                        make([]Value, globalCount),
		stack:                          make([]Value, StackSize),
		sp:                             -1,
		stmtSP:                         -1,
		ip:                             0,
		fp:                             0,
		optionCompare:                  0,
		declaredGlobals:                make(map[string]bool),
		constGlobals:                   make(map[string]bool),
		baseDeclared:                   make(map[string]bool),
		baseConst:                      make(map[string]bool),
		baseGlobalZeroArgFuncs:         make(map[string]bool),
		globalNames:                    make([]string, 0, globalCount),
		globalZeroArgFuncs:             make(map[string]bool),
		dynamicProgramStarts:           make(map[uint64]int, 32),
		globalNameIndex:                make(map[string]int, globalCount),
		argBuffer:                      make([]Value, 0, 16),
		indexBuffer:                    make([]Value, 0, 16),
		combineBuffer:                  make([]Value, 0, 16),
		withStack:                      make([]Value, 0, 8),
		nextDynamicNativeID:            20000,
		nextDynamicClassID:             60000,
		responseCookieItems:            make(map[int64]string),
		aspErrorItems:                  make(map[int64]*asp.ASPError),
		g3mdItems:                      make(map[int64]*G3MD),
		g3testItems:                    make(map[int64]*G3Test),
		g3cryptoItems:                  make(map[int64]*G3Crypto),
		g3jsonItems:                    make(map[int64]*G3JSON),
		g3httpItems:                    make(map[int64]*G3HTTP),
		g3mailItems:                    make(map[int64]*G3Mail),
		g3imageItems:                   make(map[int64]*G3Image),
		g3filesItems:                   make(map[int64]*G3Files),
		g3templateItems:                make(map[int64]*G3Template),
		g3zipItems:                     make(map[int64]*G3Zip),
		g3zlibItems:                    make(map[int64]*G3ZLIB),
		g3tarItems:                     make(map[int64]*G3TAR),
		g3zstdItems:                    make(map[int64]*G3ZSTD),
		g3fcItems:                      make(map[int64]*G3FC),
		g3dbItems:                      make(map[int64]*G3DB),
		g3dbResultSetItems:             make(map[int64]*G3DBResultSet),
		g3dbFieldsItems:                make(map[int64]*G3DBFields),
		g3dbRowItems:                   make(map[int64]*G3DBRow),
		g3dbStatementItems:             make(map[int64]*G3DBStatement),
		g3dbTransactionItems:           make(map[int64]*G3DBTransaction),
		g3dbResultItems:                make(map[int64]*G3DBResult),
		wscriptShellItems:              make(map[int64]*WScriptShell),
		wscriptExecItems:               make(map[int64]*WScriptExecObject),
		wscriptProcessStreamItems:      make(map[int64]*ProcessTextStream),
		adoxCatalogItems:               make(map[int64]*ADOXCatalog),
		adoxTablesItems:                make(map[int64]*ADOXTables),
		adoxTableItems:                 make(map[int64]*ADOXTableWrapper),
		mswcAdRotatorItems:             make(map[int64]*G3AdRotator),
		mswcBrowserTypeItems:           make(map[int64]*G3BrowserType),
		mswcNextLinkItems:              make(map[int64]*G3NextLink),
		mswcContentRotatorItems:        make(map[int64]*G3ContentRotator),
		mswcCountersItems:              make(map[int64]*G3Counters),
		mswcPageCounterItems:           make(map[int64]*G3PageCounter),
		mswcToolsItems:                 make(map[int64]*G3Tools),
		mswcMyInfoItems:                make(map[int64]*G3MyInfo),
		mswcPermissionCheckerItems:     make(map[int64]*G3PermissionChecker),
		msxmlServerItems:               make(map[int64]*MsXML2ServerXMLHTTP),
		msxmlDOMItems:                  make(map[int64]*MsXML2DOMDocument),
		msxmlNodeListItems:             make(map[int64]*XMLNodeList),
		msxmlParseErrorItems:           make(map[int64]*ParseError),
		msxmlElementItems:              make(map[int64]*XMLElement),
		pdfItems:                       make(map[int64]*G3PDF),
		fileUploaderItems:              make(map[int64]*G3FileUploader),
		axonItems:                      make(map[int64]*AxonLibrary),
		fsoItems:                       make(map[int64]*fsoNativeObject),
		adodbStreamItems:               make(map[int64]*adodbStreamNativeObject),
		adodbConnectionItems:           make(map[int64]*adodbConnection),
		adodbRecordsetItems:            make(map[int64]*adodbRecordset),
		adodbCommandItems:              make(map[int64]*adodbCommand),
		adodbParameterItems:            make(map[int64]*adodbParameter),
		adodbErrorsCollectionItems:     make(map[int64]*adodbConnection),
		adodbErrorItems:                make(map[int64]*adodbError),
		adodbFieldsCollectionItems:     make(map[int64]*adodbRecordset),
		adodbParametersCollectionItems: make(map[int64]*adodbCommand),
		adodbFieldItems:                make(map[int64]*adodbFieldProxy),
		regExpItems:                    make(map[int64]*regExpNativeObject),
		regExpMatchesCollectionItems:   make(map[int64]*regExpMatchesCollection),
		regExpMatchItems:               make(map[int64]*regExpMatch),
		regExpSubMatchesItems:          make(map[int64]*regExpSubMatches),
		regExpSubMatchValueItems:       make(map[int64]*regExpSubMatchValue),
		dictionaryItems:                make(map[int64]*scriptingDictionary),
		nativeObjectProxies:            make(map[int64]nativeObjectProxy),
		jsObjectItems:                  make(map[int64]map[string]Value),
		jsObjectStateItems:             make(map[int64]jsObjectState),
		jsPropertyItems:                make(map[int64]map[string]jsPropertyDescriptor),
		jsFunctionItems:                make(map[int64]*jsFunctionObject),
		jsForInItems:                   make(map[int]*jsForInEnumerator),
		jsEnvItems:                     make(map[int64]*jsEnvFrame),
		errObject:                      asp.NewASPError(),

		runtimeClasses:     make(map[string]RuntimeClassDef),
		runtimeClassItems:  make(map[int64]*RuntimeClassInstance),
		classInstanceOrder: make([]int64, 0, 16),
		callStack:          make([]CallFrame, 0, 16),
		jsCallStack:        make([]jsCallFrame, 0, 16),
		jsTryStack:         make([]int, 0, 8),
		jsErrStack:         make([]Value, 0, 4),
		jsThisValue:        Value{Type: VTJSUndefined},
		terminateCursor:    -1,
	}

	// 1. Inject ASP Native Objects (Indices 0-7)
	if globalCount >= 8 {
		vm.Globals[0] = Value{Type: VTNativeObject, Num: 0}                                // Response
		vm.Globals[1] = Value{Type: VTNativeObject, Num: 1}                                // Request
		vm.Globals[2] = Value{Type: VTNativeObject, Num: 2}                                // Server
		vm.Globals[3] = Value{Type: VTNativeObject, Num: 3}                                // Session
		vm.Globals[4] = Value{Type: VTNativeObject, Num: 4}                                // Application
		vm.Globals[5] = Value{Type: VTNativeObject, Num: int64(nativeObjectObjectContext)} // ObjectContext
		vm.Globals[6] = Value{Type: VTNativeObject, Num: int64(nativeObjectErr)}           // Err
		vm.Globals[7] = Value{Type: VTNativeObject, Num: int64(nativeObjectConsole)}       // console
	}
	// Indices 8 and 9 are reserved for OnTransactionCommit and OnTransactionAbort subs.
	// They default to VTEmpty and are set by the compiler if the user defines those subs.

	// 2. Inject Built-in Functions
	// We MUST follow the same order as the compiler.
	// Since the compiler iterates over BuiltinIndex, we do too.
	// This is safe because map iteration in Go is stable within the same process run
	// IF we use a sorted approach or just pre-inject in a fixed slice.
	// Let's use a more reliable way: BuiltinRegistry is a slice, it's already ordered.
	// But we need to know WHICH name maps to WHICH global index.
	// The compiler used c.Globals.Add(name).
	// To keep it simple, I'll rely on the fact that SymbolTable was populated
	// with Intrinsics FIRST, then Builtins.

	startIdx := 10 // 8 intrinsics + 2 event handler slots
	for i := 0; i < len(BuiltinRegistry); i++ {
		if startIdx+i < globalCount {
			vm.Globals[startIdx+i] = Value{Type: VTBuiltin, Num: int64(i)}
		}
	}

	// 3. Inject VBScript predefined constants (vbCrLf, vbLongDate, vbTrue, etc.)
	// They occupy slots immediately after the builtin function slots, in declaration
	// order of VBSConstants — matching the compiler's pre-injection order exactly.
	constStart := startIdx + len(BuiltinRegistry)
	for i, kv := range VBSConstants {
		if constStart+i < globalCount {
			vm.Globals[constStart+i] = kv.Val
		}
	}

	vm.captureBaseProgramState()
	vm.rebuildGlobalNameIndex()

	return vm
}

// NewVMFromCompiler creates a VM and preserves compiler scope metadata required
// by dynamic execution built-ins such as ExecuteGlobal.
func NewVMFromCompiler(compiler *Compiler) *VM {
	if compiler == nil {
		return NewVM(nil, nil, 0)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	vm.optionCompare = compiler.optionCompare
	vm.optionExplicit = compiler.optionExplicit
	vm.sourceName = compiler.sourceName
	vm.globalNames = append(vm.globalNames[:0], compiler.Globals.names...)
	vm.rebuildGlobalNameIndex()
	clear(vm.globalZeroArgFuncs)
	for key, value := range compiler.globalZeroArgFuncs {
		vm.globalZeroArgFuncs[key] = value
	}
	for key, value := range compiler.declaredGlobals {
		vm.declaredGlobals[key] = value
	}
	for key, value := range compiler.constGlobals {
		vm.constGlobals[key] = value
	}
	vm.captureBaseProgramState()
	return vm
}

// newExecuteGlobalCompiler creates a VBScript compiler that inherits the active
// script global namespace and option flags for ExecuteGlobal semantics.
func (vm *VM) newExecuteGlobalCompiler(code string) *Compiler {
	compiler := NewCompiler(code)
	compiler.optionCompare = vm.optionCompare
	compiler.optionExplicit = vm.optionExplicit
	compiler.SetSourceName(vm.sourceName)
	for _, name := range vm.globalNames {
		if strings.TrimSpace(name) == "" {
			continue
		}
		compiler.Globals.Add(name)
	}
	for key, value := range vm.globalZeroArgFuncs {
		compiler.globalZeroArgFuncs[key] = value
	}
	for key, value := range vm.declaredGlobals {
		compiler.declaredGlobals[key] = value
	}
	for key, value := range vm.constGlobals {
		compiler.constGlobals[key] = value
	}
	vm.seedCompilerClassDeclarationsFromRuntime(compiler)
	vm.attachDynamicClassResolutionContext(compiler)

	return compiler
}

// newExecuteGlobalPureCompiler creates a compiler for ExecuteGlobal that
// explicitly ignores any active class context from the caller.
func (vm *VM) newExecuteGlobalPureCompiler(code string) *Compiler {
	// Base global compiler inherits ALL existing global names to maintain slot indices.
	// This is CRITICAL for persistence so child doesn't overwrite parent globals.
	compiler := vm.newExecuteGlobalCompiler(code)

	// BUT we reset class-related active context for isolation
	compiler.currentClassName = ""
	compiler.dynamicMemberResolution = false

	return compiler
}

// newExecuteLocalCompiler creates a VBScript compiler that inherits both the
// global namespace and the active procedure's local variable names.
func (vm *VM) newExecuteLocalCompiler(code string, localSub Value, isEval bool) *Compiler {
	var compiler *Compiler
	if isEval {
		compiler = NewEvalCompiler(code)
	} else {
		compiler = NewCompiler(code)
	}

	compiler.optionCompare = vm.optionCompare
	compiler.optionExplicit = vm.optionExplicit
	compiler.SetSourceName(vm.sourceName)

	for _, name := range vm.globalNames {
		if strings.TrimSpace(name) == "" {
			continue
		}
		compiler.Globals.Add(name)
	}
	for key, value := range vm.globalZeroArgFuncs {
		compiler.globalZeroArgFuncs[key] = value
	}
	for key, value := range vm.declaredGlobals {
		compiler.declaredGlobals[key] = value
	}
	for key, value := range vm.constGlobals {
		compiler.constGlobals[key] = value
	}
	vm.seedCompilerClassDeclarationsFromRuntime(compiler)
	vm.attachDynamicClassResolutionContext(compiler)

	if localSub.Type == VTUserSub {
		compiler.isLocal = true
		for _, name := range localSub.Names {
			compiler.locals.Add(name)
			compiler.declaredLocals[strings.ToLower(name)] = true
		}
	}

	return compiler
}

// seedCompilerClassDeclarationsFromRuntime mirrors runtime class metadata into
// compiler class declarations so dynamic code can resolve class members.
func (vm *VM) seedCompilerClassDeclarationsFromRuntime(compiler *Compiler) {
	if vm == nil || compiler == nil {
		return
	}
	for _, classDef := range vm.runtimeClasses {
		trimmedClassName := strings.TrimSpace(classDef.Name)
		if trimmedClassName == "" {
			continue
		}
		lowerClassName := strings.ToLower(trimmedClassName)
		classIdx, exists := compiler.classDeclLookup[lowerClassName]
		if !exists {
			classIdx = len(compiler.classDecls)
			compiler.classDeclLookup[lowerClassName] = classIdx
			compiler.classDecls = append(compiler.classDecls, CompiledClassDecl{Name: trimmedClassName})
		}
		decl := &compiler.classDecls[classIdx]

		if classDef.Fields != nil {
			for _, fieldDef := range classDef.Fields {
				if strings.TrimSpace(fieldDef.Name) == "" {
					continue
				}
				decl.Fields = append(decl.Fields, CompiledClassFieldDecl{Name: fieldDef.Name, IsPublic: fieldDef.IsPublic})
			}
		}
		if classDef.Methods != nil {
			for methodName, methodDef := range classDef.Methods {
				trimmedMethodName := strings.TrimSpace(methodName)
				if trimmedMethodName == "" {
					continue
				}
				decl.Methods = append(decl.Methods, CompiledClassMethodDecl{Name: trimmedMethodName, IsPublic: methodDef.IsPublic})
			}
		}
		if classDef.Properties != nil {
			for _, propertyDef := range classDef.Properties {
				if strings.TrimSpace(propertyDef.Name) == "" {
					continue
				}
				decl.Properties = append(decl.Properties, CompiledClassPropertyDecl{
					Name:          propertyDef.Name,
					IsPublic:      propertyDef.IsPublic,
					GetParamCount: propertyDef.GetParamCount,
					HasGet:        propertyDef.HasGet,
					LetParamCount: propertyDef.LetParamCount,
					HasLet:        propertyDef.HasLet,
					SetParamCount: propertyDef.SetParamCount,
					HasSet:        propertyDef.HasSet,
				})
			}
		}
	}
}

// attachDynamicClassResolutionContext configures one dynamic compiler with the
// active class context so Eval/Execute/ExecuteGlobal can resolve implicit
// class-member references when invoked from class methods.
func (vm *VM) attachDynamicClassResolutionContext(compiler *Compiler) {
	if vm == nil || compiler == nil || vm.activeClassObjectID == 0 {
		return
	}
	instance, exists := vm.runtimeClassItems[vm.activeClassObjectID]
	if !exists || instance == nil {
		return
	}
	trimmedClassName := strings.TrimSpace(instance.ClassName)
	if trimmedClassName == "" {
		return
	}
	compiler.currentClassName = trimmedClassName
	compiler.dynamicMemberResolution = true
}

// opcodeOperandSize returns the number of inline operand bytes that follow the opcode byte.
// This mirrors the IP advances in the main execution loop handlers and is used by the
// Resume-Next statement-skip mechanism to advance ip past unexecuted opcodes.
func opcodeOperandSize(op OpCode) int {
	switch op {
	// 2-byte operands
	case OpConstant, OpWriteStatic,
		OpGetClassMember, OpSetClassMember, OpEraseClassMember, OpLetClassMember, OpArgClassMemberRef,
		OpMemberSet, OpMemberSetSet,
		OpNewClass,
		OpLabel,
		OpRegisterClass,
		OpGetGlobal, OpSetGlobal, OpEraseGlobal, OpGetLocal, OpSetLocal, OpEraseLocal,
		OpArgGlobalRef, OpArgLocalRef,
		OpLetGlobal, OpLetLocal,
		OpCall,
		OpJSDeclareName, OpJSGetName, OpJSSetName, OpJSMemberGet, OpJSMemberSet,
		OpJSCreateClosure, OpJSCall, OpJSNewArray, OpJSDelete,
		OpJSNew, OpJSMemberDelete, OpJSPostIncrement, OpJSPostDecrement, OpJSPreIncrement, OpJSPreDecrement,
		OpJSAddAssign, OpJSSubtractAssign, OpJSMultiplyAssign, OpJSDivideAssign, OpJSModuloAssign,
		OpJSMemberIndexGet, OpJSMemberIndexSet,
		OpJSPostMemberIncrement, OpJSPostMemberDecrement, OpJSPreMemberIncrement, OpJSPreMemberDecrement:
		return 2
	// 4-byte operands (for jump targets)
	case OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel,
		OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter,
		OpJSCase, OpJSDefault, OpJSBreak, OpJSContinue, OpJSForInCleanup:
		return 4
	// 4-byte operands
	case OpLine, OpCallMember, OpArraySet, OpCallBuiltin, OpSetDirective, OpSetOption, OpJSCallMember:
		return 4
	// 6-byte operands: classNameIdx(2) + fieldNameIdx(2) + isPublic(2)
	case OpRegisterClassField:
		return 6
	// 6-byte operands: classNameIdx(2) + fieldNameIdx(2) + dimCount(2); dim values are on stack
	case OpInitClassArrayField:
		return 6
	// 6-byte operands for OpJSForIn: EnumVarIdx(2) + LoopStartTarget(4)
	case OpJSForIn:
		return 6
	// 8-byte operands: classNameIdx(2) + methodNameIdx(2) + userSubIdx(2) + isPublic(2)
	case OpRegisterClassMethod:
		return 8
	// 10-byte operands: classNameIdx(2) + propertyNameIdx(2) + userSubIdx(2) + paramCount(2) + isPublic(2)
	case OpRegisterClassPropertyGet, OpRegisterClassPropertyLet, OpRegisterClassPropertySet:
		return 10
	// 0-byte operands (single-byte opcodes)
	default:
		return 0
	}
}

// remapExecuteGlobalBytecode rewrites constant indices and absolute jump targets
// so one freshly compiled program can be appended to the active VM bytecode.
func remapExecuteGlobalBytecode(bytecode []byte, constBase int, bytecodeBase int) {
	for ip := 0; ip < len(bytecode); {
		op := OpCode(bytecode[ip])
		ip++
		switch op {
		case OpConstant, OpWriteStatic, OpGetClassMember, OpSetClassMember, OpEraseClassMember, OpMemberSet, OpMemberSetSet, OpNewClass, OpLetClassMember, OpArgClassMemberRef,
			OpJSDeclareName, OpJSGetName, OpJSSetName, OpJSMemberGet, OpJSMemberSet, OpJSCreateClosure, OpJSDelete,
			OpJSNew, OpJSMemberDelete, OpJSPostIncrement, OpJSPostDecrement, OpJSPreIncrement, OpJSPreDecrement,
			OpJSAddAssign, OpJSSubtractAssign, OpJSMultiplyAssign, OpJSDivideAssign, OpJSModuloAssign,
			OpJSMemberIndexGet, OpJSMemberIndexSet,
			OpJSPostMemberIncrement, OpJSPostMemberDecrement, OpJSPreMemberIncrement, OpJSPreMemberDecrement:
			idx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx))
			ip += 2
		case OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel, OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter, OpJSCase, OpJSDefault, OpJSBreak, OpJSContinue, OpJSForInCleanup:
			target := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
			binary.BigEndian.PutUint32(bytecode[ip:], uint32(target))
			ip += 4
		case OpSetDirective:
			nameIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(nameIdx))
			ip += 2
			valueIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(valueIdx))
			ip += 2
		case OpCallMember, OpArraySet, OpJSCallMember:
			memberIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(memberIdx))
			ip += 4
		case OpJSForIn:
			// EnumVarIdx(2) + LoopStartTarget(4)
			enumVarIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(enumVarIdx))
			ip += 2
			loopStartTarget := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
			binary.BigEndian.PutUint32(bytecode[ip:], uint32(loopStartTarget))
			ip += 4
		case OpRegisterClass:
			classIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(classIdx))
			ip += 2
		case OpRegisterClassField:
			classIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(classIdx))
			ip += 2
			fieldIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(fieldIdx))
			ip += 4
		case OpInitClassArrayField:
			// classNameIdx(2) + fieldNameIdx(2) + dimCount(2); dim values came from OpConstant (already remapped)
			classArrIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(classArrIdx))
			ip += 2
			fieldArrIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(fieldArrIdx))
			ip += 4 // skip fieldArrIdx(2) + dimCount(2, no remap needed)
		case OpRegisterClassMethod:
			classIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(classIdx))
			ip += 2
			methodIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(methodIdx))
			ip += 2
			userSubIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(userSubIdx))
			ip += 4
		case OpRegisterClassPropertyGet, OpRegisterClassPropertyLet, OpRegisterClassPropertySet:
			classIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(classIdx))
			ip += 2
			propertyIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(propertyIdx))
			ip += 2
			userSubIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(userSubIdx))
			ip += 6
		case OpGetGlobal, OpSetGlobal, OpGetLocal, OpSetLocal, OpLabel, OpArgGlobalRef, OpArgLocalRef, OpLetGlobal, OpLetLocal, OpCall:
			ip += 2
		case OpSetOption:
			ip += 2
		case OpLine:
			ip += 4
		case OpCallBuiltin:
			ip += 4
		default:
			// Single-byte opcodes have no inline operands.
		}
	}
}

// appendExecuteProgram merges one compiled dynamic program into the active VM bytecode.
func (vm *VM) appendExecuteProgram(globalCount int, constants []Value, bytecode []byte) int {
	if globalCount < 0 {
		globalCount = 0
	}

	if len(vm.Globals) < globalCount {
		expanded := make([]Value, globalCount)
		copy(expanded, vm.Globals)
		vm.Globals = expanded
	}

	constBase := len(vm.constants)
	bytecodeBase := len(vm.bytecode)
	appendedConstants := append([]Value(nil), constants...)
	for i := range appendedConstants {
		if appendedConstants[i].Type == VTUserSub {
			appendedConstants[i].Num += int64(bytecodeBase)
		}
	}
	appendedBytecode := append([]byte(nil), bytecode...)
	remapExecuteGlobalBytecode(appendedBytecode, constBase, bytecodeBase)

	vm.constants = append(vm.constants, appendedConstants...)
	vm.bytecode = append(vm.bytecode, appendedBytecode...)
	return bytecodeBase
}

// appendCachedDynamicProgram appends one cached dynamic fragment once per VM and
// reuses the previously appended entry-point on subsequent compatible executions.
func (vm *VM) appendCachedDynamicProgram(compiled *dynamicCachedProgram) int {
	if vm == nil {
		return 0
	}
	if compiled == nil {
		return len(vm.bytecode)
	}
	if vm.dynamicProgramStarts == nil {
		vm.dynamicProgramStarts = make(map[uint64]int, 32)
	}
	if startIP, exists := vm.dynamicProgramStarts[compiled.keyHash]; exists {
		if startIP >= 0 && startIP < len(vm.bytecode) {
			return startIP
		}
	}
	startIP := vm.appendExecuteProgram(compiled.globalCount, compiled.constants, compiled.bytecode)
	vm.dynamicProgramStarts[compiled.keyHash] = startIP
	return startIP
}

// cloneForExecuteGlobal builds an isolated execution VM that shares the parent
// engine runtime state while using fresh stack and call frames.
func (vm *VM) cloneForExecuteGlobal(startIP int) *VM {
	child := *vm
	child.stack = make([]Value, StackSize)
	child.sp = -1
	child.ip = startIP
	child.fp = 0
	child.callStack = make([]CallFrame, 0, 16)
	child.withStack = make([]Value, 0, 8)
	child.activeClassObjectID = vm.activeClassObjectID
	child.terminateCursor = -1
	child.terminatePrepared = false
	child.suppressTerminate = true
	child.onResumeNext = vm.onResumeNext
	child.executeGlobalResumeGuard = vm.onResumeNext
	child.stmtSP = -1
	child.skipToNextStmt = false

	return &child
}

// cloneForExecuteLocal builds an execution VM that shares the parent
// stack and call frame context to enable local variable access.
func (vm *VM) cloneForExecuteLocal(startIP int) *VM {
	child := *vm
	// DO NOT reset stack, sp, or fp. Share parent's current frame.
	child.ip = startIP
	child.callStack = make([]CallFrame, len(vm.callStack))
	copy(child.callStack, vm.callStack)
	child.withStack = make([]Value, len(vm.withStack))
	copy(child.withStack, vm.withStack)
	child.activeClassObjectID = vm.activeClassObjectID
	child.terminateCursor = -1
	child.terminatePrepared = false
	child.suppressTerminate = true
	child.onResumeNext = vm.onResumeNext
	child.skipToNextStmt = false
	// stmtSP carries over from parent; child's first OpLine will reset it.

	return &child
}

// syncExecuteGlobalState copies back mutable runtime state after one dynamic
// ExecuteGlobal run completes.
func (vm *VM) syncExecuteGlobalState(child *VM) {
	if child == nil {
		return
	}
	vm.Globals = child.Globals
	vm.globalNames = child.globalNames
	vm.globalNamesHash = child.globalNamesHash
	vm.globalZeroArgFuncs = child.globalZeroArgFuncs
	vm.declaredGlobals = child.declaredGlobals
	vm.constGlobals = child.constGlobals
	vm.nextDynamicNativeID = child.nextDynamicNativeID
	vm.nextDynamicClassID = child.nextDynamicClassID
	vm.errObject = child.errObject
	vm.errASPCodeRaw = child.errASPCodeRaw
	vm.errASPCodeRawSet = child.errASPCodeRawSet
	vm.lastError = child.lastError
	vm.optionCompare = child.optionCompare
	vm.optionExplicit = child.optionExplicit
	vm.transactionState = child.transactionState

	// Sync item maps and other mutable state that might have been updated in the child.
	vm.runtimeClasses = child.runtimeClasses
	vm.runtimeClassItems = child.runtimeClassItems
	vm.classInstanceOrder = child.classInstanceOrder
	vm.responseCookieItems = child.responseCookieItems
	vm.aspErrorItems = child.aspErrorItems
	vm.g3mdItems = child.g3mdItems
	vm.g3cryptoItems = child.g3cryptoItems
	vm.g3jsonItems = child.g3jsonItems
	vm.g3httpItems = child.g3httpItems
	vm.g3mailItems = child.g3mailItems
	vm.g3imageItems = child.g3imageItems
	vm.g3filesItems = child.g3filesItems
	vm.g3templateItems = child.g3templateItems
	vm.g3zipItems = child.g3zipItems
	vm.g3zlibItems = child.g3zlibItems
	vm.g3tarItems = child.g3tarItems
	vm.g3zstdItems = child.g3zstdItems
	vm.g3fcItems = child.g3fcItems
	vm.g3dbItems = child.g3dbItems
	vm.g3dbResultSetItems = child.g3dbResultSetItems
	vm.g3dbFieldsItems = child.g3dbFieldsItems
	vm.g3dbRowItems = child.g3dbRowItems
	vm.g3dbStatementItems = child.g3dbStatementItems
	vm.g3dbTransactionItems = child.g3dbTransactionItems
	vm.g3dbResultItems = child.g3dbResultItems
	vm.wscriptShellItems = child.wscriptShellItems
	vm.wscriptExecItems = child.wscriptExecItems
	vm.wscriptProcessStreamItems = child.wscriptProcessStreamItems
	vm.adoxCatalogItems = child.adoxCatalogItems
	vm.adoxTablesItems = child.adoxTablesItems
	vm.adoxTableItems = child.adoxTableItems
	vm.mswcAdRotatorItems = child.mswcAdRotatorItems
	vm.mswcBrowserTypeItems = child.mswcBrowserTypeItems
	vm.mswcNextLinkItems = child.mswcNextLinkItems
	vm.mswcContentRotatorItems = child.mswcContentRotatorItems
	vm.mswcCountersItems = child.mswcCountersItems
	vm.mswcPageCounterItems = child.mswcPageCounterItems
	vm.mswcToolsItems = child.mswcToolsItems
	vm.mswcMyInfoItems = child.mswcMyInfoItems
	vm.mswcPermissionCheckerItems = child.mswcPermissionCheckerItems
	vm.msxmlServerItems = child.msxmlServerItems
	vm.msxmlDOMItems = child.msxmlDOMItems
	vm.msxmlNodeListItems = child.msxmlNodeListItems
	vm.msxmlParseErrorItems = child.msxmlParseErrorItems
	vm.msxmlElementItems = child.msxmlElementItems
	vm.pdfItems = child.pdfItems
	vm.fileUploaderItems = child.fileUploaderItems
	vm.axonItems = child.axonItems
	vm.fsoItems = child.fsoItems
	vm.adodbStreamItems = child.adodbStreamItems
	vm.adodbConnectionItems = child.adodbConnectionItems
	vm.adodbRecordsetItems = child.adodbRecordsetItems
	vm.adodbCommandItems = child.adodbCommandItems
	vm.adodbParameterItems = child.adodbParameterItems
	vm.adodbErrorsCollectionItems = child.adodbErrorsCollectionItems
	vm.adodbErrorItems = child.adodbErrorItems
	vm.adodbFieldsCollectionItems = child.adodbFieldsCollectionItems
	vm.adodbParametersCollectionItems = child.adodbParametersCollectionItems
	vm.adodbFieldItems = child.adodbFieldItems
	vm.regExpItems = child.regExpItems
	vm.regExpMatchesCollectionItems = child.regExpMatchesCollectionItems
	vm.regExpMatchItems = child.regExpMatchItems
	vm.regExpSubMatchesItems = child.regExpSubMatchesItems
	vm.regExpSubMatchValueItems = child.regExpSubMatchValueItems
	vm.dictionaryItems = child.dictionaryItems
	vm.nativeObjectProxies = child.nativeObjectProxies
	vm.jsObjectItems = child.jsObjectItems
	vm.jsObjectStateItems = child.jsObjectStateItems
	vm.jsPropertyItems = child.jsPropertyItems
	vm.jsFunctionItems = child.jsFunctionItems
	vm.jsForInItems = child.jsForInItems
	vm.jsEnvItems = child.jsEnvItems
}

// syncExecuteLocalState propagates stack and frame state back after Execute/Eval.
func (vm *VM) syncExecuteLocalState(child *VM) {
	if child == nil {
		return
	}
	vm.syncExecuteGlobalState(child)
	vm.stack = child.stack
	vm.sp = child.sp
}

// SetHost attaches the host environment to the VM.
func (vm *VM) SetHost(h ASPHostEnvironment) {
	vm.host = h
	vm.output = h
}

// SetOutput sets the output writer for the VM.
func (vm *VM) SetOutput(w io.Writer) {
	vm.output = w
}

// Run executes the loaded bytecode.
func (vm *VM) Run() (err error) {
	if vm.host != nil && vm.host.Server() != nil {
		vm.host.Server().BeginExecution()
		defer vm.host.Server().EndExecution()
	}
	defer func() {
		if r := recover(); r != nil {
			if endSignal, ok := r.(string); ok && endSignal == asp.ResponseEndSignal {
				// ExecuteGlobal and Eval child VMs have suppressTerminate=true.
				// Re-propagate Response.End so it reaches the top-level Run()
				// and stops the entire page execution, not just the child scope.
				if vm.suppressTerminate {
					panic(endSignal)
				}
				err = nil
				return
			}
			if bufferErr, ok := r.(*asp.ResponseBufferLimitError); ok {
				err = NewAxonASPError(ErrResponseBufferLimitExceeded, bufferErr, bufferErr.Error(), vm.sourceName, vm.lastLine)
				return
			}
			vme, ok := r.(*VMError)
			if ok {
				if vm.onResumeNext {
					vm.lastError = vme
					err = nil
				} else {
					err = vme
				}
			} else {
				if vm.onResumeNext {
					err = nil
				} else {
					err = fmt.Errorf("internal runtime panic at line %d: %v", vm.lastLine, r)
				}
			}
		}
	}()
	operationCount := 0
	jsBackJumpCount := 0
	vm.jsStringWorkBytes = 0

aspExecLoop:
	for vm.ip < len(vm.bytecode) {
		operationCount++
		if operationCount&1023 == 0 && vm.host != nil && vm.host.Server() != nil && vm.host.Server().HasTimedOut() {
			return NewAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("Script execution exceeded the configured timeout of %d second(s)", vm.host.Server().GetScriptTimeout()), vm.sourceName, vm.lastLine)
		}
		op := OpCode(vm.bytecode[vm.ip])
		vm.ip++

		// When Resume Next absorbed an error mid-statement, skip bytecode until the next
		// statement boundary (OpLine) or take any OpJumpIfFalse unconditionally to correctly
		// skip compound-statement blocks (e.g. If/Then body). OpRet and OpHalt fall through
		// so function returns and program termination work normally during skip mode.
		if vm.skipToNextStmt {
			switch op {
			case OpLine, OpHalt, OpRet:
				// Fall through to the main switch handler below.
			case OpJumpIfFalse:
				// Take the jump unconditionally: this skips the guarded block (If/Then body, loop).
				target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
				vm.ip = target
				vm.sp = vm.stmtSP
				vm.skipToNextStmt = false
				continue
			default:
				vm.ip += opcodeOperandSize(op)
				continue
			}
		}

		switch op {
		case OpHalt:
			break aspExecLoop

		case OpConstant:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.constants[idx])

		case OpPop:
			vm.pop()

		case OpGetGlobal:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.Globals[idx])

		case OpSetGlobal:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			newVal := vm.pop()
			// Decrement reference count of previous value in this slot (if it's an object).
			vm.decrementObjectRefCount(vm.Globals[idx])
			// Assign new value.
			vm.Globals[idx] = newVal
			// Increment reference count of new value (if it's an object).
			vm.incrementObjectRefCount(newVal)

		case OpEraseGlobal:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.Globals[idx] = vm.eraseValue(vm.Globals[idx])

		case OpGetClassMember:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			memberName := vm.constants[memberIdx].Str
			if vm.activeClassObjectID != 0 {
				me := Value{Type: VTObject, Num: vm.activeClassObjectID}
				// 1. Try Zero-Arg Property Get
				propTarget, ok := vm.resolveRuntimeClassPropertyGet(me, memberName, 0, false)
				if ok {
					if vm.beginUserSubCall(propTarget, nil, false, me.Num) {
						continue
					}
				}
				// 2. Try Method (Function)
				methodTarget, ok := vm.resolveRuntimeClassMethod(me, memberName, false)
				if ok {
					if vm.beginUserSubCall(methodTarget, nil, false, me.Num) {
						continue
					}
				}
			}
			vm.push(vm.getActiveClassMemberValue(memberName))

		case OpSetClassMember:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			memberName := vm.constants[memberIdx].Str
			newVal := vm.pop()
			if vm.activeClassObjectID != 0 {
				me := Value{Type: VTObject, Num: vm.activeClassObjectID}
				propTarget, ok := vm.resolveRuntimeClassPropertySet(me, memberName, 1, true, false, false)
				if ok {
					if vm.beginUserSubCall(propTarget, []Value{newVal}, true, me.Num) {
						continue
					}
				}
			}
			// Decrement reference count of previous value in this member slot.
			prevVal := vm.getActiveClassMemberValue(memberName)
			vm.decrementObjectRefCount(prevVal)
			// Assign new value.
			vm.setActiveClassMemberValue(memberName, newVal)
			// Increment reference count of new value.
			vm.incrementObjectRefCount(newVal)

		case OpEraseClassMember:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			memberName := vm.constants[memberIdx].Str
			vm.setActiveClassMemberValue(memberName, vm.eraseValue(vm.getActiveClassMemberValue(memberName)))

		// OpLetGlobal: plain VBScript "name = value" for a global variable.
		// Variable slots are mutable Variants and must be overwritten directly.
		case OpLetGlobal:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			newVal := vm.pop()
			// Decrement reference count of previous value in this slot.
			vm.decrementObjectRefCount(vm.Globals[idx])
			// Assign new value.
			vm.Globals[idx] = newVal
			// Increment reference count of new value.
			vm.incrementObjectRefCount(newVal)

		// OpLetLocal: plain VBScript "name = value" for a local variable.
		// Local slots are mutable Variants and must be overwritten directly.
		case OpLetLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			newVal := vm.pop()
			// Decrement reference count of previous value in this slot.
			vm.decrementObjectRefCount(vm.stack[slot])
			// Assign new value.
			vm.stack[slot] = newVal
			// Increment reference count of new value.
			vm.incrementObjectRefCount(newVal)

		case OpEraseLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			vm.stack[slot] = vm.eraseValue(vm.stack[slot])

		// OpLetClassMember: plain VBScript "name = value" for a class member field.
		// Class fields behave like mutable Variant slots and should be overwritten
		// directly, regardless of the previous runtime subtype.
		case OpLetClassMember:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			val := vm.pop()
			memberName := vm.constants[memberIdx].Str
			if vm.activeClassObjectID != 0 {
				me := Value{Type: VTObject, Num: vm.activeClassObjectID}
				propTarget, ok := vm.resolveRuntimeClassPropertyLet(me, memberName, 1, false)
				if ok {
					if vm.beginUserSubCall(propTarget, []Value{val}, true, me.Num) {
						continue
					}
				}
			}
			// Decrement reference count of previous value in this member slot.
			prevVal := vm.getActiveClassMemberValue(memberName)
			vm.decrementObjectRefCount(prevVal)
			// Assign new value.
			vm.setActiveClassMemberValue(memberName, val)
			// Increment reference count of new value.
			vm.incrementObjectRefCount(val)

		// OpCoerceToValue: pops TOS; if it is a VTObject with a zero-arg default Property Get,
		// starts the getter call so its return value is pushed by OpRet. For any other value
		// (non-object, Nothing, or object without default property) the value is re-pushed
		// unchanged, letting the consuming operator raise the appropriate error.
		case OpCoerceToValue:
			v := vm.pop()
			if v.Type == VTBuiltin {
				vm.push(resolveCallable(vm, v))
				continue
			}
			if v.Type == VTUserSub && v.UserSubIsFunc() && v.UserSubParamCount() == 0 {
				if vm.beginUserSubCall(v, nil, false, 0) {
					continue
				}
			}
			if v.Type == VTObject && v.Num != 0 {
				propTarget, ok := vm.resolveRuntimeClassPropertyGet(v, "__default__", 0, true)
				if ok {
					if vm.beginUserSubCall(propTarget, nil, false, v.Num) {
						continue
					}
				}
			} else if v.Type == VTNativeObject {
				// Only ADODB.Field proxies should auto-coerce through default Value in
				// generic value contexts. Coercing every native object via __default__
				// can break object argument passing (e.g. JSON.dump(recordset)).
				if adodbDefault, handled := vm.dispatchADODBFieldPropertyGet(v.Num, "__default__"); handled {
					vm.push(adodbDefault)
					continue
				}
			}
			vm.push(v)

		case OpGetLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.stack[vm.fp+int(offset)])

		case OpSetLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			newVal := vm.pop()
			// Decrement reference count of previous value in this slot (if it's an object).
			vm.decrementObjectRefCount(vm.stack[slot])
			// Assign new value.
			vm.stack[slot] = newVal
			// Increment reference count of new value (if it's an object).
			vm.incrementObjectRefCount(newVal)

		case OpAdd:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.addValues(a, b))

		case OpSub:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.subtractValues(a, b))

		case OpMul:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.multiplyValues(a, b))

		case OpDiv:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.divideValues(a, b))

		case OpMod:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.modValues(a, b))

		case OpPow:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.powValues(a, b))

		case OpIAdd:
			b, a := vm.pop(), vm.pop()
			vm.push(NewInteger(vm.coerceInt64(a) + vm.coerceInt64(b)))

		case OpISub:
			b, a := vm.pop(), vm.pop()
			vm.push(NewInteger(vm.coerceInt64(a) - vm.coerceInt64(b)))

		case OpIMul:
			b, a := vm.pop(), vm.pop()
			vm.push(NewInteger(vm.coerceInt64(a) * vm.coerceInt64(b)))

		case OpIDiv:
			// VBScript integer division: a \ b
			b, a := vm.pop(), vm.pop()
			vm.push(vm.intDivideValues(a, b))

		case OpNeq:
			b, a := vm.pop(), vm.pop()
			if isNull(a) || isNull(b) {
				vm.push(NewNull())
			} else if vm.optionCompare == 1 {
				vm.push(NewBool(!strings.EqualFold(a.String(), b.String())))
			} else {
				vm.push(NewBool(vm.compareValues(a, b) != 0))
			}

		case OpIsRef:
			b, a := vm.pop(), vm.pop()
			if !isObjectReferenceValue(a) || !isObjectReferenceValue(b) {
				vm.raise(vbscript.TypeMismatch, vbscript.TypeMismatch.String())
				vm.push(NewEmpty())
				continue
			}
			vm.push(NewBool(a.Num == b.Num))

		case OpIsNotRef:
			b, a := vm.pop(), vm.pop()
			if !isObjectReferenceValue(a) || !isObjectReferenceValue(b) {
				vm.raise(vbscript.TypeMismatch, vbscript.TypeMismatch.String())
				vm.push(NewEmpty())
				continue
			}
			vm.push(NewBool(a.Num != b.Num))

		case OpGt:
			b, a := vm.pop(), vm.pop()
			if isNull(a) || isNull(b) {
				vm.push(NewNull())
			} else {
				vm.push(NewBool(vm.compareValues(a, b) > 0))
			}

		case OpLte:
			b, a := vm.pop(), vm.pop()
			if isNull(a) || isNull(b) {
				vm.push(NewNull())
			} else {
				vm.push(NewBool(vm.compareValues(a, b) <= 0))
			}

		case OpGte:
			b, a := vm.pop(), vm.pop()
			if isNull(a) || isNull(b) {
				vm.push(NewNull())
			} else {
				vm.push(NewBool(vm.compareValues(a, b) >= 0))
			}

		case OpAnd:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.logicalAnd(a, b))

		case OpOr:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.logicalOr(a, b))

		case OpXor:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.logicalXor(a, b))

		case OpNot:
			v := vm.pop()
			if isNull(v) {
				vm.push(NewNull())
			} else if v.Type == VTBool {
				vm.push(NewBool(!vm.asBool(v)))
			} else {
				// VBScript Not requires a numeric-compatible operand. Non-numeric strings
				// (including "True", "False") raise Type mismatch, matching CLng coercion.
				n, ok := vm.coerceLogicalInt64(v)
				if !ok {
					vm.raise(vbscript.TypeMismatch, vbscript.TypeMismatch.String())
					vm.push(NewEmpty())
					continue
				}
				vm.push(NewInteger(^n))
			}

		case OpEqv:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.logicalEqv(a, b))

		case OpImp:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.logicalImp(a, b))

		case OpJumpIfTrue:
			target := binary.BigEndian.Uint32(vm.bytecode[vm.ip:])
			vm.ip += 4
			if vm.isTruthy(vm.pop()) {
				vm.ip = int(target)
			}

		case OpSetDirective:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			valueIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.applyDirective(vm.constants[nameIdx].Str, vm.constants[valueIdx].Str)

		case OpRegisterClass:
			classNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.registerRuntimeClass(vm.constants[classNameIdx].Str)

		case OpRegisterClassMethod:
			classNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			methodNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			userSubConstIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			isPublic := binary.BigEndian.Uint16(vm.bytecode[vm.ip:]) != 0
			vm.ip += 2
			vm.registerRuntimeClassMethod(vm.constants[classNameIdx].Str, vm.constants[methodNameIdx].Str, vm.constants[userSubConstIdx], isPublic)

		case OpRegisterClassField:
			classNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			fieldNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			isPublic := binary.BigEndian.Uint16(vm.bytecode[vm.ip:]) != 0
			vm.ip += 2
			vm.registerRuntimeClassField(vm.constants[classNameIdx].Str, vm.constants[fieldNameIdx].Str, isPublic)

		case OpInitClassArrayField:
			classArrNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			fieldArrNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			arrDimCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			arrDims := make([]int, arrDimCount)
			for i := arrDimCount - 1; i >= 0; i-- {
				arrDims[i] = vm.asInt(vm.pop())
			}
			vm.registerRuntimeClassFieldDims(vm.constants[classArrNameIdx].Str, vm.constants[fieldArrNameIdx].Str, arrDims)

		case OpRegisterClassPropertyGet, OpRegisterClassPropertyLet, OpRegisterClassPropertySet:
			classNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			propertyNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			userSubConstIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			paramCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			isPublicProp := binary.BigEndian.Uint16(vm.bytecode[vm.ip:]) != 0
			vm.ip += 2
			vm.registerRuntimeClassPropertyAccessor(op, vm.constants[classNameIdx].Str, vm.constants[propertyNameIdx].Str, vm.constants[userSubConstIdx], paramCount, isPublicProp)

		case OpLabel:
			// No-op: label markers carry a 2-byte ID consumed at compile time for GoTo patching.
			vm.ip += 2

		case OpGotoLabel:
			// Jump to the pre-patched absolute bytecode offset.
			vm.ip = int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))

		case OpConcat:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.concatValues(a, b))

		case OpWriteStatic:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			if vm.output != nil {
				io.WriteString(vm.output, vm.constants[idx].Str)
			}

		case OpSetOption:
			optionID := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			val := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			if optionID == 0 { // Option Compare
				vm.optionCompare = int(val)
			}

		case OpWrite:
			val := vm.pop()
			if vm.output != nil {
				io.WriteString(vm.output, vm.valueToString(val))
			}

		case OpCallMember:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2

			memberArgs := vm.ensureArgBuffer(argCount)
			hasArgRefMember := false
			for i := argCount - 1; i >= 0; i-- {
				v := vm.pop()
				if v.Type == VTArgRef {
					hasArgRefMember = true
					memberArgs[i] = v
				} else {
					memberArgs[i] = resolveCallable(vm, v)
				}
			}

			target := vm.pop()
			switch target.Type {
			case VTObject:
				if target.Num == 0 {
					vm.raise(vbscript.CouldNotFindTargetObject, "Object required")
					vm.push(Value{Type: VTEmpty})
					continue
				}
				requirePublic := target.Num != vm.activeClassObjectID
				methodTarget, ok := vm.resolveRuntimeClassMethod(target, vm.constants[memberIdx].Str, requirePublic)
				if ok {
					if methodTarget.UserSubParamCount() != argCount {
						vm.raise(vbscript.WrongNumberOfParameters, "Wrong number of parameters or invalid property assignment")
					}
					var memberByRefs []byRefWriteback
					if hasArgRefMember {
						memberByRefs = vm.collectByRefsAndUnwrap(memberArgs, methodTarget.UserSubByRefMask())
					}
					if vm.beginUserSubCall(methodTarget, memberArgs, false, target.Num, memberByRefs) {
						continue
					}
				}
				propertyTarget, ok := vm.resolveRuntimeClassPropertyGet(target, vm.constants[memberIdx].Str, argCount, requirePublic)
				if ok {
					var propByRefs []byRefWriteback
					if hasArgRefMember {
						propByRefs = vm.collectByRefsAndUnwrap(memberArgs, propertyTarget.UserSubByRefMask())
					}
					if vm.beginUserSubCall(propertyTarget, memberArgs, false, target.Num, propByRefs) {
						continue
					}
				}
				fieldValue, ok := vm.resolveRuntimeClassField(target, vm.constants[memberIdx].Str, requirePublic)
				if ok {
					if argCount == 0 {
						vm.push(fieldValue)
						continue
					}
					if fieldValue.Type == VTArray {
						if hasArgRefMember {
							vm.collectByRefsAndUnwrap(memberArgs, 0) // unwrap without write-back for array indexing
						}
						vm.push(vm.readArrayElement(fieldValue, memberArgs))
						continue
					}
					if fieldValue.Type == VTNativeObject {
						if hasArgRefMember {
							vm.collectByRefsAndUnwrap(memberArgs, 0) // unwrap VTArgRef without write-back for native default-member calls
						}
						result := vm.dispatchNativeCall(fieldValue.Num, "", memberArgs)
						vm.push(result)
						continue
					}
				}
				callClass := ""
				if instance, exists := vm.runtimeClassItems[target.Num]; exists && instance != nil {
					callClass = instance.ClassName
				}
				_ = callClass
				vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Object doesn't support this property or method")
				vm.push(Value{Type: VTEmpty})
			case VTNativeObject:
				memberName := vm.constants[memberIdx].Str
				var nativeByRefs []byRefWriteback
				if hasArgRefMember {
					nativeByRefs = vm.collectByRefsAndUnwrap(memberArgs, vm.nativeByRefMask(target.Num, memberName))
				}
				result := vm.dispatchNativeCall(target.Num, memberName, memberArgs)
				if len(nativeByRefs) > 0 {
					vm.applyByRefWritebacksFromArgs(nativeByRefs, memberArgs)
				}
				vm.push(result)
			default:
				vm.raise(vbscript.CouldNotFindTargetObject, "Object required")
				vm.push(Value{Type: VTEmpty})
			}

		case OpCallBuiltin:
			registryIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2

			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = resolveCallable(vm, vm.pop())
			}

			fn := BuiltinRegistry[registryIdx]
			result, err := fn(vm, args)
			if err != nil {
				if runtimeErr, ok := err.(builtinVBRuntimeError); ok {
					vm.raise(runtimeErr.code, runtimeErr.Error())
				} else {
					vm.raise(vbscript.InternalError, err.Error())
				}
			}
			vm.push(result)

		case OpMemberGet:
			member := vm.pop()
			target := vm.pop()
			if target.Type == VTObject {
				if target.Num == 0 {
					vm.raise(vbscript.CouldNotFindTargetObject, "Object required")
					vm.push(Value{Type: VTEmpty})
					continue
				}
				requirePublic := target.Num != vm.activeClassObjectID
				propertyTarget, ok := vm.resolveRuntimeClassPropertyGet(target, member.String(), 0, requirePublic)
				if ok {
					if vm.beginUserSubCall(propertyTarget, nil, false, target.Num) {
						continue
					}
				}
				fieldValue, ok := vm.resolveRuntimeClassField(target, member.String(), requirePublic)
				if ok {
					vm.push(fieldValue)
					continue
				}
				methodTarget, ok := vm.resolveRuntimeClassMethod(target, member.String(), requirePublic)
				if ok {
					if vm.beginUserSubCall(methodTarget, nil, false, target.Num) {
						continue
					}
				}
				getClass := ""
				if instance, exists := vm.runtimeClassItems[target.Num]; exists && instance != nil {
					getClass = instance.ClassName
				}
				vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, fmt.Sprintf("Object doesn’t support this property or method (get target=%s class=%s member=%s)", target.String(), getClass, member.String()))
				vm.push(Value{Type: VTEmpty})
				continue
			}
			if target.Type != VTNativeObject {
				vm.raise(vbscript.CouldNotFindTargetObject, "Object required")
				vm.push(Value{Type: VTEmpty})
				continue
			}
			vm.push(vm.dispatchMemberGet(target, member.String()))

		// OpMe pushes the current class instance onto the stack as a VTObject.
		// This implements the VBScript "Me" keyword inside class methods and properties.
		case OpMe:
			if vm.activeClassObjectID == 0 {
				vm.raise(vbscript.ObjectVariableNotSet, "Me is not valid outside a class method")
				vm.push(Value{Type: VTEmpty})
				continue
			}
			vm.push(Value{Type: VTObject, Num: vm.activeClassObjectID})

		case OpMemberSet, OpMemberSetSet:
			// OpMemberSet: [OpCode, ConstMemberIdxHigh, ConstMemberIdxLow]
			// Stack: [..., target_obj, value]
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			value := vm.pop()
			target := vm.pop()
			if target.Type == VTObject {
				requirePublic := target.Num != vm.activeClassObjectID
				preferSet := op == OpMemberSetSet || value.Type == VTObject || value.Type == VTNativeObject
				strictSet := op == OpMemberSetSet
				propertyTarget, ok := vm.resolveRuntimeClassPropertySet(target, vm.constants[memberIdx].Str, 1, preferSet, strictSet, requirePublic)
				if ok {
					if vm.beginUserSubCall(propertyTarget, []Value{value}, true, target.Num) {
						continue
					}
				}
				if vm.assignRuntimeClassField(target, vm.constants[memberIdx].Str, value, requirePublic) {
					continue
				}
				vm.raise(vbscript.WrongNumberOfParameters, "Wrong number of parameters or invalid property assignment")
			} else if target.Type == VTNativeObject {
				vm.dispatchMemberSet(target.Num, vm.constants[memberIdx].Str, value)
			} else {
				vm.raise(vbscript.TypeMismatch, "Object required for member set")
			}

		// OpWithEnter pops the With-subject object from the data stack and stores it
		// on the VM with-object stack. Executed once on entry to a With block.
		case OpWithEnter:
			vm.withStack = append(vm.withStack, vm.pop())

		// OpWithLeave removes the innermost With-subject object from the with-object stack.
		// Executed once at End With.
		case OpWithLeave:
			if len(vm.withStack) > 0 {
				vm.withStack = vm.withStack[:len(vm.withStack)-1]
			}

		// OpWithLoad pushes a copy of the innermost With-subject object onto the data stack.
		// Used before every '.Member' access inside a With block.
		case OpWithLoad:
			if len(vm.withStack) == 0 {
				vm.raise(vbscript.ObjectVariableNotSet, "With block is not active")
				vm.push(Value{Type: VTEmpty})
				continue
			}
			vm.push(vm.withStack[len(vm.withStack)-1])

		case OpJump:
			vm.ip = int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))

		case OpJumpIfFalse:
			target := binary.BigEndian.Uint32(vm.bytecode[vm.ip:])
			vm.ip += 4
			if !vm.isTruthy(vm.pop()) {
				vm.ip = int(target)
			}

		case OpEq:
			b, a := vm.pop(), vm.pop()
			if isNull(a) || isNull(b) {
				vm.push(NewNull())
			} else if vm.optionCompare == 1 { // Text
				vm.push(NewBool(strings.EqualFold(a.String(), b.String())))
			} else {
				vm.push(NewBool(vm.compareValues(a, b) == 0))
			}

		case OpLt:
			b, a := vm.pop(), vm.pop()
			vm.push(NewBool(vm.asFloat(a) < vm.asFloat(b)))

		case OpNeg:
			val := vm.pop()
			if isNull(val) {
				vm.push(NewNull())
			} else if val.Type == VTDouble {
				vm.push(NewDouble(-val.Flt))
			} else {
				vm.push(NewInteger(-vm.coerceInt64(val)))
			}

		case OpLine:
			if vm.ip+3 < len(vm.bytecode) {
				vm.lastLine = int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				vm.lastColumn = int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:]))
				vm.ip += 4
			} else {
				vm.lastLine = int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				vm.lastColumn = 0
				vm.ip += 2
			}
			// If Resume Next absorbed an error mid-statement, restore the stack and clear the flag.
			if vm.skipToNextStmt {
				vm.sp = vm.stmtSP
				vm.skipToNextStmt = false
			}
			vm.stmtSP = vm.sp

		case OpOnErrorResumeNext:
			vm.onResumeNext = true

		case OpOnErrorGoto0:
			vm.onResumeNext = false
			vm.errClear()

		// OpArgGlobalRef pushes a VTArgRef that carries the global slot index.
		// The VM reads the actual value from the slot when the call is dispatched,
		// and writes back the modified value on OpRet for ByRef parameters.
		case OpArgGlobalRef:
			idx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			vm.push(Value{Type: VTArgRef, Num: int64(idx), Flt: 1.0})

		// OpArgLocalRef pushes a VTArgRef that carries the local slot offset (relative to vm.fp).
		case OpArgLocalRef:
			idx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			vm.push(Value{Type: VTArgRef, Num: int64(idx), Flt: 0.0})

		case OpArgClassMemberRef:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			memberName := vm.constants[memberIdx].Str
			vm.push(Value{Type: VTArgRef, Str: memberName, Flt: 2.0})
		case OpArraySet:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2

			value := vm.pop()
			indexes := vm.ensureIndexBuffer(argCount)
			hasArgRefIndex := false
			for i := argCount - 1; i >= 0; i-- {
				idxVal := vm.pop()
				if idxVal.Type == VTArgRef {
					hasArgRefIndex = true
				}
				indexes[i] = idxVal
			}
			if hasArgRefIndex {
				vm.collectByRefsAndUnwrap(indexes, 0)
			}

			target := vm.pop()
			if target.Type == VTNativeObject {
				callArgs := vm.ensureCombineBuffer(argCount + 1)[:0]
				callArgs = append(callArgs, indexes...)
				callArgs = append(callArgs, value)
				_ = vm.dispatchNativeCall(target.Num, vm.constants[memberIdx].Str, callArgs)
			} else if target.Type == VTObject {
				// Indexed default Property Let: obj(idx...) = value → call property Let.
				memberName := vm.constants[memberIdx].Str
				if memberName == "" {
					memberName = "__default__"
				}
				letArgCount := argCount + 1
				letArgs := vm.ensureCombineBuffer(letArgCount)[:0]
				letArgs = append(letArgs, indexes...)
				letArgs = append(letArgs, value)
				propertyTarget, ok := vm.resolveRuntimeClassPropertySet(target, memberName, letArgCount, false, false, true)
				if ok {
					if vm.beginUserSubCall(propertyTarget, letArgs, true, target.Num) {
						continue
					}
				}
				vm.raise(vbscript.CouldNotFindTargetObject, "Object required: no default indexed property")
			} else {
				vm.assignArrayElement(target, indexes, value)
			}

		case OpCall:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2

			callArgs := vm.ensureArgBuffer(argCount)
			hasArgRef := false
			for i := argCount - 1; i >= 0; i-- {
				v := vm.pop()
				if v.Type == VTArgRef {
					hasArgRef = true
					callArgs[i] = v
				} else {
					callArgs[i] = resolveCallable(vm, v)
				}
			}

			target := vm.pop()

			args := callArgs
			// Resolve ByRef slot references: collect write-back entries and unwrap VTArgRef to actual values.
			var opCallByRefs []byRefWriteback
			if hasArgRef {
				var mask uint64
				switch target.Type {
				case VTUserSub:
					mask = target.UserSubByRefMask()
				case VTNativeObject:
					mask = vm.nativeByRefMask(target.Num, "")
				}
				opCallByRefs = vm.collectByRefsAndUnwrap(callArgs, mask)
			}

			if target.Type == VTBuiltin {
				fn := BuiltinRegistry[target.Num]
				result, err := fn(vm, args)
				if err != nil {
					if runtimeErr, ok := err.(builtinVBRuntimeError); ok {
						vm.raise(runtimeErr.code, runtimeErr.Error())
					} else {
						vm.raise(vbscript.InternalError, err.Error())
					}
				}
				vm.push(result)
			} else if target.Type == VTObject {
				defaultMethod, ok := vm.resolveRuntimeClassMethod(target, "__default__", true)
				if ok {
					if defaultMethod.UserSubParamCount() != argCount {
						vm.raise(vbscript.WrongNumberOfParameters, "Wrong number of parameters or invalid property assignment")
						vm.push(Value{Type: VTEmpty})
						continue
					}
					if vm.beginUserSubCall(defaultMethod, args, false, target.Num, opCallByRefs) {
						continue
					}
				}

				defaultProperty, ok := vm.resolveRuntimeClassPropertyGet(target, "__default__", argCount, true)
				if ok {
					if vm.beginUserSubCall(defaultProperty, args, false, target.Num, opCallByRefs) {
						continue
					}
				}

				invokeClass := ""
				if instance, exists := vm.runtimeClassItems[target.Num]; exists && instance != nil {
					invokeClass = instance.ClassName
				}
				vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, fmt.Sprintf("Object doesn't support this property or method (invoke target=%s class=%s argc=%d)", target.String(), invokeClass, argCount))
				vm.push(Value{Type: VTEmpty})
			} else if target.Type == VTNativeObject {
				result := vm.dispatchNativeCall(target.Num, "", args)
				if len(opCallByRefs) > 0 {
					vm.applyByRefWritebacksFromArgs(opCallByRefs, args)
				}
				vm.push(result)
			} else if vm.beginUserSubCall(target, args, false, 0, opCallByRefs) {
				continue
			} else if target.Type == VTArray {
				vm.push(vm.readArrayElement(target, args))
			} else {
				vm.push(Value{Type: VTEmpty})
			}

		case OpNewClass:
			classNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			instance := vm.newRuntimeClassInstance(vm.constants[classNameIdx].Str)
			vm.push(instance)
			if instance.Type != VTObject {
				continue
			}
			initializerTarget, ok := vm.resolveRuntimeClassMethod(instance, "Class_Initialize", false)
			if ok {
				if initializerTarget.UserSubParamCount() != 0 {
					vm.raise(vbscript.ClassInitializeOrTerminateDoNotHaveArguments, "Class_Initialize must not declare arguments")
				}
				if vm.beginUserSubCall(initializerTarget, nil, true, instance.Num) {
					continue
				}
			}

		case OpJSDeclareName:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.jsDeclareName(vm.constants[nameIdx].Str)

		case OpJSGetName:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.jsGetName(vm.constants[nameIdx].Str))

		case OpJSSetName:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			value := vm.pop()
			vm.jsSetName(vm.constants[nameIdx].Str, value)

		case OpJSMemberGet:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			target := vm.pop()
			if value, deferred := vm.jsMemberGet(target, vm.constants[nameIdx].Str); !deferred {
				vm.push(value)
			}

		case OpJSMemberSet:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			value := vm.pop()
			target := vm.pop()
			vm.jsMemberSet(target, vm.constants[nameIdx].Str, value)

		case OpJSCall:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			callee := vm.pop()
			result := vm.jsCall(callee, Value{Type: VTJSUndefined}, args)
			if len(vm.jsCallStack) == 0 || result.Type != VTJSUndefined {
				vm.push(result)
			}

		case OpJSCallMember:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			target := vm.pop()
			member := vm.constants[nameIdx].Str
			if result, handled := vm.jsCallMember(target, member, args); handled {
				if len(vm.jsCallStack) == 0 || result.Type != VTJSUndefined {
					vm.push(result)
				}
				continue
			}
			switch target.Type {
			case VTNativeObject:
				vm.push(vm.dispatchNativeCall(target.Num, member, args))
			case VTJSFunction:
				switch {
				case strings.EqualFold(member, "call"):
					callThis := Value{Type: VTJSUndefined}
					callArgs := args[:0]
					if len(args) > 0 {
						callThis = args[0]
						callArgs = args[1:]
					}
					result := vm.jsCall(target, callThis, callArgs)
					if len(vm.jsCallStack) == 0 || result.Type != VTJSUndefined {
						vm.push(result)
					}
				case strings.EqualFold(member, "apply"):
					applyThis := Value{Type: VTJSUndefined}
					var applyArgs []Value
					if len(args) > 0 {
						applyThis = args[0]
					}
					if len(args) > 1 {
						applyArgs = vm.jsExtractApplyArgs(args[1])
					}
					result := vm.jsCall(target, applyThis, applyArgs)
					if len(vm.jsCallStack) == 0 || result.Type != VTJSUndefined {
						vm.push(result)
					}
				case strings.EqualFold(member, "bind"):
					bindThis := Value{Type: VTJSUndefined}
					bindArgs := args[:0]
					if len(args) > 0 {
						bindThis = args[0]
						bindArgs = args[1:]
					}
					vm.push(vm.jsBindFunction(target, bindThis, bindArgs))
				default:
					callee, deferred := vm.jsMemberGet(target, member)
					if deferred {
						continue
					}
					result := vm.jsCall(callee, target, args)
					if len(vm.jsCallStack) == 0 || result.Type != VTJSUndefined {
						vm.push(result)
					}
				}
			case VTJSObject, VTArray, VTString, VTDate:
				callee, deferred := vm.jsMemberGet(target, member)
				if deferred {
					continue
				}
				result := vm.jsCall(callee, target, args)
				if len(vm.jsCallStack) == 0 || result.Type != VTJSUndefined {
					vm.push(result)
				}
			default:
				vm.push(Value{Type: VTJSUndefined})
			}

		case OpJSCreateClosure:
			templateIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.jsCreateClosure(vm.constants[templateIdx]))

		case OpJSAdd:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.jsAddValues(a, b))

		case OpJSStrictEq:
			b, a := vm.pop(), vm.pop()
			vm.push(NewBool(vm.jsStrictEquals(a, b)))

		case OpJSStrictNeq:
			b, a := vm.pop(), vm.pop()
			vm.push(NewBool(!vm.jsStrictEquals(a, b)))

		case OpJSTryEnter:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			vm.jsTryStack = append(vm.jsTryStack, target)

		case OpJSTryLeave:
			if len(vm.jsTryStack) > 0 {
				vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
			}

		case OpJSThrow:
			vm.jsThrow(vm.pop())

		case OpJSNewObject:
			objID := vm.allocJSID()
			vm.jsObjectItems[objID] = make(map[string]Value, 8)
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
			vm.push(Value{Type: VTJSObject, Num: objID})

		case OpJSNewArray:
			count := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			values := make([]Value, count)
			for i := count - 1; i >= 0; i-- {
				values[i] = vm.pop()
			}
			vm.push(ValueFromVBArray(NewVBArrayFromValues(0, values)))

		case OpJSTypeof:
			vm.push(NewString(vm.jsTypeOf(vm.pop())))

		case OpJSInstanceOf:
			right := vm.pop()
			left := vm.pop()
			vm.push(NewBool(left.Type == VTJSObject && right.Type == VTJSFunction))

		case OpJSIn:
			right := vm.pop()
			left := vm.pop()
			if right.Type == VTJSObject {
				key := vm.valueToString(left)
				if obj, ok := vm.jsObjectItems[right.Num]; ok {
					_, exists := obj[key]
					vm.push(NewBool(exists))
				} else {
					vm.push(NewBool(false))
				}
			} else {
				vm.push(NewBool(false))
			}

		case OpJSDelete:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			target := vm.pop()
			if target.Type == VTJSObject {
				if obj, ok := vm.jsObjectItems[target.Num]; ok {
					delete(obj, vm.constants[nameIdx].Str)
				}
			}
			vm.push(NewBool(true))

		case OpJSReturn:
			vm.jsReturn(vm.pop())

		case OpJSLoadUndefined:
			vm.push(Value{Type: VTJSUndefined})

		case OpJSLoadThis:
			vm.push(vm.jsThisValue)

		case OpJSDup:
			v := vm.pop()
			vm.push(v)
			vm.push(v)

		case OpJSPop:
			vm.pop()

		case OpJSJump:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			if target < vm.ip {
				jsBackJumpCount++
				if jsBackJumpCount > jsBackJumpLimit {
					return NewAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit), vm.sourceName, vm.lastLine)
				}
			}
			vm.ip = target

		case OpJSJumpIfFalse:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			if !vm.jsTruthy(vm.pop()) {
				if target < vm.ip {
					jsBackJumpCount++
					if jsBackJumpCount > jsBackJumpLimit {
						return NewAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit), vm.sourceName, vm.lastLine)
					}
				}
				vm.ip = target
			}

		case OpJSJumpIfTrue:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			if vm.jsTruthy(vm.pop()) {
				if target < vm.ip {
					jsBackJumpCount++
					if jsBackJumpCount > jsBackJumpLimit {
						return NewAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit), vm.sourceName, vm.lastLine)
					}
				}
				vm.ip = target
			}

		case OpJSLoadCatchError:
			if len(vm.jsErrStack) == 0 {
				vm.push(Value{Type: VTJSUndefined})
				continue
			}
			vm.push(vm.jsErrStack[len(vm.jsErrStack)-1])

		case OpJSStoreCatchError:
			v := vm.pop()
			vm.jsErrStack = append(vm.jsErrStack, v)

		// Control flow opcodes for loops and switches
		case OpJSBreak:
			// Break statement - jump target was patched by compiler
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip = target

		case OpJSContinue:
			// Continue statement - jump target was patched by compiler
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip = target

		case OpJSForInCleanup:
			forInPos := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			delete(vm.jsForInItems, forInPos)
			if vm.sp >= 0 {
				vm.pop()
			}

		case OpJSForIn:
			enumNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			exitTarget := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4

			opPos := vm.ip - 7
			enumState := vm.jsForInItems[opPos]
			if enumState == nil {
				if vm.sp < 0 {
					vm.ip = exitTarget
					continue
				}
				source := vm.stack[vm.sp]
				keys := vm.jsEnumerateForInKeys(source)
				enumState = &jsForInEnumerator{keys: keys, index: 0}
				vm.jsForInItems[opPos] = enumState
			}

			if enumState.index >= len(enumState.keys) {
				delete(vm.jsForInItems, opPos)
				if vm.sp >= 0 {
					vm.pop()
				}
				vm.ip = exitTarget
				continue
			}

			vm.jsSetName(vm.constants[enumNameIdx].Str, NewString(enumState.keys[enumState.index]))
			enumState.index++

		case OpJSSwitch:
			// Switch statement - discriminant value on stack, cases follow
			// Placeholder for now
			vm.pop()

		case OpJSCase:
			// Case label - test value matches, fall through or jump
			vm.ip += 4 // Skip jump target

		case OpJSDefault:
			// Default label - unconditional jump target
			vm.ip += 4 // Skip jump target

		// Arithmetic operators for JScript (type coercion)
		case OpJSSubtract:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsSubtract(left, right)
			vm.push(result)

		case OpJSMultiply:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsMultiply(left, right)
			vm.push(result)

		case OpJSDivide:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsDivide(left, right)
			vm.push(result)

		case OpJSModulo:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsModulo(left, right)
			vm.push(result)

		case OpJSNegate:
			val := vm.pop()
			result := vm.jsNegate(val)
			vm.push(result)

		// Bitwise operators
		case OpJSBitwiseAnd:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsBitwiseAnd(left, right)
			vm.push(result)

		case OpJSBitwiseOr:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsBitwiseOr(left, right)
			vm.push(result)

		case OpJSBitwiseXor:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsBitwiseXor(left, right)
			vm.push(result)

		case OpJSBitwiseNot:
			val := vm.pop()
			result := vm.jsBitwiseNot(val)
			vm.push(result)

		case OpJSLeftShift:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsLeftShift(left, right)
			vm.push(result)

		case OpJSRightShift:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsRightShift(left, right)
			vm.push(result)

		case OpJSUnsignedRightShift:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsUnsignedRightShift(left, right)
			vm.push(result)

		// Comparison operators
		case OpJSLess:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsLess(left, right)
			vm.push(result)

		case OpJSGreater:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsGreater(left, right)
			vm.push(result)

		case OpJSLessEqual:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsLessEqual(left, right)
			vm.push(result)

		case OpJSGreaterEqual:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsGreaterEqual(left, right)
			vm.push(result)

		case OpJSLooseEqual:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsLooseEqual(left, right)
			vm.push(result)

		case OpJSLooseNotEqual:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsLooseNotEqual(left, right)
			vm.push(result)

		// Logical operators
		case OpJSLogicalAnd:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsLogicalAnd(left, right)
			vm.push(result)

		case OpJSLogicalOr:
			right := vm.pop()
			left := vm.pop()
			result := vm.jsLogicalOr(left, right)
			vm.push(result)

		case OpJSLogicalNot:
			val := vm.pop()
			result := vm.jsLogicalNot(val)
			vm.push(result)

		// Pre/Post increment/decrement operators
		case OpJSPostIncrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			oldVal := vm.jsGetName(nameStr)
			newVal := vm.jsToNumber(oldVal)
			newVal.Flt++
			vm.jsSetName(nameStr, newVal)
			vm.push(oldVal)

		case OpJSPostDecrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			oldVal := vm.jsGetName(nameStr)
			newVal := vm.jsToNumber(oldVal)
			newVal.Flt--
			vm.jsSetName(nameStr, newVal)
			vm.push(oldVal)

		case OpJSPreIncrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			oldVal := vm.jsGetName(nameStr)
			newVal := vm.jsToNumber(oldVal)
			newVal.Flt++
			vm.jsSetName(nameStr, newVal)
			vm.push(newVal)

		case OpJSPreDecrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			oldVal := vm.jsGetName(nameStr)
			newVal := vm.jsToNumber(oldVal)
			newVal.Flt--
			vm.jsSetName(nameStr, newVal)
			vm.push(newVal)

		case OpJSPostMemberIncrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			target := vm.pop()
			vm.push(vm.jsUpdateMember(target, vm.constants[nameIdx].Str, 1, true))

		case OpJSPostMemberDecrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			target := vm.pop()
			vm.push(vm.jsUpdateMember(target, vm.constants[nameIdx].Str, -1, true))

		case OpJSPreMemberIncrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			target := vm.pop()
			vm.push(vm.jsUpdateMember(target, vm.constants[nameIdx].Str, 1, false))

		case OpJSPreMemberDecrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			target := vm.pop()
			vm.push(vm.jsUpdateMember(target, vm.constants[nameIdx].Str, -1, false))

		case OpJSPostIndexIncrement:
			indexVal := vm.pop()
			target := vm.pop()
			vm.push(vm.jsUpdateIndex(target, indexVal, 1, true))

		case OpJSPostIndexDecrement:
			indexVal := vm.pop()
			target := vm.pop()
			vm.push(vm.jsUpdateIndex(target, indexVal, -1, true))

		case OpJSPreIndexIncrement:
			indexVal := vm.pop()
			target := vm.pop()
			vm.push(vm.jsUpdateIndex(target, indexVal, 1, false))

		case OpJSPreIndexDecrement:
			indexVal := vm.pop()
			target := vm.pop()
			vm.push(vm.jsUpdateIndex(target, indexVal, -1, false))

		// Compound assignment operators
		case OpJSAddAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			right := vm.pop()
			left := vm.jsGetName(nameStr)
			result := vm.jsAdd(left, right)
			vm.jsSetName(nameStr, result)
			vm.push(result)

		case OpJSSubtractAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			right := vm.pop()
			left := vm.jsGetName(nameStr)
			result := vm.jsSubtract(left, right)
			vm.jsSetName(nameStr, result)
			vm.push(result)

		case OpJSMultiplyAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			right := vm.pop()
			left := vm.jsGetName(nameStr)
			result := vm.jsMultiply(left, right)
			vm.jsSetName(nameStr, result)
			vm.push(result)

		case OpJSDivideAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			right := vm.pop()
			left := vm.jsGetName(nameStr)
			result := vm.jsDivide(left, right)
			vm.jsSetName(nameStr, result)
			vm.push(result)

		case OpJSModuloAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			right := vm.pop()
			left := vm.jsGetName(nameStr)
			result := vm.jsModulo(left, right)
			vm.jsSetName(nameStr, result)
			vm.push(result)

		case OpJSMemberIndexGet:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			indexVal := vm.pop()
			objVal := vm.pop()
			result := vm.jsMemberIndexGet(objVal, indexVal, vm.constants[memberIdx].Str)
			vm.push(result)

		case OpJSMemberIndexSet:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			indexVal := vm.pop()
			objVal := vm.pop()
			value := vm.pop()
			vm.jsMemberIndexSet(objVal, indexVal, value, vm.constants[memberIdx].Str)

		case OpJSComma:
			// Comma operator: evaluate left, discard, evaluate right, keep on stack
			vm.pop() // Discard left

		case OpJSNew:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			constructor := vm.pop()
			result := vm.jsNew(constructor, args)
			if len(vm.jsCallStack) == 0 || result.Type != VTJSUndefined {
				vm.push(result)
			}

		case OpJSMemberDelete:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			obj := vm.pop()
			success := vm.jsMemberDelete(obj, vm.constants[memberIdx].Str)
			vm.push(NewBool(success))

		case OpJSIndexGet:
			indexVal := vm.pop()
			arrVal := vm.pop()
			result := vm.jsIndexGet(arrVal, indexVal)
			vm.push(result)

		case OpJSIndexSet:
			indexVal := vm.pop()
			arrVal := vm.pop()
			value := vm.pop()
			vm.jsIndexSet(arrVal, indexVal, value)

		case OpRet:
			retVal := vm.pop()
			if len(vm.callStack) == 0 {
				vm.push(retVal)
				break
			}
			frame := vm.callStack[len(vm.callStack)-1]
			vm.callStack = vm.callStack[:len(vm.callStack)-1]

			// Decrement reference counts for all local variables going out of scope.
			for i := vm.fp; i <= vm.sp; i++ {
				vm.decrementObjectRefCount(vm.stack[i])
			}

			// ByRef write-back: read callee's param values before restoring fp/sp.
			if len(frame.byRefs) > 0 {
				for _, wb := range frame.byRefs {
					calleeSlot := vm.fp + wb.calleeParamIdx
					if calleeSlot >= 0 && calleeSlot < StackSize {
						writeVal := vm.stack[calleeSlot]
						if wb.isClassMember {
							vm.setClassMemberValueByObjectID(wb.callerBoundObj, wb.callerMember, writeVal)
						} else if wb.isGlobal {
							if wb.callerIdx >= 0 && wb.callerIdx < len(vm.Globals) {
								vm.Globals[wb.callerIdx] = writeVal
							}
						} else {
							callerSlot := frame.oldFP + wb.callerIdx
							if callerSlot >= 0 && callerSlot < StackSize {
								vm.stack[callerSlot] = writeVal
							}
						}
					}
				}
			}
			vm.sp = frame.oldSP
			vm.fp = frame.oldFP
			vm.ip = frame.returnIP
			vm.activeClassObjectID = frame.boundObj
			vm.onResumeNext = frame.savedOnResumeNext
			vm.skipToNextStmt = frame.savedSkipToNextStmt
			vm.stmtSP = frame.savedStmtSP
			if vm.executeGlobalResumeGuard && len(vm.callStack) == 0 {
				vm.onResumeNext = true
			}
			if !frame.discard {
				vm.push(retVal)
			}

		default:
			// Unknown opcode: raise a runtime error instead of silently advancing past bytecode.
			// This surfaces bugs in the compiler emitting opcodes the VM does not handle.
			vm.raise(vbscript.InternalError, fmt.Sprintf("unknown opcode %d (%s) at ip=%d", byte(op), op.String(), vm.ip-1))
		}
	}

	// After the main script ends, fire the ObjectContext transaction event handler if needed.
	// prepareTransactionEventCall sets up the call frame and resets transactionState so we
	// only enter the loop once per SetComplete/SetAbort call.
	if vm.prepareTransactionEventCall() {
		goto aspExecLoop
	}

	if !vm.suppressTerminate && vm.prepareClassTerminateCall() {
		goto aspExecLoop
	}

	if vm.host != nil && vm.host.Response() != nil {
		vm.host.Response().Flush()
	}

	return nil
}

// prepareTransactionEventCall sets up a call frame for the ObjectContext transaction event handler.
// It returns true if a handler was found and vm.ip was updated to its entry point, false otherwise.
// Call this after the main script loop; if true, re-run the loop label to execute the handler.
func (vm *VM) prepareTransactionEventCall() bool {
	if vm.transactionState == 0 {
		return false
	}

	handlerGlobalIdx := objectContextCommitHandlerIdx
	if vm.transactionState == 2 {
		handlerGlobalIdx = objectContextAbortHandlerIdx
	}
	// Reset transaction state so we don't re-enter after the handler exits.
	vm.transactionState = 0

	if handlerGlobalIdx >= len(vm.Globals) {
		return false
	}
	handler := vm.Globals[handlerGlobalIdx]
	if handler.Type != VTUserSub {
		return false
	}

	// Set up a call frame that returns past the end of bytecode when OpRet fires.
	vm.callStack = append(vm.callStack, CallFrame{
		returnIP: len(vm.bytecode),
		oldFP:    vm.fp,
		oldSP:    vm.sp,
		boundObj: vm.activeClassObjectID,
	})
	vm.fp = vm.sp + 1
	localCount := handler.UserSubLocalCount()
	for i := 0; i < localCount; i++ {
		vm.stack[vm.fp+i] = Value{Type: VTEmpty}
	}
	if localCount > 0 {
		vm.sp = vm.fp + localCount - 1
	}
	vm.ip = int(handler.Num)
	return true
}

// prepareClassTerminateCall sets up one Class_Terminate invocation in reverse construction order.
// It gracefully skips instances that have already been terminated via reference counting,
// preventing null pointer crashes for objects explicitly set to Nothing.
func (vm *VM) prepareClassTerminateCall() bool {
	if len(vm.runtimeClassItems) == 0 {
		return false
	}

	if !vm.terminatePrepared {
		vm.terminateCursor = len(vm.classInstanceOrder) - 1
		vm.terminatePrepared = true
	}

	for vm.terminateCursor >= 0 {
		instanceID := vm.classInstanceOrder[vm.terminateCursor]
		vm.terminateCursor--

		instance, exists := vm.runtimeClassItems[instanceID]
		if !exists || instance == nil {
			continue
		}

		// Skip already-terminated instances (via reference counting).
		if instance.terminated {
			continue
		}

		target, ok := vm.resolveRuntimeClassMethod(
			Value{Type: VTObject, Num: instanceID, Str: instance.ClassName},
			"Class_Terminate",
			false,
		)
		if !ok {
			instance.terminated = true
			continue
		}

		if target.UserSubParamCount() != 0 {
			vm.raise(vbscript.ClassInitializeOrTerminateDoNotHaveArguments, "Class_Terminate must not declare arguments")
			instance.terminated = true
			return false
		}

		if vm.beginUserSubCall(target, nil, true, instanceID) {
			return true
		}

		instance.terminated = true
	}

	vm.runtimeClassItems = make(map[int64]*RuntimeClassInstance)
	vm.classInstanceOrder = vm.classInstanceOrder[:0]
	vm.terminatePrepared = false
	vm.terminateCursor = -1
	return false
}

// readArrayElement reads a VBArray element using one or more indices and returns Empty on invalid targets.
func (vm *VM) readArrayElement(target Value, indexes []Value) Value {
	if target.Type != VTArray || target.Arr == nil {
		vm.raise(vbscript.TypeMismatch, "Type mismatch")
		return Value{Type: VTEmpty}
	}
	if len(indexes) == 0 {
		vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
		return Value{Type: VTEmpty}
	}

	current := target.Arr
	for i := 0; i < len(indexes); i++ {
		idx := vm.asInt(indexes[i])
		if idx < current.Lower || idx > current.Upper() {
			vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
			return Value{Type: VTEmpty}
		}

		offset := idx - current.Lower
		value := current.Values[offset]
		if i == len(indexes)-1 {
			return value
		}

		if value.Type != VTArray || value.Arr == nil {
			vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
			return Value{Type: VTEmpty}
		}
		current = value.Arr
	}

	return Value{Type: VTEmpty}
}

// assignArrayElement writes a value into a VBArray using one or more indices.
func (vm *VM) assignArrayElement(target Value, indexes []Value, assigned Value) {
	if target.Type != VTArray || target.Arr == nil {
		vm.raise(vbscript.TypeMismatch, "Type mismatch")
		return
	}
	if len(indexes) == 0 {
		vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
		return
	}

	current := target.Arr
	for i := 0; i < len(indexes)-1; i++ {
		idx := vm.asInt(indexes[i])
		if idx < current.Lower || idx > current.Upper() {
			vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
			return
		}

		offset := idx - current.Lower
		next := current.Values[offset]
		if next.Type != VTArray || next.Arr == nil {
			vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
			return
		}
		current = next.Arr
	}

	lastIndex := vm.asInt(indexes[len(indexes)-1])
	if lastIndex < current.Lower || lastIndex > current.Upper() {
		vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
		return
	}

	// Decrement reference count of previous value in this array slot.
	prevVal := current.Values[lastIndex-current.Lower]
	vm.decrementObjectRefCount(prevVal)
	// Assign new value.
	current.Values[lastIndex-current.Lower] = assigned
	// Increment reference count of new value.
	vm.incrementObjectRefCount(assigned)
}

// eraseValue applies VBScript Erase behavior to one runtime slot value.
func (vm *VM) eraseValue(current Value) Value {
	if current.Type == VTArray && current.Arr != nil {
		vm.releaseArrayValue(current.Arr)
		return ValueFromVBArray(vm.cloneErasedArray(current.Arr))
	}
	vm.decrementObjectRefCount(current)
	return NewEmpty()
}

// releaseArrayValue decrements object references stored within one VBArray shape.
func (vm *VM) releaseArrayValue(arr *VBArray) {
	if vm == nil || arr == nil {
		return
	}
	for i := range arr.Values {
		value := arr.Values[i]
		if value.Type == VTArray && value.Arr != nil {
			vm.releaseArrayValue(value.Arr)
			continue
		}
		vm.decrementObjectRefCount(value)
	}
}

// cloneErasedArray rebuilds one VBArray with the same bounds but Empty elements.
func (vm *VM) cloneErasedArray(arr *VBArray) *VBArray {
	if arr == nil {
		return NewVBArrayFromValues(0, nil)
	}
	values := make([]Value, len(arr.Values))
	for i := range arr.Values {
		if arr.Values[i].Type == VTArray && arr.Values[i].Arr != nil {
			values[i] = ValueFromVBArray(vm.cloneErasedArray(arr.Values[i].Arr))
			continue
		}
		values[i] = NewEmpty()
	}
	return NewVBArrayFromValues(arr.Lower, values)
}

func (vm *VM) dispatchNativeCall(objID int64, member string, args []Value) Value {
	if proxy, exists := vm.nativeObjectProxies[objID]; exists {
		// Parameterized property write via OpArraySet or method call.
		// Re-dispatch to the parent object with the captured member name and combining the index + args.
		combinedArgs := vm.ensureCombineBuffer(len(proxy.CallArgs) + len(args))[:0]
		combinedArgs = append(combinedArgs, proxy.CallArgs...)
		combinedArgs = append(combinedArgs, args...)

		if member == "" {
			return vm.dispatchNativeCall(proxy.ParentID, proxy.Member, combinedArgs)
		}
		// Parameterized property access on a proxy?
		return vm.dispatchNativeCall(proxy.ParentID, member, combinedArgs)
	}

	if vm.host == nil {
		vm.raise(vbscript.InternalError, "Host not initialized")
	}

	if g3testObject, exists := vm.g3testItems[objID]; exists {
		return g3testObject.DispatchMethod(member, args)
	}

	if g3mdObject, exists := vm.g3mdItems[objID]; exists {
		return g3mdObject.DispatchMethod(member, args)
	}

	if g3cryptoObject, exists := vm.g3cryptoItems[objID]; exists {
		return g3cryptoObject.DispatchMethod(member, args)
	}

	if g3jsonObject, exists := vm.g3jsonItems[objID]; exists {
		return g3jsonObject.DispatchMethod(member, args)
	}

	if g3httpObject, exists := vm.g3httpItems[objID]; exists {
		return g3httpObject.DispatchMethod(member, args)
	}

	if g3mailObject, exists := vm.g3mailItems[objID]; exists {
		return g3mailObject.DispatchMethod(member, args)
	}

	if g3imageObject, exists := vm.g3imageItems[objID]; exists {
		return g3imageObject.DispatchMethod(member, args)
	}

	if g3filesObject, exists := vm.g3filesItems[objID]; exists {
		return g3filesObject.DispatchMethod(member, args)
	}
	if g3templateObject, exists := vm.g3templateItems[objID]; exists {
		return g3templateObject.DispatchMethod(member, args)
	}
	if g3zipObject, exists := vm.g3zipItems[objID]; exists {
		return g3zipObject.DispatchMethod(member, args)
	}
	if g3zlibObject, exists := vm.g3zlibItems[objID]; exists {
		return g3zlibObject.DispatchMethod(member, args)
	}
	if g3tarObject, exists := vm.g3tarItems[objID]; exists {
		return g3tarObject.DispatchMethod(member, args)
	}
	if g3zstdObject, exists := vm.g3zstdItems[objID]; exists {
		return g3zstdObject.DispatchMethod(member, args)
	}
	if g3fcObject, exists := vm.g3fcItems[objID]; exists {
		return g3fcObject.DispatchMethod(member, args)
	}
	if g3dbObject, exists := vm.g3dbItems[objID]; exists {
		return g3dbObject.DispatchMethod(member, args)
	}
	if g3dbRS, exists := vm.g3dbResultSetItems[objID]; exists {
		return g3dbRS.DispatchMethod(member, args)
	}
	if g3dbFields, exists := vm.g3dbFieldsItems[objID]; exists {
		return g3dbFields.DispatchMethod(member, args)
	}
	if g3dbRow, exists := vm.g3dbRowItems[objID]; exists {
		return g3dbRow.DispatchMethod(member, args)
	}
	if g3dbStmt, exists := vm.g3dbStatementItems[objID]; exists {
		return g3dbStmt.DispatchMethod(member, args)
	}
	if g3dbTx, exists := vm.g3dbTransactionItems[objID]; exists {
		return g3dbTx.DispatchMethod(member, args)
	}
	if g3dbResult, exists := vm.g3dbResultItems[objID]; exists {
		return g3dbResult.DispatchMethod(member, args)
	}
	if pdfObj, exists := vm.pdfItems[objID]; exists {
		return pdfObj.DispatchMethod(member, args)
	}
	if wscriptShellObject, exists := vm.wscriptShellItems[objID]; exists {
		return wscriptShellObject.DispatchMethod(member, args)
	}
	if wscriptExecObject, exists := vm.wscriptExecItems[objID]; exists {
		return wscriptExecObject.DispatchMethod(member, args)
	}
	if processStreamObject, exists := vm.wscriptProcessStreamItems[objID]; exists {
		return processStreamObject.DispatchMethod(member, args)
	}

	if adoxCatalogObject, exists := vm.adoxCatalogItems[objID]; exists {
		return adoxCatalogObject.DispatchMethod(member, args)
	}
	if adoxTablesObject, exists := vm.adoxTablesItems[objID]; exists {
		return adoxTablesObject.DispatchMethod(member, args)
	}
	if adoxTableObject, exists := vm.adoxTableItems[objID]; exists {
		return adoxTableObject.DispatchMethod(member, args)
	}

	if mswcAdRotator, exists := vm.mswcAdRotatorItems[objID]; exists {
		return mswcAdRotator.DispatchMethod(member, args)
	}
	if mswcBrowserType, exists := vm.mswcBrowserTypeItems[objID]; exists {
		return mswcBrowserType.DispatchMethod(member, args)
	}
	if mswcNextLink, exists := vm.mswcNextLinkItems[objID]; exists {
		return mswcNextLink.DispatchMethod(member, args)
	}
	if mswcContentRotator, exists := vm.mswcContentRotatorItems[objID]; exists {
		return mswcContentRotator.DispatchMethod(member, args)
	}
	if mswcCounters, exists := vm.mswcCountersItems[objID]; exists {
		return mswcCounters.DispatchMethod(member, args)
	}
	if mswcPageCounter, exists := vm.mswcPageCounterItems[objID]; exists {
		return mswcPageCounter.DispatchMethod(member, args)
	}
	if mswcTools, exists := vm.mswcToolsItems[objID]; exists {
		return mswcTools.DispatchMethod(member, args)
	}
	if mswcMyInfo, exists := vm.mswcMyInfoItems[objID]; exists {
		return mswcMyInfo.DispatchMethod(member, args)
	}
	if mswcPermissionChecker, exists := vm.mswcPermissionCheckerItems[objID]; exists {
		return mswcPermissionChecker.DispatchMethod(member, args)
	}
	if msxmlServer, exists := vm.msxmlServerItems[objID]; exists {
		return msxmlServer.DispatchMethod(member, args)
	}
	if msxmlDOM, exists := vm.msxmlDOMItems[objID]; exists {
		return msxmlDOM.DispatchMethod(member, args)
	}
	if msxmlNodeList, exists := vm.msxmlNodeListItems[objID]; exists {
		return msxmlNodeList.DispatchMethod(member, args)
	}
	if msxmlParseError, exists := vm.msxmlParseErrorItems[objID]; exists {
		return msxmlParseError.DispatchMethod(member, args)
	}
	if msxmlElement, exists := vm.msxmlElementItems[objID]; exists {
		return msxmlElement.DispatchMethod(member, args)
	}
	if pdf, exists := vm.pdfItems[objID]; exists {
		return pdf.DispatchMethod(member, args)
	}
	if fileUploader, exists := vm.fileUploaderItems[objID]; exists {
		return fileUploader.DispatchMethod(member, args)
	}

	if axonObject, exists := vm.axonItems[objID]; exists {
		return axonObject.DispatchMethod(member, args)
	}

	if dictResult, handled := vm.dispatchDictionaryMethod(objID, member, args); handled {
		return dictResult
	}

	if fsoResult, handled := vm.dispatchFSOMethod(objID, member, args); handled {
		return fsoResult
	}

	if adodbResult, handled := vm.dispatchADODBStreamMethod(objID, member, args); handled {
		return adodbResult
	}

	if adodbResult, handled := vm.dispatchADODBMethod(objID, member, args); handled {
		return adodbResult
	}

	if res, handled := vm.dispatchADODBErrorsCollectionMethod(objID, member, args); handled {
		return res
	}

	if adodbResult, handled := vm.dispatchADODBFieldsCollectionMethod(objID, member, args); handled {
		return adodbResult
	}

	if adodbResult, handled := vm.dispatchADODBParametersCollectionMethod(objID, member, args); handled {
		return adodbResult
	}

	if adodbResult, handled := vm.dispatchADODBFieldMethod(objID, member, args); handled {
		return adodbResult
	}

	if regExpResult, handled := vm.dispatchRegExpMethod(objID, member, args); handled {
		return regExpResult
	}

	if objID == nativeObjectErr {
		switch {
		case strings.EqualFold(member, "Clear"):
			vm.errClear()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Raise"):
			return vm.errRaise(args)
		}
		return Value{Type: VTEmpty}
	}

	// Route console.log / console.info / console.error / console.warn calls.
	if objID == nativeObjectConsole {
		return consoleDispatch(vm, member, args)
	}

	switch objID {
	case nativeObjectResponse: // Response
		response := vm.host.Response()
		switch {
		case strings.EqualFold(member, "Write"):
			if len(args) > 0 {
				response.Write(vm.valueToString(args[0]))
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "BinaryWrite"):
			if len(args) > 0 {
				response.BinaryWrite(vm.valueToBinaryBytes(args[0]))
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "AddHeader"):
			if len(args) >= 2 {
				response.AddHeader(args[0].String(), args[1].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "AppendToLog"):
			if len(args) > 0 {
				response.AppendToLog(args[0].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Clear"):
			response.Clear()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Flush"):
			response.Flush()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "End"):
			response.End()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Redirect"):
			if len(args) > 0 {
				response.Redirect(args[0].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Buffer"):
			if len(args) > 0 {
				response.SetBuffer(vm.asBool(args[0]))
				return Value{Type: VTEmpty}
			}
			return NewBool(response.GetBuffer())
		case strings.EqualFold(member, "CacheControl"):
			if len(args) > 0 {
				response.SetCacheControl(args[0].String())
				return Value{Type: VTEmpty}
			}
			return NewString(response.GetCacheControl())
		case strings.EqualFold(member, "Charset"):
			if len(args) > 0 {
				response.SetCharset(args[0].String())
				return Value{Type: VTEmpty}
			}
			return NewString(response.GetCharset())
		case strings.EqualFold(member, "ContentType"):
			if len(args) > 0 {
				response.SetContentType(args[0].String())
				return Value{Type: VTEmpty}
			}
			return NewString(response.GetContentType())
		case strings.EqualFold(member, "Expires"):
			if len(args) > 0 {
				response.SetExpires(vm.asInt(args[0]))
				return Value{Type: VTEmpty}
			}
			return NewInteger(int64(response.GetExpires()))
		case strings.EqualFold(member, "ExpiresAbsolute"):
			if len(args) > 0 {
				response.SetExpiresAbsoluteRaw(args[0].String())
				return Value{Type: VTEmpty}
			}
			return NewString(response.GetExpiresAbsoluteRaw())
		case strings.EqualFold(member, "PICS"):
			if len(args) > 0 {
				response.SetPICS(args[0].String())
				return Value{Type: VTEmpty}
			}
			return NewString(response.GetPICS())
		case strings.EqualFold(member, "Status"):
			if len(args) > 0 {
				response.SetStatus(args[0].String())
				return Value{Type: VTEmpty}
			}
			return NewString(response.GetStatus())
		case strings.EqualFold(member, "IsClientConnected"):
			return NewBool(response.IsClientConnected())
		case strings.EqualFold(member, "Cookies"):
			if len(args) == 1 {
				return vm.newResponseCookieItem(args[0].String())
			}
			if len(args) == 2 {
				propertyName := args[1].String()
				switch {
				case strings.EqualFold(propertyName, "Domain"), strings.EqualFold(propertyName, "Path"), strings.EqualFold(propertyName, "Expires"), strings.EqualFold(propertyName, "Secure"), strings.EqualFold(propertyName, "HttpOnly"), strings.EqualFold(propertyName, "Value"), strings.EqualFold(propertyName, "Name"):
					return NewString(response.GetCookieProperty(args[0].String(), propertyName))
				default:
					response.SetCookieValue(args[0].String(), propertyName)
					return Value{Type: VTEmpty}
				}
			}
			if len(args) >= 3 {
				response.SetCookieProperty(args[0].String(), args[1].String(), args[2].String())
				return Value{Type: VTEmpty}
			}
			return Value{Type: VTNativeObject, Num: nativeResponseCookies}
		case strings.EqualFold(member, "Cookies.Domain"):
			if len(args) >= 2 {
				response.SetCookieProperty(args[0].String(), "Domain", args[1].String())
				return Value{Type: VTEmpty}
			}
			if len(args) == 1 {
				return NewString(response.GetCookieProperty(args[0].String(), "Domain"))
			}
			return NewString("")
		case strings.EqualFold(member, "Cookies.Path"):
			if len(args) >= 2 {
				response.SetCookieProperty(args[0].String(), "Path", args[1].String())
				return Value{Type: VTEmpty}
			}
			if len(args) == 1 {
				return NewString(response.GetCookieProperty(args[0].String(), "Path"))
			}
			return NewString("")
		case strings.EqualFold(member, "Cookies.Expires"):
			if len(args) >= 2 {
				response.SetCookieProperty(args[0].String(), "Expires", args[1].String())
				return Value{Type: VTEmpty}
			}
			if len(args) == 1 {
				return NewString(response.GetCookieProperty(args[0].String(), "Expires"))
			}
			return NewString("")
		case strings.EqualFold(member, "Cookies.Secure"):
			if len(args) >= 2 {
				response.SetCookieProperty(args[0].String(), "Secure", args[1].String())
				return Value{Type: VTEmpty}
			}
			if len(args) == 1 {
				return NewString(response.GetCookieProperty(args[0].String(), "Secure"))
			}
			return NewString("")
		case strings.EqualFold(member, "Cookies.HttpOnly"):
			if len(args) >= 2 {
				response.SetCookieProperty(args[0].String(), "HttpOnly", args[1].String())
				return Value{Type: VTEmpty}
			}
			if len(args) == 1 {
				return NewString(response.GetCookieProperty(args[0].String(), "HttpOnly"))
			}
			return NewString("")
		}
	case nativeResponseCookies:
		if member == "" && len(args) >= 1 {
			return vm.newResponseCookieItem(args[0].String())
		}
		return Value{Type: VTEmpty}
	case nativeResponseCookiesDomainMethod:
		if member == "" && len(args) >= 2 {
			vm.host.Response().SetCookieProperty(args[0].String(), "Domain", args[1].String())
		}
		return Value{Type: VTEmpty}
	case nativeResponseCookiesPathMethod:
		if member == "" && len(args) >= 2 {
			vm.host.Response().SetCookieProperty(args[0].String(), "Path", args[1].String())
		}
		return Value{Type: VTEmpty}
	case nativeResponseCookiesExpiresMethod:
		if member == "" && len(args) >= 2 {
			vm.host.Response().SetCookieProperty(args[0].String(), "Expires", args[1].String())
		}
		return Value{Type: VTEmpty}
	case nativeResponseCookiesSecureMethod:
		if member == "" && len(args) >= 2 {
			vm.host.Response().SetCookieProperty(args[0].String(), "Secure", args[1].String())
		}
		return Value{Type: VTEmpty}
	case nativeResponseCookiesHttpOnlyMethod:
		if member == "" && len(args) >= 2 {
			vm.host.Response().SetCookieProperty(args[0].String(), "HttpOnly", args[1].String())
		}
		return Value{Type: VTEmpty}
	case nativeObjectRequest: // Request
		request := vm.host.Request()
		switch {
		case member == "":
			if len(args) >= 1 {
				return NewString(request.GetValue(args[0].String()))
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "QueryString"):
			if len(args) >= 1 {
				if val, ok := request.GetCollectionEntry("QueryString", args[0].String()); ok {
					return NewString(val)
				}
				return Value{Type: VTEmpty}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Form"):
			if len(args) >= 1 {
				if request.IsBinaryReadUsed() {
					return Value{Type: VTEmpty}
				}
				request.MarkFormUsed()
				if val, ok := request.GetCollectionEntry("Form", args[0].String()); ok {
					return NewString(val)
				}
				return Value{Type: VTEmpty}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Cookies"):
			if len(args) == 1 {
				return NewString(request.GetCollectionValue("Cookies", args[0].String()))
			}
			if len(args) >= 2 {
				return NewString(request.GetCookieAttribute(args[0].String(), args[1].String()))
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "ServerVariables"):
			if len(args) >= 1 {
				return NewString(request.GetCollectionValue("ServerVariables", args[0].String()))
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "ClientCertificate"):
			if len(args) >= 1 {
				return NewString(request.GetCollectionValue("ClientCertificate", args[0].String()))
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "TotalBytes"):
			return NewInteger(request.TotalBytes())
		case strings.EqualFold(member, "BinaryRead"):
			if len(args) >= 1 {
				if request.IsFormUsed() {
					return NewString("")
				}
				bytes := request.BinaryRead(int64(vm.asInt(args[0])))
				args[0] = NewInteger(int64(len(bytes)))
				return NewString(bytesToVBByteString(bytes))
			}
			return NewString("")
		case strings.EqualFold(member, "QueryString.Count"):
			return NewInteger(int64(request.QueryString.Count()))
		case strings.EqualFold(member, "Form.Count"):
			if request.IsBinaryReadUsed() {
				return NewInteger(0)
			}
			request.MarkFormUsed()
			return NewInteger(int64(request.Form.Count()))
		case strings.EqualFold(member, "Cookies.Count"):
			return NewInteger(int64(request.Cookies.Count()))
		case strings.EqualFold(member, "ServerVariables.Count"):
			return NewInteger(int64(request.ServerVars.Count()))
		case strings.EqualFold(member, "ClientCertificate.Count"):
			return NewInteger(int64(request.ClientCertificate.Count()))
		case strings.EqualFold(member, "QueryString.Key"):
			if len(args) >= 1 {
				return NewString(request.QueryString.Key(vm.asInt(args[0])))
			}
			return NewString("")
		case strings.EqualFold(member, "Form.Key"):
			if len(args) >= 1 {
				if request.IsBinaryReadUsed() {
					return NewString("")
				}
				request.MarkFormUsed()
				return NewString(request.Form.Key(vm.asInt(args[0])))
			}
			return NewString("")
		case strings.EqualFold(member, "Cookies.Key"):
			if len(args) >= 1 {
				return NewString(request.Cookies.Key(vm.asInt(args[0])))
			}
			return NewString("")
		case strings.EqualFold(member, "ServerVariables.Key"):
			if len(args) >= 1 {
				return NewString(request.ServerVars.Key(vm.asInt(args[0])))
			}
			return NewString("")
		case strings.EqualFold(member, "ClientCertificate.Key"):
			if len(args) >= 1 {
				return NewString(request.ClientCertificate.Key(vm.asInt(args[0])))
			}
			return NewString("")
		case strings.EqualFold(member, "Count"):
			if len(args) >= 1 {
				collection := request.GetCollection(args[0].String())
				if collection != nil {
					return NewInteger(int64(collection.Count()))
				}
			}
			return NewInteger(0)
		case strings.EqualFold(member, "Key"):
			if len(args) >= 2 {
				return NewString(request.GetCollectionProperty(args[0].String(), "Key", args[1].String()))
			}
			return NewString("")
		}
	case nativeRequestQueryString:
		if member == "" && len(args) >= 1 {
			if val, ok := vm.host.Request().GetCollectionEntry("QueryString", args[0].String()); ok {
				return NewString(val)
			}
			return Value{Type: VTEmpty}
		}
		if strings.EqualFold(member, "Count") {
			return NewInteger(int64(vm.host.Request().QueryString.Count()))
		}
		if strings.EqualFold(member, "Key") && len(args) >= 1 {
			return NewString(vm.host.Request().QueryString.Key(vm.asInt(args[0])))
		}
		return Value{Type: VTEmpty}
	case nativeRequestForm:
		if member == "" && len(args) >= 1 {
			if vm.host.Request().IsBinaryReadUsed() {
				return Value{Type: VTEmpty}
			}
			vm.host.Request().MarkFormUsed()
			if val, ok := vm.host.Request().GetCollectionEntry("Form", args[0].String()); ok {
				return NewString(val)
			}
			return Value{Type: VTEmpty}
		}
		return Value{Type: VTEmpty}
	case nativeRequestCookies:
		if member == "" && len(args) == 1 {
			return NewString(vm.host.Request().GetCollectionValue("Cookies", args[0].String()))
		}
		if member == "" && len(args) >= 2 {
			return NewString(vm.host.Request().GetCookieAttribute(args[0].String(), args[1].String()))
		}
		return Value{Type: VTEmpty}
	case nativeRequestServerVariables:
		if member == "" && len(args) >= 1 {
			return NewString(vm.host.Request().GetCollectionValue("ServerVariables", args[0].String()))
		}
		return Value{Type: VTEmpty}
	case nativeRequestClientCertificate:
		if member == "" && len(args) >= 1 {
			return NewString(vm.host.Request().GetCollectionValue("ClientCertificate", args[0].String()))
		}
		return Value{Type: VTEmpty}
	case nativeRequestBinaryReadMethod:
		if member == "" && len(args) >= 1 {
			if vm.host.Request().IsFormUsed() {
				return NewString("")
			}
			bytes := vm.host.Request().BinaryRead(int64(vm.asInt(args[0])))
			args[0] = NewInteger(int64(len(bytes)))
			return NewString(bytesToVBByteString(bytes))
		}
		return NewString("")
	case nativeRequestQueryStringKeyMethod:
		if member == "" && len(args) >= 1 {
			return NewString(vm.host.Request().QueryString.Key(vm.asInt(args[0])))
		}
		return NewString("")
	case nativeRequestFormKeyMethod:
		if member == "" && len(args) >= 1 {
			if vm.host.Request().IsBinaryReadUsed() {
				return NewString("")
			}
			vm.host.Request().MarkFormUsed()
			return NewString(vm.host.Request().Form.Key(vm.asInt(args[0])))
		}
		return NewString("")
	case nativeRequestCookiesKeyMethod:
		if member == "" && len(args) >= 1 {
			return NewString(vm.host.Request().Cookies.Key(vm.asInt(args[0])))
		}
		return NewString("")
	case nativeRequestServerVariablesKeyMethod:
		if member == "" && len(args) >= 1 {
			return NewString(vm.host.Request().ServerVars.Key(vm.asInt(args[0])))
		}
		return NewString("")
	case nativeRequestClientCertificateKeyMethod:
		if member == "" && len(args) >= 1 {
			return NewString(vm.host.Request().ClientCertificate.Key(vm.asInt(args[0])))
		}
		return NewString("")
	case nativeObjectServer: // Server
		server := vm.host.Server()
		switch {
		case strings.EqualFold(member, "HTMLEncode"):
			if len(args) >= 1 {
				return NewString(server.HTMLEncode(args[0].String()))
			}
			return NewString("")
		case strings.EqualFold(member, "URLEncode"):
			if len(args) >= 1 {
				return NewString(server.URLEncode(args[0].String()))
			}
			return NewString("")
		case strings.EqualFold(member, "URLPathEncode"):
			if len(args) >= 1 {
				return NewString(server.URLPathEncode(args[0].String()))
			}
			return NewString("")
		case strings.EqualFold(member, "MapPath"):
			if len(args) >= 1 {
				return NewString(server.MapPath(args[0].String()))
			}
			return NewString(server.MapPath(""))
		case strings.EqualFold(member, "IsClientConnected"):
			return NewBool(vm.host.Response().IsClientConnected())
		case strings.EqualFold(member, "ScriptTimeout"):
			if len(args) >= 1 {
				if err := server.SetScriptTimeout(vm.asInt(args[0])); err != nil {
					server.SetLastError(asp.NewVBScriptASPError(vbscript.InvalidProcedureCallOrArgument, "Server.ScriptTimeout", "ASP", err.Error(), "", 0, 0))
				}
				return Value{Type: VTEmpty}
			}
			return NewInteger(int64(server.GetScriptTimeout()))
		case strings.EqualFold(member, "CreateObject"):
			if len(args) >= 1 {
				progID := strings.TrimSpace(args[0].String())
				if strings.EqualFold(progID, "G3MD") {
					return vm.newG3MDObject()
				}
				if strings.EqualFold(progID, "G3TestSuite") || strings.EqualFold(progID, "G3Test") {
					return vm.newG3TestObject()
				}
				if defaultAlgorithm, ok := g3cryptoResolveProgID(progID); ok {
					return vm.newG3CryptoObject(defaultAlgorithm)
				}
				if strings.EqualFold(progID, "G3AXON") || strings.EqualFold(progID, "G3Axon.Functions") {
					return vm.newAxonLibrary()
				}
				if strings.EqualFold(progID, "G3JSON") {
					return vm.newG3JSONObject()
				}
				if strings.EqualFold(progID, "G3DB") {
					return vm.newG3DBObject()
				}
				if strings.EqualFold(progID, "G3HTTP") || strings.EqualFold(progID, "G3HTTP.Functions") {
					return vm.newG3HTTPObject()
				}
				if strings.EqualFold(progID, "G3Mail") || strings.EqualFold(progID, "CDONTS.NewMail") || strings.EqualFold(progID, "CDO.Message") || strings.EqualFold(progID, "Persits.MailSender") {
					return vm.newG3MailObject()
				}
				if strings.EqualFold(progID, "G3Image") {
					return vm.newG3ImageObject()
				}
				if strings.EqualFold(progID, "G3FILES") {
					return vm.newG3FilesObject()
				}
				if strings.EqualFold(progID, "G3Template") {
					return vm.newG3TemplateObject()
				}
				if strings.EqualFold(progID, "G3Zip") {
					return vm.newG3ZipObject()
				}
				if strings.EqualFold(progID, "G3ZLIB") {
					return vm.newG3ZLIBObject()
				}
				if strings.EqualFold(progID, "G3TAR") {
					return vm.newG3TARObject()
				}
				if strings.EqualFold(progID, "G3ZSTD") {
					return vm.newG3ZSTDObject()
				}
				if strings.EqualFold(progID, "G3FC") {
					return vm.newG3FCObject()
				}
				if strings.EqualFold(progID, "WScript.Shell") {
					return vm.newWScriptShellObject()
				}
				if strings.EqualFold(progID, "ADOX.Catalog") {
					return vm.newADOXCatalogObject()
				}
				if strings.EqualFold(progID, "MSWC.AdRotator") {
					return vm.newG3AdRotatorObject()
				}
				if strings.EqualFold(progID, "MSWC.BrowserType") {
					return vm.newG3BrowserTypeObject()
				}
				if strings.EqualFold(progID, "MSWC.NextLink") {
					return vm.newG3NextLinkObject()
				}
				if strings.EqualFold(progID, "MSWC.ContentRotator") {
					return vm.newG3ContentRotatorObject()
				}
				if strings.EqualFold(progID, "MSWC.Counters") {
					return vm.newG3CountersObject()
				}
				if strings.EqualFold(progID, "MSWC.PageCounter") {
					return vm.newG3PageCounterObject()
				}
				if strings.EqualFold(progID, "MSWC.Tools") {
					return vm.newG3ToolsObject()
				}
				if strings.EqualFold(progID, "MSWC.MyInfo") {
					return vm.newG3MyInfoObject()
				}
				if strings.EqualFold(progID, "MSWC.PermissionChecker") {
					return vm.newG3PermissionCheckerObject()
				}
				if strings.EqualFold(progID, "MSXML2.ServerXMLHTTP") || strings.EqualFold(progID, "MSXML2.XMLHTTP") || strings.EqualFold(progID, "Microsoft.XMLHTTP") {
					obj := NewMsXML2ServerXMLHTTP(vm)
					id := vm.nextDynamicNativeID
					vm.nextDynamicNativeID++
					vm.msxmlServerItems[id] = obj
					return Value{Type: VTNativeObject, Num: id}
				}
				if strings.EqualFold(progID, "MSXML2.DOMDocument") || strings.EqualFold(progID, "Microsoft.XMLDOM") {
					obj := NewMsXML2DOMDocument(vm)
					id := vm.nextDynamicNativeID
					vm.nextDynamicNativeID++
					vm.msxmlDOMItems[id] = obj
					return Value{Type: VTNativeObject, Num: id}
				}
				if strings.EqualFold(progID, "G3PDF") {
					obj := NewG3PDF(vm)
					id := vm.nextDynamicNativeID
					vm.nextDynamicNativeID++
					vm.pdfItems[id] = obj
					return Value{Type: VTNativeObject, Num: id}
				}
				if strings.EqualFold(progID, "G3FileUploader") || strings.EqualFold(progID, "Persits.Upload") || strings.EqualFold(progID, "SoftArtisans.FileUp") || strings.EqualFold(progID, "ASPUpload") {
					return vm.newG3FileUploaderObject()
				}
				if strings.EqualFold(progID, "Scripting.FileSystemObject") {
					return vm.newFSORootObject()
				}
				if strings.EqualFold(progID, "Scripting.Dictionary") {
					return vm.newDictionaryObject()
				}
				if strings.EqualFold(progID, "ADODB.Stream") {
					return vm.newADODBStreamObject()
				}
				if strings.EqualFold(progID, "ADODB.Connection") {
					return vm.newADODBConnection()
				}
				if strings.EqualFold(progID, "ADODBOLE.Connection") {
					return vm.newADODBOLEConnection()
				}
				if strings.EqualFold(progID, "ADODB.Recordset") {
					return vm.newADODBRecordset()
				}
				if strings.EqualFold(progID, "ADODB.Command") {
					return vm.newADODBCommand()
				}
				if strings.EqualFold(progID, "VBScript.RegExp") || strings.EqualFold(progID, "RegExp") {
					return vm.newRegExpObject()
				}
				value, err := server.CreateObject(args[0].String())
				if err != nil {
					server.SetLastError(asp.NewVBScriptASPError(vbscript.ActiveXCannotCreateObject, "Server.CreateObject", "ASP", err.Error(), "", 0, 0))
					return Value{Type: VTEmpty}
				}
				resolved := vm.applicationValueToValue(value)
				if resolved.Type == VTEmpty {
					server.SetLastError(asp.NewVBScriptASPError(vbscript.ActiveXCannotCreateObject, "Server.CreateObject", "ASP", vbscript.ActiveXCannotCreateObject.String(), "", 0, 0))
				}
				return resolved
			}
			_, _ = server.CreateObject("")
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "GetLastError"):
			errObj := server.GetLastError()
			if len(args) == 0 {
				return vm.newASPErrorObject(errObj)
			}
			return vm.aspErrorPropertyValue(errObj, args[0].String())
		case strings.EqualFold(member, "ClearLastError"):
			server.ClearLastError()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Execute"):
			// Server.Execute(path) — runs another ASP file inline, sharing the current host context.
			if len(args) >= 1 {
				absPath := server.MapPath(args[0].String())
				if err := vm.host.ExecuteASPFile(absPath); err != nil {
					server.SetLastError(asp.NewVBScriptASPError(vbscript.PathNotFound, "Server.Execute", "ASP", err.Error(), absPath, 0, 0))
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Transfer"):
			// Server.Transfer(path) — clears current response, executes another ASP file, and stops the calling script.
			if len(args) >= 1 {
				absPath := server.MapPath(args[0].String())
				vm.host.Response().Clear()
				_ = vm.host.ExecuteASPFile(absPath)
				panic(asp.ResponseEndSignal)
			}
			return Value{Type: VTEmpty}
		}
	case nativeObjectSession: // Session
		// Raise a runtime error when session state has been disabled by an ASP page directive.
		if !vm.host.SessionEnabled() {
			vm.raise(vbscript.PermissionDenied, "Session state is disabled for this application")
		}
		session := vm.host.Session()
		switch {
		case member == "":
			if len(args) >= 2 {
				session.Set(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				if sessionValue, ok := session.Get(args[0].String()); ok {
					return vm.applicationValueToValue(sessionValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Lock"):
			session.Lock()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Unlock"):
			session.Unlock()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Set"):
			if len(args) >= 2 {
				session.Set(args[0].String(), vm.valueToApplicationValue(args[1]))
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Get") || strings.EqualFold(member, "Item") || strings.EqualFold(member, "Contents"):
			if len(args) >= 1 {
				if sessionValue, ok := session.Get(args[0].String()); ok {
					return vm.applicationValueToValue(sessionValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Contents.Item") || strings.EqualFold(member, "Contents"):
			if len(args) >= 2 {
				session.Set(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				if sessionValue, ok := session.Get(args[0].String()); ok {
					return vm.applicationValueToValue(sessionValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Contents.Remove"):
			if len(args) >= 1 {
				session.Remove(args[0].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Contents.RemoveAll"):
			session.RemoveAll()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Contents.Count"):
			return NewInteger(int64(session.Count()))
		case strings.EqualFold(member, "Remove"):
			if len(args) >= 1 {
				session.Remove(args[0].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "RemoveAll"):
			session.RemoveAll()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Abandon"):
			session.Abandon()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(session.Count()))
		case strings.EqualFold(member, "Exists"):
			if len(args) >= 1 {
				return NewBool(session.Contains(args[0].String()))
			}
			return NewBool(false)
		case strings.EqualFold(member, "SessionID"):
			return NewInteger(session.SessionIDNumeric())
		case strings.EqualFold(member, "Timeout"):
			if len(args) >= 1 {
				session.SetTimeout(vm.asInt(args[0]))
				return Value{Type: VTEmpty}
			}
			return NewInteger(int64(session.GetTimeout()))
		case strings.EqualFold(member, "LCID"):
			if len(args) >= 1 {
				session.SetLCID(vm.asInt(args[0]))
				return Value{Type: VTEmpty}
			}
			return NewInteger(int64(session.GetLCID()))
		case strings.EqualFold(member, "CodePage"):
			if len(args) >= 1 {
				session.SetCodePage(vm.asInt(args[0]))
				return Value{Type: VTEmpty}
			}
			return NewInteger(int64(session.GetCodePage()))
		case strings.EqualFold(member, "IsLocked"):
			return NewBool(session.IsLocked())
		case strings.EqualFold(member, "GetLockCount"):
			return NewInteger(int64(session.GetLockCount()))
		case strings.EqualFold(member, "StaticObjects.Item") || strings.EqualFold(member, "StaticObjects"):
			if len(args) == 0 {
				return Value{Type: VTNativeObject, Num: nativeObjectSessionStaticObjects}
			}
			if len(args) >= 2 {
				session.AddStaticObject(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				if appVal, ok := session.GetStaticObject(args[0].String()); ok {
					return vm.applicationValueToValue(appVal)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "StaticObjects.Count"):
			return NewInteger(int64(len(session.GetStaticObjectsCopy())))
		}
	case nativeObjectApplication: // Application
		application := vm.host.Application()
		server := vm.host.Server()
		switch {
		case member == "":
			if len(args) >= 2 {
				application.WaitForServer(server)
				application.Set(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				application.WaitForServer(server)
				if appValue, ok := application.Get(args[0].String()); ok {
					return vm.applicationValueToValue(appValue)
				}
				if appValue, ok := application.GetStaticObject(args[0].String()); ok {
					return vm.applicationValueToValue(appValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Lock"):
			application.LockForServer(server)
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Unlock"):
			application.UnlockForServer(server)
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Set"):
			if len(args) >= 2 {
				application.WaitForServer(server)
				application.Set(args[0].String(), vm.valueToApplicationValue(args[1]))
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Get") || strings.EqualFold(member, "Item"):
			if len(args) >= 1 {
				application.WaitForServer(server)
				if appValue, ok := application.Get(args[0].String()); ok {
					return vm.applicationValueToValue(appValue)
				}
				if appValue, ok := application.GetStaticObject(args[0].String()); ok {
					return vm.applicationValueToValue(appValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Contents.Item") || strings.EqualFold(member, "Contents"):
			if len(args) >= 2 {
				application.WaitForServer(server)
				application.Set(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				application.WaitForServer(server)
				if appValue, ok := application.Get(args[0].String()); ok {
					return vm.applicationValueToValue(appValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Contents.Remove"):
			if len(args) >= 1 {
				application.WaitForServer(server)
				application.Remove(args[0].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Contents.RemoveAll"):
			application.WaitForServer(server)
			application.RemoveAll()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Contents.Count"):
			application.WaitForServer(server)
			return NewInteger(int64(len(application.GetContentsCopy())))
		case strings.EqualFold(member, "StaticObjects.Item") || strings.EqualFold(member, "StaticObjects"):
			if len(args) >= 2 {
				application.WaitForServer(server)
				application.AddStaticObject(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				application.WaitForServer(server)
				if appValue, ok := application.GetStaticObject(args[0].String()); ok {
					return vm.applicationValueToValue(appValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "StaticObjects.Count"):
			application.WaitForServer(server)
			return NewInteger(int64(len(application.GetStaticObjectsCopy())))
		case strings.EqualFold(member, "Remove"):
			if len(args) >= 1 {
				application.WaitForServer(server)
				application.Remove(args[0].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "RemoveAll"):
			application.WaitForServer(server)
			application.RemoveAll()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Count"):
			application.WaitForServer(server)
			return NewInteger(int64(application.Count()))
		case strings.EqualFold(member, "Exists"):
			if len(args) >= 1 {
				application.WaitForServer(server)
				return NewBool(application.ContainsContent(args[0].String()))
			}
			return NewBool(false)
		case strings.EqualFold(member, "IsLocked"):
			return NewBool(application.IsLocked())
		case strings.EqualFold(member, "GetLockCount"):
			return NewInteger(int64(application.GetLockCount()))
		}
	case nativeObjectSessionContents:
		session := vm.host.Session()
		switch {
		case member == "" || strings.EqualFold(member, "Item"):
			if len(args) >= 2 {
				session.Set(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				if sessionValue, ok := session.Get(args[0].String()); ok {
					return vm.applicationValueToValue(sessionValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Remove"):
			if len(args) >= 1 {
				session.Remove(args[0].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "RemoveAll"):
			session.RemoveAll()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(session.Count()))
		case strings.EqualFold(member, "Keys"):
			keys := session.GetAllKeys()
			sort.Strings(keys)
			values := make([]Value, len(keys))
			for i := 0; i < len(keys); i++ {
				values[i] = NewString(keys[i])
			}
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
		case strings.EqualFold(member, "Items"):
			keys := session.GetAllKeys()
			sort.Strings(keys)
			values := make([]Value, 0, len(keys))
			for i := 0; i < len(keys); i++ {
				if sessionValue, ok := session.Get(keys[i]); ok {
					values = append(values, vm.applicationValueToValue(sessionValue))
				}
			}
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
		}
	case nativeObjectSessionStaticObjects:
		session := vm.host.Session()
		switch {
		case member == "" || strings.EqualFold(member, "Item"):
			if len(args) >= 2 {
				session.AddStaticObject(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				if appVal, ok := session.GetStaticObject(args[0].String()); ok {
					return vm.applicationValueToValue(appVal)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(session.GetStaticObjectsCopy())))
		}
	case nativeObjectApplicationContents:
		application := vm.host.Application()
		application.WaitForServer(vm.host.Server())
		switch {
		case member == "" || strings.EqualFold(member, "Item"):
			if len(args) >= 2 {
				application.Set(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				if appValue, ok := application.Get(args[0].String()); ok {
					return vm.applicationValueToValue(appValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Remove"):
			if len(args) >= 1 {
				application.Remove(args[0].String())
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "RemoveAll"):
			application.RemoveAll()
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(application.GetContentsCopy())))
		case strings.EqualFold(member, "Keys"):
			contents := application.GetContentsCopy()
			keys := make([]string, 0, len(contents))
			for key := range contents {
				keys = append(keys, key)
			}
			sort.Strings(keys)
			values := make([]Value, len(keys))
			for i := 0; i < len(keys); i++ {
				values[i] = NewString(keys[i])
			}
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
		case strings.EqualFold(member, "Items"):
			contents := application.GetContentsCopy()
			keys := make([]string, 0, len(contents))
			for key := range contents {
				keys = append(keys, key)
			}
			sort.Strings(keys)
			values := make([]Value, 0, len(keys))
			for i := 0; i < len(keys); i++ {
				if appValue, ok := application.Get(keys[i]); ok {
					values = append(values, vm.applicationValueToValue(appValue))
				}
			}
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
		}
	case nativeObjectApplicationStaticObjects:
		application := vm.host.Application()
		application.WaitForServer(vm.host.Server())
		switch {
		case member == "" || strings.EqualFold(member, "Item"):
			if len(args) >= 2 {
				application.AddStaticObject(args[0].String(), vm.valueToApplicationValue(args[1]))
				return Value{Type: VTEmpty}
			}
			if len(args) >= 1 {
				if appValue, ok := application.GetStaticObject(args[0].String()); ok {
					return vm.applicationValueToValue(appValue)
				}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(application.GetStaticObjectsCopy())))
		}
	case nativeObjectObjectContext:
		// ObjectContext provides transaction control for COM+ transactional ASP pages.
		switch {
		case strings.EqualFold(member, "SetComplete"):
			// Marks the transaction as complete (commit on page end).
			vm.transactionState = 1
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "SetAbort"):
			// Marks the transaction as aborted (rollback on page end).
			vm.transactionState = 2
			return Value{Type: VTEmpty}
		}
	}
	return Value{Type: VTEmpty}
}

// dispatchMemberGet resolves chained member access on native objects.
func (vm *VM) dispatchMemberGet(target Value, member string) Value {
	if target.Type != VTNativeObject {
		return Value{Type: VTEmpty}
	}
	if g3mdObject, exists := vm.g3mdItems[target.Num]; exists {
		return g3mdObject.DispatchPropertyGet(member)
	}

	if g3testObject, exists := vm.g3testItems[target.Num]; exists {
		return g3testObject.DispatchPropertyGet(member)
	}

	if g3cryptoObject, exists := vm.g3cryptoItems[target.Num]; exists {
		return g3cryptoObject.DispatchPropertyGet(member)
	}

	if g3jsonObject, exists := vm.g3jsonItems[target.Num]; exists {
		return g3jsonObject.DispatchPropertyGet(member)
	}

	if g3httpObject, exists := vm.g3httpItems[target.Num]; exists {
		return g3httpObject.DispatchPropertyGet(member)
	}

	if g3mailObject, exists := vm.g3mailItems[target.Num]; exists {
		return g3mailObject.DispatchPropertyGet(member)
	}

	if g3imageObject, exists := vm.g3imageItems[target.Num]; exists {
		return g3imageObject.DispatchPropertyGet(member)
	}
	if g3filesObject, exists := vm.g3filesItems[target.Num]; exists {
		return g3filesObject.DispatchPropertyGet(member)
	}
	if g3templateObject, exists := vm.g3templateItems[target.Num]; exists {
		return g3templateObject.DispatchPropertyGet(member)
	}
	if g3zipObject, exists := vm.g3zipItems[target.Num]; exists {
		return g3zipObject.DispatchPropertyGet(member)
	}
	if g3zlibObject, exists := vm.g3zlibItems[target.Num]; exists {
		return g3zlibObject.DispatchPropertyGet(member)
	}
	if g3tarObject, exists := vm.g3tarItems[target.Num]; exists {
		return g3tarObject.DispatchPropertyGet(member)
	}
	if g3zstdObject, exists := vm.g3zstdItems[target.Num]; exists {
		return g3zstdObject.DispatchPropertyGet(member)
	}
	if g3dbObject, exists := vm.g3dbItems[target.Num]; exists {
		return g3dbObject.DispatchPropertyGet(member)
	}
	if g3dbRS, exists := vm.g3dbResultSetItems[target.Num]; exists {
		return g3dbRS.DispatchPropertyGet(member)
	}
	if g3dbFields, exists := vm.g3dbFieldsItems[target.Num]; exists {
		return g3dbFields.DispatchPropertyGet(member)
	}
	if g3dbRow, exists := vm.g3dbRowItems[target.Num]; exists {
		return g3dbRow.DispatchPropertyGet(member)
	}
	if g3dbStmt, exists := vm.g3dbStatementItems[target.Num]; exists {
		return g3dbStmt.DispatchPropertyGet(member)
	}
	if g3dbTx, exists := vm.g3dbTransactionItems[target.Num]; exists {
		return g3dbTx.DispatchPropertyGet(member)
	}
	if g3dbResult, exists := vm.g3dbResultItems[target.Num]; exists {
		return g3dbResult.DispatchPropertyGet(member)
	}
	if pdfObj, exists := vm.pdfItems[target.Num]; exists {
		return pdfObj.DispatchPropertyGet(member)
	}
	if wscriptShellObject, exists := vm.wscriptShellItems[target.Num]; exists {
		return wscriptShellObject.DispatchPropertyGet(member)
	}
	if wscriptExecObject, exists := vm.wscriptExecItems[target.Num]; exists {
		return wscriptExecObject.DispatchPropertyGet(member)
	}
	if processStreamObject, exists := vm.wscriptProcessStreamItems[target.Num]; exists {
		return processStreamObject.DispatchPropertyGet(member)
	}

	if adoxCatalogObject, exists := vm.adoxCatalogItems[target.Num]; exists {
		return adoxCatalogObject.DispatchPropertyGet(member)
	}
	if adoxTablesObject, exists := vm.adoxTablesItems[target.Num]; exists {
		return adoxTablesObject.DispatchPropertyGet(member)
	}
	if adoxTableObject, exists := vm.adoxTableItems[target.Num]; exists {
		return adoxTableObject.DispatchPropertyGet(member)
	}

	if mswcAdRotator, exists := vm.mswcAdRotatorItems[target.Num]; exists {
		return mswcAdRotator.DispatchPropertyGet(member)
	}
	if mswcBrowserType, exists := vm.mswcBrowserTypeItems[target.Num]; exists {
		return mswcBrowserType.DispatchPropertyGet(member)
	}
	if mswcNextLink, exists := vm.mswcNextLinkItems[target.Num]; exists {
		return mswcNextLink.DispatchPropertyGet(member)
	}
	if mswcContentRotator, exists := vm.mswcContentRotatorItems[target.Num]; exists {
		return mswcContentRotator.DispatchPropertyGet(member)
	}
	if mswcCounters, exists := vm.mswcCountersItems[target.Num]; exists {
		return mswcCounters.DispatchPropertyGet(member)
	}
	if mswcPageCounter, exists := vm.mswcPageCounterItems[target.Num]; exists {
		return mswcPageCounter.DispatchPropertyGet(member)
	}
	if mswcTools, exists := vm.mswcToolsItems[target.Num]; exists {
		return mswcTools.DispatchPropertyGet(member)
	}
	if mswcMyInfo, exists := vm.mswcMyInfoItems[target.Num]; exists {
		return mswcMyInfo.DispatchPropertyGet(member)
	}
	if mswcPermissionChecker, exists := vm.mswcPermissionCheckerItems[target.Num]; exists {
		return mswcPermissionChecker.DispatchPropertyGet(member)
	}
	if msxmlServer, exists := vm.msxmlServerItems[target.Num]; exists {
		return msxmlServer.DispatchPropertyGet(member)
	}
	if msxmlDOM, exists := vm.msxmlDOMItems[target.Num]; exists {
		return msxmlDOM.DispatchPropertyGet(member)
	}
	if msxmlNodeList, exists := vm.msxmlNodeListItems[target.Num]; exists {
		return msxmlNodeList.DispatchPropertyGet(member)
	}
	if msxmlParseError, exists := vm.msxmlParseErrorItems[target.Num]; exists {
		return msxmlParseError.DispatchPropertyGet(member)
	}
	if msxmlElement, exists := vm.msxmlElementItems[target.Num]; exists {
		return msxmlElement.DispatchPropertyGet(member)
	}
	if pdf, exists := vm.pdfItems[target.Num]; exists {
		return pdf.DispatchPropertyGet(member)
	}
	if fileUploader, exists := vm.fileUploaderItems[target.Num]; exists {
		return fileUploader.DispatchPropertyGet(member)
	}

	if axonObject, exists := vm.axonItems[target.Num]; exists {
		return axonObject.DispatchPropertyGet(member)
	}

	if dictResult, handled := vm.dispatchDictionaryPropertyGet(target.Num, member); handled {
		return dictResult
	}

	if fsoResult, handled := vm.dispatchFSOPropertyGet(target.Num, member); handled {
		return fsoResult
	}

	if adodbResult, handled := vm.dispatchADODBStreamPropertyGet(target.Num, member); handled {
		return adodbResult
	}

	if adodbResult, handled := vm.dispatchADODBPropertyGet(target.Num, member); handled {
		return adodbResult
	}

	if adodbResult, handled := vm.dispatchADODBErrorsCollectionPropertyGet(target.Num, member); handled {
		return adodbResult
	}

	if adodbResult, handled := vm.dispatchADODBErrorPropertyGet(target.Num, member); handled {
		return adodbResult
	}

	if adodbResult, handled := vm.dispatchADODBFieldPropertyGet(target.Num, member); handled {
		return adodbResult
	}

	// FieldsCollection: rsAccess.Fields.Count / rsAccess.Fields.Item(i)
	// These are accessed via OpMemberGet (property path), not OpCallMember.
	// Without this handler, rsAccess.Fields returns a VTNativeObject that is then
	// passed to dispatchMemberGet for ".Count" — which previously found no handler
	// and silently returned VTEmpty (0), causing Fields.Count = 0 even when columns
	// were fully populated.
	if adodbResult, handled := vm.dispatchADODBFieldsCollectionPropertyGet(target.Num, member); handled {
		return adodbResult
	}

	if adodbResult, handled := vm.dispatchADODBParametersCollectionPropertyGet(target.Num, member); handled {
		return adodbResult
	}

	if regExpResult, handled := vm.dispatchRegExpPropertyGet(target.Num, member); handled {
		return regExpResult
	}

	if target.Num == nativeObjectErr {
		return vm.errPropertyValue(member)
	}

	// The console object exposes no settable properties; method access returns Empty.
	if target.Num == nativeObjectConsole {
		return Value{Type: VTEmpty}
	}

	switch target.Num {
	case nativeObjectResponse:
		switch {
		case strings.EqualFold(member, "Buffer"):
			return NewBool(vm.host.Response().GetBuffer())
		case strings.EqualFold(member, "CacheControl"):
			return NewString(vm.host.Response().GetCacheControl())
		case strings.EqualFold(member, "Charset"):
			return NewString(vm.host.Response().GetCharset())
		case strings.EqualFold(member, "CodePage"):
			return NewInteger(int64(vm.host.Response().GetCodePage()))
		case strings.EqualFold(member, "ContentType"):
			return NewString(vm.host.Response().GetContentType())
		case strings.EqualFold(member, "Expires"):
			return NewInteger(int64(vm.host.Response().GetExpires()))
		case strings.EqualFold(member, "ExpiresAbsolute"):
			return NewString(vm.host.Response().GetExpiresAbsoluteRaw())
		case strings.EqualFold(member, "PICS"):
			return NewString(vm.host.Response().GetPICS())
		case strings.EqualFold(member, "Status"):
			return NewString(vm.host.Response().GetStatus())
		case strings.EqualFold(member, "IsClientConnected"):
			return NewBool(vm.host.Response().IsClientConnected())
		case strings.EqualFold(member, "Cookies"):
			return Value{Type: VTNativeObject, Num: nativeResponseCookies}
		}
	case nativeObjectRequest:
		switch {
		case strings.EqualFold(member, "QueryString"):
			return Value{Type: VTNativeObject, Num: nativeRequestQueryString}
		case strings.EqualFold(member, "Form"):
			return Value{Type: VTNativeObject, Num: nativeRequestForm}
		case strings.EqualFold(member, "Cookies"):
			return Value{Type: VTNativeObject, Num: nativeRequestCookies}
		case strings.EqualFold(member, "ServerVariables"):
			return Value{Type: VTNativeObject, Num: nativeRequestServerVariables}
		case strings.EqualFold(member, "ClientCertificate"):
			return Value{Type: VTNativeObject, Num: nativeRequestClientCertificate}
		case strings.EqualFold(member, "BinaryRead"):
			return Value{Type: VTNativeObject, Num: nativeRequestBinaryReadMethod}
		case strings.EqualFold(member, "TotalBytes"):
			return NewInteger(vm.host.Request().TotalBytes())
		default:
			return NewString(vm.host.Request().GetValue(member))
		}
	case nativeRequestQueryString:
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(vm.host.Request().QueryString.Count()))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeRequestQueryStringKeyMethod}
		}
	case nativeRequestForm:
		switch {
		case strings.EqualFold(member, "Count"):
			if vm.host.Request().IsBinaryReadUsed() {
				return NewInteger(0)
			}
			vm.host.Request().MarkFormUsed()
			return NewInteger(int64(vm.host.Request().Form.Count()))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeRequestFormKeyMethod}
		}
	case nativeRequestCookies:
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(vm.host.Request().Cookies.Count()))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeRequestCookiesKeyMethod}
		}
	case nativeRequestServerVariables:
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(vm.host.Request().ServerVars.Count()))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeRequestServerVariablesKeyMethod}
		}
	case nativeRequestClientCertificate:
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(vm.host.Request().ClientCertificate.Count()))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeRequestClientCertificateKeyMethod}
		}
	case nativeObjectSession:
		// Raise a runtime error when session state has been disabled by an ASP page directive.
		if !vm.host.SessionEnabled() {
			vm.raise(vbscript.PermissionDenied, "Session state is disabled for this application")
		}
		session := vm.host.Session()
		switch {
		case strings.EqualFold(member, "SessionID"):
			return NewInteger(session.SessionIDNumeric())
		case strings.EqualFold(member, "Timeout"):
			return NewInteger(int64(session.GetTimeout()))
		case strings.EqualFold(member, "LCID"):
			return NewInteger(int64(session.GetLCID()))
		case strings.EqualFold(member, "CodePage"):
			return NewInteger(int64(session.GetCodePage()))
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(session.Count()))
		case strings.EqualFold(member, "IsLocked"):
			return NewBool(session.IsLocked())
		case strings.EqualFold(member, "Contents"):
			return Value{Type: VTNativeObject, Num: nativeObjectSessionContents}
		case strings.EqualFold(member, "StaticObjects"):
			return Value{Type: VTNativeObject, Num: nativeObjectSessionStaticObjects}
		}
	case nativeObjectApplication:
		switch {
		case strings.EqualFold(member, "Contents"):
			return Value{Type: VTNativeObject, Num: nativeObjectApplicationContents}
		case strings.EqualFold(member, "StaticObjects"):
			return Value{Type: VTNativeObject, Num: nativeObjectApplicationStaticObjects}
		}
	case nativeObjectServer:
		switch {
		case strings.EqualFold(member, "ScriptTimeout"):
			return NewInteger(int64(vm.host.Server().GetScriptTimeout()))
		}
	case nativeObjectSessionContents:
		session := vm.host.Session()
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(session.Count()))
		}
	case nativeObjectSessionStaticObjects:
		session := vm.host.Session()
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(session.GetStaticObjectsCopy())))
		}
	case nativeObjectApplicationContents:
		application := vm.host.Application()
		application.WaitForServer(vm.host.Server())
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(application.GetContentsCopy())))
		}
	case nativeObjectApplicationStaticObjects:
		application := vm.host.Application()
		application.WaitForServer(vm.host.Server())
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(application.GetStaticObjectsCopy())))
		}
	}

	if cookieName, exists := vm.responseCookieItems[target.Num]; exists {
		switch {
		case strings.EqualFold(member, "Name"):
			return NewString(cookieName)
		case strings.EqualFold(member, "Value"):
			return NewString(vm.host.Response().GetCookieProperty(cookieName, "Value"))
		case strings.EqualFold(member, "Domain"):
			return NewString(vm.host.Response().GetCookieProperty(cookieName, "Domain"))
		case strings.EqualFold(member, "Path"):
			return NewString(vm.host.Response().GetCookieProperty(cookieName, "Path"))
		case strings.EqualFold(member, "Expires"):
			return NewString(vm.host.Response().GetCookieProperty(cookieName, "Expires"))
		case strings.EqualFold(member, "Secure"):
			return NewString(vm.host.Response().GetCookieProperty(cookieName, "Secure"))
		case strings.EqualFold(member, "HttpOnly"):
			return NewString(vm.host.Response().GetCookieProperty(cookieName, "HttpOnly"))
		}
	}

	if errObj, exists := vm.aspErrorItems[target.Num]; exists {
		return vm.aspErrorPropertyValue(errObj, member)
	}

	return Value{Type: VTEmpty}
}

func (vm *VM) newNativeObjectProxy(parentID int64, member string, args []Value) Value {
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	storedArgs := make([]Value, len(args))
	copy(storedArgs, args)
	vm.nativeObjectProxies[id] = nativeObjectProxy{ParentID: parentID, Member: member, CallArgs: storedArgs}
	return Value{Type: VTNativeObject, Num: id}
}

// dispatchMemberSet routes a property Let assignment to the correct native object handler.
// It is called by the OpMemberSet instruction and covers all dynamic and static native objects.
func (vm *VM) dispatchMemberSet(objID int64, member string, val Value) {
	if proxy, exists := vm.nativeObjectProxies[objID]; exists {
		// Parameterized property set via OpMemberSet (e.g. obj.Prop(idx) = val)
		// OpArraySet is handled via dispatchNativeCall (member="").
		// OpMemberSet usually happens when the compiler thinks it's a simple property set,
		// but the object returned by OpMemberGet was a proxy.
		// If member is not empty, it might be the index (if compiled as proxy.index = val).
		combinedArgs := make([]Value, 0, len(proxy.CallArgs)+2)
		combinedArgs = append(combinedArgs, proxy.CallArgs...)
		if member != "" {
			combinedArgs = append(combinedArgs, NewString(member))
		}
		combinedArgs = append(combinedArgs, val)
		vm.dispatchNativeCall(proxy.ParentID, proxy.Member, combinedArgs)
		return
	}

	// Dynamic native objects: G3MD
	if g3mdObject, exists := vm.g3mdItems[objID]; exists {
		g3mdObject.DispatchPropertySet(member, val)
		return
	}

	if g3testObject, exists := vm.g3testItems[objID]; exists {
		g3testObject.DispatchPropertySet(member, val)
		return
	}

	if g3cryptoObject, exists := vm.g3cryptoItems[objID]; exists {
		g3cryptoObject.DispatchPropertySet(member, val)
		return
	}

	if g3mailObject, exists := vm.g3mailItems[objID]; exists {
		g3mailObject.DispatchPropertySet(member, []Value{val})
		return
	}

	if g3imageObject, exists := vm.g3imageItems[objID]; exists {
		g3imageObject.DispatchPropertySet(member, []Value{val})
		return
	}

	if g3filesObject, exists := vm.g3filesItems[objID]; exists {
		g3filesObject.DispatchPropertySet(member, []Value{val})
		return
	}

	if g3dbObject, exists := vm.g3dbItems[objID]; exists {
		g3dbObject.DispatchPropertySet(member, []Value{val})
		return
	}

	if adoxCatalogObject, exists := vm.adoxCatalogItems[objID]; exists {
		adoxCatalogObject.DispatchPropertySet(member, []Value{val})
		return
	}

	if mswcAdRotator, exists := vm.mswcAdRotatorItems[objID]; exists {
		mswcAdRotator.DispatchPropertySet(member, []Value{val})
		return
	}
	if mswcMyInfo, exists := vm.mswcMyInfoItems[objID]; exists {
		mswcMyInfo.DispatchPropertySet(member, []Value{val})
		return
	}

	if msxmlServer, exists := vm.msxmlServerItems[objID]; exists {
		msxmlServer.DispatchPropertySet(member, []Value{val})
		return
	}
	if msxmlDOM, exists := vm.msxmlDOMItems[objID]; exists {
		msxmlDOM.DispatchPropertySet(member, []Value{val})
		return
	}
	if msxmlNodeList, exists := vm.msxmlNodeListItems[objID]; exists {
		msxmlNodeList.DispatchPropertySet(member, []Value{val})
		return
	}
	if msxmlParseError, exists := vm.msxmlParseErrorItems[objID]; exists {
		msxmlParseError.DispatchPropertySet(member, []Value{val})
		return
	}
	if msxmlElement, exists := vm.msxmlElementItems[objID]; exists {
		msxmlElement.DispatchPropertySet(member, []Value{val})
		return
	}
	if pdf, exists := vm.pdfItems[objID]; exists {
		pdf.DispatchPropertySet(member, []Value{val})
		return
	}
	if fileUploader, exists := vm.fileUploaderItems[objID]; exists {
		fileUploader.DispatchPropertySet(member, []Value{val})
		return
	}

	// Dynamic native objects: Scripting.FileSystemObject family
	if vm.dispatchFSOPropertySet(objID, member, val) {
		return
	}

	if vm.dispatchADODBStreamPropertySet(objID, member, val) {
		return
	}

	// Dynamic native objects: Scripting.Dictionary
	if vm.dispatchDictionaryPropertySet(objID, member, val) {
		return
	}

	if vm.dispatchADODBPropertySet(objID, member, val) {
		return
	}

	if vm.dispatchADODBFieldPropertySet(objID, member, val) {
		return
	}

	if vm.dispatchRegExpPropertySet(objID, member, val) {
		return
	}

	if objID == nativeObjectErr {
		vm.errPropertySet(member, val)
		return
	}

	// Static native object: Response

	if objID == nativeObjectResponse {
		switch {
		case strings.EqualFold(member, "Buffer"):
			vm.host.Response().SetBuffer(val.Num != 0)
		case strings.EqualFold(member, "CacheControl"):
			vm.host.Response().SetCacheControl(val.String())
		case strings.EqualFold(member, "Charset"):
			vm.host.Response().SetCharset(val.String())
		case strings.EqualFold(member, "ContentType"):
			vm.host.Response().SetContentType(val.String())
		case strings.EqualFold(member, "Expires"):
			vm.host.Response().SetExpires(int(val.Num))
		case strings.EqualFold(member, "ExpiresAbsolute"):
			vm.host.Response().SetExpiresAbsoluteRaw(val.String())
		case strings.EqualFold(member, "PICS"):
			vm.host.Response().SetPICS(val.String())
		case strings.EqualFold(member, "Status"):
			vm.host.Response().SetStatus(val.String())
		}
		return
	}

	// Static native object: Server
	if objID == nativeObjectServer {
		switch {
		case strings.EqualFold(member, "ScriptTimeout"):
			if err := vm.host.Server().SetScriptTimeout(vm.asInt(val)); err != nil {
				vm.host.Server().SetLastError(asp.NewVBScriptASPError(vbscript.InvalidProcedureCallOrArgument, "Server.ScriptTimeout", "ASP", err.Error(), "", 0, 0))
			}
		}
		return
	}

	// Unknown object ID or read-only property — silently ignore for VBScript compatibility.
}

// newResponseCookieItem creates a native cookie item object used for chained cookie member access.
func (vm *VM) newResponseCookieItem(cookieName string) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.responseCookieItems[objID] = cookieName
	return Value{Type: VTNativeObject, Num: objID}
}

// newASPErrorObject creates a native ASPError object used for Server.GetLastError().
func (vm *VM) newASPErrorObject(errObj *asp.ASPError) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.aspErrorItems[objID] = errObj.Clone()
	return Value{Type: VTNativeObject, Num: objID}
}

// newG3MDObject creates a native G3MD object used by Server.CreateObject("G3MD").
func (vm *VM) newG3MDObject() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3mdItems[objID] = NewG3MD()
	return Value{Type: VTNativeObject, Num: objID}
}

// newG3CryptoObject creates a native G3Crypto object used by Server.CreateObject aliases.
func (vm *VM) newG3CryptoObject(defaultAlgorithm string) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3cryptoItems[objID] = NewG3CryptoWithAlgorithm(defaultAlgorithm)
	return Value{Type: VTNativeObject, Num: objID}
}

// aspErrorPropertyValue reads one property from the ASPError object model.
func (vm *VM) aspErrorPropertyValue(errObj *asp.ASPError, property string) Value {
	if errObj == nil {
		return Value{Type: VTEmpty}
	}

	switch strings.ToLower(property) {
	case "aspcode":
		return NewInteger(int64(errObj.ASPCode))
	case "aspdescription":
		return NewString(errObj.ASPDescription)
	case "number":
		return NewInteger(int64(errObj.Number))
	case "source":
		return NewString(errObj.Source)
	case "description":
		return NewString(errObj.Description)
	case "helpfile":
		return NewString(errObj.HelpFile)
	case "helpcontext":
		return NewInteger(int64(errObj.HelpContext))
	case "file":
		return NewString(errObj.File)
	case "line":
		return NewInteger(int64(errObj.Line))
	case "column":
		return NewInteger(int64(errObj.Column))
	case "category":
		return NewString(errObj.Category)
	default:
		return Value{Type: VTEmpty}
	}
}

// valueToString converts VM values to string and resolves dynamic native values when needed.
func (vm *VM) valueToString(v Value) string {
	v = resolveCallable(vm, v)
	if v.Type == VTNativeObject {
		if cookieName, exists := vm.responseCookieItems[v.Num]; exists {
			return vm.host.Response().GetCookieValue(cookieName)
		}
		if errObj, exists := vm.aspErrorItems[v.Num]; exists {
			return errObj.Description
		}
		if subMatchValue, exists := vm.regExpSubMatchValueItems[v.Num]; exists {
			return subMatchValue.value
		}
		if v.Num == nativeObjectErr {
			return vm.errPropertyValue("Description").String()
		}
		// ADODB.Field proxy: coerce to its default Value property in string context.
		if adodbDefault, handled := vm.dispatchADODBFieldPropertyGet(v.Num, "__default__"); handled {
			return vm.valueToString(adodbDefault)
		}
	}
	// Handle date formatting according to locale when converting to string
	if v.Type == VTDate {
		return vm.dateToLocalizedString(v)
	}
	return v.String()
}

// valueToBinaryBytes normalizes a VM value into raw bytes for Response.BinaryWrite.
func (vm *VM) valueToBinaryBytes(v Value) []byte {
	v = resolveCallable(vm, v)
	if v.Type == VTArray && v.Arr != nil {
		values := v.Arr.Values
		if len(values) == 0 {
			return []byte{}
		}
		out := make([]byte, len(values))
		for i := 0; i < len(values); i++ {
			n := vm.asInt(values[i])
			out[i] = byte(n)
		}
		return out
	}
	return vbByteStringToBytes(vm.valueToString(v))
}

// dateToLocalizedString formats a date value according to the current locale.
// When a date is converted to string (e.g., via CStr or concatenation with &),
// it displays:
// - Only date part if time is 00:00:00 (e.g., Date() function result)
// - Date and time if time is non-zero (e.g., Now() function result)
func (vm *VM) dateToLocalizedString(v Value) string {
	if v.Type != VTDate {
		return ""
	}
	dateValue := time.Unix(0, v.Num).In(builtinCurrentLocation(vm))
	return localizedDateString(dateValue, builtinLocaleProfileForVM(vm))
}

// errPropertyValue reads one property from the intrinsic Err object.
func (vm *VM) errPropertyValue(property string) Value {
	errObj := vm.errObject
	if errObj == nil {
		errObj = asp.NewASPError()
		vm.errObject = errObj
	}

	switch strings.ToLower(property) {
	case "aspcode":
		if vm.errASPCodeRawSet {
			return NewString(vm.errASPCodeRaw)
		}
		return NewInteger(int64(errObj.ASPCode))
	case "aspdescription":
		return NewString(errObj.ASPDescription)
	case "number":
		return NewInteger(int64(errObj.Number))
	case "source":
		return NewString(errObj.Source)
	case "description":
		return NewString(errObj.Description)
	case "helpfile":
		return NewString(errObj.HelpFile)
	case "helpcontext":
		return NewInteger(int64(errObj.HelpContext))
	case "file":
		return NewString(errObj.File)
	case "line":
		return NewInteger(int64(errObj.Line))
	case "column":
		return NewInteger(int64(errObj.Column))
	case "category":
		return NewString(errObj.Category)
	default:
		return Value{Type: VTEmpty}
	}
}

// errPropertySet writes one property on the intrinsic Err object.
func (vm *VM) errPropertySet(property string, val Value) {
	errObj := vm.errObject
	if errObj == nil {
		errObj = asp.NewASPError()
		vm.errObject = errObj
	}

	switch strings.ToLower(property) {
	case "aspcode":
		if val.Type == VTString {
			trimmed := strings.TrimSpace(val.Str)
			var parsed int
			if _, err := fmt.Sscanf(trimmed, "%d", &parsed); err == nil {
				errObj.ASPCode = parsed
				vm.errASPCodeRaw = ""
				vm.errASPCodeRawSet = false
			} else {
				vm.errASPCodeRaw = val.Str
				vm.errASPCodeRawSet = true
			}
		} else {
			errObj.ASPCode = vm.asInt(val)
			vm.errASPCodeRaw = ""
			vm.errASPCodeRawSet = false
		}
	case "aspdescription":
		errObj.ASPDescription = vm.valueToString(val)
	case "number":
		errObj.Number = vm.asInt(val)
	case "source":
		errObj.Source = vm.valueToString(val)
	case "description":
		errObj.Description = vm.valueToString(val)
	case "helpfile":
		errObj.HelpFile = vm.valueToString(val)
	case "helpcontext":
		errObj.HelpContext = vm.asInt(val)
	case "file":
		errObj.File = vm.valueToString(val)
	case "line":
		errObj.Line = vm.asInt(val)
	case "column":
		errObj.Column = vm.asInt(val)
	case "category":
		errObj.Category = vm.valueToString(val)
	}
	errObj.Normalize()
}

// errClear resets the intrinsic Err object to its default state.
func (vm *VM) errClear() {
	vm.errObject = asp.NewASPError()
	vm.errASPCodeRaw = ""
	vm.errASPCodeRawSet = false
	vm.lastError = nil
}

// errRaise implements Err.Raise(number, source, description, helpfile, helpcontext).
// It updates the intrinsic Err object and either suppresses or propagates the
// runtime error according to the current On Error mode.
func (vm *VM) errRaise(args []Value) Value {
	if len(args) == 0 {
		vm.raise(vbscript.InvalidProcedureCallOrArgument, "Invalid procedure call or argument")
		return Value{Type: VTEmpty}
	}

	number := vm.asInt(resolveCallable(vm, args[0]))
	source := "VBScript runtime error"
	if len(args) >= 2 {
		source = strings.TrimSpace(vm.valueToString(resolveCallable(vm, args[1])))
		if source == "" {
			source = "VBScript runtime error"
		}
	}

	description := ""
	if len(args) >= 3 {
		description = strings.TrimSpace(vm.valueToString(resolveCallable(vm, args[2])))
	}
	helpFile := ""
	if len(args) >= 4 {
		helpFile = strings.TrimSpace(vm.valueToString(resolveCallable(vm, args[3])))
	}
	helpContext := 0
	if len(args) >= 5 {
		helpContext = vm.asInt(resolveCallable(vm, args[4]))
	}
	if description == "" {
		description = vbscript.VBSyntaxErrorCode(number).String()
		if description == "" {
			description = "Application-defined or object-defined error"
		}
	}

	if number == 2 && strings.EqualFold(source, "JSONobject") && strings.EqualFold(description, "A property already exists with the name: data.") {
		// Compatibility: legacy JSONobject scripts may use "data" both as a
		// default-property alias and as an explicit key. IIS-compatible execution
		// should not hard-abort or dirty Err state for this specific pattern.
		return Value{Type: VTEmpty}
	}

	vme := &VMError{
		Code:           vbscript.VBSyntaxErrorCode(number),
		Line:           vm.lastLine,
		Column:         vm.lastColumn,
		Msg:            description,
		ASPCode:        number,
		ASPDescription: description,
		Category:       "VBScript runtime",
		Description:    description,
		Number:         number,
		Source:         source,
		HelpFile:       helpFile,
		HelpContext:    helpContext,
	}

	vm.errSetFromVMError(vme)
	if vm.onResumeNext {
		vm.lastError = vme
		return Value{Type: VTEmpty}
	}

	panic(vme)
}

// errSetFromVMError mirrors a runtime VM error into the intrinsic Err object.
func (vm *VM) errSetFromVMError(vme *VMError) {
	if vme == nil {
		return
	}
	vm.errASPCodeRaw = ""
	vm.errASPCodeRawSet = false
	vm.errObject = (&asp.ASPError{
		ASPCode:        vme.ASPCode,
		ASPDescription: vme.ASPDescription,
		Number:         vme.Number,
		Source:         vme.Source,
		Description:    vme.Description,
		HelpFile:       vme.HelpFile,
		HelpContext:    vme.HelpContext,
		File:           vme.File,
		Line:           vme.Line,
		Column:         vme.Column,
		Category:       vme.Category,
	}).Normalize()
}

// raiseFromSyntaxError propagates one dynamic compilation error using the same
// Err/on-error semantics as runtime faults.
func (vm *VM) raiseFromSyntaxError(syntaxErr *vbscript.VBSyntaxError) {
	if syntaxErr == nil {
		return
	}
	vme := &VMError{
		Code:           syntaxErr.Code,
		Line:           syntaxErr.Line,
		Column:         syntaxErr.Column,
		File:           syntaxErr.File,
		Msg:            syntaxErr.ASPDescription,
		ASPCode:        syntaxErr.ASPCode,
		ASPDescription: syntaxErr.ASPDescription,
		Category:       syntaxErr.Category,
		Description:    syntaxErr.Description,
		Number:         syntaxErr.Number,
		Source:         syntaxErr.Source,
	}
	vm.errSetFromVMError(vme)
	if vm.onResumeNext || (vm.terminatePrepared && !vm.suppressTerminate) {
		vm.lastError = vme
		return
	}
	panic(vme)
}

// valueToApplicationValue converts a VM value into a typed application storage value.
func (vm *VM) valueToApplicationValue(v Value) asp.ApplicationValue {
	switch v.Type {
	case VTBool:
		return asp.NewApplicationBool(v.Num != 0)
	case VTInteger:
		return asp.NewApplicationInteger(v.Num)
	case VTDouble:
		return asp.NewApplicationDouble(v.Flt)
	case VTString:
		return asp.NewApplicationString(v.Str)
	case VTEmpty:
		return asp.NewApplicationEmpty()
	case VTArray:
		if v.Arr != nil {
			return vm.vbArrayToApplicationValue(v.Arr)
		}
		return asp.NewApplicationEmpty()
	default:
		return asp.NewApplicationString(v.String())
	}
}

// vbArrayToApplicationValue recursively converts a VBArray into an ApplicationValue tree.
func (vm *VM) vbArrayToApplicationValue(arr *VBArray) asp.ApplicationValue {
	elements := make([]asp.ApplicationValue, len(arr.Values))
	for i, elem := range arr.Values {
		elements[i] = vm.valueToApplicationValue(elem)
	}
	return asp.NewApplicationArray(arr.Lower, elements)
}

// applicationValueToVBArray recursively converts an ApplicationValue array tree back into a Value.
func (vm *VM) applicationValueToVBArray(v asp.ApplicationValue) Value {
	elements := make([]Value, len(v.Arr))
	for i, elem := range v.Arr {
		elements[i] = vm.applicationValueToValue(elem)
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(v.ArrLower, elements)}
}

// applicationValueToValue converts a typed application storage value into a VM value.
func (vm *VM) applicationValueToValue(v asp.ApplicationValue) Value {
	switch v.Type {
	case asp.ApplicationValueBool:
		return NewBool(v.Bool())
	case asp.ApplicationValueInteger:
		return NewInteger(v.Num)
	case asp.ApplicationValueDouble:
		return NewDouble(v.Flt)
	case asp.ApplicationValueString:
		if strings.HasPrefix(v.Str, staticObjectProgIDPrefix) {
			return vm.materializeStaticObjectFromMarker(v.Str)
		}
		return NewString(v.Str)
	case asp.ApplicationValueArray:
		return vm.applicationValueToVBArray(v)
	default:
		return Value{Type: VTEmpty}
	}
}

// isStaticObjectApplicationValue reports whether one Application/Session static entry comes from a <object> marker.
func (vm *VM) isStaticObjectApplicationValue(v asp.ApplicationValue) bool {
	return v.Type == asp.ApplicationValueString && strings.HasPrefix(v.Str, staticObjectProgIDPrefix)
}

// materializeStaticObjectFromMarker creates a runtime object from a stored static-object ProgID marker.
func (vm *VM) materializeStaticObjectFromMarker(marker string) Value {
	progID := strings.TrimSpace(strings.TrimPrefix(marker, staticObjectProgIDPrefix))
	if progID == "" {
		return Value{Type: VTEmpty}
	}

	if strings.EqualFold(progID, "G3MD") {
		return vm.newG3MDObject()
	}
	if strings.EqualFold(progID, "G3TestSuite") || strings.EqualFold(progID, "G3Test") {
		return vm.newG3TestObject()
	}
	if defaultAlgorithm, ok := g3cryptoResolveProgID(progID); ok {
		return vm.newG3CryptoObject(defaultAlgorithm)
	}
	if strings.EqualFold(progID, "Scripting.FileSystemObject") {
		return vm.newFSORootObject()
	}
	if strings.EqualFold(progID, "G3FILES") || strings.EqualFold(progID, "G3Files") {
		return vm.newG3FilesObject()
	}
	if strings.EqualFold(progID, "ADODB.Stream") {
		return vm.newADODBStreamObject()
	}
	if strings.EqualFold(progID, "ADODB.Connection") {
		return vm.newADODBConnection()
	}
	if strings.EqualFold(progID, "ADODBOLE.Connection") {
		return vm.newADODBOLEConnection()
	}
	if strings.EqualFold(progID, "ADODB.Recordset") {
		return vm.newADODBRecordset()
	}
	if strings.EqualFold(progID, "ADODB.Command") {
		return vm.newADODBCommand()
	}
	if strings.EqualFold(progID, "VBScript.RegExp") || strings.EqualFold(progID, "RegExp") {
		return vm.newRegExpObject()
	}
	if strings.EqualFold(progID, "Scripting.Dictionary") {
		return vm.newDictionaryObject()
	}

	if vm.host != nil && vm.host.Server() != nil {
		resolved := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString(progID)})
		if resolved.Type != VTEmpty {
			return resolved
		}
	}

	return Value{Type: VTEmpty}
}

func (vm *VM) push(v Value) {
	if vm.sp+1 >= StackSize {
		vm.raise(vbscript.StackOverflow, "Stack overflow")
	}
	vm.sp++
	vm.stack[vm.sp] = v
}

// bumpRuntimeClassVersion advances one monotonic class metadata version used by dynamic caches.
func (vm *VM) bumpRuntimeClassVersion() {
	if vm == nil {
		return
	}
	vm.runtimeClassVersion++
	if vm.runtimeClassVersion == 0 {
		vm.runtimeClassVersion = 1
	}
}

// registerRuntimeClass stores one class definition name for New object allocation.
func (vm *VM) registerRuntimeClass(className string) {
	trimmedName := strings.TrimSpace(className)
	if trimmedName == "" {
		return
	}
	lowerName := strings.ToLower(trimmedName)
	if _, exists := vm.runtimeClasses[lowerName]; exists {
		return
	}
	vm.runtimeClasses[lowerName] = RuntimeClassDef{
		Name:       trimmedName,
		Fields:     make(map[string]RuntimeClassFieldDef),
		Methods:    make(map[string]RuntimeClassMethodDef),
		Properties: make(map[string]RuntimeClassPropertyDef),
	}
	vm.bumpRuntimeClassVersion()
}

// registerRuntimeClassMethod stores one method entry for one class definition.
func (vm *VM) registerRuntimeClassMethod(className string, methodName string, methodTarget Value, isPublic bool) {
	trimmedClassName := strings.TrimSpace(className)
	trimmedMethodName := strings.TrimSpace(methodName)
	if trimmedClassName == "" || trimmedMethodName == "" {
		return
	}

	lowerClassName := strings.ToLower(trimmedClassName)
	classDef, exists := vm.runtimeClasses[lowerClassName]
	if !exists {
		classDef = RuntimeClassDef{
			Name:       trimmedClassName,
			Fields:     make(map[string]RuntimeClassFieldDef),
			Methods:    make(map[string]RuntimeClassMethodDef),
			Properties: make(map[string]RuntimeClassPropertyDef),
		}
	} else if classDef.Methods == nil {
		classDef.Methods = make(map[string]RuntimeClassMethodDef)
		if classDef.Fields == nil {
			classDef.Fields = make(map[string]RuntimeClassFieldDef)
		}
	}

	classDef.Methods[strings.ToLower(trimmedMethodName)] = RuntimeClassMethodDef{Target: methodTarget, IsPublic: isPublic}
	vm.runtimeClasses[lowerClassName] = classDef
	vm.bumpRuntimeClassVersion()
}

// registerRuntimeClassField stores one direct field entry for one class definition.
func (vm *VM) registerRuntimeClassField(className string, fieldName string, isPublic bool) {
	trimmedClassName := strings.TrimSpace(className)
	trimmedFieldName := strings.TrimSpace(fieldName)
	if trimmedClassName == "" || trimmedFieldName == "" {
		return
	}

	lowerClassName := strings.ToLower(trimmedClassName)
	classDef, exists := vm.runtimeClasses[lowerClassName]
	if !exists {
		classDef = RuntimeClassDef{
			Name:       trimmedClassName,
			Fields:     make(map[string]RuntimeClassFieldDef),
			Methods:    make(map[string]RuntimeClassMethodDef),
			Properties: make(map[string]RuntimeClassPropertyDef),
		}
	} else if classDef.Fields == nil {
		classDef.Fields = make(map[string]RuntimeClassFieldDef)
	}

	classDef.Fields[strings.ToLower(trimmedFieldName)] = RuntimeClassFieldDef{Name: trimmedFieldName, IsPublic: isPublic}
	vm.runtimeClasses[lowerClassName] = classDef
	vm.bumpRuntimeClassVersion()
}

// registerRuntimeClassFieldDims updates an already-registered class field with fixed-size array dimension
// bounds. Called by OpInitClassArrayField after the field is registered via OpRegisterClassField.
// Each entry in dims is the upper bound for that dimension (matching VBScript Dim arr(N) semantics).
func (vm *VM) registerRuntimeClassFieldDims(className string, fieldName string, dims []int) {
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	lowerFieldName := strings.ToLower(strings.TrimSpace(fieldName))
	if lowerClassName == "" || lowerFieldName == "" {
		return
	}
	classDef, exists := vm.runtimeClasses[lowerClassName]
	if !exists {
		return
	}
	fieldDef, exists := classDef.Fields[lowerFieldName]
	if !exists {
		return
	}
	fieldDef.Dims = dims
	classDef.Fields[lowerFieldName] = fieldDef
	vm.runtimeClasses[lowerClassName] = classDef
	vm.bumpRuntimeClassVersion()
}

// registerRuntimeClassPropertyAccessor stores one Property Get/Let/Set accessor for one class definition.
func (vm *VM) registerRuntimeClassPropertyAccessor(op OpCode, className string, propertyName string, accessorTarget Value, paramCount int, isPublic bool) {
	trimmedClassName := strings.TrimSpace(className)
	trimmedPropertyName := strings.TrimSpace(propertyName)
	if trimmedClassName == "" || trimmedPropertyName == "" {
		return
	}

	lowerClassName := strings.ToLower(trimmedClassName)
	classDef, exists := vm.runtimeClasses[lowerClassName]
	if !exists {
		classDef = RuntimeClassDef{
			Name:       trimmedClassName,
			Fields:     make(map[string]RuntimeClassFieldDef),
			Methods:    make(map[string]RuntimeClassMethodDef),
			Properties: make(map[string]RuntimeClassPropertyDef),
		}
	} else if classDef.Properties == nil {
		classDef.Properties = make(map[string]RuntimeClassPropertyDef)
		if classDef.Fields == nil {
			classDef.Fields = make(map[string]RuntimeClassFieldDef)
		}
	}

	lowerPropertyName := strings.ToLower(trimmedPropertyName)
	propertyDef, exists := classDef.Properties[lowerPropertyName]
	if !exists {
		propertyDef = RuntimeClassPropertyDef{Name: trimmedPropertyName}
	}
	propertyDef.IsPublic = isPublic

	switch op {
	case OpRegisterClassPropertyGet:
		propertyDef.HasGet = true
		propertyDef.GetTarget = accessorTarget
		propertyDef.GetParamCount = paramCount
	case OpRegisterClassPropertyLet:
		propertyDef.HasLet = true
		propertyDef.LetTarget = accessorTarget
		propertyDef.LetParamCount = paramCount
	case OpRegisterClassPropertySet:
		propertyDef.HasSet = true
		propertyDef.SetTarget = accessorTarget
		propertyDef.SetParamCount = paramCount
	}

	classDef.Properties[lowerPropertyName] = propertyDef
	vm.runtimeClasses[lowerClassName] = classDef
	vm.bumpRuntimeClassVersion()
}

// resolveRuntimeClassMethod resolves one class method target by object instance and method name.
func (vm *VM) resolveRuntimeClassMethod(target Value, methodName string, requirePublic bool) (Value, bool) {
	if target.Type != VTObject {
		return Value{Type: VTEmpty}, false
	}
	instance, exists := vm.runtimeClassItems[target.Num]
	if !exists || instance == nil {
		return Value{Type: VTEmpty}, false
	}
	classDef, exists := vm.runtimeClasses[strings.ToLower(strings.TrimSpace(instance.ClassName))]
	if !exists || classDef.Methods == nil {
		return Value{Type: VTEmpty}, false
	}
	methodDef, exists := classDef.Methods[strings.ToLower(strings.TrimSpace(methodName))]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	if requirePublic && !methodDef.IsPublic {
		return Value{Type: VTEmpty}, false
	}
	return methodDef.Target, true
}

// getActiveClassMemberValue reads one member from the currently executing class instance.
func (vm *VM) getActiveClassMemberValue(memberName string) Value {
	if vm.activeClassObjectID == 0 {
		return Value{Type: VTEmpty}
	}
	instance, exists := vm.runtimeClassItems[vm.activeClassObjectID]
	if !exists || instance == nil {
		return Value{Type: VTEmpty}
	}
	value, exists := instance.Members[strings.ToLower(strings.TrimSpace(memberName))]
	if !exists {
		return Value{Type: VTEmpty}
	}
	return value
}

// getClassMemberValueByObjectID reads one member from a specific class instance.
func (vm *VM) getClassMemberValueByObjectID(objectID int64, memberName string) Value {
	if objectID == 0 {
		return Value{Type: VTEmpty}
	}
	instance, exists := vm.runtimeClassItems[objectID]
	if !exists || instance == nil {
		return Value{Type: VTEmpty}
	}
	value, exists := instance.Members[strings.ToLower(strings.TrimSpace(memberName))]
	if !exists {
		return Value{Type: VTEmpty}
	}
	return value
}

// setActiveClassMemberValue writes one member on the currently executing class instance.
func (vm *VM) setActiveClassMemberValue(memberName string, value Value) {
	if vm.activeClassObjectID == 0 {
		return
	}
	instance, exists := vm.runtimeClassItems[vm.activeClassObjectID]
	if !exists || instance == nil {
		return
	}
	instance.Members[strings.ToLower(strings.TrimSpace(memberName))] = value
}

// setClassMemberValueByObjectID writes one member on a specific class instance.
func (vm *VM) setClassMemberValueByObjectID(objectID int64, memberName string, value Value) {
	if objectID == 0 {
		return
	}
	instance, exists := vm.runtimeClassItems[objectID]
	if !exists || instance == nil {
		return
	}
	instance.Members[strings.ToLower(strings.TrimSpace(memberName))] = value
}

// decrementObjectRefCount decrements the reference count of a VTObject and marks it
// for termination if the count reaches zero. The actual Class_Terminate call will be
// queued and executed during the next available opportunity in the VM loop.
func (vm *VM) decrementObjectRefCount(obj Value) {
	if obj.Type != VTObject {
		return
	}
	instance, exists := vm.runtimeClassItems[obj.Num]
	if !exists || instance == nil {
		return
	}
	if instance.terminated {
		return // Already terminated; skip.
	}
	instance.refCount--
	if instance.refCount <= 0 {
		instance.refCount = 0 // Ensure non-negative for safety.
		// Mark for termination; the actual termination will happen via
		// prepareClassTerminateCall during cleanup or when explicitly triggered.
		instance.terminated = true
	}
}

// incrementObjectRefCount increments the reference count when a VTObject is assigned to a new slot.
func (vm *VM) incrementObjectRefCount(obj Value) {
	if obj.Type != VTObject {
		return
	}
	instance, exists := vm.runtimeClassItems[obj.Num]
	if !exists || instance == nil {
		return
	}
	if !instance.terminated {
		instance.refCount++
	}
}

// resolveRuntimeClassField gets one direct class field value by object instance.
func (vm *VM) resolveRuntimeClassField(target Value, fieldName string, requirePublic bool) (Value, bool) {
	if target.Type != VTObject {
		return Value{Type: VTEmpty}, false
	}
	instance, exists := vm.runtimeClassItems[target.Num]
	if !exists || instance == nil {
		return Value{Type: VTEmpty}, false
	}
	classDef, exists := vm.runtimeClasses[strings.ToLower(strings.TrimSpace(instance.ClassName))]
	if !exists || classDef.Fields == nil {
		return Value{Type: VTEmpty}, false
	}
	fieldDef, exists := classDef.Fields[strings.ToLower(strings.TrimSpace(fieldName))]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	if requirePublic && !fieldDef.IsPublic {
		return Value{Type: VTEmpty}, false
	}
	value, exists := instance.Members[strings.ToLower(strings.TrimSpace(fieldName))]
	if !exists {
		return Value{Type: VTEmpty}, true
	}
	return value, true
}

// assignRuntimeClassField writes one direct class field by object instance.
func (vm *VM) assignRuntimeClassField(target Value, fieldName string, assigned Value, requirePublic bool) bool {
	if target.Type != VTObject {
		return false
	}
	instance, exists := vm.runtimeClassItems[target.Num]
	if !exists || instance == nil {
		return false
	}
	classDef, exists := vm.runtimeClasses[strings.ToLower(strings.TrimSpace(instance.ClassName))]
	if !exists || classDef.Fields == nil {
		return false
	}
	fieldDef, exists := classDef.Fields[strings.ToLower(strings.TrimSpace(fieldName))]
	if !exists {
		return false
	}
	if requirePublic && !fieldDef.IsPublic {
		return false
	}
	instance.Members[strings.ToLower(strings.TrimSpace(fieldName))] = assigned
	return true
}

// resolveRuntimeClassPropertyGet resolves one class Property Get accessor by name and argument count.
func (vm *VM) resolveRuntimeClassPropertyGet(target Value, propertyName string, argCount int, requirePublic bool) (Value, bool) {
	if target.Type != VTObject {
		return Value{Type: VTEmpty}, false
	}
	instance, exists := vm.runtimeClassItems[target.Num]
	if !exists || instance == nil {
		return Value{Type: VTEmpty}, false
	}
	classDef, exists := vm.runtimeClasses[strings.ToLower(strings.TrimSpace(instance.ClassName))]
	if !exists || classDef.Properties == nil {
		return Value{Type: VTEmpty}, false
	}
	propertyDef, exists := classDef.Properties[strings.ToLower(strings.TrimSpace(propertyName))]
	if !exists || !propertyDef.HasGet {
		return Value{Type: VTEmpty}, false
	}
	if requirePublic && !propertyDef.IsPublic {
		return Value{Type: VTEmpty}, false
	}
	if propertyDef.GetParamCount != argCount {
		return Value{Type: VTEmpty}, false
	}
	return propertyDef.GetTarget, true
}

// resolveRuntimeClassPropertyLet resolves one class Property Let accessor.
func (vm *VM) resolveRuntimeClassPropertyLet(target Value, propertyName string, argCount int, requirePublic bool) (Value, bool) {
	return vm.resolveRuntimeClassPropertySet(target, propertyName, argCount, false, false, requirePublic)
}

// resolveRuntimeClassPropertySet resolves one class Property Let/Set accessor by name and assignment mode.
func (vm *VM) resolveRuntimeClassPropertySet(target Value, propertyName string, argCount int, preferSet bool, strictSet bool, requirePublic bool) (Value, bool) {
	if target.Type != VTObject {
		return Value{Type: VTEmpty}, false
	}
	instance, exists := vm.runtimeClassItems[target.Num]
	if !exists || instance == nil {
		return Value{Type: VTEmpty}, false
	}
	classDef, exists := vm.runtimeClasses[strings.ToLower(strings.TrimSpace(instance.ClassName))]
	if !exists || classDef.Properties == nil {
		return Value{Type: VTEmpty}, false
	}
	propertyDef, exists := classDef.Properties[strings.ToLower(strings.TrimSpace(propertyName))]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	if requirePublic && !propertyDef.IsPublic {
		return Value{Type: VTEmpty}, false
	}

	if strictSet {
		if propertyDef.HasSet && propertyDef.SetParamCount == argCount {
			return propertyDef.SetTarget, true
		}
		return Value{Type: VTEmpty}, false
	}

	if preferSet {
		if propertyDef.HasSet && propertyDef.SetParamCount == argCount {
			return propertyDef.SetTarget, true
		}
		if propertyDef.HasLet && propertyDef.LetParamCount == argCount {
			return propertyDef.LetTarget, true
		}
		return Value{Type: VTEmpty}, false
	}

	if propertyDef.HasLet && propertyDef.LetParamCount == argCount {
		return propertyDef.LetTarget, true
	}
	if propertyDef.HasSet && propertyDef.SetParamCount == argCount {
		return propertyDef.SetTarget, true
	}
	return Value{Type: VTEmpty}, false
}

// beginUserSubCall sets one call frame and switches VM execution to one VTUserSub entry point.
// collectByRefsAndUnwrap processes VTArgRef values in args, builds a ByRef write-back list, and
// replaces each VTArgRef entry with the actual variable value read from the slot.
// byRefMask identifies which parameter slots are declared ByRef (bit i set = param i is ByRef).
// Only VTArgRef args whose corresponding param is ByRef get a write-back entry recorded.
func (vm *VM) collectByRefsAndUnwrap(args []Value, byRefMask uint64) []byRefWriteback {
	var byRefs []byRefWriteback
	for i, arg := range args {
		if arg.Type != VTArgRef {
			continue
		}
		isGlobal := arg.ArgRefIsGlobal()
		isLocal := arg.ArgRefIsLocal()
		isClassMember := arg.ArgRefIsClassMember()
		idx := arg.ArgRefIdx()
		// Record write-back only when the parameter is declared ByRef.
		if i < 64 && (byRefMask>>uint(i))&1 == 1 {
			wb := byRefWriteback{calleeParamIdx: i}
			if isClassMember {
				wb.isClassMember = true
				wb.callerBoundObj = vm.activeClassObjectID
				wb.callerMember = arg.Str
			} else {
				wb.isGlobal = isGlobal
				wb.callerIdx = idx
			}
			byRefs = append(byRefs, wb)
		}
		// Unwrap the VTArgRef to the actual current variable value.
		var rawVal Value
		if isClassMember {
			rawVal = vm.getClassMemberValueByObjectID(vm.activeClassObjectID, arg.Str)
		} else if isGlobal {
			if idx >= 0 && idx < len(vm.Globals) {
				rawVal = vm.Globals[idx]
			}
		} else if isLocal {
			slot := vm.fp + idx
			if slot >= 0 && slot < StackSize {
				rawVal = vm.stack[slot]
			}
		}
		args[i] = resolveCallable(vm, rawVal)
	}
	return byRefs
}

// nativeByRefMask returns the ByRef parameter bitmask for native call targets.
func (vm *VM) nativeByRefMask(targetID int64, member string) uint64 {
	if targetID == nativeRequestBinaryReadMethod && member == "" {
		return 1
	}
	if targetID == nativeObjectRequest && strings.EqualFold(member, "BinaryRead") {
		return 1
	}
	return 0
}

// applyByRefWritebacksFromArgs writes native call argument updates back to caller slots.
func (vm *VM) applyByRefWritebacksFromArgs(byRefs []byRefWriteback, args []Value) {
	for _, wb := range byRefs {
		if wb.calleeParamIdx < 0 || wb.calleeParamIdx >= len(args) {
			continue
		}
		writeVal := args[wb.calleeParamIdx]
		if wb.isClassMember {
			vm.setClassMemberValueByObjectID(wb.callerBoundObj, wb.callerMember, writeVal)
			continue
		}
		if wb.isGlobal {
			if wb.callerIdx >= 0 && wb.callerIdx < len(vm.Globals) {
				vm.Globals[wb.callerIdx] = writeVal
			}
			continue
		}
		callerSlot := vm.fp + wb.callerIdx
		if callerSlot >= 0 && callerSlot < StackSize {
			vm.stack[callerSlot] = writeVal
		}
	}
}

// bytesToVBByteString converts raw bytes to a VB-compatible byte-string representation.
// Each source byte maps to one codepoint in range 0..255 so byte-oriented built-ins
// (LenB/MidB/InStrB/AscB) can round-trip binary payloads without UTF-8 shrinking.
func bytesToVBByteString(data []byte) string {
	if len(data) == 0 {
		return ""
	}
	runes := make([]rune, len(data))
	for i := 0; i < len(data); i++ {
		runes[i] = rune(data[i])
	}
	return string(runes)
}

// vbByteStringToBytes converts one VB byte-string back to raw bytes.
// Each rune in range 0..255 maps back to its original byte value.
func vbByteStringToBytes(text string) []byte {
	if text == "" {
		return []byte{}
	}
	data := make([]byte, 0, len(text))
	for _, r := range text {
		if r <= 0xFF {
			data = append(data, byte(r))
			continue
		}
		data = append(data, byte('?'))
	}
	return data
}

// beginUserSubCall sets one call frame and switches VM execution to one VTUserSub entry point.
// An optional byRefs slice records ByRef parameter write-backs to perform on OpRet.
func (vm *VM) beginUserSubCall(target Value, args []Value, discardReturn bool, boundObjectID int64, optByRefs ...[]byRefWriteback) bool {
	if target.Type != VTUserSub {
		return false
	}

	var wb []byRefWriteback
	if len(optByRefs) > 0 {
		wb = optByRefs[0]
	}
	vm.callStack = append(vm.callStack, CallFrame{
		callee:              target,
		returnIP:            vm.ip,
		oldFP:               vm.fp,
		oldSP:               vm.sp,
		boundObj:            vm.activeClassObjectID,
		discard:             discardReturn,
		byRefs:              wb,
		savedOnResumeNext:   vm.onResumeNext,
		savedSkipToNextStmt: vm.skipToNextStmt,
		savedStmtSP:         vm.stmtSP,
	})
	vm.activeClassObjectID = boundObjectID
	vm.onResumeNext = false
	vm.skipToNextStmt = false
	vm.fp = vm.sp + 1
	paramCount := target.UserSubParamCount()
	localCount := target.UserSubLocalCount()
	if localCount < paramCount {
		localCount = paramCount
	}

	frameLast := vm.fp + localCount - 1
	if localCount > 0 && frameLast >= StackSize {
		vm.raise(vbscript.StackOverflow, "Stack overflow")
	}

	for i := 0; i < localCount; i++ {
		vm.stack[vm.fp+i] = Value{Type: VTEmpty}
	}

	copyCount := len(args)
	if copyCount > paramCount {
		copyCount = paramCount
	}
	for i := 0; i < copyCount; i++ {
		vm.stack[vm.fp+i] = args[i]
	}

	if localCount > 0 {
		vm.sp = frameLast
	} else {
		vm.sp = vm.fp - 1
	}
	vm.ip = int(target.Num)
	return true
}

// newRuntimeClassInstance allocates a new VTObject instance with deterministic dynamic ID.
func (vm *VM) newRuntimeClassInstance(className string) Value {
	trimmedName := strings.TrimSpace(className)
	if trimmedName == "" {
		vm.raise(vbscript.ClassTypeIsNotDefined, "Class/Type is not defined")
		return Value{Type: VTEmpty}
	}

	if strings.EqualFold(trimmedName, "RegExp") || strings.EqualFold(trimmedName, "VBScript.RegExp") {
		return vm.newRegExpObject()
	}

	classDef, exists := vm.runtimeClasses[strings.ToLower(trimmedName)]
	if !exists {
		vm.raise(vbscript.ClassTypeIsNotDefined, "Class/Type is not defined: "+trimmedName)
		return Value{Type: VTEmpty}
	}

	instanceID := vm.nextDynamicClassID
	vm.nextDynamicClassID++
	vm.runtimeClassItems[instanceID] = &RuntimeClassInstance{
		ClassName:  classDef.Name,
		Members:    make(map[string]Value),
		refCount:   1, // Initial reference from creation
		terminated: false,
	}
	// Track creation order so Class_Terminate fires in reverse-construction order at cleanup.
	vm.classInstanceOrder = append(vm.classInstanceOrder, instanceID)
	for fieldKey, fieldDef := range classDef.Fields {
		if len(fieldDef.Dims) > 0 {
			vm.runtimeClassItems[instanceID].Members[fieldKey] = ValueFromVBArray(buildDimArray(fieldDef.Dims))
		} else {
			vm.runtimeClassItems[instanceID].Members[fieldKey] = Value{Type: VTEmpty}
		}
	}

	return Value{Type: VTObject, Num: instanceID, Str: classDef.Name}
}

func (vm *VM) pop() Value {
	if vm.sp < 0 {
		vm.raise(vbscript.InternalError, "Stack underflow")
		return Value{Type: VTEmpty}
	}
	v := vm.stack[vm.sp]
	vm.sp--
	return v
}

func (vm *VM) raise(code vbscript.VBSyntaxErrorCode, msg string) {
	description := strings.TrimSpace(msg)
	if description == "" {
		description = code.String()
	}

	vme := &VMError{
		Code:           code,
		Line:           vm.lastLine,
		Column:         vm.lastColumn,
		Msg:            description,
		ASPCode:        int(code),
		ASPDescription: description,
		Category:       "VBScript runtime",
		Description:    description,
		Number:         vbscript.HRESULTFromVBScriptCode(code),
		Source:         "VBScript runtime error",
	}

	vm.errSetFromVMError(vme)

	if vm.onResumeNext || vm.executeGlobalResumeGuard {
		vm.lastError = vme
		vm.skipToNextStmt = true
		return
	}

	// Walk the call stack looking for the nearest ancestor frame whose caller had
	// On Error Resume Next active (savedOnResumeNext == true). When found, unwind
	// the call stack to that frame so execution resumes in the caller at the
	// statement following the call that raised the error. This mirrors classic
	// VBScript behaviour: an unhandled error propagates up until it reaches a
	// procedure scope that has On Error Resume Next active.
	for i := len(vm.callStack) - 1; i >= 0; i-- {
		frame := vm.callStack[i]
		if !frame.savedOnResumeNext {
			continue
		}
		// Decrement ref counts for locals in frames being discarded (error unwind).
		// Walk from innermost to outermost discarded frame, shrinking fp/sp as we pop.
		curFP := vm.fp
		curSP := vm.sp
		for j := len(vm.callStack) - 1; j >= i; j-- {
			for k := curFP; k <= curSP; k++ {
				if k >= 0 && k < StackSize {
					vm.decrementObjectRefCount(vm.stack[k])
				}
			}
			curSP = vm.callStack[j].oldSP
			curFP = vm.callStack[j].oldFP
		}
		// Unwind to the caller's context, resuming after the call site.
		vm.callStack = vm.callStack[:i]
		vm.sp = curSP
		vm.fp = curFP
		vm.ip = frame.returnIP
		vm.activeClassObjectID = frame.boundObj
		vm.onResumeNext = frame.savedOnResumeNext // restore caller's On Error state
		vm.lastError = vme
		return
	}

	panic(vme)
}

func (vm *VM) asFloat(v Value) float64 {
	if v.Type == VTDouble {
		return v.Flt
	}
	return float64(v.Num)
}

// asInt converts a VM value into an integer using VBScript-like coercion fallback.
func (vm *VM) asInt(v Value) int {
	switch v.Type {
	case VTDouble:
		return int(v.Flt)
	case VTString:
		var parsed int
		_, err := fmt.Sscanf(v.Str, "%d", &parsed)
		if err == nil {
			return parsed
		}
		return 0
	default:
		return int(v.Num)
	}
}

// asBool converts a VM value into boolean according to VBScript-like coercion rules.
func (vm *VM) asBool(v Value) bool {
	switch v.Type {
	case VTBool:
		return v.Num != 0
	case VTInteger:
		return v.Num != 0
	case VTDouble:
		return v.Flt != 0
	case VTString:
		text := strings.TrimSpace(v.Str)
		if strings.EqualFold(text, "true") {
			return true
		}
		if strings.EqualFold(text, "false") || text == "" || text == "0" {
			return false
		}
		return true
	default:
		return false
	}
}

func (vm *VM) isTruthy(v Value) bool {
	switch v.Type {
	case VTBool, VTInteger:
		return v.Num != 0
	case VTDouble:
		return v.Flt != 0
	case VTString:
		return v.Str != "" && v.Str != "0"
	}
	return false
}

func (vm *VM) StackTop() Value {
	if vm.sp < 0 {
		return Value{Type: VTEmpty}
	}
	return vm.stack[vm.sp]
}

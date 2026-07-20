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
 // Package axonvm provides the core functions of AxonASP Server, it provides the virtual machine and the runtime environment for executing VBScript and JScript code, as well as the built-in objects and libraries used by ASP pages.
package axonvm

import (
	"bufio"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"maps"
	"math"
	"os"
	"path/filepath"
	"runtime"
	"runtime/debug"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

	"golang.org/x/text/collate"
	"golang.org/x/text/language"

	"g3pix.com.br/axonasp/axonvm/asp"
	"g3pix.com.br/axonasp/jscript"
	"g3pix.com.br/axonasp/vbscript"
)

const StackSize = 4096
const jsBackJumpLimit = 1000000

const staticObjectProgIDPrefix = "__AXON_STATIC_OBJECT_PROGID__:"

// ExecutionMode indicates the context in which ASP code is executing.
// Used to control caching behavior - interactive modes bypass cache to prevent stalls.
type ExecutionMode uint8

const (
	ExecutionModeServer ExecutionMode = iota // Normal HTTP/FastCGI server execution - uses cache
	ExecutionModeCLI                         // Interactive CLI REPL execution - bypasses cache
	ExecutionModeTUI                         // Interactive TUI execution - bypasses cache
	ExecutionModeEval                        // Dynamic eval() execution - bypasses cache
)

// EngineMode indicates the language mode for the current execution.
type EngineMode uint8

const (
	EngineModeDefault    EngineMode = iota // Standard ASP with delimiters (<% %>)
	EngineModeVBScript                     // Pure VBScript (no delimiters)
	EngineModeJavaScript                   // Pure JavaScript (no delimiters)
)

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
	nativeResponseCookiesKeyMethod
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
	nativeObjectSessionContentsKeyMethod
	nativeObjectSessionStaticObjectsKeyMethod
	nativeObjectApplicationContentsKeyMethod
	nativeObjectApplicationStaticObjectsKeyMethod
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
	Name       string
	IsPublic   bool
	WithEvents bool
	Dims       []int // nil for plain variables; non-nil upper bounds for fixed-size arrays
	Type       ValueType
	ClassType  string
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
	Events     map[string]bool
	Interfaces []string
}

// EventObserver represents one registered event handler.
type EventObserver struct {
	Handler     Value  // VTUserSub or VTNativeObject (for future extensibility)
	ContainerID int64  // ID of the RuntimeClassInstance that owns the handler (0 for global)
	Prefix      string // The "objname_" prefix
}

// RuntimeClassInstance stores one allocated class instance state.
type RuntimeClassInstance struct {
	ClassName       string
	Members         map[string]Value
	Observers       map[string][]EventObserver // Key: EventName (lowercase)
	WithEventsNames map[string]bool            // Key: MemberName (lowercase) that has WithEvents
	refCount        int                        // Reference count for deterministic termination
	terminated      bool                       // True if Class_Terminate has already been called
}

type callMemberICKind uint8

const (
	callMemberICNone callMemberICKind = iota
	callMemberICClassMethod
	callMemberICClassPropertyGet
	callMemberICNativeMember
)

type callMemberICEntry struct {
	kind         callMemberICKind
	expectedType ValueType
	expectedNum  int64
	target       Value
	nativeMember string
}

// InlineCacheSlot holds monomorphic inline cache state for a single JScript
// property access site. It is stored in a VM-local flat array indexed by the
// ICNodeID assigned during compilation, isolating IC state from the shared
// AST/bytecode and eliminating concurrent mutation races.
type InlineCacheSlot struct {
	ShapeID uint32
	Slot    uint16
	Flags   uint16
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
	requestCollectionValueItems    map[int64]asp.RequestCollectionValue
	aspErrorItems                  map[int64]*asp.ASPError
	g3mdItems                      map[int64]*G3MD
	g3dateItems                    map[int64]*G3Date
	g3searchItems                  map[int64]*G3Search
	g3stringBuilderItems           map[int64]*G3StringBuilder
	g3testItems                    map[int64]*G3Test
	g3cryptoItems                  map[int64]*G3Crypto
	g3jsonItems                    map[int64]*G3JSON
	g3httpItems                    map[int64]*G3HTTP
	g3mailItems                    map[int64]*G3Mail
	g3imageItems                   map[int64]*G3Image
	g3imageCanvasItems             map[int64]*G3ImageCanvas
	g3imageFontItems               map[int64]*G3ImageFont
	g3imagePenItems                map[int64]*G3ImagePen
	g3filesItems                   map[int64]*G3Files
	g3templateItems                map[int64]*G3Template
	g3zipItems                     map[int64]*G3Zip
	g3zlibItems                    map[int64]*G3ZLIB
	g3tarItems                     map[int64]*G3TAR
	g3zstdItems                    map[int64]*G3ZSTD
	g3fcItems                      map[int64]*G3FC
	fileIOItems                    map[int]*os.File
	fileIOBufReaders               map[int]*bufio.Reader
	fileIOBufWriters               map[int]*bufio.Writer
	g3axonliveItems                map[int64]*G3AXONLIVE
	g3axonliveProxyItems           map[int64]*G3ALComponentProxy
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
	wscriptEnvironmentItems        map[int64]*WshEnvironment
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
	pdfDocItems                    map[int64]*PdfDocument
	pdfPageItems                   map[int64]*PdfPage
	pdfCanvasItems                 map[int64]*PdfCanvas
	pdfFontItems                   map[int64]*PdfFont
	pdfPagesItems                  map[int64]*PdfPagesCollection
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
	jsRegExpItems                  map[int64]*jsRegExpObject
	regExpMatchesCollectionItems   map[int64]*regExpMatchesCollection
	regExpMatchItems               map[int64]*regExpMatch
	regExpSubMatchesItems          map[int64]*regExpSubMatches
	regExpSubMatchValueItems       map[int64]*regExpSubMatchValue
	dictionaryItems                map[int64]*scriptingDictionary
	collectionItems                map[int64]*vbsCollection
	collectionEnumeratorItems      map[int64]*vbsCollectionEnumerator
	nativeObjectProxies            map[int64]nativeObjectProxy
	jsObjectItems                  map[int64]map[string]Value
	jsObjectKeyOrder               map[int64][]string
	jsObjectSlots                  map[int64][]Value
	jsObjectSlotIndex              map[int64]map[string]uint16
	jsObjectShape                  map[int64]uint32
	jsShapeSlots                   map[uint32][]string
	jsShapeBySignature             map[string]uint32
	jsNextShapeID                  uint32
	jsObjectStateItems             map[int64]jsObjectState
	jsSymbolStateItems             map[int64]jsObjectState
	jsPropertyItems                map[int64]map[string]jsPropertyDescriptor
	jsFunctionItems                map[int64]*jsFunctionObject
	jsForInItems                   map[int]*jsForInEnumerator
	jsForOfItems                   map[int]*jsForOfEnumerator
	jsEnvItems                     map[int64]*jsEnvFrame
	jsArgumentsItems               map[int64]*jsArgumentsBinding
	jsSetItems                     map[int64]map[string]Value
	jsMapItems                     map[int64]map[string]Value
	jsWeakRefItems                 map[int64]*jsWeakRef
	jsFinalizationRegistryItems    map[int64]*jsFinalizationRegistry
	jsArrayIterators               map[int64]*jsArrayIterator
	jsStringIterators              map[int64]*jsStringIterator
	jsRegExpStringIterators        map[int64]*jsRegExpStringIterator
	jsArrayBuffers                 map[int64][]byte       // backing byte slices for ArrayBuffer objects
	jsSharedArrayBuffers           map[int64][]byte       // backing byte slices for SharedArrayBuffer objects
	jsModuleInstances              map[string]*jsEnvFrame // Subphase 8.3: Request-local module registry
	jsModuleLoading                map[string]struct{}    // Tracks modules currently executing for circular import handling
	jsIntlDateTimeFormatItems      map[int64]*jsIntlDateTimeFormatObject
	jsIntlNumberFormatItems        map[int64]*jsIntlNumberFormatObject
	jsIntlCollatorItems            map[int64]*jsIntlCollatorObject
	jsIntlPluralRulesItems         map[int64]*jsIntlPluralRulesObject
	jsIntlRelativeTimeFormatItems  map[int64]*jsIntlRelativeTimeFormatObject
	jsPromiseItems                 map[int64]*jsPromiseObject
	jsGeneratorItems               map[int64]*jsGeneratorObject
	jsProxyItems                   map[int64]*jsProxyObject
	jsStreamHookItems              map[int64]*jsNodeStreamHookResource
	jsAsyncFSReadResults           chan jsAsyncFSReadResult
	jsTimerItems                   map[int64]*jsTimerItem  // active setTimeout/setInterval handles
	jsTimerResultQueue             chan jsTimerFiredResult // goroutine -> VM thread timer completions
	jsImmediateQueue               []jsImmediateItem       // setImmediate callbacks
	jsNextTickQueue                []jsNextTickItem        // process.nextTick callbacks
	jsPumpingNodeTasks             bool                    // re-entrancy guard for jsPumpNodeAsyncTasks
	jsMicrotaskQueue               []func()
	jsProcessingMicrotasks         bool
	jsSymbolGlobalRegistry         map[string]Value // Symbol.for global registry: description -> Symbol Value
	jsRegisteredSymbolIDs          map[int64]struct{}
	jsBufferItems                  map[int64]*jsBuffer // Node.js Buffer instances
	jsProcessObjectID              int64               // ID of the process global object
	jsNextSymbolID                 int64
	jsRootEnvID                    int64                 // ID of the JScript root environment frame
	jsStrictMode                   bool                  // Current strict mode state
	jsFunctionStrictModes          map[int64]bool        // Maps function IDs to strict mode status
	jsBlockScopes                  []map[string]Value    // Stack of block-scoped (let/const) variable values
	jsBlockScopeConst              []map[string]struct{} // Per block scope: which names are declared const
	jsBlockScopeTDZ                []map[string]struct{} // Per block scope: which names are in TDZ (const before init)
	jsBlockScopeDepth              int                   // Current block scope depth
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

	stack []Value
	sp    int
	ip    int
	fp    int
	// localTypes stores the declared VB6 type (if any) for each stack slot.
	// 0 (VTEmpty) = no declared type (Variant). Non-zero = declared type for type enforcement.
	localTypes [StackSize]ValueType
	// globalTypes stores the declared VB6 type (if any) for each global slot.
	// 0 (VTEmpty) = no declared type (Variant). Non-zero = declared type for type enforcement.
	globalTypes []ValueType
	// globalClassTypes stores the declared Class/Interface name for global VTObject slots.
	globalClassTypes map[uint16]string
	// globalWithEvents stores the indices of global variables declared WithEvents.
	globalWithEvents map[uint16]bool
	// funcLocalTypes maps function entry point bytecode offsets to local variable type
	// declarations for VB6 As Type support. The inner map key is the local slot offset,
	// the value is the declared ValueType (VTEmpty = Variant/no constraint).
	funcLocalTypes map[int]map[int]ValueType
	// funcLocalClassTypes maps function entry point to local slot Class/Interface names.
	funcLocalClassTypes map[int]map[int]string
	callStack           []CallFrame
	jsCallStack         []jsCallFrame
	// withStack holds the subject object for each active With...End With block.
	// OpWithEnter appends, OpWithLeave shrinks, OpWithLoad peeks at the top.
	withStack         []Value
	jsTryStack        []int
	jsErrStack        []Value
	jsActiveEnvID     int64
	jsThisValue       Value
	startTime         time.Time
	jsNewTarget       Value
	consoleTimerItems map[string]time.Time

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

	// executionMode tracks whether this VM is running in an interactive context (CLI/TUI/eval)
	// where caching should be bypassed to prevent stalls. Server mode uses cache normally.
	executionMode ExecutionMode
	// engineMode indicates the language mode for the current execution (ASP, VBScript, JavaScript).
	engineMode EngineMode

	// Runtime Options
	optionCompare        int               // 0: Binary, 1: Text
	optionExplicit       bool              // Mirrors compiler Option Explicit for dynamic execution APIs.
	collator             *collate.Collator // cached locale-aware collator for Option Compare Text
	collatorLCID         int               // LCID for which the collator was built
	globalNames          []string          // Global symbol names aligned with Globals slot order.
	globalNamesHash      uint64            // Fingerprint of globalNames for fast dynamic-cache scope validation.
	globalZeroArgFuncs   map[string]bool   // Known zero-arg global Functions for dynamic autocall compatibility.
	globalZeroArgSubs    map[string]bool   // Known zero-arg global Subs for bare-call compilation.
	runtimeClassVersion  uint64            // Monotonic version for runtime class metadata invalidation.
	declaredGlobals      map[string]bool   // Global Dim/Const declaration state for dynamic compilation.
	constGlobals         map[string]bool   // Global Const protection state for dynamic compilation.
	sourceName           string            // Source file path used for dynamic execution error reporting.
	sourceMap            SourceMap         // Sparse merged-to-original source line mapping for include-aware errors.
	dynamicProgramStarts map[uint64]int    // Per-VM start offsets for already-appended cached dynamic fragments.
	jsStringWorkBytes    int64             // Per-run cumulative bytes produced by JScript string operations.

	RecordDecls      []CompiledRecordDecl
	RecordDeclLookup map[string]int

	// funcParamDefaults maps function entry point -> per-parameter constant pool indices
	// for Optional parameter default values. -1 means no default for that parameter.
	funcParamDefaults map[int][]int

	baseBytecode         []byte
	baseConstants        []Value
	baseGlobals          []Value
	baseGlobalTypes      []ValueType
	baseOptionCompare    int
	baseOptionExplicit   bool
	baseGlobalNames      []string
	baseGlobalNamesLower []string
	baseGlobalNamesHash  uint64
	baseEngineMode       EngineMode

	baseGlobalZeroArgFuncs  map[string]bool
	baseGlobalZeroArgSubs   map[string]bool
	baseRuntimeClassVersion uint64
	globalNameIndex         map[string]int
	baseDeclared            map[string]bool
	baseConst               map[string]bool
	baseSourceName          string
	baseSourceMap           SourceMap

	argBuffer     []Value
	indexBuffer   []Value
	combineBuffer []Value
	// stringWorkBuffer is a per-VM scratch byte slice used by concatValues to build
	// concatenated strings without triggering an extra intermediate Go string allocation
	// from the built-in '+' operator. It is reset to length zero before each use and
	// cleared after each VM run so the backing array can be collected by the GC when
	// the VM is returned to its pool.
	stringWorkBuffer []byte

	pooledFrom      *vmProgramPool
	pooledSlot      chan struct{}
	comInitialized  bool
	comThreadLocked bool
	staTaskChan     chan func()
	quitSTA         chan struct{}
	parentVM        *VM
	inSTATask       bool
	runDepth        int

	// jsDirectCallHaltIP is a lazily allocated OpHalt trampoline used by
	// internal direct callback execution paths that must stop at function return.
	jsDirectCallHaltIP int

	// cloneForExecuteLocalCount tracks cloneForExecuteLocal invocations for
	// perf regression tests in clone-sensitive JScript paths.
	cloneForExecuteLocalCount uint64

	nextCallMemberICID uint32
	callMemberIC       map[uint32]callMemberICEntry
	icState            []InlineCacheSlot // VM-local inline cache state indexed by ICNodeID
}

func (vm *VM) mapRuntimeLocation(line int, column int) (string, int, int) {
	file := vm.sourceName
	mappedLine := line
	if mappedFile, resolvedLine, ok := vm.sourceMap.ResolveLine(line); ok {
		if strings.TrimSpace(mappedFile) != "" {
			file = mappedFile
		}
		if resolvedLine > 0 {
			mappedLine = resolvedLine
		}
	}
	return file, mappedLine, column
}

func (vm *VM) mappedCurrentLocation() (string, int) {
	file, line, _ := vm.mapRuntimeLocation(vm.lastLine, vm.lastColumn)
	return file, line
}

func (vm *VM) newMappedAxonASPError(code AxonASPErrorCode, err error, description string) *AxonASPError {
	file, line := vm.mappedCurrentLocation()
	return NewAxonASPError(code, err, description, file, line)
}

// NewVM creates and initializes a new VM instance.
func NewVM(bytecode []byte, constants []Value, globalCount int) *VM {
	vm := &VM{
		bytecode:                       bytecode,
		constants:                      constants,
		Globals:                        make([]Value, globalCount),
		globalTypes:                    make([]ValueType, globalCount),
		globalClassTypes:               make(map[uint16]string),
		globalWithEvents:               make(map[uint16]bool),
		funcLocalTypes:                 make(map[int]map[int]ValueType),
		funcLocalClassTypes:            make(map[int]map[int]string),
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
		baseGlobalZeroArgSubs:          make(map[string]bool),
		globalNames:                    make([]string, 0, globalCount),
		globalZeroArgFuncs:             make(map[string]bool),
		globalZeroArgSubs:              make(map[string]bool),
		dynamicProgramStarts:           make(map[uint64]int, 32),
		globalNameIndex:                make(map[string]int, globalCount),
		RecordDeclLookup:               make(map[string]int),
		funcParamDefaults:              make(map[int][]int),
		startTime:                      time.Now(),
		argBuffer:                      make([]Value, 0, 16),
		indexBuffer:                    make([]Value, 0, 16),
		combineBuffer:                  make([]Value, 0, 16),
		stringWorkBuffer:               make([]byte, 0, 256),
		withStack:                      make([]Value, 0, 8),
		engineMode:                     EngineModeDefault,
		nextDynamicNativeID:            20000,
		nextDynamicClassID:             60000,
		responseCookieItems:            make(map[int64]string),
		requestCollectionValueItems:    make(map[int64]asp.RequestCollectionValue),
		aspErrorItems:                  make(map[int64]*asp.ASPError),
		g3mdItems:                      make(map[int64]*G3MD),
		g3dateItems:                    make(map[int64]*G3Date),
		g3searchItems:                  make(map[int64]*G3Search),
		g3stringBuilderItems:           make(map[int64]*G3StringBuilder),
		g3testItems:                    make(map[int64]*G3Test),
		g3cryptoItems:                  make(map[int64]*G3Crypto),
		g3jsonItems:                    make(map[int64]*G3JSON),
		g3httpItems:                    make(map[int64]*G3HTTP),
		g3mailItems:                    make(map[int64]*G3Mail),
		g3imageItems:                   make(map[int64]*G3Image),
		g3imageCanvasItems:             make(map[int64]*G3ImageCanvas),
		g3imageFontItems:               make(map[int64]*G3ImageFont),
		g3imagePenItems:                make(map[int64]*G3ImagePen),
		g3filesItems:                   make(map[int64]*G3Files),
		g3templateItems:                make(map[int64]*G3Template),
		g3zipItems:                     make(map[int64]*G3Zip),
		g3zlibItems:                    make(map[int64]*G3ZLIB),
		g3tarItems:                     make(map[int64]*G3TAR),
		fileIOItems:                    make(map[int]*os.File),
		fileIOBufReaders:               make(map[int]*bufio.Reader),
		fileIOBufWriters:               make(map[int]*bufio.Writer),
		g3axonliveItems:                make(map[int64]*G3AXONLIVE),
		g3axonliveProxyItems:           make(map[int64]*G3ALComponentProxy),
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
		wscriptEnvironmentItems:        make(map[int64]*WshEnvironment),
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
		pdfDocItems:                    make(map[int64]*PdfDocument),
		pdfPageItems:                   make(map[int64]*PdfPage),
		pdfCanvasItems:                 make(map[int64]*PdfCanvas),
		pdfFontItems:                   make(map[int64]*PdfFont),
		pdfPagesItems:                  make(map[int64]*PdfPagesCollection),
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
		jsRegExpItems:                  make(map[int64]*jsRegExpObject),
		regExpMatchesCollectionItems:   make(map[int64]*regExpMatchesCollection),
		regExpMatchItems:               make(map[int64]*regExpMatch),
		regExpSubMatchesItems:          make(map[int64]*regExpSubMatches),
		regExpSubMatchValueItems:       make(map[int64]*regExpSubMatchValue),
		dictionaryItems:                make(map[int64]*scriptingDictionary),
		collectionItems:                make(map[int64]*vbsCollection),
		collectionEnumeratorItems:      make(map[int64]*vbsCollectionEnumerator),
		nativeObjectProxies:            make(map[int64]nativeObjectProxy),
		jsObjectItems:                  make(map[int64]map[string]Value),
		jsObjectKeyOrder:               make(map[int64][]string),
		jsObjectSlots:                  make(map[int64][]Value),
		jsObjectSlotIndex:              make(map[int64]map[string]uint16),
		jsObjectShape:                  make(map[int64]uint32),
		jsShapeSlots:                   make(map[uint32][]string),
		jsShapeBySignature:             make(map[string]uint32),
		jsNextShapeID:                  1,
		jsObjectStateItems:             make(map[int64]jsObjectState),
		jsSymbolStateItems:             make(map[int64]jsObjectState),
		jsPropertyItems:                make(map[int64]map[string]jsPropertyDescriptor),
		jsFunctionItems:                make(map[int64]*jsFunctionObject),
		jsForInItems:                   make(map[int]*jsForInEnumerator),
		jsForOfItems:                   make(map[int]*jsForOfEnumerator),
		jsEnvItems:                     make(map[int64]*jsEnvFrame),
		jsArgumentsItems:               make(map[int64]*jsArgumentsBinding),
		jsSetItems:                     make(map[int64]map[string]Value),
		jsMapItems:                     make(map[int64]map[string]Value),
		jsWeakRefItems:                 make(map[int64]*jsWeakRef),
		jsFinalizationRegistryItems:    make(map[int64]*jsFinalizationRegistry),
		jsArrayIterators:               make(map[int64]*jsArrayIterator),
		jsStringIterators:              make(map[int64]*jsStringIterator),
		jsRegExpStringIterators:        make(map[int64]*jsRegExpStringIterator),
		jsArrayBuffers:                 make(map[int64][]byte),
		jsSharedArrayBuffers:           make(map[int64][]byte),
		jsModuleInstances:              make(map[string]*jsEnvFrame),
		jsModuleLoading:                make(map[string]struct{}),
		jsIntlDateTimeFormatItems:      make(map[int64]*jsIntlDateTimeFormatObject),
		jsIntlNumberFormatItems:        make(map[int64]*jsIntlNumberFormatObject),
		jsIntlCollatorItems:            make(map[int64]*jsIntlCollatorObject),
		jsIntlPluralRulesItems:         make(map[int64]*jsIntlPluralRulesObject),
		jsIntlRelativeTimeFormatItems:  make(map[int64]*jsIntlRelativeTimeFormatObject),
		jsPromiseItems:                 make(map[int64]*jsPromiseObject),
		jsGeneratorItems:               make(map[int64]*jsGeneratorObject),
		jsProxyItems:                   make(map[int64]*jsProxyObject),
		jsStreamHookItems:              make(map[int64]*jsNodeStreamHookResource),
		jsAsyncFSReadResults:           make(chan jsAsyncFSReadResult, jsAsyncFSReadResultQueueSize),
		jsTimerItems:                   make(map[int64]*jsTimerItem),
		jsTimerResultQueue:             make(chan jsTimerFiredResult, jsTimerResultQueueSize),
		jsImmediateQueue:               make([]jsImmediateItem, 0, 8),
		jsNextTickQueue:                make([]jsNextTickItem, 0, 8),
		jsSymbolGlobalRegistry:         make(map[string]Value),
		jsRegisteredSymbolIDs:          make(map[int64]struct{}),
		jsNextSymbolID:                 1,
		jsFunctionStrictModes:          make(map[int64]bool),
		jsBlockScopes:                  make([]map[string]Value, 0, 32),
		jsBlockScopeConst:              make([]map[string]struct{}, 0, 32),
		jsBlockScopeTDZ:                make([]map[string]struct{}, 0, 32),
		jsBlockScopeDepth:              0,
		jsStrictMode:                   false,
		errObject:                      asp.NewASPError(),

		runtimeClasses:     make(map[string]RuntimeClassDef),
		runtimeClassItems:  make(map[int64]*RuntimeClassInstance),
		classInstanceOrder: make([]int64, 0, 16),
		callStack:          make([]CallFrame, 0, 16),
		jsCallStack:        make([]jsCallFrame, 0, 16),
		jsTryStack:         make([]int, 0, 8),
		jsErrStack:         make([]Value, 0, 4),
		jsThisValue:        Value{Type: VTJSUndefined},
		consoleTimerItems:  make(map[string]time.Time),
		terminateCursor:    -1,
		jsDirectCallHaltIP: -1,
		nextCallMemberICID: 1,
		callMemberIC:       make(map[uint32]callMemberICEntry, 32),
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

	vm.startSTAWorker()
	runtime.SetFinalizer(vm, func(v *VM) {
		v.stopSTAWorker()
	})

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
	vm.sourceMap = compiler.sourceMap.Clone()
	vm.globalNames = append(vm.globalNames[:0], compiler.Globals.names...)
	vm.rebuildGlobalNameIndex()
	for i, name := range vm.globalNames {
		if strings.HasPrefix(name, "__static_") {
			if i < len(vm.Globals) && vm.Globals[i].Type == VTEmpty {
				vm.Globals[i].Interface = "static"
			}
		}
	}
	clear(vm.globalZeroArgFuncs)
	maps.Copy(vm.globalZeroArgFuncs, compiler.globalZeroArgFuncs)
	clear(vm.globalZeroArgSubs)
	maps.Copy(vm.globalZeroArgSubs, compiler.globalZeroArgSubs)
	maps.Copy(vm.declaredGlobals, compiler.declaredGlobals)
	maps.Copy(vm.constGlobals, compiler.constGlobals)
	// Apply VB6 As Type declarations for global variables.
	vm.applyGlobalVarTypes(compiler.GlobalVarTypes(), compiler.GlobalRecordTypes())
	// Apply VB6 As Type declarations for local variables (function-scoped).
	vm.applyLocalVarTypes(compiler)

	// Copy UDT declarations
	vm.RecordDecls = append(vm.RecordDecls[:0], compiler.recordDecls...)
	maps.Copy(vm.RecordDeclLookup, compiler.recordDeclLookup)

	// Copy parameter default values for Optional parameters.
	clear(vm.funcParamDefaults)
	for entryPoint, defaults := range compiler.funcParamDefaults {
		defs := make([]int, len(defaults))
		copy(defs, defaults)
		vm.funcParamDefaults[entryPoint] = defs
	}

	// Pre-allocate inline cache state from compiler IC node count.
	if compiler.jsICNodeCount > 0 {
		vm.icState = make([]InlineCacheSlot, compiler.jsICNodeCount)
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
	maps.Copy(compiler.globalZeroArgFuncs, vm.globalZeroArgFuncs)
	maps.Copy(compiler.globalZeroArgSubs, vm.globalZeroArgSubs)
	maps.Copy(compiler.declaredGlobals, vm.declaredGlobals)
	maps.Copy(compiler.constGlobals, vm.constGlobals)
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
	maps.Copy(compiler.globalZeroArgFuncs, vm.globalZeroArgFuncs)
	maps.Copy(compiler.globalZeroArgSubs, vm.globalZeroArgSubs)
	maps.Copy(compiler.declaredGlobals, vm.declaredGlobals)
	maps.Copy(compiler.constGlobals, vm.constGlobals)
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
func opcodeOperandSize(op OpCode, bytecode []byte, ip int) int {
	switch op {
	// 2-byte operands
	case OpConstant, OpWriteStatic, OpWriteN,
		OpGetClassMember, OpSetClassMember, OpEraseClassMember, OpLetClassMember, OpArgClassMemberRef,
		OpMemberSet, OpMemberSetSet,
		OpNewClass,
		OpLabel,
		OpRegisterClass,
		OpGetGlobal, OpSetGlobal, OpEraseGlobal, OpGetLocal, OpSetLocal, OpEraseLocal,
		OpArgGlobalRef, OpArgLocalRef,
		OpLetGlobal, OpLetLocal,
		OpCall,
		OpJSDeclareName, OpJSGetName, OpJSSetName, OpJSGetLocal, OpJSSetLocal, OpJSIncLocal, OpJSDecLocal,
		OpJSRootFrameEnter, OpJSRootFrameLeave,
		OpJSCreateClosure, OpJSCall, OpJSTailCall, OpJSCallComputedMember, OpJSTailCallComputedMember, OpJSNewArray, OpJSSuperCall,
		OpJSNew, OpJSMemberDelete, OpJSPostIncrement, OpJSPostDecrement, OpJSPreIncrement, OpJSPreDecrement,
		OpJSAddAssign, OpJSSubtractAssign, OpJSMultiplyAssign, OpJSDivideAssign, OpJSModuloAssign,
		OpJSExponentAssign, OpJSLogicalAndAssign, OpJSLogicalOrAssign, OpJSCoalesceAssign,
		OpJSMemberIndexGet, OpJSMemberIndexSet,
		OpJSPostMemberIncrement, OpJSPostMemberDecrement, OpJSPreMemberIncrement, OpJSPreMemberDecrement,
		OpJSLetDeclare, OpJSTDZRegisterLet, OpJSTDZRegisterConst, OpJSConstInitialize, OpJSRot,
		OpJSIncLocalInt, OpJSDecLocalInt,
		OpIncLocalInt, OpDecLocalInt, OpIncGlobalInt, OpDecGlobalInt,
		OpJSSuperMemberGet, OpJSSuperMemberSet, OpJSSuperCallComputedMember,
		OpJSExportAll:
		return 2
	case OpJSObjectRest:
		// countH(1), countL(1) + 2*count operand indices + dynamicCountH(1), dynamicCountL(1)
		count := int(binary.BigEndian.Uint16(bytecode[ip+1:]))
		return 2 + 2*count + 2
	// 4-byte operands
	case OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel, OpSet,
		OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter,
		OpJSJumpIfNullish, OpJSJumpIfNotNullish, OpJSJumpIfNotUndefined,
		OpJSCase, OpJSDefault, OpJSBreak, OpJSContinue, OpJSForInCleanup, OpJSForOfCleanup,
		OpJumpIfNotEq, OpJumpIfEq, OpJumpIfNotLt, OpJumpIfLte, OpJumpIfNotIs,
		OpJSJumpIfLooseNotEq, OpJSJumpIfLooseEq, OpJSJumpIfStrictNotEq, OpJSJumpIfStrictEq, OpJSJumpIfNotLess, OpJSJumpIfLessEqual:
		return 4
	case OpForNextFastInt, OpForNextFastGlobalInt:
		return 9
	case OpJSForFastIntEnter:
		return 8
	case OpJSForFastInt:
		return 8
	// 1-byte opcodes (none)
	case OpJSSetProto, OpJSSetThis, OpJSSuperIndexGet, OpJSSuperIndexSet:
		return 0
	case OpExtPrefix:
		extOp := ExtOpCode(bytecode[ip+1])
		switch extOp {
		case ExtOpRegisterClassEvent, ExtOpRaiseEvent, ExtOpWithEventsRegister, ExtOpRegisterClassInterface:
			return 5
		case ExtOpJumpLocalIfFalse, ExtOpJumpGlobalIfFalse, ExtOpJSJumpNameIfFalse:
			return 7
		case ExtOpAddLocalConst, ExtOpSubGlobalConst, ExtOpConcatLocalConst:
			return 5
		case ExtOpConstant2:
			return 5
		case ExtOpConstant3:
			return 7
		case ExtOpConstant4:
			return 9
		case ExtOpFilePrint, ExtOpFileWrite:
			return 3
		case ExtOpFileOpen, ExtOpFileClose, ExtOpFileLineInput, ExtOpFilePut, ExtOpFileGet, ExtOpFileFreeFile, ExtOpAxonASP, ExtOpJSReThrow, ExtOpCloneRecord, ExtOpShiftLeft, ExtOpShiftRight:
			return 1
		case ExtOpJSMathSin, ExtOpJSMathCos, ExtOpJSMathTan, ExtOpJSMathAbs, ExtOpJSMathFloor, ExtOpJSMathCeil, ExtOpJSMathRound, ExtOpJSMathSqrt, ExtOpJSMathMin, ExtOpJSMathMax:
			return 1
		default:
			return 3
		}
	// 4-byte operands
	case OpLine, OpArraySet, OpCallBuiltin, OpSetDirective, OpSetOption, OpJSCallMember, OpJSTailCallMember, OpJSDefineProperty, OpJSSuperCallMember, OpJSExport:
		return 4
	// 8-byte operands
	// 8-byte operands
	case OpCallMember:
		return 8
	// 4-byte operands: nameConstIdx(2) + icNodeID(2)
	case OpJSMemberGet, OpJSMemberSet:
		return 4
	// 9-byte operands: classNameIdx(2) + fieldNameIdx(2) + isPublic(2) + type(1) + classTypeIdx(2)
	case OpRegisterClassField:
		return 9
	// 6-byte operands: classNameIdx(2) + fieldNameIdx(2) + dimCount(2); dim values are on stack
	case OpInitClassArrayField:
		return 6
	// 6-byte operands for OpJSForIn: EnumVarIdx(2) + LoopStartTarget(4)
	case OpJSForIn:
		return 6
	// 6-byte operands for OpJSForOf: EnumVarIdx(2) + exitTarget(4)
	case OpJSForOf:
		return 6
	// 8-byte operands: classNameIdx(2) + methodNameIdx(2) + userSubIdx(2) + isPublic(2)
	case OpRegisterClassMethod:
		return 8
	// 8-byte operands: nameConstIdx(2) + limitConstIdx(2) + exitTarget(4)
	case OpJSJumpIfLessFast:
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
			OpJSDeclareName, OpJSGetName, OpJSSetName, OpJSCreateClosure,
			OpJSMemberDelete, OpJSPostIncrement, OpJSPostDecrement, OpJSPreIncrement, OpJSPreDecrement,
			OpJSAddAssign, OpJSSubtractAssign, OpJSMultiplyAssign, OpJSDivideAssign, OpJSModuloAssign,
			OpJSExponentAssign, OpJSLogicalAndAssign, OpJSLogicalOrAssign, OpJSCoalesceAssign,
			OpJSMemberIndexGet, OpJSMemberIndexSet,
			OpJSPostMemberIncrement, OpJSPostMemberDecrement, OpJSPreMemberIncrement, OpJSPreMemberDecrement,
			OpJSLetDeclare, OpJSTDZRegisterLet, OpJSTDZRegisterConst, OpJSConstInitialize, OpJSIncLocalInt, OpJSDecLocalInt,
			OpJSSuperMemberGet, OpJSSuperMemberSet, OpJSExportAll:
			idx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx))
			ip += 2
		case OpJSMemberGet, OpJSMemberSet:
			idx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx))
			ip += 4
		case OpExtPrefix:
			ext := ExtOpCode(bytecode[ip])
			ip++
			switch ext {
			case ExtOpInitRecord, ExtOpGetRecordMember, ExtOpSetRecordMember:
				ip += 2
			case ExtOpAxonASP, ExtOpJSMathSin, ExtOpJSMathCos, ExtOpJSMathTan, ExtOpJSMathAbs, ExtOpJSMathFloor, ExtOpJSMathCeil, ExtOpJSMathRound, ExtOpJSMathSqrt, ExtOpJSMathMin, ExtOpJSMathMax,
				ExtOpFileOpen, ExtOpFileClose, ExtOpFileLineInput, ExtOpFilePut, ExtOpFileGet, ExtOpFileFreeFile,
				ExtOpJSReThrow, ExtOpCloneRecord, ExtOpShiftLeft, ExtOpShiftRight:
				// No operands to remap or skip
			case ExtOpFilePrint, ExtOpFileWrite:
				ip += 2
			case ExtOpJumpLocalIfFalse, ExtOpJumpGlobalIfFalse:
				ip += 2
				target := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
				binary.BigEndian.PutUint32(bytecode[ip:], uint32(target))
				ip += 4
			case ExtOpJSJumpNameIfFalse:
				idx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx))
				ip += 2
				target := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
				binary.BigEndian.PutUint32(bytecode[ip:], uint32(target))
				ip += 4
			case ExtOpAddLocalConst, ExtOpConcatLocalConst:
				ip += 2
				idx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx))
				ip += 2
			case ExtOpSubGlobalConst:
				gidx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(gidx))
				ip += 2
				idx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx))
				ip += 2
			case ExtOpConstant2:
				idx1 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx1))
				ip += 2
				idx2 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx2))
				ip += 2
			case ExtOpConstant3:
				idx1 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx1))
				ip += 2
				idx2 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx2))
				ip += 2
				idx3 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx3))
				ip += 2
			case ExtOpConstant4:
				idx1 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx1))
				ip += 2
				idx2 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx2))
				ip += 2
				idx3 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx3))
				ip += 2
				idx4 := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(idx4))
				ip += 2
			default:
				panic("unsupported extended opcode in remapExecuteGlobalBytecode")
			}
		case OpJSRot, OpJSSuperCall, OpJSCall, OpJSTailCall, OpJSCallComputedMember, OpJSTailCallComputedMember, OpJSNewArray, OpJSSuperCallComputedMember:
			ip += 2
		case OpJSDefineProperty:
			nameIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(nameIdx))
			ip += 4
		case OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel, OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter, OpJSCase, OpJSDefault, OpJSBreak, OpJSContinue, OpJSForInCleanup, OpJSForOfCleanup,
			OpJSJumpIfNullish, OpJSJumpIfNotNullish, OpJSJumpIfNotUndefined,
			OpJumpIfNotEq, OpJumpIfEq, OpJumpIfNotLt, OpJumpIfLte, OpJumpIfNotIs,
			OpJSJumpIfLooseNotEq, OpJSJumpIfLooseEq, OpJSJumpIfStrictNotEq, OpJSJumpIfStrictEq, OpJSJumpIfNotLess, OpJSJumpIfLessEqual:
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
		case OpCallMember, OpJSCallMember, OpJSTailCallMember, OpJSSuperCallMember:
			memberIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(memberIdx))
			if op == OpCallMember {
				ip += 8
			} else {
				ip += 4
			}
		case OpArraySet:
			ip += 4
		case OpJSForIn:
			// EnumVarIdx(2) + LoopStartTarget(4)
			enumVarIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(enumVarIdx))
			ip += 2
			loopStartTarget := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
			binary.BigEndian.PutUint32(bytecode[ip:], uint32(loopStartTarget))
			ip += 4
		case OpJSForOf:
			// EnumVarIdx(2) + exitTarget(4)
			foVarIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(foVarIdx))
			ip += 2
			foExitTarget := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
			binary.BigEndian.PutUint32(bytecode[ip:], uint32(foExitTarget))
			ip += 4
		case OpJSForFastIntEnter:
			// counterSlot(2) + limitSlot(2) + exitTarget(4)
			ip += 4
			target := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
			binary.BigEndian.PutUint32(bytecode[ip:], uint32(target))
			ip += 4
		case OpJSForFastInt:
			// counterSlot(2) + limitSlot(2) + jumpOffset(4)
			ip += 4
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
			ip += 4 // skip fieldIdx(2) + isPublic(2)
			ip += 1 // skip type(1)
			classTypeIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(classTypeIdx))
			ip += 2
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
		case OpGetGlobal, OpSetGlobal, OpGetLocal, OpSetLocal, OpLabel, OpArgGlobalRef, OpArgLocalRef, OpLetGlobal, OpLetLocal, OpCall, OpIncLocalInt, OpDecLocalInt, OpIncGlobalInt, OpDecGlobalInt, OpWriteN, OpJSGetLocal, OpJSSetLocal, OpJSIncLocal, OpJSDecLocal, OpJSRootFrameEnter, OpJSRootFrameLeave:
			ip += 2
		case OpForNextFastInt:
			// varLocalIdx(2) + endLocalIdx(2) + stepSign(1): local indices, no remapping needed.
			// bodyTarget(4): absolute bytecode offset that must be rebased.
			ip += 5
			target := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
			binary.BigEndian.PutUint32(bytecode[ip:], uint32(target))
			ip += 4
		case OpForNextFastGlobalInt:
			// varGlobalIdx(2) + endGlobalIdx(2) + stepSign(1): global slots that must be rebased.
			globalIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(globalIdx))
			ip += 2
			endIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(endIdx))
			ip += 3
			target := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
			binary.BigEndian.PutUint32(bytecode[ip:], uint32(target))
			ip += 4
		case OpJSJumpIfLessFast:
			// nameConstIdx(2): remap into merged constant table.
			nameIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(nameIdx))
			ip += 2
			// limitConstIdx(2): remap into merged constant table.
			limitIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(limitIdx))
			ip += 2
			// exitTarget(4): absolute bytecode offset that must be rebased.
			exitTgt := int(binary.BigEndian.Uint32(bytecode[ip:])) + bytecodeBase
			binary.BigEndian.PutUint32(bytecode[ip:], uint32(exitTgt))
			ip += 4
		case OpJSForIterEnter, OpJSForIterExit:
			// Variable-length: [numVars(2), nameIdx1(2), nameIdx2(2), ...]
			if ip+2 > len(bytecode) {
				break
			}
			numVars := int(binary.BigEndian.Uint16(bytecode[ip:]))
			ip += 2
			for j := 0; j < numVars && ip+2 <= len(bytecode); j++ {
				varIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(varIdx))
				ip += 2
			}
		case OpJSImport:
			// Variable-length: moduleIdx(2) + specCount(2) + (importedIdx(2), localIdx(2))*N
			if ip+4 > len(bytecode) {
				break
			}
			moduleIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(moduleIdx))
			ip += 2
			specCount := int(binary.BigEndian.Uint16(bytecode[ip:]))
			ip += 2
			for j := 0; j < specCount && ip+4 <= len(bytecode); j++ {
				importedIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(importedIdx))
				ip += 2
				localIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
				binary.BigEndian.PutUint16(bytecode[ip:], uint16(localIdx))
				ip += 2
			}
		case OpJSExport:
			localIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(localIdx))
			ip += 2
			exportIdx := int(binary.BigEndian.Uint16(bytecode[ip:])) + constBase
			binary.BigEndian.PutUint16(bytecode[ip:], uint16(exportIdx))
			ip += 2
		case OpSetOption:
			ip += 4
		case OpLine:
			ip += 4
		case OpCallBuiltin:
			ip += 4
		default:
			// Keep scan aligned for opcodes that do not need remapping but still carry operands.
			ip += opcodeOperandSize(op, bytecode, ip-1)
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
			continue
		}
		if appendedConstants[i].Type == VTJSFunctionTemplate || appendedConstants[i].Type == VTJSArrowFunctionTemplate {
			appendedConstants[i].Num += int64(bytecodeBase)
			appendedConstants[i].Flt += float64(bytecodeBase)
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
	// Copy declared type arrays for VB6 As Type support.
	child.localTypes = vm.localTypes
	child.globalTypes = make([]ValueType, len(vm.globalTypes))
	copy(child.globalTypes, vm.globalTypes)
	child.funcLocalTypes = vm.funcLocalTypes // map reference is shared (read-only after init)
	child.activeClassObjectID = vm.activeClassObjectID
	child.terminateCursor = -1
	child.terminatePrepared = false
	child.suppressTerminate = true
	child.onResumeNext = vm.onResumeNext
	child.executeGlobalResumeGuard = vm.onResumeNext
	child.stmtSP = -1
	child.skipToNextStmt = false
	// Global dynamic execution (for CommonJS modules and ExecuteGlobal paths)
	// must not inherit caller lexical block scopes, otherwise TDZ bindings from
	// the caller can leak and break valid module-local identifiers.
	child.jsBlockScopes = make([]map[string]Value, 0)
	child.jsBlockScopeConst = make([]map[string]struct{}, 0)
	child.jsBlockScopeTDZ = make([]map[string]struct{}, 0)
	child.jsBlockScopeDepth = 0
	// Deep-copy icState for isolated inline cache state in the child VM.
	if len(vm.icState) > 0 {
		child.icState = make([]InlineCacheSlot, len(vm.icState))
		copy(child.icState, vm.icState)
	}

	return &child
}

// cloneForExecuteLocal builds an execution VM that shares the parent
// stack and call frame context to enable local variable access.
func (vm *VM) cloneForExecuteLocal(startIP int) *VM {
	vm.cloneForExecuteLocalCount++
	child := *vm
	child.parentVM = vm
	// Copy the active stack so the child can suspend or resume without being
	// clobbered by the caller continuing execution on the parent VM.
	child.stack = make([]Value, len(vm.stack))
	copy(child.stack, vm.stack)
	child.jsCallStack = make([]jsCallFrame, len(vm.jsCallStack))
	copy(child.jsCallStack, vm.jsCallStack)
	// Copy the declared type arrays.
	child.localTypes = vm.localTypes
	child.globalTypes = make([]ValueType, len(vm.globalTypes))
	copy(child.globalTypes, vm.globalTypes)
	child.funcLocalTypes = vm.funcLocalTypes
	child.ip = startIP
	child.callStack = make([]CallFrame, len(vm.callStack))
	copy(child.callStack, vm.callStack)
	child.activeClassObjectID = vm.activeClassObjectID
	child.terminateCursor = -1
	child.terminatePrepared = false
	child.suppressTerminate = true
	child.onResumeNext = vm.onResumeNext
	child.skipToNextStmt = false
	child.jsEnvItems = make(map[int64]*jsEnvFrame, len(vm.jsEnvItems))
	for id, env := range vm.jsEnvItems {
		if env == nil {
			child.jsEnvItems[id] = nil
			continue
		}
		bindings := make(map[string]Value, len(env.bindings))
		maps.Copy(bindings, env.bindings)
		child.jsEnvItems[id] = &jsEnvFrame{parentID: env.parentID, bindings: bindings}
	}
	child.jsArgumentsItems = make(map[int64]*jsArgumentsBinding, len(vm.jsArgumentsItems))
	for id, binding := range vm.jsArgumentsItems {
		if binding == nil {
			child.jsArgumentsItems[id] = nil
			continue
		}
		indexToParam := make(map[string]string, len(binding.indexToParam))
		maps.Copy(indexToParam, binding.indexToParam)
		paramToIndex := make(map[string]string, len(binding.paramToIndex))
		maps.Copy(paramToIndex, binding.paramToIndex)
		child.jsArgumentsItems[id] = &jsArgumentsBinding{
			envID:        binding.envID,
			indexToParam: indexToParam,
			paramToIndex: paramToIndex,
		}
	}
	child.jsBlockScopes = make([]map[string]Value, len(vm.jsBlockScopes))
	for i, scope := range vm.jsBlockScopes {
		if scope == nil {
			continue
		}
		cloned := make(map[string]Value, len(scope))
		maps.Copy(cloned, scope)
		child.jsBlockScopes[i] = cloned
	}
	child.jsBlockScopeConst = make([]map[string]struct{}, len(vm.jsBlockScopeConst))
	for i, scopeConst := range vm.jsBlockScopeConst {
		if scopeConst == nil {
			continue
		}
		cloned := make(map[string]struct{}, len(scopeConst))
		for name := range scopeConst {
			cloned[name] = struct{}{}
		}
		child.jsBlockScopeConst[i] = cloned
	}
	child.jsBlockScopeTDZ = make([]map[string]struct{}, len(vm.jsBlockScopeTDZ))
	for i, scopeTDZ := range vm.jsBlockScopeTDZ {
		if scopeTDZ == nil {
			continue
		}
		cloned := make(map[string]struct{}, len(scopeTDZ))
		for name := range scopeTDZ {
			cloned[name] = struct{}{}
		}
		child.jsBlockScopeTDZ[i] = cloned
	}
	child.jsBlockScopeDepth = len(child.jsBlockScopes)
	// Deep-copy icState so the child VM has its own isolated inline cache slots.
	if len(vm.icState) > 0 {
		child.icState = make([]InlineCacheSlot, len(vm.icState))
		copy(child.icState, vm.icState)
	}
	child.classInstanceOrder = append(make([]int64, 0, len(vm.classInstanceOrder)), vm.classInstanceOrder...)
	child.jsTryStack = make([]int, 0, 8)
	child.jsErrStack = make([]Value, 0, 4)
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
	vm.globalZeroArgSubs = child.globalZeroArgSubs
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
	vm.requestCollectionValueItems = child.requestCollectionValueItems
	vm.aspErrorItems = child.aspErrorItems
	vm.g3mdItems = child.g3mdItems
	vm.g3dateItems = child.g3dateItems
	vm.g3searchItems = child.g3searchItems
	vm.g3stringBuilderItems = child.g3stringBuilderItems
	vm.g3cryptoItems = child.g3cryptoItems
	vm.g3jsonItems = child.g3jsonItems
	vm.g3httpItems = child.g3httpItems
	vm.g3mailItems = child.g3mailItems
	vm.g3imageItems = child.g3imageItems
	vm.g3imageCanvasItems = child.g3imageCanvasItems
	vm.g3imageFontItems = child.g3imageFontItems
	vm.g3imagePenItems = child.g3imagePenItems
	vm.g3filesItems = child.g3filesItems
	vm.g3templateItems = child.g3templateItems
	vm.g3zipItems = child.g3zipItems
	vm.g3zlibItems = child.g3zlibItems
	vm.g3tarItems = child.g3tarItems
	vm.g3zstdItems = child.g3zstdItems
	vm.g3fcItems = child.g3fcItems
	vm.g3axonliveItems = child.g3axonliveItems
	vm.g3axonliveProxyItems = child.g3axonliveProxyItems
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
	vm.wscriptEnvironmentItems = child.wscriptEnvironmentItems
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
	vm.pdfDocItems = child.pdfDocItems
	vm.pdfPageItems = child.pdfPageItems
	vm.pdfCanvasItems = child.pdfCanvasItems
	vm.pdfFontItems = child.pdfFontItems
	vm.pdfPagesItems = child.pdfPagesItems
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
	vm.collectionItems = child.collectionItems
	vm.collectionEnumeratorItems = child.collectionEnumeratorItems
	vm.nativeObjectProxies = child.nativeObjectProxies
	vm.jsObjectItems = child.jsObjectItems
	vm.jsObjectKeyOrder = child.jsObjectKeyOrder
	vm.jsObjectStateItems = child.jsObjectStateItems
	vm.jsPropertyItems = child.jsPropertyItems
	vm.jsFunctionItems = child.jsFunctionItems
	vm.jsForInItems = child.jsForInItems
	vm.jsForOfItems = child.jsForOfItems
	vm.jsEnvItems = child.jsEnvItems
	vm.jsArgumentsItems = child.jsArgumentsItems
	vm.jsSetItems = child.jsSetItems
	vm.jsMapItems = child.jsMapItems
	vm.jsArrayBuffers = child.jsArrayBuffers
	vm.jsSharedArrayBuffers = child.jsSharedArrayBuffers
	vm.jsImmediateQueue = child.jsImmediateQueue
	vm.jsNextTickQueue = child.jsNextTickQueue
	vm.jsMicrotaskQueue = child.jsMicrotaskQueue
	vm.jsSymbolGlobalRegistry = child.jsSymbolGlobalRegistry
	vm.jsNextSymbolID = child.jsNextSymbolID
	vm.jsProxyItems = child.jsProxyItems
	vm.jsStreamHookItems = child.jsStreamHookItems
	vm.jsModuleInstances = child.jsModuleInstances
	vm.jsModuleLoading = child.jsModuleLoading
	vm.jsRootEnvID = child.jsRootEnvID
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

// EngineMode returns the current language mode of the VM.
func (vm *VM) EngineMode() EngineMode {
	if vm == nil {
		return EngineModeDefault
	}
	return vm.engineMode
}

// SetEngineMode updates the language mode for the current execution.
func (vm *VM) SetEngineMode(mode EngineMode) {
	if vm == nil {
		return
	}
	vm.engineMode = mode
}

// SetExecutionMode sets the execution context for this VM instance.
// Interactive modes (CLI, TUI, Eval) bypass caching to prevent stalls.
func (vm *VM) SetExecutionMode(mode ExecutionMode) {
	if vm != nil {
		vm.executionMode = mode
	}
}

// GetExecutionMode returns the current execution mode.
func (vm *VM) GetExecutionMode() ExecutionMode {
	if vm != nil {
		return vm.executionMode
	}
	return ExecutionModeServer
}

// IsInteractiveMode reports whether this VM is in an interactive execution context.
func (vm *VM) IsInteractiveMode() bool {
	if vm == nil {
		return false
	}
	return vm.executionMode == ExecutionModeCLI ||
		vm.executionMode == ExecutionModeTUI ||
		vm.executionMode == ExecutionModeEval
}

// Run executes the loaded bytecode.
func (vm *VM) Run() (err error) {
	vm.runDepth++
	isRootRun := vm.runDepth == 1
	defer func() {
		vm.runDepth--
	}()

	if isRootRun && vm.host != nil && vm.host.Server() != nil {
		vm.host.Server().BeginExecution()
		defer vm.host.Server().EndExecution()
	}
	if isRootRun {
		defer vm.closeAllFiles()
	}
	defer func() {
		if isRootRun && !vm.suppressTerminate {
			vm.jsCleanupCollections()
		}
	}()
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
			if are, ok := r.(*jsAsyncRejectionError); ok {
				if vm.suppressTerminate || !isRootRun {
					panic(are)
				}
				vm.jsActiveEnvID = 0
				vm.jsRootEnvID = 0
				vm.jsCallStack = nil
				vm.jsTryStack = nil
				vm.jsErrStack = nil

				description := "Unhandled JScript exception: " + vm.valueToString(are.reason)
				code := vbscript.InternalError
				hresult := int64(vbscript.HRESULTFromVBScriptCode(code))
				if numVal, ok := vm.jsMemberGet(are.reason, "number"); ok {
					switch numVal.Type {
					case VTInteger:
						hresult = numVal.Num
					case VTDouble:
						hresult = int64(numVal.Flt)
					}
				}

				file, line, column := vm.mapRuntimeLocation(vm.lastLine, vm.lastColumn)
				vme := &VMError{
					Code:           code,
					Line:           line,
					Column:         column,
					File:           file,
					Msg:            description,
					ASPCode:        int(code),
					ASPDescription: description,
					Category:       "VBScript runtime",
					Description:    description,
					Number:         int(hresult),
					Source:         "VBScript runtime error",
				}
				vm.errSetFromVMError(vme)
				vm.lastError = vme

				if vm.onResumeNext {
					err = nil
				} else {
					err = vme
				}
				return
			}
			if bufferErr, ok := r.(*asp.ResponseBufferLimitError); ok {
				err = vm.newMappedAxonASPError(ErrResponseBufferLimitExceeded, bufferErr, bufferErr.Error())
				return
			}
			if ye, ok := r.(*jsYieldError); ok {
				if vm.suppressTerminate {
					panic(ye)
				}
				err = ye
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
			} else if re, ok := r.(error); ok {
				nextOp := "<out-of-range>"
				if vm.ip >= 0 && vm.ip < len(vm.bytecode) {
					nextOp = OpCode(vm.bytecode[vm.ip]).String()
				}
				err = fmt.Errorf("%w\nVM context: ip=%d, nextOp=%s, lastLine=%d, bytecodeLen=%d, globalsLen=%d, constantsLen=%d\n%s", re, vm.ip, nextOp, vm.lastLine, len(vm.bytecode), len(vm.Globals), len(vm.constants), debug.Stack())
			} else {
				if vm.onResumeNext {
					err = nil
				} else {
					err = fmt.Errorf("internal runtime panic at line %d: %v\n%s", vm.lastLine, r, debug.Stack())
				}
			}
		}
	}()
	operationCount := 0
	jsBackJumpCount := 0
	if isRootRun {
		vm.jsStringWorkBytes = 0
		// Reset the concat scratch buffer so leftover capacity from a previous run does not
		// pin a large backing array unnecessarily across request boundaries.
		vm.stringWorkBuffer = vm.stringWorkBuffer[:0]
	}

aspExecLoop:
	for vm.ip < len(vm.bytecode) {
		operationCount++
		if operationCount&63 == 0 {
			vm.jsPumpNodeAsyncTasks(32)
		}
		if operationCount&1023 == 0 && vm.host != nil && vm.host.Server() != nil && vm.host.Server().HasTimedOut() {
			return vm.newMappedAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("Script execution exceeded the configured timeout of %d second(s)", vm.host.Server().GetScriptTimeout()))
		}
		op := OpCode(vm.bytecode[vm.ip])
		vm.ip++

		// When Resume Next absorbed an error mid-statement, skip bytecode until the next
		// statement boundary (OpLine) or take any OpJumpIfFalse unconditionally to correctly
		// skip compound-statement blocks (e.g. If/Then body). OpRet and OpHalt fall through
		// so function returns and program termination work normally during skip mode.
		if vm.skipToNextStmt {
			// ... (rest of skip logic)
		}

		switch op {
		case OpHalt:
			vm.jsDrainNodeAsyncOnExit()
			break aspExecLoop

		case OpConstant:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.constants[idx])

		case OpPop:
			vm.pop()

		case OpNop:
			// No operation: consumed only the single opcode byte; advance to next instruction.

		case OpGetGlobal:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			val := vm.Globals[idx]
			if val.Type == VTEmpty && int(idx) < len(vm.globalNames) && strings.HasPrefix(vm.globalNames[idx], "__static_") {
				val.Interface = "static"
			}
			if val.Type == VTObject {
				if className, ok := vm.globalClassTypes[idx]; ok {
					val.Interface = className
				}
			}
			vm.push(val)

		case OpSetGlobal:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			newVal := vm.pop()
			// Type enforcement for VB6 As Type declarations.
			if int(idx) < len(vm.globalTypes) && vm.globalTypes[idx] != VTEmpty {
				coerced, err := vm.coerceToDeclaredType(newVal, vm.globalTypes[idx])
				if err != nil {
					vm.raise(vbscript.TypeMismatch, err.Error())
					if vm.skipToNextStmt {
						continue
					}
				}
				newVal = coerced
				// Phase 5: Preserve Interface name from global metadata.
				if newVal.Type == VTObject {
					if className, ok := vm.globalClassTypes[idx]; ok {
						newVal.Interface = className
					}
				}
			} else {
				newVal.Interface = ""
			}

			// Phase 4: WithEvents binding
			if vm.globalWithEvents[idx] {
				prevVal := vm.Globals[idx]
				vm.unbindWithEvents(nil, vm.globalNames[idx], prevVal)
				vm.bindWithEvents(nil, vm.globalNames[idx], newVal)
			}

			// Decrement reference count only for object slots to avoid hot-path call overhead.
			if vm.Globals[idx].Type == VTObject {
				vm.decrementObjectRefCount(vm.Globals[idx])
			}
			// Assign new value.
			vm.Globals[idx] = newVal
			// Increment reference count only for object values to avoid hot-path call overhead.
			if newVal.Type == VTObject {
				vm.incrementObjectRefCount(newVal)
			}

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

			// Phase 4: WithEvents binding
			if vm.activeClassObjectID != 0 {
				instance := vm.runtimeClassItems[vm.activeClassObjectID]
				if instance != nil && instance.WithEventsNames[strings.ToLower(memberName)] {
					prevVal := vm.getActiveClassMemberValue(memberName)
					vm.unbindWithEvents(instance, memberName, prevVal)
					vm.bindWithEvents(instance, memberName, newVal)
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
			// Type enforcement for VB6 As Type declarations.
			if int(idx) < len(vm.globalTypes) && vm.globalTypes[idx] != VTEmpty {
				coerced, err := vm.coerceToDeclaredType(newVal, vm.globalTypes[idx])
				if err != nil {
					vm.raise(vbscript.TypeMismatch, err.Error())
					if vm.skipToNextStmt {
						continue
					}
				}
				newVal = coerced
			} else {
				newVal.Interface = ""
			}
			// Decrement reference count only for object slots to avoid hot-path call overhead.
			if vm.Globals[idx].Type == VTObject {
				vm.decrementObjectRefCount(vm.Globals[idx])
			}
			// Assign new value.
			vm.Globals[idx] = newVal
			// Increment reference count only for object values to avoid hot-path call overhead.
			if newVal.Type == VTObject {
				vm.incrementObjectRefCount(newVal)
			}

		// OpLetLocal: plain VBScript "name = value" for a local variable.
		// Local slots are mutable Variants and must be overwritten directly.
		case OpLetLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			newVal := vm.pop()
			// Type enforcement for VB6 As Type declarations.
			if declaredType := vm.localTypes[slot]; declaredType != VTEmpty {
				coerced, err := vm.coerceToDeclaredType(newVal, declaredType)
				if err != nil {
					vm.raise(vbscript.TypeMismatch, err.Error())
					if vm.skipToNextStmt {
						continue
					}
				}
				newVal = coerced
				// Phase 5: Preserve Interface name from local metadata.
				if newVal.Type == VTObject {
					if vm.funcLocalClassTypes != nil {
						if frame, ok := vm.funcLocalClassTypes[vm.getCurrentFuncEntryPoint()]; ok {
							if className, ok2 := frame[int(offset)]; ok2 {
								newVal.Interface = className
							}
						}
					}
				}
			} else {
				newVal.Interface = ""
			}
			// Decrement reference count only for object slots to avoid hot-path call overhead.
			if vm.stack[slot].Type == VTObject {
				vm.decrementObjectRefCount(vm.stack[slot])
			}
			// Assign new value.
			vm.stack[slot] = newVal
			// Increment reference count only for object values to avoid hot-path call overhead.
			if newVal.Type == VTObject {
				vm.incrementObjectRefCount(newVal)
			}

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
				if collectionValue, exists := vm.requestCollectionValueItems[v.Num]; exists {
					vm.push(NewString(collectionValue.Joined()))
					continue
				}
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
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			val := vm.stack[vm.fp+int(idx)]
			if val.Type == VTObject {
				if vm.funcLocalClassTypes != nil {
					if frame, ok := vm.funcLocalClassTypes[vm.getCurrentFuncEntryPoint()]; ok {
						if className, ok2 := frame[int(idx)]; ok2 {
							val.Interface = className
						}
					}
				}
			}
			vm.push(val)

		case OpSetLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			newVal := vm.pop()
			// Type enforcement for VB6 As Type declarations.
			if declaredType := vm.localTypes[slot]; declaredType != VTEmpty {
				coerced, err := vm.coerceToDeclaredType(newVal, declaredType)
				if err != nil {
					vm.raise(vbscript.TypeMismatch, err.Error())
					if vm.skipToNextStmt {
						continue
					}
				}
				newVal = coerced
				// Phase 5: Preserve Interface name from local metadata.
				if newVal.Type == VTObject {
					if vm.funcLocalClassTypes != nil {
						if frame, ok := vm.funcLocalClassTypes[vm.getCurrentFuncEntryPoint()]; ok {
							if className, ok2 := frame[int(offset)]; ok2 {
								newVal.Interface = className
							}
						}
					}
				}
			} else {
				newVal.Interface = ""
			}
			// Decrement reference count only for object slots to avoid hot-path call overhead.
			if vm.stack[slot].Type == VTObject {
				vm.decrementObjectRefCount(vm.stack[slot])
			}
			// Assign new value.
			vm.stack[slot] = newVal
			// Increment reference count only for object values to avoid hot-path call overhead.
			if newVal.Type == VTObject {
				vm.incrementObjectRefCount(newVal)
			}

		case OpSet:
			// OpSet destOp(16) destIdx(16)
			destOp := OpCode(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			destIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			newVal := vm.pop()

			if destOp == OpSetGlobal {
				if int(destIdx) < len(vm.globalTypes) && vm.globalTypes[destIdx] != VTEmpty {
					coerced, err := vm.coerceToDeclaredType(newVal, vm.globalTypes[destIdx])
					if err != nil {
						vm.raise(vbscript.TypeMismatch, err.Error())
						continue
					}
					newVal = coerced
					// Phase 5: Preserve Interface name from global metadata.
					if newVal.Type == VTObject {
						if className, ok := vm.globalClassTypes[destIdx]; ok {
							newVal.Interface = className
						}
					}
				}
				if vm.Globals[destIdx].Type == VTObject {
					vm.decrementObjectRefCount(vm.Globals[destIdx])
				}
				vm.Globals[destIdx] = newVal
				if newVal.Type == VTObject {
					vm.incrementObjectRefCount(newVal)
				}
			} else if destOp == OpSetLocal {
				slot := vm.fp + int(destIdx)
				if declaredType := vm.localTypes[slot]; declaredType != VTEmpty {
					coerced, err := vm.coerceToDeclaredType(newVal, declaredType)
					if err != nil {
						vm.raise(vbscript.TypeMismatch, err.Error())
						continue
					}
					newVal = coerced
					// Phase 5: Preserve Interface name from local metadata.
					if newVal.Type == VTObject {
						if vm.funcLocalClassTypes != nil {
							if frame, ok := vm.funcLocalClassTypes[vm.getCurrentFuncEntryPoint()]; ok {
								if className, ok2 := frame[int(destIdx)]; ok2 {
									newVal.Interface = className
								}
							}
						}
					}
				}
				if vm.stack[slot].Type == VTObject {
					vm.decrementObjectRefCount(vm.stack[slot])
				}
				vm.stack[slot] = newVal
				if newVal.Type == VTObject {
					vm.incrementObjectRefCount(newVal)
				}
			} else {
				vm.raise(vbscript.InternalError, "Invalid OpSet destination")
			}

		case OpIncLocalInt:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			current := &vm.stack[slot]
			switch current.Type {
			case VTInteger:
				current.Num++
			case VTDouble:
				current.Flt++
			default:
				vm.stack[slot] = vm.addValues(*current, NewInteger(1))
			}

		case OpDecLocalInt:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			current := &vm.stack[slot]
			switch current.Type {
			case VTInteger:
				current.Num--
			case VTDouble:
				current.Flt--
			default:
				vm.stack[slot] = vm.subtractValues(*current, NewInteger(1))
			}

		case OpJSIncLocalInt:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			current := vm.jsGetName(nameStr)
			vm.jsSetName(nameStr, vm.jsIncrementNumberValue(current))

		case OpJSDecLocalInt:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			current := vm.jsGetName(nameStr)
			vm.jsSetName(nameStr, vm.jsDecrementNumberValue(current))

		case OpIncGlobalInt:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			current := vm.Globals[idx]
			switch current.Type {
			case VTInteger:
				current.Num++
				vm.Globals[idx] = current
			case VTDouble:
				current.Flt++
				vm.Globals[idx] = current
			default:
				next := vm.addValues(current, NewInteger(1))
				if current.Type == VTObject {
					vm.decrementObjectRefCount(current)
				}
				vm.Globals[idx] = next
				if next.Type == VTObject {
					vm.incrementObjectRefCount(next)
				}
			}

		case OpDecGlobalInt:
			idx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			current := vm.Globals[idx]
			switch current.Type {
			case VTInteger:
				current.Num--
				vm.Globals[idx] = current
			case VTDouble:
				current.Flt--
				vm.Globals[idx] = current
			default:
				next := vm.subtractValues(current, NewInteger(1))
				if current.Type == VTObject {
					vm.decrementObjectRefCount(current)
				}
				vm.Globals[idx] = next
				if next.Type == VTObject {
					vm.incrementObjectRefCount(next)
				}
			}

		// OpForNextFastInt — fused increment/decrement + bounds-check + conditional back-jump.
		// The loop variable and limit are both local frame slots; step direction is encoded in
		// a single sign byte.  The fast integer branch executes with zero heap allocations.
		case OpForNextFastInt:
			varIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			endIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			stepSign := int8(vm.bytecode[vm.ip])
			vm.ip++
			bodyTarget := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4

			varSlot := vm.fp + varIdx
			endSlot := vm.fp + endIdx
			curr := vm.stack[varSlot]
			limit := vm.stack[endSlot]

			// Fastest path: both slots hold plain integers — pure arithmetic, no dispatch.
			if curr.Type == VTInteger && limit.Type == VTInteger {
				if stepSign > 0 {
					curr.Num++
					vm.stack[varSlot] = curr
					if curr.Num <= limit.Num {
						vm.ip = bodyTarget
					}
				} else {
					curr.Num--
					vm.stack[varSlot] = curr
					if curr.Num >= limit.Num {
						vm.ip = bodyTarget
					}
				}
				continue
			}

			// Float / mixed fallback: handles VTDouble counter or limit without allocation.
			if stepSign > 0 {
				switch curr.Type {
				case VTDouble:
					curr.Flt++
				default:
					curr = vm.addValues(curr, NewInteger(1))
				}
				vm.stack[varSlot] = curr
				if vm.asFloat(curr) <= vm.asFloat(limit) {
					vm.ip = bodyTarget
				}
			} else {
				switch curr.Type {
				case VTDouble:
					curr.Flt--
				default:
					curr = vm.subtractValues(curr, NewInteger(1))
				}
				vm.stack[varSlot] = curr
				if vm.asFloat(curr) >= vm.asFloat(limit) {
					vm.ip = bodyTarget
				}
			}

		case OpJSForFastInt:
			bc := vm.bytecode
			counterOffset := int(uint16(bc[vm.ip])<<8 | uint16(bc[vm.ip+1]))
			limitOffset := int(uint16(bc[vm.ip+2])<<8 | uint16(bc[vm.ip+3]))
			jumpOffset := int(uint32(bc[vm.ip+4])<<24 | uint32(bc[vm.ip+5])<<16 | uint32(bc[vm.ip+6])<<8 | uint32(bc[vm.ip+7]))
			vm.ip += 8

			counter := &vm.stack[vm.fp+counterOffset]
			limit := &vm.stack[vm.fp+limitOffset]

			// Hot-path is intentionally type-blind: OpJSForFastIntEnter establishes VTInteger contract.
			counter.Num++
			if counter.Num < limit.Num {
				jsBackJumpCount++
				if jsBackJumpCount > jsBackJumpLimit {
					return vm.newMappedAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit))
				}
				vm.ip -= jumpOffset
				continue
			}

		case OpJSForFastIntEnter:
			counterOffset := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			limitOffset := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:]))
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip+4:]))
			vm.ip += 8

			counter := vm.stack[vm.fp+counterOffset]
			limit := vm.stack[vm.fp+limitOffset]
			if counter.Type != VTInteger || limit.Type != VTInteger {
				vm.jsThrowTypeError("Fast integer loop requires VTInteger counter and limit")
			}
			if counter.Num >= limit.Num {
				vm.ip = target
			}

		case OpForNextFastGlobalInt:
			varIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			endIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			stepSign := int8(vm.bytecode[vm.ip])
			vm.ip++
			bodyTarget := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4

			curr := vm.Globals[varIdx]
			limit := vm.Globals[endIdx]

			if curr.Type == VTInteger && limit.Type == VTInteger {
				if stepSign > 0 {
					curr.Num++
					vm.Globals[varIdx] = curr
					if curr.Num <= limit.Num {
						vm.ip = bodyTarget
					}
				} else {
					curr.Num--
					vm.Globals[varIdx] = curr
					if curr.Num >= limit.Num {
						vm.ip = bodyTarget
					}
				}
				continue
			}

			if curr.Type == VTDouble && limit.Type == VTDouble {
				if stepSign > 0 {
					curr.Flt++
					vm.Globals[varIdx] = curr
					if curr.Flt <= limit.Flt {
						vm.ip = bodyTarget
					}
				} else {
					curr.Flt--
					vm.Globals[varIdx] = curr
					if curr.Flt >= limit.Flt {
						vm.ip = bodyTarget
					}
				}
				continue
			}

			if stepSign > 0 {
				switch curr.Type {
				case VTDouble:
					curr.Flt++
				default:
					curr = vm.addValues(curr, NewInteger(1))
				}
				if curr.Type == VTObject {
					vm.decrementObjectRefCount(vm.Globals[varIdx])
				}
				vm.Globals[varIdx] = curr
				if vm.asFloat(curr) <= vm.asFloat(limit) {
					vm.ip = bodyTarget
				}
			} else {
				switch curr.Type {
				case VTDouble:
					curr.Flt--
				default:
					curr = vm.subtractValues(curr, NewInteger(1))
				}
				if curr.Type == VTObject {
					vm.decrementObjectRefCount(vm.Globals[varIdx])
				}
				vm.Globals[varIdx] = curr
				if vm.asFloat(curr) >= vm.asFloat(limit) {
					vm.ip = bodyTarget
				}
			}

		case OpAdd:
			// Reduce the top two stack values in-place to avoid extra Value copies.
			vm.stack[vm.sp-1] = vm.addValues(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpSub:
			vm.stack[vm.sp-1] = vm.subtractValues(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpMul:
			vm.stack[vm.sp-1] = vm.multiplyValues(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpDiv:
			vm.stack[vm.sp-1] = vm.divideValues(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpMod:
			vm.stack[vm.sp-1] = vm.modValues(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpPow:
			vm.stack[vm.sp-1] = vm.powValues(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpIAdd:
			vm.stack[vm.sp-1] = NewInteger(vm.coerceInt64(vm.stack[vm.sp-1]) + vm.coerceInt64(vm.stack[vm.sp]))
			vm.sp--

		case OpISub:
			vm.stack[vm.sp-1] = NewInteger(vm.coerceInt64(vm.stack[vm.sp-1]) - vm.coerceInt64(vm.stack[vm.sp]))
			vm.sp--

		case OpIMul:
			vm.stack[vm.sp-1] = NewInteger(vm.coerceInt64(vm.stack[vm.sp-1]) * vm.coerceInt64(vm.stack[vm.sp]))
			vm.sp--

		case OpIDiv:
			// VBScript integer division: a \ b
			vm.stack[vm.sp-1] = vm.intDivideValues(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpIRightShift:
			right := vm.coerceInt64(resolveCallable(vm, vm.stack[vm.sp]))
			left := vm.coerceInt64(resolveCallable(vm, vm.stack[vm.sp-1]))
			if right <= 0 {
				vm.stack[vm.sp-1] = NewInteger(left)
			} else if right >= 63 {
				if left < 0 {
					vm.stack[vm.sp-1] = NewInteger(-1)
				} else {
					vm.stack[vm.sp-1] = NewInteger(0)
				}
			} else {
				vm.stack[vm.sp-1] = NewInteger(left >> uint(right))
			}
			vm.sp--

		case OpMathSin:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewDouble(math.Sin(input))

		case OpMathCos:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewDouble(math.Cos(input))

		case OpMathTan:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewDouble(math.Tan(input))

		case OpMathAtn:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewDouble(math.Atan(input))

		case OpMathSqr:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewDouble(math.Sqrt(input))

		case OpMathAbs:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			if arg.Type == VTDouble {
				vm.stack[vm.sp] = NewDouble(math.Abs(arg.Flt))
			} else {
				val := arg.Num
				if val < 0 {
					val = -val
				}
				vm.stack[vm.sp] = NewInteger(val)
			}

		case OpMathExp:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewDouble(math.Exp(input))

		case OpMathLog:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewDouble(math.Log(input))

		case OpMathRound:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewDouble(math.RoundToEven(input))

		case OpMathInt:
			arg := resolveCallable(vm, vm.stack[vm.sp])
			input := float64(arg.Num)
			if arg.Type == VTDouble {
				input = arg.Flt
			}
			vm.stack[vm.sp] = NewInteger(int64(math.Floor(input)))

		case OpNeq:
			a := vm.stack[vm.sp-1]
			b := vm.stack[vm.sp]
			if isNull(a) || isNull(b) {
				vm.stack[vm.sp-1] = NewNull()
			} else if vm.optionCompare == 1 {
				vm.stack[vm.sp-1] = NewBool(!vm.textEqual(a.String(), b.String()))
			} else {
				vm.stack[vm.sp-1] = NewBool(vm.compareValues(a, b) != 0)
			}
			vm.sp--

		case OpIsRef:
			a := vm.stack[vm.sp-1]
			b := vm.stack[vm.sp]
			if !IsObjectReferenceValue(a) || !IsObjectReferenceValue(b) {
				vm.raise(vbscript.TypeMismatch, vbscript.TypeMismatch.String())
				vm.stack[vm.sp-1] = NewEmpty()
				vm.sp--
				continue
			}
			isNothingA := a.Type == VTNothing || (a.Type == VTEmpty && a.Interface == "static") || ((a.Type == VTObject || a.Type == VTNativeObject) && a.Num == 0)
			isNothingB := b.Type == VTNothing || (b.Type == VTEmpty && b.Interface == "static") || ((b.Type == VTObject || b.Type == VTNativeObject) && b.Num == 0)
			if isNothingA && isNothingB {
				vm.stack[vm.sp-1] = NewBool(true)
			} else if isNothingA || isNothingB {
				vm.stack[vm.sp-1] = NewBool(false)
			} else {
				vm.stack[vm.sp-1] = NewBool(a.Type == b.Type && a.Num == b.Num)
			}
			vm.sp--

		case OpIsNotRef:
			a := vm.stack[vm.sp-1]
			b := vm.stack[vm.sp]
			if !IsObjectReferenceValue(a) || !IsObjectReferenceValue(b) {
				vm.raise(vbscript.TypeMismatch, vbscript.TypeMismatch.String())
				vm.stack[vm.sp-1] = NewEmpty()
				vm.sp--
				continue
			}
			isNothingA := a.Type == VTNothing || (a.Type == VTEmpty && a.Interface == "static") || ((a.Type == VTObject || a.Type == VTNativeObject) && a.Num == 0)
			isNothingB := b.Type == VTNothing || (b.Type == VTEmpty && b.Interface == "static") || ((b.Type == VTObject || b.Type == VTNativeObject) && b.Num == 0)
			if isNothingA && isNothingB {
				vm.stack[vm.sp-1] = NewBool(false)
			} else if isNothingA || isNothingB {
				vm.stack[vm.sp-1] = NewBool(true)
			} else {
				vm.stack[vm.sp-1] = NewBool(a.Type != b.Type || a.Num != b.Num)
			}
			vm.sp--

		case OpGt:
			a := vm.stack[vm.sp-1]
			b := vm.stack[vm.sp]
			if isNull(a) || isNull(b) {
				vm.stack[vm.sp-1] = NewNull()
			} else {
				vm.stack[vm.sp-1] = NewBool(vm.compareValues(a, b) > 0)
			}
			vm.sp--

		case OpLte:
			a := vm.stack[vm.sp-1]
			b := vm.stack[vm.sp]
			if isNull(a) || isNull(b) {
				vm.stack[vm.sp-1] = NewNull()
			} else {
				vm.stack[vm.sp-1] = NewBool(vm.compareValues(a, b) <= 0)
			}
			vm.sp--

		case OpGte:
			a := vm.stack[vm.sp-1]
			b := vm.stack[vm.sp]
			if isNull(a) || isNull(b) {
				vm.stack[vm.sp-1] = NewNull()
			} else {
				vm.stack[vm.sp-1] = NewBool(vm.compareValues(a, b) >= 0)
			}
			vm.sp--

		case OpAnd:
			vm.stack[vm.sp-1] = vm.logicalAnd(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpOr:
			vm.stack[vm.sp-1] = vm.logicalOr(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpXor:
			vm.stack[vm.sp-1] = vm.logicalXor(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

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
			vm.stack[vm.sp-1] = vm.logicalEqv(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpImp:
			vm.stack[vm.sp-1] = vm.logicalImp(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

		case OpJumpIfTrue:
			target := binary.BigEndian.Uint32(vm.bytecode[vm.ip:])
			vm.ip += 4
			if vm.isTruthy(vm.pop()) {
				vm.ip = int(target)
			}

		case OpJumpIfNotEq:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			var eq bool
			if isNull(a) || isNull(b) {
				eq = false
			} else if vm.optionCompare == 1 {
				eq = vm.textEqual(a.String(), b.String())
			} else {
				eq = vm.compareValues(a, b) == 0
			}
			if !eq {
				vm.ip = target
			}

		case OpJumpIfEq:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			var neq bool
			if isNull(a) || isNull(b) {
				neq = false
			} else if vm.optionCompare == 1 {
				neq = !vm.textEqual(a.String(), b.String())
			} else {
				neq = vm.compareValues(a, b) != 0
			}
			if !neq {
				vm.ip = target
			}

		case OpJumpIfNotLt:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			if isNull(a) || isNull(b) || !(vm.asFloat(a) < vm.asFloat(b)) {
				vm.ip = target
			}

		case OpJumpIfLte:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			// Fuses OpGt + OpJumpIfFalse -> Jump if !(a > b) -> Jump if a <= b
			if isNull(a) || isNull(b) || (vm.asFloat(a) <= vm.asFloat(b)) {
				vm.ip = target
			}

		case OpJumpIfNotIs:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			if !IsObjectReferenceValue(a) || !IsObjectReferenceValue(b) {
				vm.raise(vbscript.TypeMismatch, vbscript.TypeMismatch.String())
				vm.ip = target
				continue
			}
			isNothingA := a.Type == VTNothing || (a.Type == VTEmpty && a.Interface == "static") || ((a.Type == VTObject || a.Type == VTNativeObject) && a.Num == 0)
			isNothingB := b.Type == VTNothing || (b.Type == VTEmpty && b.Interface == "static") || ((b.Type == VTObject || b.Type == VTNativeObject) && b.Num == 0)
			var eq bool
			if isNothingA && isNothingB {
				eq = true
			} else if isNothingA || isNothingB {
				eq = false
			} else {
				eq = (a.Type == b.Type && a.Num == b.Num)
			}
			if !eq {
				vm.ip = target
			}

		case OpJSJumpIfLooseNotEq:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			if !vm.isTruthy(vm.jsLooseEqual(a, b)) {
				vm.ip = target
			}

		case OpJSJumpIfLooseEq:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			if !vm.isTruthy(vm.jsLooseNotEqual(a, b)) {
				vm.ip = target
			}

		case OpJSJumpIfStrictNotEq:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			if !vm.jsStrictEquals(a, b) {
				vm.ip = target
			}

		case OpJSJumpIfStrictEq:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			if vm.jsStrictEquals(a, b) {
				vm.ip = target
			}

		case OpJSJumpIfNotLess:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			if !vm.isTruthy(vm.jsLess(a, b)) {
				vm.ip = target
			}

		case OpJSJumpIfLessEqual:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			b, a := vm.pop(), vm.pop()
			if !vm.isTruthy(vm.jsGreater(a, b)) {
				vm.ip = target
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
			fieldType := ValueType(vm.bytecode[vm.ip])
			vm.ip += 1
			classTypeIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2

			classType := ""
			if classTypeIdx != 0xFFFF {
				classType = vm.constants[classTypeIdx].Str
			}
			vm.registerRuntimeClassField(vm.constants[classNameIdx].Str, vm.constants[fieldNameIdx].Str, isPublic, fieldType, classType)

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
			vm.stack[vm.sp-1] = vm.concatValues(vm.stack[vm.sp-1], vm.stack[vm.sp])
			vm.sp--

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

		case OpWriteN:
			// OpWriteN pops N values (pushed in left-to-right evaluation order) and
			// writes each as a string directly to the Response without creating
			// intermediate concatenated Value objects.  All N parts are accumulated
			// into vm.stringWorkBuffer and flushed with a single host.WriteString call
			// so the Response mutex is acquired only once.
			n := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			if n > 0 && vm.output != nil {
				// Ensure indexBuffer has capacity for N entries.
				if cap(vm.indexBuffer) < n {
					vm.indexBuffer = make([]Value, n)
				}
				parts := vm.indexBuffer[:n]
				// Pop in reverse order (TOS = last operand) and store in left-to-right order.
				for i := n - 1; i >= 0; i-- {
					parts[i] = vm.pop()
				}
				// Accumulate all parts into the reusable scratch buffer, then write once.
				vm.stringWorkBuffer = vm.stringWorkBuffer[:0]
				for i := range n {
					vm.stringWorkBuffer = append(vm.stringWorkBuffer, vm.valueToString(parts[i])...)
				}
				if len(vm.stringWorkBuffer) > 0 {
					_, _ = vm.output.Write(vm.stringWorkBuffer)
				}
			} else {
				// Drain stack even if output is nil.
				for range n {
					vm.pop()
				}
			}

		case OpCallMember:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			cacheOffset := vm.ip
			cacheID := binary.BigEndian.Uint32(vm.bytecode[cacheOffset:])
			vm.ip += 4

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
			if target.Type == VTArgRef {
				target = vm.unwrapArgRefValue(target)
			}
			if cacheID != 0 {
				if cacheEntry, exists := vm.callMemberIC[cacheID]; exists && cacheEntry.expectedType == target.Type && cacheEntry.expectedNum == target.Num {
					switch cacheEntry.kind {
					case callMemberICClassMethod:
						if cacheEntry.target.UserSubParamCount() != argCount {
							vm.raise(vbscript.WrongNumberOfParameters, "Wrong number of parameters or invalid property assignment")
						}
						var memberByRefs []byRefWriteback
						if hasArgRefMember {
							memberByRefs = vm.collectByRefsAndUnwrap(memberArgs, cacheEntry.target.UserSubByRefMask())
						}
						if vm.beginUserSubCall(cacheEntry.target, memberArgs, false, target.Num, memberByRefs) {
							continue
						}
					case callMemberICClassPropertyGet:
						var propByRefs []byRefWriteback
						if hasArgRefMember {
							propByRefs = vm.collectByRefsAndUnwrap(memberArgs, cacheEntry.target.UserSubByRefMask())
						}
						if vm.beginUserSubCall(cacheEntry.target, memberArgs, false, target.Num, propByRefs) {
							continue
						}
					case callMemberICNativeMember:
						var nativeByRefs []byRefWriteback
						if hasArgRefMember {
							nativeByRefs = vm.collectByRefsAndUnwrap(memberArgs, vm.nativeByRefMask(target.Num, cacheEntry.nativeMember))
						}
						result := vm.dispatchNativeCall(target.Num, cacheEntry.nativeMember, memberArgs)
						if len(nativeByRefs) > 0 {
							vm.applyByRefWritebacksFromArgs(nativeByRefs, memberArgs)
						}
						vm.push(result)
						continue
					}
				}
			}
			switch target.Type {
			case VTObject:
				if target.Num == 0 {
					vm.raise(vbscript.CouldNotFindTargetObject, "Object required")
					vm.push(Value{Type: VTEmpty})
					continue
				}
				requirePublic := target.Num != vm.activeClassObjectID
				interfaceName := target.Interface
				if interfaceName != "" {
					if instance, exists := vm.runtimeClassItems[target.Num]; exists && instance != nil {
						if strings.EqualFold(interfaceName, instance.ClassName) {
							interfaceName = ""
						}
					}
				}
				methodName := vm.constants[memberIdx].Str
				if interfaceName != "" {
					methodName = interfaceName + "_" + methodName
				}
				methodTarget, ok := vm.resolveRuntimeClassMethod(target, methodName, requirePublic)
				if ok {
					if cacheID == 0 {
						cacheID = vm.allocateCallMemberIC(callMemberICEntry{
							kind:         callMemberICClassMethod,
							expectedType: target.Type,
							expectedNum:  target.Num,
							target:       methodTarget,
						})
						binary.BigEndian.PutUint32(vm.bytecode[cacheOffset:], cacheID)
					}
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
				propertyName := vm.constants[memberIdx].Str
				if interfaceName != "" {
					propertyName = interfaceName + "_" + propertyName
				}
				propertyTarget, ok := vm.resolveRuntimeClassPropertyGet(target, propertyName, argCount, requirePublic)
				if ok {
					if cacheID == 0 {
						cacheID = vm.allocateCallMemberIC(callMemberICEntry{
							kind:         callMemberICClassPropertyGet,
							expectedType: target.Type,
							expectedNum:  target.Num,
							target:       propertyTarget,
						})
						binary.BigEndian.PutUint32(vm.bytecode[cacheOffset:], cacheID)
					}
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
				if cacheID == 0 {
					cacheID = vm.allocateCallMemberIC(callMemberICEntry{
						kind:         callMemberICNativeMember,
						expectedType: target.Type,
						expectedNum:  target.Num,
						nativeMember: memberName,
					})
					binary.BigEndian.PutUint32(vm.bytecode[cacheOffset:], cacheID)
				}
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
				interfaceName := target.Interface
				if interfaceName != "" {
					if instance, exists := vm.runtimeClassItems[target.Num]; exists && instance != nil {
						if strings.EqualFold(interfaceName, instance.ClassName) {
							interfaceName = ""
						}
					}
				}
				memberName := member.String()
				if interfaceName != "" {
					memberName = interfaceName + "_" + memberName
				}
				propertyTarget, ok := vm.resolveRuntimeClassPropertyGet(target, memberName, 0, requirePublic)
				if ok {
					if vm.beginUserSubCall(propertyTarget, nil, false, target.Num) {
						continue
					}
				}
				fieldValue, ok := vm.resolveRuntimeClassField(target, memberName, requirePublic)
				if ok {
					vm.push(fieldValue)
					continue
				}
				methodTarget, ok := vm.resolveRuntimeClassMethod(target, memberName, requirePublic)
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

			if target.Type == VTRecord && target.Rec != nil {
				memberName := member.String()
				def := vm.RecordDecls[target.Rec.DefIdx]
				found := false
				for i, m := range def.Members {
					if strings.EqualFold(m.Name, memberName) {
						vm.push(target.Rec.Members[i])
						found = true
						break
					}
				}
				if found {
					continue
				}
				vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Member not found in UDT: "+memberName)
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

		case OpExtPrefix:
			ext := ExtOpCode(vm.bytecode[vm.ip])
			vm.ip++
			switch ext {
			case ExtOpJumpLocalIfFalse:
				offset := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip+2:]))
				vm.ip += 6
				if !vm.isTruthy(vm.stack[vm.fp+offset]) {
					vm.ip = target
				}

			case ExtOpJumpGlobalIfFalse:
				idx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip+2:]))
				vm.ip += 6
				if !vm.isTruthy(vm.Globals[idx]) {
					vm.ip = target
				}

			case ExtOpJSJumpNameIfFalse:
				nameConstIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip+2:]))
				vm.ip += 6
				name := vm.constants[nameConstIdx].String()
				val := vm.jsGetName(name)
				if !vm.jsTruthy(val) {
					if target < vm.ip {
						jsBackJumpCount++
						if jsBackJumpCount > jsBackJumpLimit {
							return vm.newMappedAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit))
						}
					}
					vm.ip = target
				}

			case ExtOpAddLocalConst:
				offset := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				constIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:]))
				vm.ip += 4
				vm.stack[vm.fp+offset] = vm.addValues(vm.stack[vm.fp+offset], vm.constants[constIdx])

			case ExtOpSubGlobalConst:
				idx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				constIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:]))
				vm.ip += 4
				vm.Globals[idx] = vm.subtractValues(vm.Globals[idx], vm.constants[constIdx])

			case ExtOpConcatLocalConst:
				offset := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				constIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:]))
				vm.ip += 4
				vm.stack[vm.fp+offset] = vm.concatValues(vm.stack[vm.fp+offset], vm.constants[constIdx])

			case ExtOpConstant2:
				idx1 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				idx2 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:]))
				vm.ip += 4
				vm.stack[vm.sp+1] = vm.constants[idx1]
				vm.stack[vm.sp+2] = vm.constants[idx2]
				vm.sp += 2

			case ExtOpConstant3:
				idx1 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				idx2 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:]))
				idx3 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+4:]))
				vm.ip += 6
				vm.stack[vm.sp+1] = vm.constants[idx1]
				vm.stack[vm.sp+2] = vm.constants[idx2]
				vm.stack[vm.sp+3] = vm.constants[idx3]
				vm.sp += 3

			case ExtOpConstant4:
				idx1 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				idx2 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:]))
				idx3 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+4:]))
				idx4 := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip+6:]))
				vm.ip += 8
				vm.stack[vm.sp+1] = vm.constants[idx1]
				vm.stack[vm.sp+2] = vm.constants[idx2]
				vm.stack[vm.sp+3] = vm.constants[idx3]
				vm.stack[vm.sp+4] = vm.constants[idx4]
				vm.sp += 4

			case ExtOpInitRecord:
				defIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				vm.ip += 2
				if int(defIdx) >= len(vm.RecordDecls) {
					vm.raise(vbscript.InternalError, "Invalid UDT definition index")
					vm.push(Value{Type: VTEmpty})
					continue
				}
				decl := vm.RecordDecls[defIdx]
				vm.push(vm.initRecordValueRecursive(decl.Name))

			case ExtOpGetRecordMember:
				memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				vm.ip += 2
				recVal := vm.pop()
				if recVal.Type != VTRecord || recVal.Rec == nil {
					vm.raise(vbscript.CouldNotFindTargetObject, "UDT required for GetRecordMember")
					vm.push(Value{Type: VTEmpty})
					continue
				}
				if int(memberIdx) >= len(recVal.Rec.Members) {
					vm.raise(vbscript.InternalError, "Invalid UDT member index")
					vm.push(Value{Type: VTEmpty})
					continue
				}
				vm.push(recVal.Rec.Members[memberIdx])

			case ExtOpSetRecordMember:
				memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				vm.ip += 2
				rhs := vm.pop()
				recVal := vm.pop()
				if recVal.Type != VTRecord || recVal.Rec == nil {
					vm.raise(vbscript.CouldNotFindTargetObject, "UDT required for SetRecordMember")
					continue
				}
				if int(memberIdx) >= len(recVal.Rec.Members) {
					vm.raise(vbscript.InternalError, "Invalid UDT member index")
					continue
				}
				recVal.Rec.Members[memberIdx] = rhs

			case ExtOpFileOpen:
				vm.ensureCLIMode()
				numVal := vm.pop()
				modeVal := vm.pop()
				pathVal := vm.pop()
				num := int(numVal.Num)
				mode := int(modeVal.Num)
				path := filepath.FromSlash(vm.valueToString(pathVal))

				if num < 1 || num > 511 {
					vm.raise(vbscript.BadFileNameOrNumber, "Invalid file number")
					continue
				}

				var f *os.File
				var err error
				switch mode {
				case 1: // Input
					f, err = os.Open(path)
				case 2: // Output
					f, err = os.Create(path)
				case 3: // Append
					f, err = os.OpenFile(path, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
				case 4, 5: // Binary/Random
					f, err = os.OpenFile(path, os.O_RDWR|os.O_CREATE, 0644)
				default:
					vm.raise(vbscript.InternalError, "Invalid file mode")
					continue
				}

				if err != nil {
					vm.raise(vbscript.FileNotFound, fmt.Sprintf("Error opening file '%s': %v", path, err))
					continue
				}

				vm.fileIOItems[num] = f
				if mode == 1 {
					vm.fileIOBufReaders[num] = bufio.NewReader(f)
				} else {
					vm.fileIOBufWriters[num] = bufio.NewWriter(f)
				}

			case ExtOpFileClose:
				numVal := vm.pop()
				num := int(numVal.Num)
				if num == 0 {
					vm.closeAllFiles()
				} else {
					if f, ok := vm.fileIOItems[num]; ok {
						if w, ok2 := vm.fileIOBufWriters[num]; ok2 {
							w.Flush()
							delete(vm.fileIOBufWriters, num)
						}
						delete(vm.fileIOBufReaders, num)
						f.Close()
						delete(vm.fileIOItems, num)
					}
				}

			case ExtOpFilePrint, ExtOpFileWrite:
				vm.ensureCLIMode()
				argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
				vm.ip += 2
				args := make([]Value, argCount)
				for i := argCount - 1; i >= 0; i-- {
					args[i] = vm.pop()
				}
				numVal := vm.pop()
				num := int(numVal.Num)

				w, ok := vm.fileIOBufWriters[num]
				if !ok {
					vm.raise(vbscript.BadFileNameOrNumber, "File not open for output")
					continue
				}

				isWrite := ext == ExtOpFileWrite
				for i, arg := range args {
					s := vm.valueToString(arg)
					if isWrite {
						if arg.Type == VTString {
							w.WriteString("\"")
							w.WriteString(s)
							w.WriteString("\"")
						} else {
							w.WriteString(s)
						}
						if i < argCount-1 {
							w.WriteString(",")
						}
					} else {
						w.WriteString(s)
					}
				}
				w.WriteString("\n")

			case ExtOpFileLineInput:
				vm.ensureCLIMode()
				numVal := vm.pop()
				num := int(numVal.Num)
				r, ok := vm.fileIOBufReaders[num]
				if !ok {
					vm.raise(vbscript.BadFileNameOrNumber, "File not open for input")
					continue
				}
				line, err := r.ReadString('\n')
				if err != nil && err != io.EOF {
					vm.raise(vbscript.InternalError, err.Error())
					continue
				}
				line = strings.TrimSuffix(line, "\n")
				line = strings.TrimSuffix(line, "\r")
				vm.push(NewString(line))

			case ExtOpFilePut:
				vm.ensureCLIMode()
				val := vm.pop()
				posVal := vm.pop()
				numVal := vm.pop()
				num := int(numVal.Num)

				f, ok := vm.fileIOItems[num]
				if !ok {
					vm.raise(vbscript.BadFileNameOrNumber, "File not open")
					continue
				}

				if posVal.Type != VTEmpty {
					pos := int64(posVal.Num)
					f.Seek(pos-1, 0)
				}

				var data []byte
				switch val.Type {
				case VTString:
					data = []byte(val.Str)
				case VTInteger:
					data = make([]byte, 8)
					binary.LittleEndian.PutUint64(data, uint64(val.Num))
				case VTDouble:
					data = make([]byte, 8)
					binary.LittleEndian.PutUint64(data, math.Float64bits(val.Flt))
				default:
					vm.raise(vbscript.TypeMismatch, "Unsupported type for Put")
					continue
				}
				f.Write(data)

			case ExtOpFileGet:
				vm.ensureCLIMode()
				posVal := vm.pop()
				numVal := vm.pop()
				num := int(numVal.Num)

				f, ok := vm.fileIOItems[num]
				if !ok {
					vm.raise(vbscript.BadFileNameOrNumber, "File not open")
					continue
				}

				if posVal.Type != VTEmpty {
					pos := int64(posVal.Num)
					f.Seek(pos-1, 0)
				}

				buf := make([]byte, 1024)
				n, err := f.Read(buf)
				if err != nil && err != io.EOF {
					vm.raise(vbscript.InternalError, err.Error())
					continue
				}
				vm.push(NewString(string(buf[:n])))

			case ExtOpFileFreeFile:
				vm.ensureCLIMode()
				nextNum := 1
				for {
					if _, ok := vm.fileIOItems[nextNum]; !ok {
						break
					}
					nextNum++
					if nextNum > 511 {
						vm.raise(vbscript.InternalError, "No more file numbers available")
						break
					}
				}
				vm.push(NewInteger(int64(nextNum)))

			case ExtOpAxonASP:
				// vm.ip points to next instruction. No operands to skip.
				vm.push(NewString("G3pix AxonASP VBScript Engine"))

			case ExtOpJSReThrow:
				// vm.ip points to next instruction. No operands to skip.
				if len(vm.jsErrStack) > 0 {
					errVal := vm.jsErrStack[len(vm.jsErrStack)-1]
					vm.jsErrStack = vm.jsErrStack[:len(vm.jsErrStack)-1]
					vm.jsThrow(errVal)
				}

			case ExtOpRegisterClassEvent:
				classNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				eventNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:])
				vm.ip += 4
				className := vm.constants[classNameIdx].Str
				eventName := vm.constants[eventNameIdx].Str
				if def, ok := vm.runtimeClasses[strings.ToLower(className)]; ok {
					if def.Events == nil {
						def.Events = make(map[string]bool)
					}
					def.Events[strings.ToLower(eventName)] = true
					vm.runtimeClasses[strings.ToLower(className)] = def
				}

			case ExtOpRaiseEvent:
				eventNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				argCount := binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:])
				vm.ip += 4
				eventName := vm.constants[eventNameIdx].Str
				vm.handleRaiseEvent(eventName, int(argCount))

			case ExtOpWithEventsRegister:
				classNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				varNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:])
				vm.ip += 4
				className := ""
				if classNameIdx != 0xFFFF {
					className = vm.constants[classNameIdx].Str
				}
				varName := vm.constants[varNameIdx].Str
				vm.handleWithEventsRegister(className, varName)

			case ExtOpRegisterClassInterface:
				classNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				interfaceNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip+2:])
				vm.ip += 4
				className := vm.constants[classNameIdx].Str
				interfaceName := vm.constants[interfaceNameIdx].Str
				if def, ok := vm.runtimeClasses[strings.ToLower(className)]; ok {
					def.Interfaces = append(def.Interfaces, interfaceName)
					vm.runtimeClasses[strings.ToLower(className)] = def
				}

			case ExtOpJSMathSin:
				vm.push(NewDouble(math.Sin(vm.jsToNumber(vm.pop()).Flt)))

			case ExtOpJSMathCos:
				vm.push(NewDouble(math.Cos(vm.jsToNumber(vm.pop()).Flt)))

			case ExtOpJSMathTan:
				vm.push(NewDouble(math.Tan(vm.jsToNumber(vm.pop()).Flt)))

			case ExtOpJSMathAbs:
				v := vm.pop()
				if v.Type == VTInteger {
					if v.Num < 0 {
						vm.push(NewInteger(-v.Num))
					} else {
						vm.push(v)
					}
				} else {
					vm.push(NewDouble(math.Abs(vm.jsToNumber(v).Flt)))
				}

			case ExtOpJSMathFloor:
				vm.push(NewDouble(math.Floor(vm.jsToNumber(vm.pop()).Flt)))

			case ExtOpJSMathCeil:
				vm.push(NewDouble(math.Ceil(vm.jsToNumber(vm.pop()).Flt)))

			case ExtOpJSMathRound:
				vm.push(NewDouble(vm.jsMathRound(vm.jsToNumber(vm.pop()).Flt)))

			case ExtOpJSMathSqrt:
				vm.push(NewDouble(math.Sqrt(vm.jsToNumber(vm.pop()).Flt)))

			case ExtOpJSMathMin:
				b := vm.jsToNumber(vm.pop()).Flt
				a := vm.jsToNumber(vm.pop()).Flt
				vm.push(NewDouble(math.Min(a, b)))

			case ExtOpJSMathMax:
				b := vm.jsToNumber(vm.pop()).Flt
				a := vm.jsToNumber(vm.pop()).Flt
				vm.push(NewDouble(math.Max(a, b)))

			case ExtOpCloneRecord:
				val := vm.pop()
				if val.Type == VTRecord {
					vm.push(vm.cloneValue(val))
				} else {
					vm.push(val)
				}

			case ExtOpShiftLeft:
				right := vm.pop()
				left := vm.pop()
				if isNull(left) || isNull(right) {
					vm.push(NewNull())
					continue
				}
				shift := uint64(vm.coerceInt64(right))
				val := uint64(vm.coerceInt64(left))
				if shift >= 64 {
					vm.push(NewInteger(0))
				} else {
					vm.push(NewInteger(int64(val << shift)))
				}

			case ExtOpShiftRight:
				right := vm.pop()
				left := vm.pop()
				if isNull(left) || isNull(right) {
					vm.push(NewNull())
					continue
				}
				shift := uint64(vm.coerceInt64(right))
				val := uint64(vm.coerceInt64(left))
				if shift >= 64 {
					vm.push(NewInteger(0))
				} else {
					vm.push(NewInteger(int64(val >> shift)))
				}

			default:
				vm.raise(vbscript.InternalError, "Unsupported extended opcode")
			}
			continue

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
				interfaceName := target.Interface
				if interfaceName != "" {
					if instance, exists := vm.runtimeClassItems[target.Num]; exists && instance != nil {
						if strings.EqualFold(interfaceName, instance.ClassName) {
							interfaceName = ""
						}
					}
				}
				memberName := vm.constants[memberIdx].Str
				if interfaceName != "" {
					memberName = interfaceName + "_" + memberName
				}
				propertyTarget, ok := vm.resolveRuntimeClassPropertySet(target, memberName, 1, preferSet, strictSet, requirePublic)
				if ok {
					if vm.beginUserSubCall(propertyTarget, []Value{value}, true, target.Num) {
						continue
					}
				}
				if vm.assignRuntimeClassField(target, memberName, value, requirePublic) {
					continue
				}
				vm.raise(vbscript.WrongNumberOfParameters, "Wrong number of parameters or invalid property assignment")
			} else if target.Type == VTRecord && target.Rec != nil {
				memberName := vm.constants[memberIdx].Str
				def := vm.RecordDecls[target.Rec.DefIdx]
				found := false
				for i, m := range def.Members {
					if strings.EqualFold(m.Name, memberName) {
						target.Rec.Members[i] = value
						found = true
						break
					}
				}
				if !found {
					vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Member not found in UDT: "+memberName)
				}
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
			a := vm.stack[vm.sp-1]
			b := vm.stack[vm.sp]
			if isNull(a) || isNull(b) {
				vm.stack[vm.sp-1] = NewNull()
			} else if vm.optionCompare == 1 { // Text
				vm.stack[vm.sp-1] = NewBool(vm.textEqual(a.String(), b.String()))
			} else {
				vm.stack[vm.sp-1] = NewBool(vm.compareValues(a, b) == 0)
			}
			vm.sp--

		case OpLt:
			a := vm.stack[vm.sp-1]
			b := vm.stack[vm.sp]
			if isNull(a) || isNull(b) {
				vm.stack[vm.sp-1] = NewNull()
			} else {
				vm.stack[vm.sp-1] = NewBool(vm.compareValues(a, b) < 0)
			}
			vm.sp--

		case OpNeg:
			val := vm.pop()
			if isNull(val) {
				vm.push(NewNull())
			} else if val.Type == VTDouble {
				vm.push(NewDouble(-val.Flt))
			} else if val.Type == VTString {
				strVal := strings.TrimSpace(val.Str)
				parsedFloat, err := strconv.ParseFloat(strVal, 64)
				if err != nil {
					vm.raise(vbscript.TypeMismatch, "Type mismatch")
					vm.push(NewEmpty())
					continue
				}
				if strings.Contains(strVal, ".") || strings.Contains(strings.ToLower(strVal), "e") {
					vm.push(NewDouble(-parsedFloat))
				} else {
					parsedInt, intErr := strconv.ParseInt(strVal, 10, 64)
					if intErr == nil {
						vm.push(NewInteger(-parsedInt))
					} else {
						vm.push(NewDouble(-parsedFloat))
					}
				}
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
				// Indexed access: obj.member(idx...) = value
				// First try Property Set with (index, value) arguments.
				requirePublic := target.Num != vm.activeClassObjectID
				memberName := vm.constants[memberIdx].Str
				if memberName == "" {
					memberName = "__default__"
				}
				letArgCount := argCount + 1
				letArgs := vm.ensureCombineBuffer(letArgCount)[:0]
				letArgs = append(letArgs, indexes...)
				letArgs = append(letArgs, value)
				propertyTarget, ok := vm.resolveRuntimeClassPropertySet(target, memberName, letArgCount, false, false, requirePublic)
				if ok {
					if vm.beginUserSubCall(propertyTarget, letArgs, true, target.Num) {
						continue
					}
				}
				// If no Property Set exists, try to resolve as a field and
				// assign to an array element (e.g. obj.Arr(0) = val).
				if fieldValue, ok := vm.resolveRuntimeClassField(target, memberName, requirePublic); ok {
					if fieldValue.Type == VTArray {
						vm.assignArrayElement(fieldValue, indexes, value)
						// Write the modified array back to the field so the
						// change is visible through the class member.
						vm.assignRuntimeClassField(target, memberName, fieldValue, requirePublic)
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

			// __AXON_ENUM_VALUES must receive native objects (e.g.
			// RequestCollectionValue) uncoerced so For Each can enumerate
			// sub-keys. The compiler removes OpCoerceToValue before the
			// call, so we also skip resolveCallable here.
			targetPeek := vm.stack[vm.sp-argCount]
			skipResolve := targetPeek.Type == VTBuiltin && targetPeek.Num == builtinEnumValuesIdx

			callArgs := vm.ensureArgBuffer(argCount)
			hasArgRef := false
			for i := argCount - 1; i >= 0; i-- {
				v := vm.pop()
				if v.Type == VTArgRef {
					hasArgRef = true
					callArgs[i] = v
				} else if skipResolve {
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
			} else if target.Type == VTJSFunction || target.Type == VTJSObject || target.Type == VTJSProxy {
				// Bridge VBScript -> JScript function calls for cross-language invocation.
				// Functions defined in <script language="JScript" runat="server"> blocks
				// are stored as VTJSFunction/VTJSObject values in the global scope.
				result := vm.jsCall(target, Value{Type: VTJSUndefined}, args)
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

		case OpJSImport:
			savedIP := vm.ip - 1
			moduleIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			specCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			moduleSpecifier := vm.constants[moduleIdx].Str
			moduleEnv, ok := vm.jsImportModule(moduleSpecifier)
			if !ok {
				// If jsImportModule already threw a JScript exception, vm.ip was moved.
				// We must not advance it further.
				if vm.ip != savedIP+5 {
					goto aspExecLoop
				}
				for range specCount {
					vm.ip += 4
				}
				continue
			}
			moduleLoading := vm.jsIsModuleLoading(moduleSpecifier)
			for range specCount {
				importedIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				vm.ip += 2
				localIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				vm.ip += 2
				importedName := vm.constants[importedIdx].Str
				localName := vm.constants[localIdx].Str
				var value Value
				if importedName == "*" {
					value = vm.jsGetModuleNamespace(moduleEnv)
				} else {
					exportKey := vm.jsModuleExportKey(importedName)
					var exists bool
					value, exists = moduleEnv.bindings[exportKey]
					if !exists && !moduleLoading {
						vm.jsThrowReferenceError("The module '" + vm.constants[moduleIdx].Str + "' does not provide an export named '" + importedName + "'")
					} else if !exists {
						value = Value{Type: VTJSUndefined}
					}
				}
				vm.jsDeclareName(localName)
				vm.jsSetName(localName, value)
			}

		case OpJSExport:
			localIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			exportIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			localName := vm.constants[localIdx].Str
			exportName := vm.constants[exportIdx].Str
			vm.jsSetModuleExport(exportName, vm.jsGetName(localName))

		case OpJSExportAll:
			moduleIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.jsExportAllFromModule(vm.constants[moduleIdx].Str)

		case OpJSRootFrameEnter:
			localCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			if localCount > 0 {
				base := max(vm.fp, 0)
				end := base + localCount - 1
				if end >= len(vm.stack) {
					vm.jsThrowOutOfStackSpace()
					break
				}
				for i := base; i <= end; i++ {
					vm.stack[i] = Value{Type: VTJSUndefined}
				}
				vm.sp = end
			}

		case OpJSRootFrameLeave:
			localCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			if localCount > 0 {
				vm.sp = vm.fp - 1
			}

		case OpJSGetLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.stack[vm.fp+int(offset)])

		case OpJSSetLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			val := vm.pop()
			vm.stack[slot] = val
			vm.jsSyncAliasedLocalSlot(int(offset), val)

		case OpJSIncLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			var val Value
			v := &vm.stack[slot]
			switch v.Type {
			case VTInteger:
				if next, ok := jsAddIntegersNoOverflow(v.Num, 1); ok {
					v.Num = next
					val = *v
				} else {
					val = NewDouble(float64(v.Num) + 1)
					vm.stack[slot] = val
				}
			case VTDouble:
				v.Flt++
				val = *v
			default:
				val = vm.jsIncrementNumberValue(*v)
				vm.stack[slot] = val
			}
			vm.jsSyncAliasedLocalSlot(int(offset), val)

		case OpJSDecLocal:
			offset := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			slot := vm.fp + int(offset)
			var val Value
			v := &vm.stack[slot]
			switch v.Type {
			case VTInteger:
				if next, ok := jsSubtractIntegersNoOverflow(v.Num, 1); ok {
					v.Num = next
					val = *v
				} else {
					val = NewDouble(float64(v.Num) - 1)
					vm.stack[slot] = val
				}
			case VTDouble:
				v.Flt--
				val = *v
			default:
				val = vm.jsDecrementNumberValue(*v)
				vm.stack[slot] = val
			}
			vm.jsSyncAliasedLocalSlot(int(offset), val)

		case OpJSMemberGet:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			icNodeID := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			target := vm.pop()
			if int(icNodeID) < len(vm.icState) {
				cached := vm.icState[icNodeID]
				if value, ok := vm.jsICMemberGet(target, vm.constants[nameIdx].Str, cached.ShapeID, cached.Slot, cached.Flags); ok {
					vm.push(value)
					continue
				}
			}
			vm.jsICPopulate(icNodeID, target, vm.constants[nameIdx].Str)
			if value, deferred := vm.jsMemberGet(target, vm.constants[nameIdx].Str); !deferred {
				vm.push(value)
			}

		case OpJSMemberSet:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			icNodeID := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			value := vm.pop()
			target := vm.pop()
			if int(icNodeID) < len(vm.icState) {
				cached := vm.icState[icNodeID]
				if vm.jsICMemberSet(target, vm.constants[nameIdx].Str, value, cached.ShapeID, cached.Slot, cached.Flags) {
					continue
				}
			}
			vm.jsMemberSet(target, vm.constants[nameIdx].Str, value)
			vm.jsICPopulate(icNodeID, target, vm.constants[nameIdx].Str)

		case OpJSCall:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			callee := vm.pop()
			if vm.jsBeginDirectCall(callee, Value{Type: VTJSUndefined}, args) {
				continue
			}
			stackLen := len(vm.jsCallStack)
			result := vm.jsCall(callee, Value{Type: VTJSUndefined}, args)
			if len(vm.jsCallStack) == stackLen || result.Type != VTJSUndefined {
				vm.push(result)
			}

		case OpJSTailCall:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			callee := vm.pop()
			if vm.jsTailCallValue(callee, Value{Type: VTJSUndefined}, args) {
				continue
			}
			vm.jsReturn(vm.jsCall(callee, Value{Type: VTJSUndefined}, args))

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
			stackLen := len(vm.jsCallStack)
			if result, handled := vm.jsCallMember(target, member, args); handled {
				if len(vm.jsCallStack) == stackLen || result.Type != VTJSUndefined {
					vm.push(result)
				}
			} else {
				vm.push(Value{Type: VTJSUndefined})
			}

		case OpJSCallComputedMember:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			keyVal := vm.pop()
			target := vm.pop()
			key := vm.jsPropertyKeyFromValue(keyVal)
			stackLen := len(vm.jsCallStack)
			if result, handled := vm.jsCallMember(target, key, args); handled {
				if len(vm.jsCallStack) == stackLen || result.Type != VTJSUndefined {
					vm.push(result)
				}
			} else {
				vm.push(Value{Type: VTJSUndefined})
			}

		case OpJSTailCallMember:
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
			if callee, thisVal, ok, deferred := vm.jsPrepareMemberCallee(target, member); deferred {
				vm.jsReturn(Value{Type: VTJSUndefined})
			} else if ok {
				if vm.jsTailCallValue(callee, thisVal, args) {
					continue
				}
				vm.jsReturn(vm.jsCall(callee, thisVal, args))
			} else if result, handled := vm.jsCallMember(target, member, args); handled {
				vm.jsReturn(result)
			} else {
				vm.jsReturn(Value{Type: VTJSUndefined})
			}

		case OpJSTailCallComputedMember:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			keyVal := vm.pop()
			target := vm.pop()
			key := vm.jsPropertyKeyFromValue(keyVal)
			if callee, thisVal, ok, deferred := vm.jsPrepareMemberCallee(target, key); deferred {
				vm.jsReturn(Value{Type: VTJSUndefined})
			} else if ok {
				if vm.jsTailCallValue(callee, thisVal, args) {
					continue
				}
				vm.jsReturn(vm.jsCall(callee, thisVal, args))
			} else if result, handled := vm.jsCallMember(target, key, args); handled {
				vm.jsReturn(result)
			} else {
				vm.jsReturn(Value{Type: VTJSUndefined})
			}

		case OpJSCreateClosure:
			templateIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.jsCreateClosure(vm.constants[templateIdx]))

		case OpJSDefineProperty:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			kind := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			val := vm.pop()
			target := vm.pop()
			name := vm.constants[nameIdx].Str

			// Set homeObjID for functions to support super member access
			if val.Type == VTJSFunction {
				if fn := vm.jsFunctionItems[val.Num]; fn != nil {
					fn.homeObjID = target.Num
				}
			}

			vm.jsDefineProperty(target, name, int(kind), val)

		case OpJSSetProto:
			proto := vm.pop()
			target := vm.pop()
			vm.jsSetProto(target, proto)

		case OpJSClassInherit:
			subclass := vm.pop()
			superclass := vm.pop()
			vm.push(vm.jsClassInherit(subclass, superclass))

		case OpJSSuperCall:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			stackLen := len(vm.jsCallStack)
			result := vm.jsSuperCall(args)
			if len(vm.jsCallStack) == stackLen || result.Type != VTJSUndefined {
				vm.push(result)
			}

		case OpJSSuperMemberGet:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			vm.push(vm.jsSuperGet(vm.constants[nameIdx].Str))

		case OpJSSuperMemberSet:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			val := vm.pop()
			vm.jsSuperSet(vm.constants[nameIdx].Str, val)
			vm.push(val)

		case OpJSSuperCallMember:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			stackLen := len(vm.jsCallStack)
			result := vm.jsSuperCallMember(vm.constants[nameIdx].Str, args)
			if len(vm.jsCallStack) == stackLen || result.Type != VTJSUndefined {
				vm.push(result)
			}

		case OpJSSuperIndexGet:
			index := vm.pop()
			vm.push(vm.jsSuperIndexGet(index))

		case OpJSSuperIndexSet:
			index := vm.pop()
			val := vm.pop()
			vm.jsSuperIndexSet(index, val)
			vm.push(val)

		case OpJSSuperCallComputedMember:
			argCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			args := vm.ensureArgBuffer(argCount)
			for i := argCount - 1; i >= 0; i-- {
				args[i] = vm.pop()
			}
			index := vm.pop()
			stackLen := len(vm.jsCallStack)
			result := vm.jsSuperCallMember(vm.jsPropertyKeyFromValue(index), args)
			if len(vm.jsCallStack) == stackLen || result.Type != VTJSUndefined {
				vm.push(result)
			}

		case OpJSAdd:
			b, a := vm.pop(), vm.pop()
			if a.Type == VTInteger && b.Type == VTInteger {
				if sum, ok := jsAddIntegersNoOverflow(a.Num, b.Num); ok {
					vm.push(NewInteger(sum))
					break
				}
				vm.push(NewDouble(float64(a.Num) + float64(b.Num)))
				break
			}
			vm.push(vm.jsAdd(a, b))

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
			obj := make(map[string]Value, 8)
			if proto := vm.jsGetIntrinsicPrototype("Object"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsObjectSlots[objID] = make([]Value, 0, 8)
			vm.jsObjectSlotIndex[objID] = make(map[string]uint16, 8)
			vm.jsObjectShape[objID] = 0
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
			vm.push(NewBool(vm.jsInstanceOf(left, right)))

		case OpJSIn:
			right := vm.pop()
			left := vm.pop()
			switch right.Type {
			case VTJSObject:
				key := vm.valueToString(left)
				if obj, ok := vm.jsObjectItems[right.Num]; ok {
					_, exists := obj[key]
					vm.push(NewBool(exists))
				} else {
					vm.push(NewBool(false))
				}
			case VTJSProxy:
				vm.push(NewBool(vm.jsProxyHas(right, vm.valueToString(left))))
			default:
				vm.push(NewBool(false))
			}

		case OpJSDelete:
			keyVal := vm.pop()
			target := vm.pop()
			key := vm.valueToString(keyVal)
			success := false
			if target.Type == VTJSProxy {
				success = vm.jsProxyDelete(target, key)
			} else {
				success = vm.jsMemberDelete(target, key)
			}
			vm.push(NewBool(success))

		case OpJSReturn:
			vm.jsReturn(vm.pop())

		case OpJSLoadUndefined:
			vm.push(Value{Type: VTJSUndefined})

		case OpJSLoadNewTarget:
			vm.push(vm.jsNewTarget)

		case OpJSLoadThis:
			if vm.jsThisValue.Type == VTJSUninitialized {
				vm.jsThrowReferenceError("OpJSLoadThis: Must call super constructor in derived class before accessing 'this'")
				vm.push(Value{Type: VTJSUndefined})
			} else {
				vm.push(vm.jsThisValue)
			}

		case OpJSAwait:
			p := vm.pop()
			if p.Type == VTJSPromise {
				for vm.jsGetPromiseState(p) == jsPromisePending {
					vm.jsProcessMicrotasks()
					if vm.jsGetPromiseState(p) == jsPromisePending {
						time.Sleep(1 * time.Millisecond)
					}
				}
				res := vm.jsGetPromiseResult(p)
				if vm.jsGetPromiseState(p) == jsPromiseRejected {
					vm.jsThrow(res)
					goto aspExecLoop
				}
				vm.push(res)
			} else {
				vm.push(p)
			}

		case OpJSYield:
			val := vm.pop()
			vm.jsYield(val, false)

		case OpJSYieldDelegate:
			val := vm.pop()
			vm.jsYield(val, true)

		case OpJSSetThis:
			v := vm.pop()
			vm.jsThisValue = v
			if len(vm.jsCallStack) > 0 {
				vm.jsCallStack[len(vm.jsCallStack)-1].thisVal = v
				if vm.jsCallStack[len(vm.jsCallStack)-1].ctorObj.Type == VTJSUninitialized {
					vm.jsCallStack[len(vm.jsCallStack)-1].ctorObj = v
				}
			}

		case OpJSDup:
			v := vm.pop()
			vm.push(v)
			vm.push(v)

		case OpJSRequireObject:
			v := vm.pop()
			if v.Type == VTNull || v.Type == VTJSUndefined {
				vm.jsThrowTypeError("Cannot destructure null or undefined")
			}
			vm.push(v)

		case OpJSGetIterator:
			source := vm.pop()
			vm.push(vm.jsGetIterator(source))

		case OpJSIteratorNext:
			itObj := vm.stack[vm.sp]
			vm.push(vm.jsIteratorNextValue(itObj))

		case OpJSCollectRest:
			itObj := vm.pop()
			vm.push(vm.jsCollectRest(itObj))

		case OpJSObjectRest:
			staticCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			excluded := make(map[string]struct{}, staticCount+4)
			for range staticCount {
				keyIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				vm.ip += 2
				excluded[vm.constants[keyIdx].Str] = struct{}{}
			}
			dynamicCount := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			for range dynamicCount {
				keyVal := vm.pop()
				excluded[vm.jsPropertyKeyFromValue(keyVal)] = struct{}{}
			}
			source := vm.pop()
			vm.push(vm.jsObjectRest(source, excluded))

		case OpJSPop:
			vm.pop()

		case OpJSRot:
			n := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			if n > 1 && vm.sp >= n-1 {
				top := vm.stack[vm.sp]
				copy(vm.stack[vm.sp-n+2:vm.sp+1], vm.stack[vm.sp-n+1:vm.sp])
				vm.stack[vm.sp-n+1] = top
			}

		case OpJSJump:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			if target < vm.ip {
				jsBackJumpCount++
				if jsBackJumpCount > jsBackJumpLimit {
					return vm.newMappedAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit))
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
						return vm.newMappedAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit))
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
						return vm.newMappedAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit))
					}
				}
				vm.ip = target
			}

		// OpJSJumpIfLessFast — fused bounds check for `identifier < numericLiteral` loop tests.
		// Reads the loop variable from the JS environment, compares it to the stored constant limit,
		// and jumps to the exit target when the variable is NOT less (condition false).  Falls through
		// to the loop body when the variable IS less — zero stack mutation on the hot path.
		// The exit target is always a forward jump, so no back-jump counter increment is needed here.
		case OpJSJumpIfLessFast:
			nameIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			limitIdx := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			exitTarget := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4

			varVal := vm.jsGetName(vm.constants[nameIdx].Str)
			limit := vm.constants[limitIdx]

			// Numeric fast path: avoid full jsLess dispatch when both sides are plain numbers.
			var varF, limitF float64
			fastOK := true
			switch varVal.Type {
			case VTDouble:
				varF = varVal.Flt
			case VTInteger:
				varF = float64(varVal.Num)
			default:
				fastOK = false
			}
			if fastOK {
				switch limit.Type {
				case VTDouble:
					limitF = limit.Flt
				case VTInteger:
					limitF = float64(limit.Num)
				default:
					fastOK = false
				}
			}
			if fastOK {
				// varF >= limitF means the loop condition (var < limit) is false → exit.
				if varF >= limitF {
					vm.ip = exitTarget
				}
			} else {
				// General fallback: delegate to the full JScript less-than comparison.
				if !vm.jsTruthy(vm.jsLess(varVal, limit)) {
					vm.ip = exitTarget
				}
			}

		case OpJSJumpIfNotUndefined:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			v := vm.stack[vm.sp]
			if v.Type != VTJSUndefined {
				if target < vm.ip {
					jsBackJumpCount++
					if jsBackJumpCount > jsBackJumpLimit {
						return vm.newMappedAxonASPError(ErrScriptTimeout, nil, fmt.Sprintf("JScript loop iteration limit (%d back-jumps) exceeded", jsBackJumpLimit))
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

		case OpJSForOfCleanup:
			// Remove stale for-of enumerator on early exit (break/throw).
			forOfPos := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			delete(vm.jsForOfItems, forOfPos)
			if vm.sp >= 0 {
				vm.pop()
			}

		case OpJSForOf:
			// Iterate over the VALUES of an iterable (array, string, Set, Map, etc.).
			// Format: nameConstIdx(2) + exitTarget(4)
			foNameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			foExitTarget := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4

			opPos := vm.ip - 7
			foState := vm.jsForOfItems[opPos]
			if foState == nil {
				// First encounter: collect iterable values and pop the source.
				if vm.sp < 0 {
					vm.ip = foExitTarget
					continue
				}
				source := vm.stack[vm.sp]
				values := vm.jsEnumerateForOfValues(source)
				foState = &jsForOfEnumerator{values: values, index: 0}
				vm.jsForOfItems[opPos] = foState
			}

			if foState.index >= len(foState.values) {
				// Exhausted: clean up and jump past the loop.
				delete(vm.jsForOfItems, opPos)
				if vm.sp >= 0 {
					vm.pop()
				}
				vm.ip = foExitTarget
				continue
			}

			// Assign the next value to the loop variable.
			vm.jsSetName(vm.constants[foNameIdx].Str, foState.values[foState.index])
			foState.index++

		case OpJSComputedPropertySet:
			// Pops key (top), value (next), object (next); sets object[key] = value.
			// The outer object reference below remains on the stack.
			key := vm.pop()
			val := vm.pop()
			obj := vm.pop()
			vm.jsIndexSet(obj, key, val)

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
			if left.Type == VTInteger && right.Type == VTInteger {
				if diff, ok := jsSubtractIntegersNoOverflow(left.Num, right.Num); ok {
					vm.push(NewInteger(diff))
					break
				}
				vm.push(NewDouble(float64(left.Num) - float64(right.Num)))
				break
			}
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

		case OpJSUnaryPlus:
			val := vm.pop()
			vm.push(vm.jsToNumber(val))

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
			newVal := oldVal
			switch oldVal.Type {
			case VTInteger:
				if next, ok := jsAddIntegersNoOverflow(oldVal.Num, 1); ok {
					newVal.Num = next
				} else {
					newVal = NewDouble(float64(oldVal.Num) + 1)
				}
			case VTDouble:
				newVal.Flt++
			default:
				newVal = vm.jsIncrementNumberValue(oldVal)
			}
			vm.jsSetName(nameStr, newVal)
			vm.push(oldVal)

		case OpJSPostDecrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			oldVal := vm.jsGetName(nameStr)
			newVal := oldVal
			switch oldVal.Type {
			case VTInteger:
				if next, ok := jsSubtractIntegersNoOverflow(oldVal.Num, 1); ok {
					newVal.Num = next
				} else {
					newVal = NewDouble(float64(oldVal.Num) - 1)
				}
			case VTDouble:
				newVal.Flt--
			default:
				newVal = vm.jsDecrementNumberValue(oldVal)
			}
			vm.jsSetName(nameStr, newVal)
			vm.push(oldVal)

		case OpJSPreIncrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			oldVal := vm.jsGetName(nameStr)
			newVal := oldVal
			switch oldVal.Type {
			case VTInteger:
				if next, ok := jsAddIntegersNoOverflow(oldVal.Num, 1); ok {
					newVal.Num = next
				} else {
					newVal = NewDouble(float64(oldVal.Num) + 1)
				}
			case VTDouble:
				newVal.Flt++
			default:
				newVal = vm.jsIncrementNumberValue(oldVal)
			}
			vm.jsSetName(nameStr, newVal)
			vm.push(newVal)

		case OpJSPreDecrement:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			oldVal := vm.jsGetName(nameStr)
			newVal := oldVal
			switch oldVal.Type {
			case VTInteger:
				if next, ok := jsSubtractIntegersNoOverflow(oldVal.Num, 1); ok {
					newVal.Num = next
				} else {
					newVal = NewDouble(float64(oldVal.Num) - 1)
				}
			case VTDouble:
				newVal.Flt--
			default:
				newVal = vm.jsDecrementNumberValue(oldVal)
			}
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

		case OpJSExponent:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.jsExponent(a, b))

		case OpJSCoalesce:
			b, a := vm.pop(), vm.pop()
			vm.push(vm.jsCoalesce(a, b))

		case OpJSJumpIfNullish:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			val := vm.pop()
			if val.Type == VTNull || val.Type == VTJSUndefined {
				vm.ip = target
			}

		case OpJSJumpIfNotNullish:
			target := int(binary.BigEndian.Uint32(vm.bytecode[vm.ip:]))
			vm.ip += 4
			val := vm.pop()
			if val.Type != VTNull && val.Type != VTJSUndefined {
				vm.ip = target
			}

		case OpJSExponentAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			right := vm.pop()
			left := vm.jsGetName(nameStr)
			result := vm.jsExponent(left, right)
			vm.jsSetName(nameStr, result)
			vm.push(result)

		case OpJSLogicalAndAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			val := vm.jsGetName(nameStr)
			right := vm.pop()
			if vm.jsTruthy(val) {
				val = vm.jsLogicalAnd(val, right)
				vm.jsSetName(nameStr, val)
			}
			vm.push(val)

		case OpJSLogicalOrAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			val := vm.jsGetName(nameStr)
			right := vm.pop()
			if !vm.jsTruthy(val) {
				val = vm.jsLogicalOr(val, right)
				vm.jsSetName(nameStr, val)
			}
			vm.push(val)

		case OpJSCoalesceAssign:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			nameStr := vm.constants[nameIdx].Str
			val := vm.jsGetName(nameStr)
			right := vm.pop()
			if val.Type == VTNull || val.Type == VTJSUndefined {
				val = right
				vm.jsSetName(nameStr, val)
			}
			vm.push(val)

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
			stackLen := len(vm.jsCallStack)
			result := vm.jsNew(constructor, args)
			if len(vm.jsCallStack) == stackLen || result.Type != VTJSUndefined {
				vm.push(result)
			}

		case OpJSMemberDelete:
			memberIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			obj := vm.pop()
			success := false
			if obj.Type == VTJSProxy {
				success = vm.jsProxyDelete(obj, vm.constants[memberIdx].Str)
			} else {
				success = vm.jsMemberDelete(obj, vm.constants[memberIdx].Str)
			}
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

		case OpJSStrictModeEnter:
			vm.jsStrictMode = true

		case OpJSStrictModeExit:
			vm.jsStrictMode = false

		case OpJSBlockScopeEnter:
			// Push a new block scope for let/const declarations
			vm.jsBlockScopes = append(vm.jsBlockScopes, make(map[string]Value, 4))
			vm.jsBlockScopeConst = append(vm.jsBlockScopeConst, make(map[string]struct{}, 2))
			vm.jsBlockScopeTDZ = append(vm.jsBlockScopeTDZ, make(map[string]struct{}, 2))
			vm.jsBlockScopeDepth++

		case OpJSBlockScopeExit:
			// Pop the current block scope
			if vm.jsBlockScopeDepth > 0 {
				vm.jsBlockScopes = vm.jsBlockScopes[:len(vm.jsBlockScopes)-1]
				vm.jsBlockScopeConst = vm.jsBlockScopeConst[:len(vm.jsBlockScopeConst)-1]
				vm.jsBlockScopeTDZ = vm.jsBlockScopeTDZ[:len(vm.jsBlockScopeTDZ)-1]
				vm.jsBlockScopeDepth--
			}

		case OpJSTDZRegisterLet:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			name := vm.constants[nameIdx].Str
			if vm.jsBlockScopeDepth > 0 && len(vm.jsBlockScopes) > 0 {
				vm.jsBlockScopes[len(vm.jsBlockScopes)-1][name] = Value{Type: VTJSUndefined}
				vm.jsBlockScopeTDZ[len(vm.jsBlockScopeTDZ)-1][name] = struct{}{}
			}

		case OpJSTDZRegisterConst:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			name := vm.constants[nameIdx].Str
			if vm.jsBlockScopeDepth > 0 && len(vm.jsBlockScopes) > 0 {
				vm.jsBlockScopes[len(vm.jsBlockScopes)-1][name] = Value{Type: VTJSUndefined}
				vm.jsBlockScopeConst[len(vm.jsBlockScopeConst)-1][name] = struct{}{}
				vm.jsBlockScopeTDZ[len(vm.jsBlockScopeTDZ)-1][name] = struct{}{}
			}

		case OpJSLetDeclare:
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			name := vm.constants[nameIdx].Str
			for i := len(vm.jsBlockScopes) - 1; i >= 0; i-- {
				if _, inTDZ := vm.jsBlockScopeTDZ[i][name]; inTDZ {
					delete(vm.jsBlockScopeTDZ[i], name)
					break
				}
			}

		case OpJSConstInitialize:
			// Pop value from stack, set const variable and clear its TDZ marker
			nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
			vm.ip += 2
			name := vm.constants[nameIdx].Str
			val := vm.pop()
			// Find the innermost block scope that has this const name and is still in TDZ
			for i := len(vm.jsBlockScopes) - 1; i >= 0; i-- {
				if _, inTDZ := vm.jsBlockScopeTDZ[i][name]; inTDZ {
					vm.jsBlockScopes[i][name] = val
					delete(vm.jsBlockScopeTDZ[i], name)
					break
				}
				if _, exists := vm.jsBlockScopes[i][name]; exists {
					vm.jsBlockScopes[i][name] = val
					break
				}
			}

		case OpJSForIterEnter:
			// Create a child env frame with per-iteration copies of the loop variable for closure capture.
			// Each iteration gets its own env frame so that closures created in the body capture
			// the iteration-specific value rather than a shared mutable slot.
			numVars := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			vm.ensureJSRootEnv()
			parentID := vm.jsActiveEnvID
			childID := vm.allocJSID()
			bindings := make(map[string]Value, numVars+2)
			for range numVars {
				nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				vm.ip += 2
				varName := vm.constants[nameIdx].Str
				// Read from block scopes first (where let/const vars live), then env chain.
				if varVal, found := vm.jsGetBlockScopeValue(varName); found {
					bindings[varName] = varVal
				} else {
					bindings[varName] = vm.jsGetNameFromEnv(parentID, varName)
				}
			}
			// Do NOT delete this frame on exit: closures created in the body reference it.
			vm.jsEnvItems[childID] = &jsEnvFrame{parentID: parentID, bindings: bindings}
			vm.jsActiveEnvID = childID

		case OpJSForIterExit:
			// Write the loop variable back from the per-iteration env frame into the block scope,
			// then restore the outer env frame WITHOUT deleting the per-iteration frame so that
			// any closures created during this iteration can still read the captured value.
			numVars := int(binary.BigEndian.Uint16(vm.bytecode[vm.ip:]))
			vm.ip += 2
			vm.ensureJSRootEnv()
			currentEnv := vm.jsEnvItems[vm.jsActiveEnvID]
			if currentEnv == nil {
				vm.ip += numVars * 2
				break
			}
			parentID := currentEnv.parentID
			for range numVars {
				nameIdx := binary.BigEndian.Uint16(vm.bytecode[vm.ip:])
				vm.ip += 2
				varName := vm.constants[nameIdx].Str
				// Write updated value from the per-iteration frame back to the block scope.
				if val, ok := currentEnv.bindings[varName]; ok {
					vm.jsSetBlockScopeValue(varName, val)
				}
			}
			// Restore the outer env without deleting the per-iteration frame: closures hold a
			// reference to it and must still be able to read the captured loop variable.
			vm.jsActiveEnvID = parentID

		case OpJSForIterEnterFast:
			// Non-capturing lexical iteration path: no child env allocation needed.

		case OpJSForIterExitFast:
			// Non-capturing lexical iteration path: no env write-back needed.

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

		case OpSwap:
			if vm.sp >= 1 {
				vm.stack[vm.sp], vm.stack[vm.sp-1] = vm.stack[vm.sp-1], vm.stack[vm.sp]
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

	if len(vm.callStack) == 0 && len(vm.jsCallStack) == 0 && len(vm.jsMicrotaskQueue) > 0 {
		vm.jsProcessMicrotasks()
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
	for i := range localCount {
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
	for i := range indexes {
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
	// Assign new value — clone VTRecord values so array elements own their
	// own copy (value semantics), matching VB6 UDT array behaviour.
	if assigned.Type == VTRecord && assigned.Rec != nil {
		assigned = vm.cloneValue(assigned)
	}
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
	switch current.Type {
	case VTInteger, VTDouble, VTBool, VTDate:
		return NewInteger(0)
	case VTString:
		return NewString("")
	case VTObject, VTNativeObject:
		return Value{Type: VTNothing}
	default:
		return NewEmpty()
	}
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

	if collectionValue, exists := vm.requestCollectionValueItems[objID]; exists {
		switch {
		case member == "" || strings.EqualFold(member, "Item"):
			if len(args) >= 1 {
				selector := args[0].String()
				if index, err := strconv.Atoi(selector); err == nil {
					if index < 1 || index > collectionValue.Count() {
						vm.raiseASPIndexOutOfRange()
						return Value{Type: VTEmpty}
					}
					return NewString(collectionValue.Values[index-1])
				}
				return NewString(collectionValue.Item(selector))
			}
			return NewString(collectionValue.Joined())
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(collectionValue.Count()))
		default:
			vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Object doesn't support this property or method")
			return Value{Type: VTEmpty}
		}
	}

	// Response cookie items: Response.Cookies(name)(subkey) = value (set) / get.
	if cookieName, exists := vm.responseCookieItems[objID]; exists {
		switch {
		case member == "":
			if len(args) >= 2 {
				// Sub-key set: Response.Cookies(name)(subkey) = value
				vm.host.Response().SetCookieSubKey(cookieName, args[0].String(), args[len(args)-1].String())
				return Value{Type: VTEmpty}
			}
			if len(args) == 1 {
				// Sub-key get: Response.Cookies(name)(subkey)
				return NewString(vm.host.Response().GetCookieSubKey(cookieName, args[0].String()))
			}
			// Default get: Response.Cookies(name)
			return NewString(vm.host.Response().GetCookieValue(cookieName))
		case strings.EqualFold(member, "Domain"), strings.EqualFold(member, "Path"),
			strings.EqualFold(member, "Expires"), strings.EqualFold(member, "Secure"),
			strings.EqualFold(member, "HttpOnly"), strings.EqualFold(member, "Value"),
			strings.EqualFold(member, "Name"):
			if len(args) >= 1 {
				vm.host.Response().SetCookieProperty(cookieName, member, args[len(args)-1].String())
				return Value{Type: VTEmpty}
			}
			return NewString(vm.host.Response().GetCookieProperty(cookieName, member))
		}
		return Value{Type: VTEmpty}
	}

	if g3dateObject, exists := vm.g3dateItems[objID]; exists {
		return g3dateObject.DispatchMethod(member, args)
	}

	if g3testObject, exists := vm.g3testItems[objID]; exists {
		return g3testObject.DispatchMethod(member, args)
	}

	if g3mdObject, exists := vm.g3mdItems[objID]; exists {
		return g3mdObject.DispatchMethod(member, args)
	}

	if g3searchObject, exists := vm.g3searchItems[objID]; exists {
		return g3searchObject.DispatchMethod(member, args)
	}

	if g3stringBuilderObject, exists := vm.g3stringBuilderItems[objID]; exists {
		return g3stringBuilderObject.DispatchMethod(member, args)
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
	if g3axonliveObject, exists := vm.g3axonliveItems[objID]; exists {
		return g3axonliveObject.DispatchMethod(member, args)
	}
	if g3axonliveProxy, exists := vm.g3axonliveProxyItems[objID]; exists {
		return g3axonliveProxy.DispatchMethod(member, args)
	}
	if g3filesObject, exists := vm.g3filesItems[objID]; exists {
		return g3filesObject.DispatchMethod(member, args)
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
	if pdfDoc, exists := vm.pdfDocItems[objID]; exists {
		return pdfDoc.DispatchMethod(member, args)
	}
	if pdfPage, exists := vm.pdfPageItems[objID]; exists {
		return pdfPage.DispatchMethod(member, args)
	}
	if pdfCanvas, exists := vm.pdfCanvasItems[objID]; exists {
		return pdfCanvas.DispatchMethod(member, args)
	}
	if pdfFont, exists := vm.pdfFontItems[objID]; exists {
		return pdfFont.DispatchMethod(member, args)
	}
	if pdfPages, exists := vm.pdfPagesItems[objID]; exists {
		return pdfPages.DispatchMethod(member, args)
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

	if wscriptEnvObject, exists := vm.wscriptEnvironmentItems[objID]; exists {
		return wscriptEnvObject.DispatchMethod(member, args)
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
	if colResult, handled := vm.dispatchCollectionMethod(objID, member, args); handled {
		return colResult
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

	if canvasObj, exists := vm.g3imageCanvasItems[objID]; exists {
		return vm.dispatchG3ImageCanvasMethod(canvasObj, member, args)
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

	// Route console object calls shared by VBScript and JScript.
	if objID == nativeObjectConsole {
		return consoleDispatch(vm, member, args)
	}

	switch objID {
	case nativeObjectResponse: // Response
		response := vm.host.Response()
		switch {
		case strings.EqualFold(member, "Write"):
			if len(args) > 0 {
				response.Write(vm.valueToResponseString(args[0]))
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
		case strings.EqualFold(member, "LCID"):
			if len(args) >= 1 {
				session := vm.host.Session()
				if session != nil {
					session.SetLCID(vm.asInt(args[0]))
				}
				return Value{Type: VTEmpty}
			}
			session := vm.host.Session()
			if session != nil {
				return NewInteger(int64(session.GetLCID()))
			}
			return NewInteger(0)
		case strings.EqualFold(member, "CodePage"):
			if len(args) >= 1 {
				session := vm.host.Session()
				if session != nil {
					session.SetCodePage(vm.asInt(args[0]))
				}
				return Value{Type: VTEmpty}
			}
			session := vm.host.Session()
			if session != nil {
				return NewInteger(int64(session.GetCodePage()))
			}
			return NewInteger(0)
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
		if strings.EqualFold(member, "Count") {
			return NewInteger(int64(vm.host.Response().GetCookieCount()))
		}
		if strings.EqualFold(member, "Key") {
			if len(args) >= 1 {
				return NewString(vm.host.Response().GetCookieKey(vm.asInt(args[0])))
			}
			return Value{Type: VTNativeObject, Num: nativeResponseCookiesKeyMethod}
		}
		if member == "" && len(args) >= 1 {
			return vm.newResponseCookieItem(args[0].String())
		}
		return Value{Type: VTEmpty}
	case nativeResponseCookiesKeyMethod:
		if member == "" && len(args) >= 1 {
			return NewString(vm.host.Response().GetCookieKey(vm.asInt(args[0])))
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
				if value, ok := request.QueryString.GetValue(args[0].String()); ok {
					return vm.newRequestCollectionValueItem(value)
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
				if value, ok := request.Form.GetValue(args[0].String()); ok {
					return vm.newRequestCollectionValueItem(value)
				}
				return Value{Type: VTEmpty}
			}
			return Value{Type: VTEmpty}
		case strings.EqualFold(member, "Cookies"):
			if len(args) == 1 {
				if value, ok := request.Cookies.GetValue(args[0].String()); ok {
					return vm.newRequestCollectionValueItem(value)
				}
				return Value{Type: VTEmpty}
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
		if (member == "" || strings.EqualFold(member, "Item")) && len(args) >= 1 {
			value, _ := vm.host.Request().QueryString.GetValue(args[0].String())
			return vm.newRequestCollectionValueItem(value)
		}
		if strings.EqualFold(member, "Count") {
			return NewInteger(int64(vm.host.Request().QueryString.Count()))
		}
		if strings.EqualFold(member, "Key") && len(args) >= 1 {
			return NewString(vm.host.Request().QueryString.Key(vm.asInt(args[0])))
		}
		return Value{Type: VTEmpty}
	case nativeRequestForm:
		if (member == "" || strings.EqualFold(member, "Item")) && len(args) >= 1 {
			if vm.host.Request().IsBinaryReadUsed() {
				return vm.newRequestCollectionValueItem(asp.RequestCollectionValue{})
			}
			vm.host.Request().MarkFormUsed()
			value, _ := vm.host.Request().Form.GetValue(args[0].String())
			return vm.newRequestCollectionValueItem(value)
		}
		if strings.EqualFold(member, "Count") {
			if vm.host.Request().IsBinaryReadUsed() {
				return NewInteger(0)
			}
			vm.host.Request().MarkFormUsed()
			return NewInteger(int64(vm.host.Request().Form.Count()))
		}
		if strings.EqualFold(member, "Key") && len(args) >= 1 {
			if vm.host.Request().IsBinaryReadUsed() {
				return NewString("")
			}
			vm.host.Request().MarkFormUsed()
			return NewString(vm.host.Request().Form.Key(vm.asInt(args[0])))
		}
		return Value{Type: VTEmpty}
	case nativeRequestCookies:
		if (member == "" || strings.EqualFold(member, "Item")) && len(args) == 1 {
			value, _ := vm.host.Request().Cookies.GetValue(args[0].String())
			return vm.newRequestCollectionValueItem(value)
		}
		if (member == "" || strings.EqualFold(member, "Item")) && len(args) >= 2 {
			return NewString(vm.host.Request().GetCookieAttribute(args[0].String(), args[1].String()))
		}
		if strings.EqualFold(member, "Count") {
			return NewInteger(int64(vm.host.Request().Cookies.Count()))
		}
		if strings.EqualFold(member, "Key") && len(args) >= 1 {
			return NewString(vm.host.Request().Cookies.Key(vm.asInt(args[0])))
		}
		return Value{Type: VTEmpty}
	case nativeRequestServerVariables:
		if (member == "" || strings.EqualFold(member, "Item")) && len(args) >= 1 {
			value, _ := vm.host.Request().ServerVars.GetValue(args[0].String())
			return vm.newRequestCollectionValueItem(value)
		}
		if strings.EqualFold(member, "Count") {
			return NewInteger(int64(vm.host.Request().ServerVars.Count()))
		}
		if strings.EqualFold(member, "Key") && len(args) >= 1 {
			return NewString(vm.host.Request().ServerVars.Key(vm.asInt(args[0])))
		}
		return Value{Type: VTEmpty}
	case nativeRequestClientCertificate:
		if (member == "" || strings.EqualFold(member, "Item")) && len(args) >= 1 {
			value, _ := vm.host.Request().ClientCertificate.GetValue(args[0].String())
			return vm.newRequestCollectionValueItem(value)
		}
		if strings.EqualFold(member, "Count") {
			return NewInteger(int64(vm.host.Request().ClientCertificate.Count()))
		}
		if strings.EqualFold(member, "Key") && len(args) >= 1 {
			return NewString(vm.host.Request().ClientCertificate.Key(vm.asInt(args[0])))
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
				// VBScript compatibility: each CreateObject attempt starts with a clean Err state.
				// If object creation fails, vm.raise repopulates Err for Resume Next checks.
				vm.errClear()
				progID := strings.TrimSpace(args[0].String())
				progIDKey := strings.ToLower(progID)
				if progIDKey == "g3stringbuilder" {
					return vm.newG3StringBuilderObject()
				}
				if progIDKey == "g3search" {
					return vm.newG3SearchObject()
				}
				if progIDKey == "g3md" {
					return vm.newG3MDObject()
				}
				if progIDKey == "g3date" {
					return vm.newG3DateObject()
				}
				if progIDKey == "g3testsuite" || progIDKey == "g3test" {
					return vm.newG3TestObject()
				}
				if defaultAlgorithm, ok := g3cryptoResolveProgID(progID); ok {
					return vm.newG3CryptoObject(defaultAlgorithm)
				}
				if progIDKey == "g3axon" || progIDKey == "g3axon.functions" {
					return vm.newAxonLibrary()
				}
				if progIDKey == "g3json" {
					return vm.newG3JSONObject()
				}
				if progIDKey == "g3db" {
					return vm.newG3DBObject()
				}
				if progIDKey == "g3http" || progIDKey == "g3http.functions" {
					return vm.newG3HTTPObject()
				}
				if progIDKey == "g3mail" || progIDKey == "cdonts.newmail" || progIDKey == "cdo.message" || progIDKey == "persits.mailsender" || progIDKey == "smtpsvg.mailer" {
					return vm.newG3MailObjectWithProgID(progIDKey)
				}
				if progIDKey == "g3image" {
					return vm.newG3ImageObject()
				}
				if progIDKey == "persits.jpeg" {
					return vm.newPersitsJpegObject()
				}
				if progIDKey == "g3files" {
					return vm.newG3FilesObject()
				}
				if progIDKey == "g3template" {
					return vm.newG3TemplateObject()
				}
				if progIDKey == "g3zip" {
					return vm.newG3ZipObject()
				}
				if progIDKey == "g3zlib" {
					return vm.newG3ZLIBObject()
				}
				if progIDKey == "g3tar" {
					return vm.newG3TARObject()
				}
				if progIDKey == "g3zstd" {
					return vm.newG3ZSTDObject()
				}
				if progIDKey == "g3fc" {
					return vm.newG3FCObject()
				}
				if progIDKey == "g3axonlive" {
					return vm.newG3AxonLiveObject()
				}
				if progIDKey == "wscript.shell" {
					return vm.newWScriptShellObject()
				}
				if progIDKey == "adox.catalog" {
					return vm.newADOXCatalogObject()
				}
				if progIDKey == "mswc.adrotator" {
					return vm.newG3AdRotatorObject()
				}
				if progIDKey == "mswc.browsertype" {
					return vm.newG3BrowserTypeObject()
				}
				if progIDKey == "mswc.nextlink" {
					return vm.newG3NextLinkObject()
				}
				if progIDKey == "mswc.contentrotator" {
					return vm.newG3ContentRotatorObject()
				}
				if progIDKey == "mswc.counters" {
					return vm.newG3CountersObject()
				}
				if progIDKey == "mswc.pagecounter" {
					return vm.newG3PageCounterObject()
				}
				if progIDKey == "mswc.tools" {
					return vm.newG3ToolsObject()
				}
				if progIDKey == "mswc.myinfo" {
					return vm.newG3MyInfoObject()
				}
				if progIDKey == "mswc.permissionchecker" {
					return vm.newG3PermissionCheckerObject()
				}
				if progIDKey == "msxml2.serverxmlhttp" || progIDKey == "msxml2.xmlhttp" || progIDKey == "microsoft.xmlhttp" {
					obj := NewMsXML2ServerXMLHTTP(vm)
					id := vm.nextDynamicNativeID
					vm.nextDynamicNativeID++
					vm.msxmlServerItems[id] = obj
					return Value{Type: VTNativeObject, Num: id}
				}
				if progIDKey == "msxml2.domdocument" || progIDKey == "microsoft.xmldom" {
					obj := NewMsXML2DOMDocument(vm)
					id := vm.nextDynamicNativeID
					vm.nextDynamicNativeID++
					vm.msxmlDOMItems[id] = obj
					return Value{Type: VTNativeObject, Num: id}
				}
				if progIDKey == "g3pdf" || progIDKey == "persits.pdf" || progIDKey == "asp.pdf" {
					obj := NewG3PDF(vm)
					if progIDKey == "persits.pdf" || progIDKey == "asp.pdf" {
						obj.isPersitsMode = true
					}
					id := vm.nextDynamicNativeID
					vm.nextDynamicNativeID++
					vm.pdfItems[id] = obj
					return Value{Type: VTNativeObject, Num: id}
				}
				if progIDKey == "g3fileuploader" || progIDKey == "persits.upload" || progIDKey == "softartisans.fileup" || progIDKey == "aspupload" {
					return vm.newG3FileUploaderObjectWithProgID(progIDKey)
				}
				if progIDKey == "scripting.filesystemobject" {
					return vm.newFSORootObject()
				}
				if progIDKey == "scripting.dictionary" {
					return vm.newDictionaryObject()
				}
				if progIDKey == "collection" {
					return vm.newCollectionObject()
				}
				if progIDKey == "adodb.stream" {
					return vm.newADODBStreamObject()
				}
				if progIDKey == "adodb.connection" {
					return vm.newADODBConnection()
				}
				if progIDKey == "adodbole.connection" {
					return vm.newADODBOLEConnection()
				}
				if progIDKey == "adodb.recordset" {
					return vm.newADODBRecordset()
				}
				if progIDKey == "adodb.command" {
					return vm.newADODBCommand()
				}
				if progIDKey == "vbscript.regexp" || progIDKey == "regexp" {
					return vm.newRegExpObject()
				}
				value, err := server.CreateObject(progID)
				if err != nil {
					aspErr := asp.NewVBScriptASPError(vbscript.ActiveXCannotCreateObject, "Server.CreateObject", "ASP", "Invalid class string", "", 0, 0)
					aspErr.Number = asp.InvalidProgIDHRESULT
					server.SetLastError(aspErr)
					vm.raise(vbscript.ActiveXCannotCreateObject, "Invalid class string")
					return Value{Type: VTEmpty}
				}
				resolved := vm.applicationValueToValue(value)
				if resolved.Type == VTEmpty {
					aspErr := asp.NewVBScriptASPError(vbscript.ActiveXCannotCreateObject, "Server.CreateObject", "ASP", "Invalid class string", "", 0, 0)
					aspErr.Number = asp.InvalidProgIDHRESULT
					server.SetLastError(aspErr)
					vm.raise(vbscript.ActiveXCannotCreateObject, "Invalid class string")
					return Value{Type: VTEmpty}
				}
				return resolved
			}
			_, _ = server.CreateObject("")
			vm.raise(vbscript.ActiveXCannotCreateObject, "Invalid class string")
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
					var aspErr *asp.ASPError
					if os.IsNotExist(err) || errors.Is(err, os.ErrNotExist) {
						aspErr = &asp.ASPError{
							ASPCode:        173,
							ASPDescription: "An error occurred when processing the target. The file was not found.",
							Number:         -2147467259, // 0x80004005
							Source:         "Active Server Pages",
							Description:    "An error occurred when processing the target. The file was not found.",
							File:           absPath,
							Category:       "Active Server Pages",
						}
					} else {
						var jsSyntaxErr *jscript.JSSyntaxError
						var vbSyntaxErr *vbscript.VBSyntaxError
						if errors.As(err, &jsSyntaxErr) || errors.As(err, &vbSyntaxErr) {
							aspErr = CompilerErrorToASPError(err, absPath)
						} else {
							aspErr = RuntimeErrorToASPError(err, absPath)
						}
					}
					server.SetLastError(aspErr)
					vme := &VMError{
						Code:           vbscript.VBSyntaxErrorCode(aspErr.ASPCode),
						Line:           aspErr.Line,
						Column:         aspErr.Column,
						File:           aspErr.File,
						Msg:            aspErr.ASPDescription,
						ASPCode:        aspErr.ASPCode,
						ASPDescription: aspErr.ASPDescription,
						Category:       aspErr.Category,
						Description:    aspErr.Description,
						Number:         aspErr.Number,
						Source:         aspErr.Source,
					}
					vm.raiseVMError(vme)
				}
				// If the child script ended the response (Response.End/Redirect),
				// propagate the termination signal to the parent VM so execution
				// of the calling page stops immediately (IIS Classic ASP behavior).
				if vm.host.Response().IsEnded() {
					panic(asp.ResponseEndSignal)
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
			for i := range keys {
				values[i] = NewString(keys[i])
			}
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
		case strings.EqualFold(member, "Items"):
			keys := session.GetAllKeys()
			sort.Strings(keys)
			values := make([]Value, 0, len(keys))
			for i := range keys {
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
	case nativeObjectSessionContentsKeyMethod:
		if member == "" && len(args) >= 1 {
			idx := vm.asInt(args[0]) - 1
			keys := vm.host.Session().GetAllKeys()
			sort.Strings(keys)
			if idx >= 0 && idx < len(keys) {
				return NewString(keys[idx])
			}
			return NewString("")
		}
		return NewString("")
	case nativeObjectSessionStaticObjectsKeyMethod:
		if member == "" && len(args) >= 1 {
			idx := vm.asInt(args[0]) - 1
			static := vm.host.Session().GetStaticObjectsCopy()
			keys := make([]string, 0, len(static))
			for k := range static {
				keys = append(keys, k)
			}
			sort.Strings(keys)
			if idx >= 0 && idx < len(keys) {
				return NewString(keys[idx])
			}
			return NewString("")
		}
		return NewString("")
	case nativeObjectApplicationContentsKeyMethod:
		if member == "" && len(args) >= 1 {
			idx := vm.asInt(args[0]) - 1
			contents := vm.host.Application().GetContentsCopy()
			keys := make([]string, 0, len(contents))
			for k := range contents {
				keys = append(keys, k)
			}
			sort.Strings(keys)
			if idx >= 0 && idx < len(keys) {
				return NewString(keys[idx])
			}
			return NewString("")
		}
		return NewString("")
	case nativeObjectApplicationStaticObjectsKeyMethod:
		if member == "" && len(args) >= 1 {
			idx := vm.asInt(args[0]) - 1
			static := vm.host.Application().GetStaticObjectsCopy()
			keys := make([]string, 0, len(static))
			for k := range static {
				keys = append(keys, k)
			}
			sort.Strings(keys)
			if idx >= 0 && idx < len(keys) {
				return NewString(keys[idx])
			}
			return NewString("")
		}
		return NewString("")
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
	if g3dateObject, exists := vm.g3dateItems[target.Num]; exists {
		return g3dateObject.DispatchPropertyGet(member)
	}

	if g3mdObject, exists := vm.g3mdItems[target.Num]; exists {
		return g3mdObject.DispatchPropertyGet(member)
	}

	if g3dateObject, exists := vm.g3dateItems[target.Num]; exists {
		return g3dateObject.DispatchPropertyGet(member)
	}

	if g3searchObject, exists := vm.g3searchItems[target.Num]; exists {
		return g3searchObject.DispatchPropertyGet(member)
	}

	if g3stringBuilderObject, exists := vm.g3stringBuilderItems[target.Num]; exists {
		return g3stringBuilderObject.DispatchPropertyGet(member)
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
	if canvasObj, exists := vm.g3imageCanvasItems[target.Num]; exists {
		return vm.dispatchG3ImageCanvasPropertyGet(canvasObj, member)
	}
	if fontObj, exists := vm.g3imageFontItems[target.Num]; exists {
		return vm.dispatchG3ImageFontPropertyGet(fontObj, member)
	}
	if penObj, exists := vm.g3imagePenItems[target.Num]; exists {
		return vm.dispatchG3ImagePenPropertyGet(penObj, member)
	}
	if g3axonliveObject, exists := vm.g3axonliveItems[target.Num]; exists {
		return g3axonliveObject.DispatchPropertyGet(member)
	}
	if g3axonliveProxy, exists := vm.g3axonliveProxyItems[target.Num]; exists {
		return g3axonliveProxy.DispatchPropertyGet(member)
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
	if pdfDoc, exists := vm.pdfDocItems[target.Num]; exists {
		return vm.dispatchPdfDocumentPropertyGet(pdfDoc, member)
	}
	if pdfPage, exists := vm.pdfPageItems[target.Num]; exists {
		return vm.dispatchPdfPagePropertyGet(pdfPage, member)
	}
	if pdfCanvas, exists := vm.pdfCanvasItems[target.Num]; exists {
		return vm.dispatchPdfCanvasPropertyGet(pdfCanvas, member)
	}
	if pdfFont, exists := vm.pdfFontItems[target.Num]; exists {
		return vm.dispatchPdfFontPropertyGet(pdfFont, member)
	}
	if pdfPages, exists := vm.pdfPagesItems[target.Num]; exists {
		return vm.dispatchPdfPagesPropertyGet(pdfPages, member)
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

	if wscriptEnvObject, exists := vm.wscriptEnvironmentItems[target.Num]; exists {
		return wscriptEnvObject.DispatchPropertyGet(member)
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
	if colResult, handled := vm.dispatchCollectionPropertyGet(target.Num, member); handled {
		return colResult
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
		if vm.engineMode == EngineModeJavaScript || len(vm.jsCallStack) > 0 || vm.jsActiveEnvID != 0 {
			switch {
			case strings.EqualFold(member, "log"):
				return vm.jsCreateIntrinsicFunction("console.log", "ConsoleLog")
			case strings.EqualFold(member, "warn"):
				return vm.jsCreateIntrinsicFunction("console.warn", "ConsoleWarn")
			case strings.EqualFold(member, "error"):
				return vm.jsCreateIntrinsicFunction("console.error", "ConsoleError")
			case strings.EqualFold(member, "info"):
				return vm.jsCreateIntrinsicFunction("console.info", "ConsoleInfo")
			case strings.EqualFold(member, "debug"):
				return vm.jsCreateIntrinsicFunction("console.debug", "ConsoleDebug")
			case strings.EqualFold(member, "trace"):
				return vm.jsCreateIntrinsicFunction("console.trace", "ConsoleTrace")
			case strings.EqualFold(member, "clear"):
				return vm.jsCreateIntrinsicFunction("console.clear", "ConsoleClear")
			}
		}
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
		case strings.EqualFold(member, "LCID"):
			if vm.host != nil && vm.host.Session() != nil {
				return NewInteger(int64(vm.host.Session().GetLCID()))
			}
			return NewInteger(0)
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
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeObjectSessionContentsKeyMethod}
		}
	case nativeObjectSessionStaticObjects:
		session := vm.host.Session()
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(session.GetStaticObjectsCopy())))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeObjectSessionStaticObjectsKeyMethod}
		}
	case nativeObjectApplicationContents:
		application := vm.host.Application()
		application.WaitForServer(vm.host.Server())
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(application.GetContentsCopy())))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeObjectApplicationContentsKeyMethod}
		}
	case nativeObjectApplicationStaticObjects:
		application := vm.host.Application()
		application.WaitForServer(vm.host.Server())
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(len(application.GetStaticObjectsCopy())))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeObjectApplicationStaticObjectsKeyMethod}
		}
	case nativeResponseCookies:
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(vm.host.Response().GetCookieCount()))
		case strings.EqualFold(member, "Key"):
			return Value{Type: VTNativeObject, Num: nativeResponseCookiesKeyMethod}
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

	if collectionValue, exists := vm.requestCollectionValueItems[target.Num]; exists {
		switch {
		case strings.EqualFold(member, "Count"):
			return NewInteger(int64(collectionValue.Count()))
		case strings.EqualFold(member, "HasKeys"):
			return NewBool(collectionValue.HasKeys())
		case strings.EqualFold(member, "Item"):
			return vm.newNativeObjectProxy(target.Num, "Item", nil)
		default:
			vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Object doesn't support this property or method")
			return Value{Type: VTEmpty}
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

	// Dynamic native objects: Response cookie items
	if cookieName, exists := vm.responseCookieItems[objID]; exists {
		vm.host.Response().SetCookieProperty(cookieName, member, val.String())
		return
	}

	// Dynamic native objects: G3MD
	if g3mdObject, exists := vm.g3mdItems[objID]; exists {
		g3mdObject.DispatchPropertySet(member, val)
		return
	}

	if g3searchObject, exists := vm.g3searchItems[objID]; exists {
		g3searchObject.DispatchPropertySet(member, val)
		return
	}

	if g3stringBuilderObject, exists := vm.g3stringBuilderItems[objID]; exists {
		g3stringBuilderObject.DispatchPropertySet(member, val)
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
	if fontObj, exists := vm.g3imageFontItems[objID]; exists {
		vm.dispatchG3ImageFontPropertySet(fontObj, member, val)
		return
	}
	if penObj, exists := vm.g3imagePenItems[objID]; exists {
		vm.dispatchG3ImagePenPropertySet(penObj, member, val)
		return
	}

	if g3axonliveProxy, exists := vm.g3axonliveProxyItems[objID]; exists {
		g3axonliveProxy.DispatchPropertySet(member, []Value{val})
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
	if pdfDoc, exists := vm.pdfDocItems[objID]; exists {
		vm.dispatchPdfDocumentPropertySet(pdfDoc, member, val)
		return
	}
	if pdfPage, exists := vm.pdfPageItems[objID]; exists {
		vm.dispatchPdfPagePropertySet(pdfPage, member, val)
		return
	}
	if pdfCanvas, exists := vm.pdfCanvasItems[objID]; exists {
		vm.dispatchPdfCanvasPropertySet(pdfCanvas, member, val)
		return
	}
	if pdfFont, exists := vm.pdfFontItems[objID]; exists {
		vm.dispatchPdfFontPropertySet(pdfFont, member, val)
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
	if vm.dispatchCollectionPropertySet(objID, member, val) {
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
		case strings.EqualFold(member, "LCID"):
			if vm.host != nil && vm.host.Session() != nil {
				vm.host.Session().SetLCID(vm.asInt(val))
			}
		case strings.EqualFold(member, "CodePage"):
			if vm.host != nil && vm.host.Session() != nil {
				vm.host.Session().SetCodePage(vm.asInt(val))
			}
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

// newG3SearchObject creates a native G3SEARCH object used by Server.CreateObject("G3SEARCH").
func (vm *VM) newG3SearchObject() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3searchItems[objID] = NewG3Search(vm)
	return Value{Type: VTNativeObject, Num: objID}
}

// newG3StringBuilderObject creates a native G3STRINGBUILDER object used by Server.CreateObject("G3STRINGBUILDER").
func (vm *VM) newG3StringBuilderObject() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3stringBuilderItems[objID] = NewG3StringBuilder()
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
	if v.Type == VTArgRef {
		v = vm.stack[int(v.Num)]
	}
	v = resolveCallable(vm, v)
	if v.Type == VTJSObject || v.Type == VTJSFunction {
		if vm.jsObjectStringProperty(v, "__js_type") == "Error" {
			name := vm.jsObjectStringProperty(v, "name")
			msg := vm.jsObjectStringProperty(v, "message")
			if name == "" {
				name = vm.jsObjectStringProperty(v, "__js_ctor")
			}
			if name == "" {
				name = "Error"
			}
			if msg == "" {
				return name
			}
			return name + ": " + msg
		}
	}
	if v.Type == VTNativeObject {
		if cookieName, exists := vm.responseCookieItems[v.Num]; exists {
			return vm.host.Response().GetCookieValue(cookieName)
		}
		if collectionValue, exists := vm.requestCollectionValueItems[v.Num]; exists {
			if len(collectionValue.Values) == 0 {
				isJS := len(vm.jsCallStack) > 0 || vm.jsActiveEnvID != 0 || vm.jsRootEnvID != 0 || len(vm.jsTryStack) > 0 || len(vm.jsErrStack) > 0 || vm.engineMode == EngineModeJavaScript
				if isJS {
					return "undefined"
				}
			}
			return collectionValue.Joined()
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
		if v.Num == nativeRequestQueryString {
			return vm.host.Request().QueryString.String()
		}
		if v.Num == nativeRequestForm {
			return vm.host.Request().Form.String()
		}
		if v.Num == nativeRequestCookies {
			return vm.host.Request().Cookies.String()
		}
		if v.Num == nativeRequestServerVariables {
			return vm.host.Request().ServerVars.String()
		}
		if v.Num == nativeRequestClientCertificate {
			return vm.host.Request().ClientCertificate.String()
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

// valueToResponseString applies Response.Write coercion rules for mixed VBScript/JScript values.
func (vm *VM) valueToResponseString(v Value) string {
	v = resolveCallable(vm, v)
	isJS := vm.engineMode == EngineModeJavaScript || len(vm.jsCallStack) > 0 || vm.jsActiveEnvID != 0 || vm.jsRootEnvID != 0
	if !isJS {
		return vm.valueToString(v)
	}

	switch v.Type {
	case VTArray:
		return vm.jsArrayToString(v)
	case VTJSFunction:
		if out, handled := vm.jsCallMember(v, "toString", nil); handled {
			return vm.valueToString(out)
		}
		return vm.jsToString(v)
	case VTJSProxy, VTJSUndefined:
		return vm.jsToString(v)
	case VTJSObject:
		objType := vm.jsObjectStringProperty(v, "__js_type")
		if objType == "Boolean" {
			prim, handled := vm.jsCallMember(v, "valueOf", nil)
			if handled && prim.Type == VTBool {
				return prim.String()
			}
		}
		if out, handled := vm.jsCallMember(v, "toString", nil); handled {
			return vm.valueToString(out)
		}
		return vm.jsToString(v)
	default:
		return vm.valueToString(v)
	}
}

// newRequestCollectionValueItem creates one native object wrapper for one Request collection entry value.
func (vm *VM) newRequestCollectionValueItem(value asp.RequestCollectionValue) Value {
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.requestCollectionValueItems[id] = value
	return Value{Type: VTNativeObject, Num: id}
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
		for i := range values {
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

// getCollator returns the locale-aware collator for Option Compare Text,
// lazily building it from the current VM LCID. The collator is cached and
// rebuilt only when the effective LCID changes.
func (vm *VM) getCollator() *collate.Collator {
	lcid := builtinCurrentLCID(vm)
	if vm.collator != nil && vm.collatorLCID == lcid {
		return vm.collator
	}
	tag := language.Make(GetGoLanguageFromMSLCID(MSLCID(lcid)))
	vm.collator = collate.New(tag, collate.IgnoreCase)
	vm.collatorLCID = lcid
	return vm.collator
}

// textEqual returns true when two strings compare equal under Option Compare Text
// using the VM's locale-aware collator.
func (vm *VM) textEqual(a, b string) bool {
	return vm.getCollator().CompareString(a, b) == 0
}

// textCompare returns -1, 0, 1 for locale-aware case-insensitive string comparison
// using the VM's configured LCID.
func (vm *VM) textCompare(a, b string) int {
	return vm.getCollator().CompareString(a, b)
}

// textContains returns true when needle is a locale-aware case-insensitive substring
// of haystack, using the VM's configured LCID.
func (vm *VM) textContains(haystack, needle string) bool {
	if needle == "" {
		return true
	}
	hayRunes := []rune(haystack)
	ndlRunes := []rune(needle)
	if len(ndlRunes) > len(hayRunes) {
		return false
	}
	for i := 0; i+len(ndlRunes) <= len(hayRunes); i++ {
		if vm.textEqual(string(hayRunes[i:i+len(ndlRunes)]), needle) {
			return true
		}
	}
	return false
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

	file, line, column := vm.mapRuntimeLocation(vm.lastLine, vm.lastColumn)
	vme := &VMError{
		Code:           vbscript.VBSyntaxErrorCode(number),
		Line:           line,
		Column:         column,
		File:           file,
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

	for i := len(vm.callStack) - 1; i >= 0; i-- {
		frame := vm.callStack[i]
		if !frame.savedOnResumeNext {
			continue
		}
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
		vm.callStack = vm.callStack[:i]
		vm.sp = curSP
		vm.fp = curFP
		vm.ip = frame.returnIP
		vm.activeClassObjectID = frame.boundObj
		vm.onResumeNext = frame.savedOnResumeNext
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
	if strings.EqualFold(progID, "G3SEARCH") {
		return vm.newG3SearchObject()
	}
	if strings.EqualFold(progID, "G3STRINGBUILDER") {
		return vm.newG3StringBuilderObject()
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
		if len(vm.jsCallStack) > 0 || vm.jsActiveEnvID != 0 || vm.jsRootEnvID != 0 || len(vm.jsTryStack) > 0 || len(vm.jsErrStack) > 0 || vm.engineMode == EngineModeJavaScript {
			vm.jsThrowOutOfStackSpace()
			return
		}
		vm.raise(vbscript.StackOverflow, "Stack overflow")
		return
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
		Events:     make(map[string]bool),
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
func (vm *VM) registerRuntimeClassField(className string, fieldName string, isPublic bool, fieldType ValueType, classType string) {
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

	classDef.Fields[strings.ToLower(trimmedFieldName)] = RuntimeClassFieldDef{
		Name:      trimmedFieldName,
		IsPublic:  isPublic,
		Type:      fieldType,
		ClassType: classType,
	}
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

// getCurrentFuncEntryPoint returns the bytecode offset of the first instruction in the current procedure.
func (vm *VM) getCurrentFuncEntryPoint() int {
	if len(vm.callStack) == 0 {
		return 0
	}
	callee := vm.callStack[len(vm.callStack)-1].callee
	if callee.Type == VTUserSub {
		return int(callee.Num)
	}
	return 0
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

	lowerMethodName := strings.ToLower(strings.TrimSpace(methodName))
	methodDef, exists := classDef.Methods[lowerMethodName]
	if exists {
		if requirePublic && !methodDef.IsPublic {
			return Value{Type: VTEmpty}, false
		}
		return methodDef.Target, true
	}

	// Phase 5: Interface Polymorphism
	// If the method is not found directly, check for InterfaceName_MethodName.
	for _, interfaceName := range classDef.Interfaces {
		interfaceMethodName := strings.ToLower(interfaceName + "_" + methodName)
		if methodDef, exists := classDef.Methods[interfaceMethodName]; exists {
			return methodDef.Target, true
		}
	}

	return Value{Type: VTEmpty}, false
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
	resolvedPropertyName := strings.ToLower(strings.TrimSpace(propertyName))
	if resolvedPropertyName == "__default__" {
		defaultProp, hasDefault := classDef.Properties["__default__"]
		if hasDefault {
			var targetNum int64 = -1
			if defaultProp.HasGet {
				targetNum = defaultProp.GetTarget.Num
			} else if defaultProp.HasLet {
				targetNum = defaultProp.LetTarget.Num
			} else if defaultProp.HasSet {
				targetNum = defaultProp.SetTarget.Num
			}
			if targetNum != -1 {
				for name, propDef := range classDef.Properties {
					if name == "__default__" {
						continue
					}
					if (propDef.HasGet && propDef.GetTarget.Num == targetNum) ||
						(propDef.HasLet && propDef.LetTarget.Num == targetNum) ||
						(propDef.HasSet && propDef.SetTarget.Num == targetNum) {
						resolvedPropertyName = name
						break
					}
				}
			}
		}
	}
	propertyDef, exists := classDef.Properties[resolvedPropertyName]
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
	resolvedPropertyName := strings.ToLower(strings.TrimSpace(propertyName))
	if resolvedPropertyName == "__default__" {
		defaultProp, hasDefault := classDef.Properties["__default__"]
		if hasDefault {
			var targetNum int64 = -1
			if defaultProp.HasGet {
				targetNum = defaultProp.GetTarget.Num
			} else if defaultProp.HasLet {
				targetNum = defaultProp.LetTarget.Num
			} else if defaultProp.HasSet {
				targetNum = defaultProp.SetTarget.Num
			}
			if targetNum != -1 {
				for name, propDef := range classDef.Properties {
					if name == "__default__" {
						continue
					}
					if (propDef.HasGet && propDef.GetTarget.Num == targetNum) ||
						(propDef.HasLet && propDef.LetTarget.Num == targetNum) ||
						(propDef.HasSet && propDef.SetTarget.Num == targetNum) {
						resolvedPropertyName = name
						break
					}
				}
			}
		}
	}
	propertyDef, exists := classDef.Properties[resolvedPropertyName]
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
		args[i] = vm.unwrapArgRefValue(arg)
	}
	return byRefs
}

// unwrapArgRefValue resolves one VTArgRef to its current underlying slot value.
func (vm *VM) unwrapArgRefValue(arg Value) Value {
	if arg.Type != VTArgRef {
		return resolveCallable(vm, arg)
	}

	isGlobal := arg.ArgRefIsGlobal()
	isLocal := arg.ArgRefIsLocal()
	isClassMember := arg.ArgRefIsClassMember()
	idx := arg.ArgRefIdx()

	var rawVal Value
	if isClassMember {
		rawVal = vm.getClassMemberValueByObjectID(vm.activeClassObjectID, arg.Str)
	} else if isGlobal {
		if idx >= 0 && idx < len(vm.Globals) {
			rawVal = vm.Globals[idx]
			if rawVal.Type == VTEmpty && idx < len(vm.globalNames) && strings.HasPrefix(vm.globalNames[idx], "__static_") {
				rawVal.Interface = "static"
			}
		}
	} else if isLocal {
		slot := vm.fp + idx
		if slot >= 0 && slot < StackSize {
			rawVal = vm.stack[slot]
		}
	}

	return resolveCallable(vm, rawVal)
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

// allocateCallMemberIC stores one call-site cache entry and returns a stable 32-bit cache ID.
func (vm *VM) allocateCallMemberIC(entry callMemberICEntry) uint32 {
	if vm == nil {
		return 0
	}
	if vm.callMemberIC == nil {
		vm.callMemberIC = make(map[uint32]callMemberICEntry, 32)
	}
	id := vm.nextCallMemberICID
	if id == 0 {
		id = 1
	}
	vm.callMemberIC[id] = entry
	vm.nextCallMemberICID = id + 1
	if vm.nextCallMemberICID == 0 {
		vm.nextCallMemberICID = 1
	}
	return id
}

// bytesToVBByteString converts raw bytes to a VB-compatible byte-string representation.
// Each source byte maps to one codepoint in range 0..255 so byte-oriented built-ins
// (LenB/MidB/InStrB/AscB) can round-trip binary payloads without UTF-8 shrinking.
func bytesToVBByteString(data []byte) string {
	if len(data) == 0 {
		return ""
	}
	runes := make([]rune, len(data))
	for i := range data {
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
	localCount := max(target.UserSubLocalCount(), paramCount)

	frameLast := vm.fp + localCount - 1
	if localCount > 0 && frameLast >= StackSize {
		vm.raise(vbscript.StackOverflow, "Stack overflow")
	}

	for i := range localCount {
		vm.stack[vm.fp+i] = Value{Type: VTEmpty}
		vm.localTypes[vm.fp+i] = VTEmpty // Clear any stale declared type from previous frames
	}

	// Apply VB6 As Type declarations for this function's local variables.
	entryPoint := int(target.Num)
	if slotTypes, exists := vm.funcLocalTypes[entryPoint]; exists {
		slotClassTypes := vm.funcLocalClassTypes[entryPoint]
		for offset, declaredType := range slotTypes {
			if offset < localCount {
				idx := vm.fp + offset
				vm.localTypes[idx] = declaredType
				val := vm.zeroValueForType(declaredType)
				if declaredType == VTObject && slotClassTypes != nil {
					if className, ok := slotClassTypes[offset]; ok {
						val.Interface = className
					}
				}
				vm.stack[idx] = val
			}
		}
	}

	// Handle Optional parameters with default values and ParamArray.
	paramArrayIdx := target.UserSubParamArrayIdx()
	optionalMask := target.UserSubOptionalMask()
	defaults, hasDefaults := vm.funcParamDefaults[entryPoint]
	hasParamArray := paramArrayIdx >= 0 && paramArrayIdx < paramCount

	// Copy provided arguments to callee frame slots.
	argIdx := 0
	paramIdx := 0
	for paramIdx < paramCount {
		// If this is the ParamArray slot, pack remaining args into an array.
		if hasParamArray && paramIdx == paramArrayIdx {
			// Copy remaining args into a new array to avoid aliasing the argBuffer.
			remaining := max(len(args)-argIdx, 0)
			vba := NewVBArray(0, remaining)
			for j := range remaining {
				vba.Values[j] = args[argIdx+j]
			}
			vm.stack[vm.fp+paramIdx] = Value{Type: VTArray, Arr: vba}
			paramIdx++
			break // ParamArray must be last param
		}

		if argIdx < len(args) {
			// Normal argument provided.
			vm.stack[vm.fp+paramIdx] = args[argIdx]
			argIdx++
		} else if (optionalMask>>uint(paramIdx))&1 == 1 {
			// Optional parameter with no argument provided - use default value.
			if hasDefaults && paramIdx < len(defaults) && defaults[paramIdx] >= 0 {
				defaultIdx := defaults[paramIdx]
				if defaultIdx >= 0 && defaultIdx < len(vm.constants) {
					vm.stack[vm.fp+paramIdx] = vm.constants[defaultIdx]
				} else {
					vm.stack[vm.fp+paramIdx] = Value{Type: VTEmpty}
				}
			} else {
				// Optional without explicit default = VTEmpty.
				vm.stack[vm.fp+paramIdx] = Value{Type: VTEmpty}
			}
		} else {
			// Missing required parameter: leave as VTEmpty (already initialized).
			// In strict mode this could raise an error, but VBScript allows it,
			// passing Empty for missing non-optional params at runtime.
		}
		paramIdx++
	}

	// If we have a ParamArray slot and no args were consumed for it (the loop above
	// only reaches this if paramArrayIdx was already handled), ensure the slot gets
	// at minimum an empty array. This handles the case where ParamArray is the only
	// parameter and no args were passed.
	if hasParamArray && paramArrayIdx >= 0 && paramArrayIdx < paramCount {
		slot := vm.fp + paramArrayIdx
		if vm.stack[slot].Type != VTArray {
			vm.stack[slot] = Value{Type: VTArray, Arr: NewVBArrayFromValues(0, nil)}
		}
	}

	if localCount > 0 {
		vm.sp = frameLast
	} else {
		vm.sp = vm.fp - 1
	}
	vm.ip = int(target.Num)
	return true
}

// handleWithEventsRegister registers a variable as having WithEvents.
func (vm *VM) handleWithEventsRegister(className, varName string) {
	lowerVarName := strings.ToLower(varName)
	if className == "" {
		// Global variable
		for i, name := range vm.globalNames {
			if strings.EqualFold(name, varName) {
				vm.globalWithEvents[uint16(i)] = true
				break
			}
		}
	} else {
		// Class member
		if def, ok := vm.runtimeClasses[strings.ToLower(className)]; ok {
			if field, ok2 := def.Fields[lowerVarName]; ok2 {
				field.WithEvents = true
				def.Fields[lowerVarName] = field
				vm.runtimeClasses[strings.ToLower(className)] = def
			}
		}
	}
}

// bindWithEvents scans for handlers and registers them in the target object.
func (vm *VM) bindWithEvents(container *RuntimeClassInstance, varName string, obj Value) {
	if obj.Type != VTObject || obj.Num == 0 {
		return
	}
	target := vm.runtimeClassItems[obj.Num]
	if target == nil {
		return
	}

	targetDef, ok := vm.runtimeClasses[strings.ToLower(target.ClassName)]
	if !ok {
		return
	}

	prefix := strings.ToLower(varName) + "_"
	containerID := int64(0)
	if container != nil {
		// Find container ID
		for id, inst := range vm.runtimeClassItems {
			if inst == container {
				containerID = id
				break
			}
		}
	}

	for eventName := range targetDef.Events {
		handlerName := prefix + eventName
		var handler Value
		found := false

		if container != nil {
			// Look in container class methods
			containerDef := vm.runtimeClasses[strings.ToLower(container.ClassName)]
			if methodDef, ok2 := containerDef.Methods[handlerName]; ok2 {
				handler = methodDef.Target
				found = true
			}
		} else {
			// Look in global scope
			for i, name := range vm.globalNames {
				if strings.EqualFold(name, handlerName) {
					handler = vm.Globals[i]
					if handler.Type == VTUserSub {
						found = true
					}
					break
				}
			}
		}

		if found {
			if target.Observers == nil {
				target.Observers = make(map[string][]EventObserver)
			}
			target.Observers[eventName] = append(target.Observers[eventName], EventObserver{
				Handler:     handler,
				ContainerID: containerID,
				Prefix:      prefix,
			})
		}
	}
}

// unbindWithEvents removes handlers registered for a specific variable from an object.
func (vm *VM) unbindWithEvents(container *RuntimeClassInstance, varName string, obj Value) {
	if obj.Type != VTObject || obj.Num == 0 {
		return
	}
	target := vm.runtimeClassItems[obj.Num]
	if target == nil {
		return
	}

	prefix := strings.ToLower(varName) + "_"
	containerID := int64(0)
	if container != nil {
		for id, inst := range vm.runtimeClassItems {
			if inst == container {
				containerID = id
				break
			}
		}
	}

	for eventName, observers := range target.Observers {
		newObservers := observers[:0]
		for _, obs := range observers {
			if obs.ContainerID == containerID && obs.Prefix == prefix {
				continue
			}
			newObservers = append(newObservers, obs)
		}
		if len(newObservers) == 0 {
			delete(target.Observers, eventName)
		} else {
			target.Observers[eventName] = newObservers
		}
	}
}

// handleRaiseEvent dispatches an event to all registered observers.
func (vm *VM) handleRaiseEvent(eventName string, argCount int) {
	if vm.activeClassObjectID == 0 {
		// RaiseEvent only allowed inside classes
		for range argCount {
			vm.pop()
		}
		return
	}

	instance := vm.runtimeClassItems[vm.activeClassObjectID]
	if instance == nil {
		for range argCount {
			vm.pop()
		}
		return
	}

	lowerEventName := strings.ToLower(eventName)
	observers, ok := instance.Observers[lowerEventName]
	if !ok || len(observers) == 0 {
		for range argCount {
			vm.pop()
		}
		return
	}

	args := make([]Value, argCount)
	for i := argCount - 1; i >= 0; i-- {
		args[i] = vm.pop()
	}

	for _, obs := range observers {
		// obs.Handler is VTUserSub
		// We need to call it.
		vm.beginUserSubCall(obs.Handler, args, true, obs.ContainerID)
		// Note: beginUserSubCall will change vm.ip and continue the loop.
		// For now, let's support only ONE observer per event for simplicity.
		break
	}
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
		ClassName:       classDef.Name,
		Members:         make(map[string]Value),
		Observers:       make(map[string][]EventObserver),
		WithEventsNames: make(map[string]bool),
		refCount:        1, // Initial reference from creation
		terminated:      false,
	}
	// Track creation order so Class_Terminate fires in reverse-construction order at cleanup.
	vm.classInstanceOrder = append(vm.classInstanceOrder, instanceID)
	for fieldKey, fieldDef := range classDef.Fields {
		if fieldDef.WithEvents {
			vm.runtimeClassItems[instanceID].WithEventsNames[fieldKey] = true
		}
		if len(fieldDef.Dims) > 0 {
			vm.runtimeClassItems[instanceID].Members[fieldKey] = ValueFromVBArray(buildDimArray(fieldDef.Dims))
		} else {
			if fieldDef.Type == VTObject {
				val := Value{Type: VTNothing}
				if fieldDef.ClassType != "" {
					val.Interface = fieldDef.ClassType
				}
				vm.runtimeClassItems[instanceID].Members[fieldKey] = val
			} else if fieldDef.Type == VTRecord {
				vm.runtimeClassItems[instanceID].Members[fieldKey] = vm.initRecordValueRecursive(fieldDef.ClassType)
			} else if fieldDef.Type != VTEmpty {
				vm.runtimeClassItems[instanceID].Members[fieldKey] = vm.zeroValueForType(fieldDef.Type)
			} else {
				vm.runtimeClassItems[instanceID].Members[fieldKey] = Value{Type: VTEmpty}
			}
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

// applyLocalVarTypes applies VB6 As Type declarations from the compiler to the VM's
// funcLocalTypes map.
func (vm *VM) applyLocalVarTypes(compiler *Compiler) {
	if vm == nil || compiler == nil || vm.funcLocalTypes == nil {
		return
	}
	for entryPoint, typesMap := range compiler.FuncLocalTypes() {
		slotTypes := make(map[int]ValueType)
		slotClassTypes := make(map[int]string)
		var localNames []string
		for _, constVal := range vm.constants {
			if constVal.Type == VTUserSub && int(constVal.Num) == entryPoint {
				localNames = constVal.Names
				break
			}
		}
		if len(localNames) == 0 {
			continue
		}
		for offset, name := range localNames {
			lower := strings.ToLower(name)
			if declaredType, exists := typesMap[lower]; exists && declaredType != VTEmpty {
				slotTypes[offset] = declaredType
				if declaredType == VTObject || declaredType == VTRecord {
					if className, ok := compiler.FuncLocalRecordTypes()[entryPoint][lower]; ok {
						slotClassTypes[offset] = className
					}
				}
			}
		}
		if len(slotTypes) > 0 {
			vm.funcLocalTypes[entryPoint] = slotTypes
		}
		if len(slotClassTypes) > 0 {
			if vm.funcLocalClassTypes == nil {
				vm.funcLocalClassTypes = make(map[int]map[int]string)
			}
			vm.funcLocalClassTypes[entryPoint] = slotClassTypes
		}
	}
}

// applyLocalVarTypesFromMaps applies VB6 As Type declarations from provided maps to the VM's
// funcLocalTypes and funcLocalClassTypes maps. It scans each VTUserSub constant, matches its
// local variable names against the maps, and stores the resulting slot-to-type mapping.
func (vm *VM) applyLocalVarTypesFromMaps(localVarTypes map[string]ValueType, localClassTypes map[string]string) {
	if vm == nil || vm.funcLocalTypes == nil || len(localVarTypes) == 0 {
		return
	}
	// Scan constants for VTUserSub entries to find function entry points and their local names.
	for _, constVal := range vm.constants {
		if constVal.Type != VTUserSub {
			continue
		}
		entryPoint := int(constVal.Num)
		localNames := constVal.Names
		if len(localNames) == 0 {
			continue
		}
		slotTypes := make(map[int]ValueType)
		slotClassTypes := make(map[int]string)
		for offset, name := range localNames {
			lower := strings.ToLower(name)
			if declaredType, exists := localVarTypes[lower]; exists && declaredType != VTEmpty {
				slotTypes[offset] = declaredType
				if declaredType == VTObject || declaredType == VTRecord {
					if className, ok := localClassTypes[lower]; ok {
						slotClassTypes[offset] = className
					}
				}
			}
		}
		if len(slotTypes) > 0 {
			vm.funcLocalTypes[entryPoint] = slotTypes
		}
		if len(slotClassTypes) > 0 {
			if vm.funcLocalClassTypes == nil {
				vm.funcLocalClassTypes = make(map[int]map[int]string)
			}
			vm.funcLocalClassTypes[entryPoint] = slotClassTypes
		}
	}
}

// applyGlobalVarTypes applies VB6 As Type declarations to the VM's global variable slots.
// Called during VM initialization from compiler metadata.
func (vm *VM) applyGlobalVarTypes(globalTypes map[string]ValueType, globalClassTypes map[string]string) {
	if vm == nil || len(globalTypes) == 0 {
		return
	}
	for name, declaredType := range globalTypes {
		if declaredType == VTEmpty {
			continue
		}
		// Find the global slot index by name.
		lower := strings.ToLower(name)
		for idx, gname := range vm.globalNames {
			if strings.EqualFold(gname, lower) {
				if idx < len(vm.globalTypes) {
					vm.globalTypes[idx] = declaredType
					if declaredType == VTObject || declaredType == VTRecord {
						if className, ok := globalClassTypes[lower]; ok {
							vm.globalClassTypes[uint16(idx)] = className
						}
					}
				}
				if idx < len(vm.Globals) {
					vm.Globals[idx] = vm.zeroValueForType(declaredType)
					if declaredType == VTObject {
						vm.Globals[idx].Interface = globalClassTypes[lower]
					}
				}
				break
			}
		}
	}
}

// zeroValueForType returns the zero-initialized Value for a given declared VB6 type.
func (vm *VM) zeroValueForType(t ValueType) Value {
	switch t {
	case VTInteger:
		return Value{Type: VTInteger, Num: 0}
	case VTDouble:
		return Value{Type: VTDouble, Flt: 0}
	case VTString:
		return Value{Type: VTString, Str: ""}
	case VTBool:
		return Value{Type: VTBool, Num: 0}
	case VTObject:
		return Value{Type: VTNothing}
	case VTRecord:
		return Value{Type: VTRecord} // Requires initialization via ExtOpInitRecord
	default:
		return Value{Type: VTEmpty}
	}
}

// coerceToDeclaredType attempts to coerce a Value to match a declared VB6 type.
// If coercion is impossible, returns an error describing the type mismatch.
func (vm *VM) coerceToDeclaredType(v Value, declaredType ValueType) (Value, error) {
	if v.Type == VTEmpty || v.Type == VTNull {
		// Empty/Null can be assigned to any typed variable as the zero value.
		return vm.zeroValueForType(declaredType), nil
	}

	if declaredType == VTRecord {
		if v.Type == VTRecord {
			return vm.cloneValue(v), nil
		}
		return Value{}, fmt.Errorf("Type mismatch: expected Record")
	}

	switch declaredType {
	case VTInteger:
		switch v.Type {
		case VTInteger:
			return v, nil
		case VTDouble:
			return Value{Type: VTInteger, Num: int64(v.Flt)}, nil
		case VTBool:
			return Value{Type: VTInteger, Num: v.Num}, nil
		case VTString:
			// Try to parse as number
			n, err := parseInt64(v.Str)
			if err != nil {
				return Value{}, fmt.Errorf("Type mismatch: cannot convert '%s' to Integer", v.Str)
			}
			return Value{Type: VTInteger, Num: n}, nil
		default:
			return Value{}, fmt.Errorf("Type mismatch: cannot convert to Integer")
		}

	case VTDouble:
		switch v.Type {
		case VTDouble:
			return v, nil
		case VTInteger:
			return Value{Type: VTDouble, Flt: float64(v.Num)}, nil
		case VTBool:
			return Value{Type: VTDouble, Flt: float64(v.Num)}, nil
		case VTString:
			f, err := parseFloat64(v.Str)
			if err != nil {
				return Value{}, fmt.Errorf("Type mismatch: cannot convert '%s' to Double", v.Str)
			}
			return Value{Type: VTDouble, Flt: f}, nil
		default:
			return Value{}, fmt.Errorf("Type mismatch: cannot convert to Double")
		}

	case VTString:
		switch v.Type {
		case VTString:
			return v, nil
		case VTInteger, VTDouble, VTBool:
			return Value{Type: VTString, Str: v.String()}, nil
		default:
			return Value{}, fmt.Errorf("Type mismatch: cannot convert to String")
		}

	case VTBool:
		switch v.Type {
		case VTBool:
			return v, nil
		case VTInteger:
			return Value{Type: VTBool, Num: v.Num}, nil
		case VTDouble:
			boolVal := int64(0)
			if v.Flt != 0 {
				boolVal = 1
			}
			return Value{Type: VTBool, Num: boolVal}, nil
		case VTString:
			lower := strings.ToLower(strings.TrimSpace(v.Str))
			switch lower {
			case "true":
				return Value{Type: VTBool, Num: 1}, nil
			case "false":
				return Value{Type: VTBool, Num: 0}, nil
			}
			return Value{}, fmt.Errorf("Type mismatch: cannot convert '%s' to Boolean", v.Str)
		default:
			return Value{}, fmt.Errorf("Type mismatch: cannot convert to Boolean")
		}

	case VTObject:
		if v.Type == VTObject || v.Type == VTNativeObject || v.Type == VTNothing {
			// Phase 5: Preserve or set Interface name from the variable's declared class/interface.
			// We store the class name in the Value struct to route calls correctly.
			// This is especially important for local variables and class members.
			if v.Type == VTObject {
				// We don't overwrite if it already has one? No, we SHOULD use the
				// declared type of the variable to enforce interface-based calling.
				// In VB6, if you do 'Dim x As IAnimal: Set x = New Dog',
				// calls through 'x' are routed to IAnimal_XXX.
				// But we need to know the interface name.
				// For globals/locals, it's in vm.globalClassTypes / vm.funcLocalClassTypes.
				// But coerceToDeclaredType doesn't know WHICH slot it's assigning to.
				// Wait, the caller (OpSetGlobal/OpSetLocal) knows.
				// I'll handle .Interface setting in the opcodes instead.
			}
			return v, nil
		}
		return Value{}, fmt.Errorf("Type mismatch: cannot convert to Object")

	default:
		// Unknown declared type, pass through
		return v, nil
	}
}

// parseInt64 attempts to parse a string as a 64-bit integer.
func parseInt64(s string) (int64, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, fmt.Errorf("empty string")
	}
	return strconv.ParseInt(s, 10, 64)
}

// parseFloat64 attempts to parse a string as a 64-bit float.
func parseFloat64(s string) (float64, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, fmt.Errorf("empty string")
	}
	return strconv.ParseFloat(s, 64)
}

func (vm *VM) raise(code vbscript.VBSyntaxErrorCode, msg string) {
	description := strings.TrimSpace(msg)
	if description == "" {
		description = code.String()
	}

	file, line, column := vm.mapRuntimeLocation(vm.lastLine, vm.lastColumn)
	vme := &VMError{
		Code:           code,
		Line:           line,
		Column:         column,
		File:           file,
		Msg:            description,
		ASPCode:        int(code),
		ASPDescription: description,
		Category:       "VBScript runtime",
		Description:    description,
		Number:         vbscript.HRESULTFromVBScriptCode(code),
		Source:         "VBScript runtime error",
	}
	if code == vbscript.ActiveXCannotCreateObject {
		vme.Number = asp.InvalidProgIDHRESULT
	}

	vm.raiseVMError(vme)
}

func (vm *VM) raiseVMError(vme *VMError) {
	vm.errSetFromVMError(vme)

	isJS := len(vm.jsCallStack) > 0 || vm.jsActiveEnvID != 0 || vm.jsRootEnvID != 0 || len(vm.jsTryStack) > 0 || len(vm.jsErrStack) > 0 || vm.engineMode == EngineModeJavaScript
	if isJS {
		vm.lastError = vme
		errObj := vm.jsCreateErrorObject("Error", vme.Msg)
		vm.jsMemberSet(errObj, "number", NewInteger(int64(vme.Number)))
		vm.jsMemberSet(errObj, "description", NewString(vme.Msg))
		vm.jsMemberSet(errObj, "message", NewString(vme.Msg))

		if len(vm.jsTryStack) > 0 {
			target := vm.jsTryStack[len(vm.jsTryStack)-1]
			vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
			vm.jsErrStack = append(vm.jsErrStack, errObj)
			vm.ip = target
			return
		}
		panic(&jsAsyncRejectionError{reason: errObj})
	}

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

func (vm *VM) raiseASPIndexOutOfRange() {
	file, line, column := vm.mapRuntimeLocation(vm.lastLine, vm.lastColumn)
	vme := &VMError{
		Code:           vbscript.SubscriptOutOfRange,
		Line:           line,
		Column:         column,
		File:           file,
		Msg:            "007~ASP 0105~Index out of range~An array index is out of range.",
		ASPCode:        105,
		ASPDescription: "An array index is out of range.",
		Category:       "ASP",
		Description:    "Index out of range",
		Number:         -2147467259,
		Source:         "ASP",
	}
	vm.raiseVMError(vme)
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

var recordPool = sync.Pool{
	New: func() any {
		return &VBRecord{}
	},
}

func (vm *VM) acquireRecord(size int) *VBRecord {
	rec := recordPool.Get().(*VBRecord)
	rec.releaseMark = false
	if cap(rec.Members) < size {
		rec.Members = make([]Value, size)
	} else {
		rec.Members = rec.Members[:size]
		// Clear values to avoid leaks or stale data
		for i := range rec.Members {
			rec.Members[i] = Value{}
		}
	}
	return rec
}

// releaseRecordValue releases one VTRecord value (including nested records) back to the record pool.
func (vm *VM) releaseRecordValue(v Value) {
	if v.Type != VTRecord || v.Rec == nil {
		return
	}
	vm.releaseRecord(v.Rec)
}

// releaseRecord returns one UDT record instance and nested members to the record pool.
func (vm *VM) releaseRecord(rec *VBRecord) {
	if rec == nil {
		return
	}
	if rec.releaseMark {
		return
	}
	rec.releaseMark = true
	for i := range rec.Members {
		if rec.Members[i].Type == VTRecord && rec.Members[i].Rec != nil {
			vm.releaseRecord(rec.Members[i].Rec)
		}
		rec.Members[i] = Value{}
	}
	rec.DefIdx = 0
	rec.Members = rec.Members[:0]
	recordPool.Put(rec)
}

// cloneValue deep-copies a Value representing a UDT/record.
func (vm *VM) cloneValue(v Value) Value {
	if v.Type == VTRecord && v.Rec != nil {
		rec := vm.acquireRecord(len(v.Rec.Members))
		rec.DefIdx = v.Rec.DefIdx
		for i, m := range v.Rec.Members {
			rec.Members[i] = vm.cloneValue(m)
		}
		return Value{Type: VTRecord, Rec: rec}
	} else if v.Type == VTArray && v.Arr != nil {
		return Value{Type: VTArray, Arr: vm.cloneArray(v.Arr)}
	}
	return v
}

// initRecordValueRecursive recursively allocates and initializes UDT fields and sub-UDTs to their zero values.
func (vm *VM) initRecordValueRecursive(udtName string) Value {
	lowerUDT := strings.ToLower(udtName)
	if defIdx, exists := vm.RecordDeclLookup[lowerUDT]; exists {
		decl := vm.RecordDecls[defIdx]
		rec := vm.acquireRecord(len(decl.Members))
		rec.DefIdx = defIdx
		for i, m := range decl.Members {
			if m.Type == VTRecord {
				rec.Members[i] = vm.initRecordValueRecursive(m.UDTName)
			} else {
				rec.Members[i] = vm.zeroValueForType(m.Type)
			}
		}
		return Value{Type: VTRecord, Rec: rec}
	}
	return Value{Type: VTRecord}
}

// cloneArray deep-copies a VBArray.
func (vm *VM) cloneArray(arr *VBArray) *VBArray {
	if arr == nil {
		return nil
	}
	values := make([]Value, len(arr.Values))
	for i := range arr.Values {
		values[i] = vm.cloneValue(arr.Values[i])
	}
	return NewVBArrayFromValues(arr.Lower, values)
}

func (vm *VM) ensureCLIMode() {
	if vm.executionMode != ExecutionModeCLI && vm.executionMode != ExecutionModeTUI {
		vm.raise(vbscript.PermissionDenied, "Native File I/O operations are restricted to CLI environment only")
	}
}

func (vm *VM) closeAllFiles() {
	if vm.fileIOItems != nil {
		for _, f := range vm.fileIOItems {
			if f != nil {
				f.Close()
			}
		}
		clear(vm.fileIOItems)
	}
	if vm.fileIOBufReaders != nil {
		clear(vm.fileIOBufReaders)
	}
	if vm.fileIOBufWriters != nil {
		for _, w := range vm.fileIOBufWriters {
			if w != nil {
				w.Flush()
			}
		}
		clear(vm.fileIOBufWriters)
	}
}

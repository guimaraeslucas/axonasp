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
	"errors"
	"fmt"
	"maps"
	"math"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"sync/atomic"

	"g3pix.com.br/axonasp/jscript"
	jsast "g3pix.com.br/axonasp/jscript/ast"
	"g3pix.com.br/axonasp/vbscript"
)

// dumpPreprocessedSourceEnabled controls whether compiled source is written to ./temp/ for debugging.
// Defaults to false. Enabled via SetDumpPreprocessedSourceEnabled or the DUMP_PREPROCESSED_SOURCE env var.
var dumpPreprocessedSourceEnabled atomic.Bool

// SetDumpPreprocessedSourceEnabled enables or disables the preprocessed source dump feature.
// Called during server/cli startup from the loaded configuration.
func SetDumpPreprocessedSourceEnabled(enabled bool) {
	dumpPreprocessedSourceEnabled.Store(enabled)
}

type SymbolTable struct {
	names  []string
	lookup map[string]int
}

// SourceMapEntry stores one sparse merged-line boundary and original source origin.
type SourceMapEntry struct {
	MergedLineStart int
	OriginalFile    string
	OriginalLine    int
}

// SourceMap maps merged lines to original source lines using sparse boundaries.
type SourceMap struct {
	entries []SourceMapEntry
}

// jscriptCompileLineAnchor maps generated JScript program lines back to merged source lines.
type jscriptCompileLineAnchor struct {
	GeneratedLineStart int
	MergedLineStart    int
}

// AddBoundary records one sparse mapping boundary.
func (m *SourceMap) AddBoundary(mergedLineStart int, originalFile string, originalLine int) {
	if m == nil || mergedLineStart <= 0 {
		return
	}
	entry := SourceMapEntry{MergedLineStart: mergedLineStart, OriginalFile: originalFile, OriginalLine: originalLine}
	if len(m.entries) == 0 {
		m.entries = append(m.entries, entry)
		return
	}
	lastIdx := len(m.entries) - 1
	last := m.entries[lastIdx]
	if last.MergedLineStart == mergedLineStart {
		m.entries[lastIdx] = entry
		return
	}
	if last.MergedLineStart > mergedLineStart {
		idx := sort.Search(len(m.entries), func(i int) bool {
			return m.entries[i].MergedLineStart >= mergedLineStart
		})
		if idx < len(m.entries) && m.entries[idx].MergedLineStart == mergedLineStart {
			m.entries[idx] = entry
			return
		}
		m.entries = append(m.entries, SourceMapEntry{})
		copy(m.entries[idx+1:], m.entries[idx:])
		m.entries[idx] = entry
		return
	}
	m.entries = append(m.entries, entry)
}

// ResolveLine maps one merged line to original file and local line.
func (m *SourceMap) ResolveLine(mergedLine int) (string, int, bool) {
	if m == nil || mergedLine <= 0 || len(m.entries) == 0 {
		return "", 0, false
	}
	idx := sort.Search(len(m.entries), func(i int) bool {
		return m.entries[i].MergedLineStart > mergedLine
	})
	if idx == 0 {
		return "", 0, false
	}
	entry := m.entries[idx-1]
	resolvedLine := entry.OriginalLine + (mergedLine - entry.MergedLineStart)
	if resolvedLine <= 0 {
		resolvedLine = entry.OriginalLine
	}
	return entry.OriginalFile, resolvedLine, true
}

// Entries returns a copy of sparse source-map entries.
func (m *SourceMap) Entries() []SourceMapEntry {
	if m == nil || len(m.entries) == 0 {
		return nil
	}
	out := make([]SourceMapEntry, len(m.entries))
	copy(out, m.entries)
	return out
}

// Clone returns a deep copy of source-map entries.
func (m *SourceMap) Clone() SourceMap {
	if m == nil || len(m.entries) == 0 {
		return SourceMap{}
	}
	cloned := make([]SourceMapEntry, len(m.entries))
	copy(cloned, m.entries)
	return SourceMap{entries: cloned}
}

type includeResolveOptions struct {
	siteRoot        string
	caseInsensitive bool
}

func defaultIncludeResolveOptions() includeResolveOptions {
	return includeResolveOptions{siteRoot: "", caseInsensitive: true}
}

func NewSymbolTable() *SymbolTable {
	return &SymbolTable{
		names:  make([]string, 0),
		lookup: make(map[string]int),
	}
}

func (s *SymbolTable) Add(name string) int {
	lower := strings.ToLower(name)
	if idx, exists := s.lookup[lower]; exists {
		return idx
	}
	idx := len(s.names)
	s.names = append(s.names, name)
	s.lookup[lower] = idx
	return idx
}

func (s *SymbolTable) Get(name string) (int, bool) {
	lower := strings.ToLower(name)
	idx, exists := s.lookup[lower]
	return idx, exists
}

func (s *SymbolTable) Names() []string {
	return s.names
}

func (s *SymbolTable) Count() int {
	return len(s.names)
}

// Compiler handles single-pass compilation of VBScript tokens into VM bytecode.
type Compiler struct {
	lexer                *vbscript.Lexer
	lexerMode            vbscript.LexerMode
	sourceCode           string
	next                 vbscript.Token
	tokenIndex           int
	bytecode             []byte
	constants            []Value
	constantMap          map[string]int // Deduplication map: content→index, populated for VTString only
	Globals              *SymbolTable
	locals               *SymbolTable                 // Current function scope
	declaredGlobals      map[string]bool              // Variables declared via Dim in global scope
	implicitGlobals      map[string]bool              // Variables implicitly created at page scope
	declaredLocals       map[string]bool              // Variables declared via Dim in local scope
	globalVarTypes       map[string]ValueType         // VB6 As Type declarations for global variables (VTEmpty = Variant)
	localVarTypes        map[string]ValueType         // VB6 As Type declarations for local variables (VTEmpty = Variant)
	globalRecordTypes    map[string]string            // UDT names for global variables
	localRecordTypes     map[string]string            // UDT names for local variables
	funcLocalTypes       map[int]map[string]ValueType // entryPoint -> localName -> type
	funcLocalRecordTypes map[int]map[string]string    // entryPoint -> localName -> record/class name
	constGlobals         map[string]bool              // Constants declared via Const in global scope
	constLocals          map[string]bool              // Constants declared via Const in local scope
	staticLocals         map[string]int               // Static local variables mapping (localName -> globalIndex)
	constLiteralGlobals  map[string]Value             // Compile-time known global constant values
	isLocal              bool                         // True if currently compiling a Sub/Function

	// Compilation Options
	optionExplicit         bool // Requires variables to be Dim'ed
	optionCompare          int  // 0: Binary (default), 1: Text
	optionBase             int  // 0 or 1 (default 0)
	optionStrict           bool // (Future) Enforces strict typing
	optionInfer            bool // (Future) Allows type inference
	sourceName             string
	includeSiteRoot        string
	includeCaseInsensitive bool
	sourceMap              SourceMap
	includeDeps            []string

	forwardCallPatches  map[string][]int
	forwardConstPatches map[string][]int
	lastCallTargetName  string
	lastCallTargetPos   int
	lastCallIsGlobal    bool
	lastCoercePos       int
	lastDebugLine       int
	lastDebugColumn     int
	prevToken           vbscript.Token
	lastToken           vbscript.Token
	tempCounter         int
	globalZeroArgFuncs  map[string]bool
	globalZeroArgSubs   map[string]bool
	classDecls          []CompiledClassDecl
	classDeclLookup     map[string]int
	recordDecls         []CompiledRecordDecl
	recordDeclLookup    map[string]int
	enumTypeLookup      map[string]map[string]Value // Enum type name -> member name -> constant value
	ObjectDeclarations  []*vbscript.ASPObjectToken
	currentClassName    string
	currentFunctionName string
	// dynamicMemberResolution enables implicit class-member lookup for dynamic
	// code compiled by Eval/Execute/ExecuteGlobal while keeping global assignment semantics.
	dynamicMemberResolution bool
	loopContexts            []loopContext
	jsLoopContexts          []*jsLoopContext // Loop contexts for JScript
	jsBreakContexts         []*jsBreakContext
	jsStrictMode            bool // Current JScript strict mode state
	jsIsDerivedConstructor  bool // True if currently compiling a derived class constructor
	jsTryDepth              int  // Current JScript try/catch/finally nesting depth
	jsClassFields           []jsast.ClassElement

	jsFunctionStrictModes map[int]bool      // Maps function start IP to strict mode
	jsBlockScopeStack     []map[string]bool // Stack of declared block-scoped variables (let/const)
	jsForIterScopes       []jsForIterScope  // Stack of active per-iteration for-let scopes
	jsOptionalChainExits  []int             // Stack of placeholder positions for ?. short-circuiting
	jsLocalScopeStack     []jsLocalScope    // Lexical stack for JScript local slot resolution
	jsLocalEnabled        bool              // True when local slot lowering is enabled for current function
	jsLocalSlotCount      int               // Number of local slots allocated for current function
	jsInGeneratorFunction bool              // True when compiling a generator body.
	jsCompileLineAnchors  []jscriptCompileLineAnchor
	jsNextICNodeID        uint32 // Next available inline cache node ID for JScript AST nodes
	jsICNodeCount         uint32 // Total inline cache nodes assigned across the program
	// withDepth tracks nesting level of With...End With blocks at compile time.
	// A value > 0 enables the leading-dot '.' statement and expression syntax.
	withDepth          int
	activeVBSConstants []VBSConstant
	// userGlobalsStart is the index of the first user-declared global variable slot.
	// Slots below this index are read-only pre-injected intrinsics, built-ins, or VBScript constants.
	// Only global slots at or above this index are eligible for ByRef argument write-back.
	userGlobalsStart int
	isEval           bool // True if compiling a VBScript expression for Eval()
	isJSModule       bool // True if compiling a pure JScript module

	labelMap            map[string]int
	forwardLabelPatches map[string][]int

	// funcParamDefaults maps function entry point -> per-parameter constant pool indices
	// for Optional parameter default values. -1 means no default for that parameter.
	funcParamDefaults map[int][]int

	lastEmittedType    ValueType
	lastEmittedUDTName string
}

func (c *Compiler) updateLastEmittedType(vt ValueType, udt string) {
	c.lastEmittedType = vt
	c.lastEmittedUDTName = udt
}

func (c *Compiler) lastEmittedUDT() (string, bool) {
	if c.lastEmittedType == VTRecord {
		return c.lastEmittedUDTName, true
	}
	return "", false
}

func (c *Compiler) getUDTMemberIndex(udtName, memberName string) (int, ValueType, string, bool) {
	lowerUDT := strings.ToLower(udtName)
	udtIdx, ok := c.recordDeclLookup[lowerUDT]
	if !ok {
		return -1, VTEmpty, "", false
	}
	decl := c.recordDecls[udtIdx]
	for i, m := range decl.Members {
		if strings.EqualFold(m.Name, memberName) {
			return i, m.Type, m.UDTName, true
		}
	}
	return -1, VTEmpty, "", false
}

func (c *Compiler) lastEmittedUDTNameFromOp() (string, bool) {
	if len(c.bytecode) < 3 {
		return "", false
	}
	op := OpCode(c.bytecode[len(c.bytecode)-3])
	idx := int(binary.BigEndian.Uint16(c.bytecode[len(c.bytecode)-2:]))

	var varName string
	switch op {
	case OpGetGlobal:
		if idx >= 0 && idx < len(c.Globals.names) {
			varName = strings.ToLower(c.Globals.names[idx])
			if udt, ok := c.globalRecordTypes[varName]; ok {
				return udt, true
			}
		}
	case OpGetLocal:
		if idx >= 0 && idx < len(c.locals.names) {
			varName = strings.ToLower(c.locals.names[idx])
			if udt, ok := c.localRecordTypes[varName]; ok {
				return udt, true
			}
		}
	case OpGetClassMember:
		if idx >= 0 && idx < len(c.constants) {
			fieldName := strings.ToLower(c.constants[idx].Str)
			if c.currentClassName != "" {
				if field, ok := c.getClassFieldDeclaration(c.currentClassName, fieldName); ok {
					if field.Type == VTRecord {
						return field.ClassType, true
					}
				}
			}
		}
	}
	return "", false
}

type loopContext struct {
	kind      string
	exitJumps []int
}

// jsLoopContext tracks break and continue targets for JScript loops.
type jsLoopContext struct {
	continueTargets     []int // Jump positions that need patching to the loop continuation
	loopStart           int   // Bytecode position of loop start
	forIterDepthAtStart int   // jsForIterDepth at the time this loop was created
}

// jsBreakContext tracks break targets for breakable JScript constructs.
type jsBreakContext struct {
	breakTargets        []int
	forIterDepthAtStart int // jsForIterDepth at the time this break context was created
}

// jsForIterScope records the per-iteration env scope context for a for-let loop.
type jsForIterScope struct {
	nameIdxs []int // constant indices for loop variable names
	fast     bool  // true when using non-alloc fast enter/exit opcodes
}

type jsType int

const (
	jsTypeUnknown jsType = iota
	jsTypeInteger
)

// jsLocalScope holds compiler-time name resolution data for JScript local slots.
// entries maps one identifier to either a local slot index (>=0) or a lexical
// barrier marker (-1) that blocks outer local-slot capture.
type jsLocalScope struct {
	entries    map[string]int
	types      map[string]jsType // inferred types for local variables
	isFunction bool
}

type definitionTokenBound struct {
	start int
	end   int
}

// CompiledRecordDecl stores metadata for one User-Defined Type (UDT) declaration.
type CompiledRecordDecl struct {
	Name    string
	Members []CompiledRecordMemberDecl
}

// CompiledRecordMemberDecl stores one UDT member metadata entry.
type CompiledRecordMemberDecl struct {
	Name    string
	Type    ValueType
	UDTName string // If Type is VTRecord, this specifies which UDT it is
}

// CompiledClassDecl stores bootstrap metadata for one class declaration.
type CompiledClassDecl struct {
	Name       string
	Fields     []CompiledClassFieldDecl
	Methods    []CompiledClassMethodDecl
	Properties []CompiledClassPropertyDecl
	Events     []CompiledClassEventDecl
	Interfaces []string
}

// CompiledClassFieldDecl stores one compiled class field metadata entry.
type CompiledClassFieldDecl struct {
	Name       string
	IsPublic   bool
	WithEvents bool
	Type       ValueType
	ClassType  string
}

// CompiledClassEventDecl stores one compiled class event metadata entry.
type CompiledClassEventDecl struct {
	Name       string
	ParamCount int // number of formal parameters declared for this event
}

// CompiledClassMethodDecl stores one compiled class method metadata entry.
type CompiledClassMethodDecl struct {
	Name           string
	IsFunction     bool
	IsPublic       bool
	UserSubConstID int
}

// CompiledClassPropertyDecl stores one compiled class property metadata entry.
type CompiledClassPropertyDecl struct {
	Name              string
	IsPublic          bool
	GetUserSubConstID int
	GetParamCount     int
	HasGet            bool
	GetPreScanned     bool
	LetUserSubConstID int
	LetParamCount     int
	HasLet            bool
	LetPreScanned     bool
	SetUserSubConstID int
	SetParamCount     int
	HasSet            bool
	SetPreScanned     bool
}

// addClassMethodDeclaration attaches one class method metadata entry to one class declaration.
func (c *Compiler) addClassMethodDeclaration(className string, method CompiledClassMethodDecl) {
	if c == nil {
		return
	}
	if c.hasClassMethodDeclaration(className, method.Name) {
		return
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return
	}
	c.classDecls[classIdx].Methods = append(c.classDecls[classIdx].Methods, method)
}

// hasClassFieldDeclaration reports whether one class field is known in compile metadata.
func (c *Compiler) hasClassFieldDeclaration(className string, fieldName string) bool {
	if c == nil {
		return false
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return false
	}
	trimmedFieldName := strings.TrimSpace(fieldName)
	for i := range c.classDecls[classIdx].Fields {
		if strings.EqualFold(c.classDecls[classIdx].Fields[i].Name, trimmedFieldName) {
			return true
		}
	}
	return false
}

// getClassFieldDeclaration finds one class field metadata entry by field name.
func (c *Compiler) getClassFieldDeclaration(className string, fieldName string) (*CompiledClassFieldDecl, bool) {
	if c == nil {
		return nil, false
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return nil, false
	}
	trimmedFieldName := strings.TrimSpace(fieldName)
	for i := range c.classDecls[classIdx].Fields {
		if strings.EqualFold(c.classDecls[classIdx].Fields[i].Name, trimmedFieldName) {
			return &c.classDecls[classIdx].Fields[i], true
		}
	}
	return nil, false
}

// addClassEventDeclaration attaches one class event metadata entry to one class declaration.
func (c *Compiler) addClassEventDeclaration(className string, event CompiledClassEventDecl) {
	if c == nil {
		return
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return
	}
	c.classDecls[classIdx].Events = append(c.classDecls[classIdx].Events, event)
}

// addClassInterface attaches one interface name to one class declaration.
func (c *Compiler) addClassInterface(className string, interfaceName string) {
	if c == nil {
		return
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return
	}
	c.classDecls[classIdx].Interfaces = append(c.classDecls[classIdx].Interfaces, interfaceName)
}

// hasClassEventDeclaration reports whether one class event is known in compile metadata.
func (c *Compiler) hasClassEventDeclaration(className string, eventName string) bool {
	if c == nil {
		return false
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return false
	}
	lowerEventName := strings.ToLower(strings.TrimSpace(eventName))
	for _, e := range c.classDecls[classIdx].Events {
		if strings.EqualFold(e.Name, lowerEventName) {
			return true
		}
	}
	return false
}

// hasClassPropertyDeclaration reports whether one class property is known in compile metadata.
func (c *Compiler) hasClassPropertyDeclaration(className string, propertyName string) bool {
	if c == nil {
		return false
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return false
	}
	trimmedPropertyName := strings.TrimSpace(propertyName)
	for i := range c.classDecls[classIdx].Properties {
		if strings.EqualFold(c.classDecls[classIdx].Properties[i].Name, trimmedPropertyName) {
			return true
		}
	}
	return false
}

// hasClassMethodDeclaration reports whether one class method is known in compile metadata.
func (c *Compiler) hasClassMethodDeclaration(className string, methodName string) bool {
	if c == nil {
		return false
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return false
	}
	trimmedMethodName := strings.TrimSpace(methodName)
	for i := range c.classDecls[classIdx].Methods {
		if strings.EqualFold(c.classDecls[classIdx].Methods[i].Name, trimmedMethodName) {
			return true
		}
	}
	return false
}

// hasClassZeroArgPropertyGetDeclaration reports whether one class property has a zero-arg getter.
func (c *Compiler) hasClassZeroArgPropertyGetDeclaration(className string, propertyName string) bool {
	property, ok := c.getClassPropertyDeclaration(className, propertyName)
	if !ok || property == nil {
		return false
	}
	return property.HasGet && property.GetParamCount == 0
}

// addClassFieldDeclaration stores one class field metadata entry.
func (c *Compiler) addClassFieldDeclaration(className string, field CompiledClassFieldDecl) {
	if c == nil {
		return
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return
	}
	for i := range c.classDecls[classIdx].Fields {
		if strings.EqualFold(c.classDecls[classIdx].Fields[i].Name, field.Name) {
			c.classDecls[classIdx].Fields[i] = field
			return
		}
	}
	c.classDecls[classIdx].Fields = append(c.classDecls[classIdx].Fields, field)
}

// getClassPropertyDeclaration finds one class property metadata entry by property name.
func (c *Compiler) getClassPropertyDeclaration(className string, propertyName string) (*CompiledClassPropertyDecl, bool) {
	if c == nil {
		return nil, false
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return nil, false
	}
	lowerPropertyName := strings.ToLower(strings.TrimSpace(propertyName))
	for i := range c.classDecls[classIdx].Properties {
		if strings.EqualFold(c.classDecls[classIdx].Properties[i].Name, lowerPropertyName) {
			return &c.classDecls[classIdx].Properties[i], true
		}
	}
	return nil, false
}

// setClassPropertyDeclaration stores one class property metadata entry.
func (c *Compiler) setClassPropertyDeclaration(className string, property CompiledClassPropertyDecl) {
	if c == nil {
		return
	}
	lowerClassName := strings.ToLower(strings.TrimSpace(className))
	classIdx, exists := c.classDeclLookup[lowerClassName]
	if !exists || classIdx < 0 || classIdx >= len(c.classDecls) {
		return
	}
	for i := range c.classDecls[classIdx].Properties {
		if strings.EqualFold(c.classDecls[classIdx].Properties[i].Name, property.Name) {
			c.classDecls[classIdx].Properties[i] = property
			return
		}
	}
	c.classDecls[classIdx].Properties = append(c.classDecls[classIdx].Properties, property)
}

// NewCompiler creates a new Compiler instance for pure VBScript.
func NewCompiler(code string) *Compiler {
	return createCompiler(code, vbscript.ModeVBScript)
}

// NewJSModuleCompiler creates a new Compiler instance for pure JScript modules.
func NewJSModuleCompiler(code string) *Compiler {
	c := createCompiler(code, vbscript.ModeVBScript)
	c.isJSModule = true
	c.sourceCode = code
	return c
}

// NewJavaScriptCompiler creates a new Compiler instance for pure JScript.
func NewJavaScriptCompiler(code string) *Compiler {
	c := createCompiler(code, vbscript.ModeVBScript)
	c.isJSModule = true
	c.sourceCode = code
	return c
}

// NewASPCompiler creates a new Compiler instance for ASP files (starting in HTML mode).
func NewASPCompiler(code string) *Compiler {
	return createCompiler(code, vbscript.ModeASP)
}

// NewEvalCompiler creates a new Compiler instance for VBScript expressions.
func NewEvalCompiler(code string) *Compiler {
	c := createCompiler(code, vbscript.ModeEval)
	c.isEval = true
	return c
}

// createCompiler initializes the compiler with the specified mode.
func createCompiler(code string, mode vbscript.LexerMode) *Compiler {
	lexer := vbscript.NewLexer(code)
	lexer.Mode = mode
	if mode == vbscript.ModeASP {
		lexer.InASPBlock = false
	} else {
		lexer.InASPBlock = true
	}

	c := &Compiler{
		lexer:                  lexer,
		lexerMode:              mode,
		sourceCode:             code,
		tokenIndex:             -1,
		bytecode:               make([]byte, 0),
		constants:              make([]Value, 0, 1024),
		constantMap:            make(map[string]int, 1024),
		Globals:                NewSymbolTable(),
		locals:                 NewSymbolTable(),
		declaredGlobals:        make(map[string]bool),
		implicitGlobals:        make(map[string]bool),
		declaredLocals:         make(map[string]bool),
		globalVarTypes:         make(map[string]ValueType),
		localVarTypes:          make(map[string]ValueType),
		globalRecordTypes:      make(map[string]string),
		localRecordTypes:       make(map[string]string),
		funcLocalTypes:         make(map[int]map[string]ValueType),
		funcLocalRecordTypes:   make(map[int]map[string]string),
		constGlobals:           make(map[string]bool),
		constLocals:            make(map[string]bool),
		staticLocals:           make(map[string]int),
		constLiteralGlobals:    make(map[string]Value),
		isLocal:                false,
		optionExplicit:         false,
		optionCompare:          0, // Binary
		includeCaseInsensitive: true,
		forwardCallPatches:     make(map[string][]int),
		forwardConstPatches:    make(map[string][]int),
		lastCallTargetPos:      -1,
		lastCoercePos:          -1,
		lastDebugLine:          -1,
		lastDebugColumn:        -1,
		tempCounter:            0,
		globalZeroArgFuncs:     make(map[string]bool),
		globalZeroArgSubs:      make(map[string]bool),
		classDecls:             make([]CompiledClassDecl, 0),
		classDeclLookup:        make(map[string]int),
		recordDecls:            make([]CompiledRecordDecl, 0),
		recordDeclLookup:       make(map[string]int),
		enumTypeLookup:         make(map[string]map[string]Value),
		loopContexts:           make([]loopContext, 0),
		jsStrictMode:           false,
		jsFunctionStrictModes:  make(map[int]bool),
		jsBlockScopeStack:      make([]map[string]bool, 0, 32),
		jsLocalScopeStack:      make([]jsLocalScope, 0, 16),
		jsLocalEnabled:         false,
		jsLocalSlotCount:       0,
		activeVBSConstants:     make([]VBSConstant, 0, len(VBSConstants)),
		labelMap:               make(map[string]int),
		forwardLabelPatches:    make(map[string][]int),
		funcParamDefaults:      make(map[int][]int),
	}
	c.activeVBSConstants = append(c.activeVBSConstants, VBSConstants...)

	// Pre-inject ASP Intrinsic Objects as "declared"
	intrinsics := []string{"Response", "Request", "Server", "Session", "Application", "ObjectContext", "Err", "console"}
	for _, name := range intrinsics {
		c.Globals.Add(name)
		c.declaredGlobals[strings.ToLower(name)] = true
	}

	// Pre-declare ObjectContext transaction event handler sub names at fixed global indices 8 and 9.
	// If the script defines Sub OnTransactionCommit / Sub OnTransactionAbort, the compiler will
	// assign the UserSub value to these known slots. The VM reads them after script execution.
	eventHandlers := []string{"OnTransactionCommit", "OnTransactionAbort"}
	for _, name := range eventHandlers {
		c.Globals.Add(name)
		c.declaredGlobals[strings.ToLower(name)] = true
	}

	// Pre-inject Built-in Functions
	// This ensures they have fixed indices that the VM can also pre-populate.
	for _, name := range BuiltinNames {
		c.Globals.Add(name)
		c.declaredGlobals[strings.ToLower(name)] = true
	}

	// Pre-inject VBScript predefined constants (e.g. vbCrLf, vbLongDate, vbTrue) and
	// type library constants (e.g. adInteger, adOpenKeyset from ADODB 2.5).
	// They are added after builtins so their global-slot indices stay stable across
	// every compilation in the same process (same order as VBSConstants slice).
	for _, kv := range c.activeVBSConstants {
		c.Globals.Add(kv.Name)
		c.declaredGlobals[strings.ToLower(kv.Name)] = true
		c.constGlobals[strings.ToLower(kv.Name)] = true
		c.constLiteralGlobals[strings.ToLower(kv.Name)] = kv.Val
	}

	c.move()
	// Record where user-declared globals begin; slots below this are read-only pre-injected.
	c.userGlobalsStart = c.Globals.Count()
	return c
}

func normalizeIncludeSiteRoot(rootDir string) string {
	trimmed := strings.TrimSpace(rootDir)
	if trimmed == "" {
		return ""
	}
	absRoot, err := filepath.Abs(trimmed)
	if err != nil {
		return filepath.Clean(trimmed)
	}
	return filepath.Clean(absRoot)
}

func pathInsideRoot(root string, target string) bool {
	absRoot, err := filepath.Abs(root)
	if err != nil {
		return false
	}
	absTarget, err := filepath.Abs(target)
	if err != nil {
		return false
	}
	rel, err := filepath.Rel(absRoot, absTarget)
	if err != nil {
		return false
	}
	if rel == "." {
		return true
	}
	if rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return false
	}
	return true
}

func resolvePathCaseInsensitive(path string) (string, error) {
	cleaned := filepath.Clean(path)
	if cleaned == "" {
		return "", fmt.Errorf("empty include path")
	}

	vol := filepath.VolumeName(cleaned)
	remainder := strings.TrimPrefix(cleaned, vol)
	startsWithSep := strings.HasPrefix(remainder, string(filepath.Separator))
	parts := strings.Split(strings.TrimPrefix(remainder, string(filepath.Separator)), string(filepath.Separator))

	current := vol
	if startsWithSep {
		current += string(filepath.Separator)
	}

	if len(parts) == 1 && parts[0] == "" {
		if _, err := os.Stat(current); err != nil {
			return "", err
		}
		return current, nil
	}

	for _, part := range parts {
		if part == "" || part == "." {
			continue
		}
		entries, err := os.ReadDir(current)
		if err != nil {
			return "", err
		}
		matched := ""
		for _, entry := range entries {
			if strings.EqualFold(entry.Name(), part) {
				matched = entry.Name()
				break
			}
		}
		if matched == "" {
			return "", os.ErrNotExist
		}
		current = filepath.Join(current, matched)
	}

	if _, err := os.Stat(current); err != nil {
		return "", err
	}
	return current, nil
}

func resolveExistingPath(candidate string, caseInsensitive bool) (string, error) {
	if _, err := os.Stat(candidate); err == nil {
		return filepath.Clean(candidate), nil
	}
	if !caseInsensitive || runtime.GOOS == "windows" {
		return "", os.ErrNotExist
	}
	matched, err := resolvePathCaseInsensitive(candidate)
	if err != nil {
		return "", os.ErrNotExist
	}
	return filepath.Clean(matched), nil
}

// injectTypeLibConstants registers extra read-only typelibrary constants for the current compilation.
// It reserves global slots before user code is parsed so ByRef and const-write protection remain correct.
func (c *Compiler) injectTypeLibConstants(constants []VBSConstant) []VBSConstant {
	if c == nil || len(constants) == 0 {
		return nil
	}

	added := make([]VBSConstant, 0, len(constants))
	for _, kv := range constants {
		lower := strings.ToLower(strings.TrimSpace(kv.Name))
		if lower == "" {
			continue
		}
		if _, exists := c.constGlobals[lower]; exists {
			continue
		}
		c.Globals.Add(kv.Name)
		c.declaredGlobals[lower] = true
		c.constGlobals[lower] = true
		c.constLiteralGlobals[lower] = kv.Val
		c.activeVBSConstants = append(c.activeVBSConstants, kv)
		added = append(added, kv)
	}

	if len(added) > 0 {
		c.userGlobalsStart = c.Globals.Count()
	}
	return added
}

// emitInjectedConstantInitializers writes one bytecode preamble that initializes typelibrary
// constants in their reserved global slots, keeping compiler/VM slot mapping deterministic.
func (c *Compiler) emitInjectedConstantInitializers(constants []VBSConstant) {
	if c == nil || len(constants) == 0 {
		return
	}

	for _, kv := range constants {
		gidx, exists := c.Globals.Get(kv.Name)
		if !exists {
			continue
		}
		cidx := c.addConstant(kv.Val)
		c.emit(OpConstant, cidx)
		c.emit(OpSetGlobal, gidx)
	}
}

// stripEmptyASPBlocks removes empty ASP blocks that should not reach compilation.
// Optimization: Uses manual scanning instead of regexp to avoid heap allocations.
func stripEmptyASPBlocks(source string) string {
	if source == "" {
		return source
	}

	var builder strings.Builder
	builder.Grow(len(source))

	cursor := 0
	for {
		start := strings.Index(source[cursor:], "<%")
		if start == -1 {
			builder.WriteString(source[cursor:])
			break
		}

		absStart := cursor + start
		builder.WriteString(source[cursor:absStart])

		end := strings.Index(source[absStart:], "%>")
		if end == -1 {
			builder.WriteString(source[absStart:])
			break
		}

		absEnd := absStart + end
		content := source[absStart+2 : absEnd]

		// Check if block is effectively empty (only whitespace and optional '=')
		isEmpty := true
		for i := 0; i < len(content); i++ {
			c := content[i]
			if c == ' ' || c == '\t' || c == '\n' || c == '\r' || c == '=' {
				continue
			}
			isEmpty = false
			break
		}

		if !isEmpty {
			builder.WriteString(source[absStart : absEnd+2])
		}

		cursor = absEnd + 2
	}

	return builder.String()
}

// stripUTF8BOM removes UTF-8 BOM (EF BB BF) from byte slice if present.
// BOM should not appear in the middle of included files and will cause output corruption.
func stripUTF8BOM(data []byte) []byte {
	if len(data) >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF {
		return data[3:]
	}
	return data
}

// preprocessASPIncludesWithDeps expands includes and optionally records resolved dependency files.
// Optimization: Uses manual scanning instead of regexp to avoid heap allocations.
func preprocessASPIncludesWithDeps(source string, sourceName string, visited map[string]bool, depth int, dependencies *[]string) (string, SourceMap, error) {
	return preprocessASPIncludesWithDepsWithOptions(source, sourceName, visited, depth, dependencies, defaultIncludeResolveOptions())
}

func preprocessASPIncludesWithDepsWithOptions(source string, sourceName string, visited map[string]bool, depth int, dependencies *[]string, options includeResolveOptions) (string, SourceMap, error) {
	if depth > 32 {
		return "", SourceMap{}, fmt.Errorf("include recursion limit exceeded")
	}

	var builder strings.Builder
	builder.Grow(len(source))
	sourceMap := SourceMap{}
	currentMergedLine := 1
	currentSourceLine := 1
	sourceMap.AddBoundary(currentMergedLine, sourceName, currentSourceLine)

	cursor := 0
	for cursor < len(source) {
		scriptStart := strings.Index(source[cursor:], "<%")
		scanLimit := len(source)
		if scriptStart != -1 {
			scanLimit = cursor + scriptStart
		}

		// Manual scan for <!-- #include ... --> in raw HTML only.
		startIdx := -1
		endIdx := -1
		kind := ""
		pathVal := ""
		directive := ""

		searchCursor := cursor
		for {
			if searchCursor >= scanLimit {
				break
			}
			idx := strings.Index(source[searchCursor:scanLimit], "<!--")
			if idx == -1 {
				break
			}
			absStart := searchCursor + idx
			if absStart >= scanLimit {
				break
			}

			// Find closing tag
			idxClose := strings.Index(source[absStart:scanLimit], "-->")
			if idxClose == -1 {
				break
			}
			absEnd := absStart + idxClose + 3

			comment := source[absStart:absEnd]
			upperComment := strings.ToUpper(comment)

			if strings.Contains(upperComment, "#INCLUDE") {
				// Parse directive
				kindIdx := -1
				if strings.Contains(upperComment, "VIRTUAL") {
					kind = "virtual"
					kindIdx = strings.Index(upperComment, "VIRTUAL")
				} else if strings.Contains(upperComment, "FILE") {
					kind = "file"
					kindIdx = strings.Index(upperComment, "FILE")
				}

				if kindIdx != -1 {
					valPart := comment[kindIdx+len(kind):]
					eqIdx := strings.Index(valPart, "=")
					if eqIdx != -1 {
						valPart = valPart[eqIdx+1:]
						valPart = strings.TrimSpace(valPart)
						if len(valPart) > 0 {
							quote := valPart[0]
							if quote == '"' || quote == '\'' {
								valPart = valPart[1:]
								before, _, ok := strings.Cut(valPart, string(quote))
								if ok {
									pathVal = before
									startIdx = absStart
									endIdx = absEnd
									directive = comment
									break
								}
							}
						}
					}
				}
				searchCursor = absStart + 4
				continue
			}

			searchCursor = absStart + 4
		}

		if startIdx == -1 {
			if scanLimit > cursor {
				appendMappedSegment(&builder, &sourceMap, source[cursor:scanLimit], sourceName, &currentMergedLine, &currentSourceLine)
				cursor = scanLimit
			}
			if scriptStart == -1 {
				break
			}
			scriptEndRel := strings.Index(source[cursor:], "%>")
			if scriptEndRel == -1 {
				appendMappedSegment(&builder, &sourceMap, source[cursor:], sourceName, &currentMergedLine, &currentSourceLine)
				break
			}
			scriptEnd := cursor + scriptEndRel + 2
			appendMappedSegment(&builder, &sourceMap, source[cursor:scriptEnd], sourceName, &currentMergedLine, &currentSourceLine)
			cursor = scriptEnd
			continue
		}

		replaceStart := startIdx
		replaceEnd := endIdx
		if lineOnlyIncludeDirective(source, startIdx, endIdx) {
			replaceStart = lineStartIndex(source, startIdx)
			replaceEnd = lineEndIndexIncludingNewline(source, endIdx)
		}

		appendMappedSegment(&builder, &sourceMap, source[cursor:replaceStart], sourceName, &currentMergedLine, &currentSourceLine)

		resolvedPath, err := resolveIncludePathWithOptions(sourceName, pathVal, kind == "virtual", options)
		if err != nil {
			return "", SourceMap{}, fmt.Errorf("include resolve failed for %s: %w", directive, err)
		}

		norm := strings.ToLower(filepath.Clean(resolvedPath))
		if visited[norm] {
			return "", SourceMap{}, fmt.Errorf("circular include detected: %s", resolvedPath)
		}

		contentBytes, err := os.ReadFile(resolvedPath)
		if err != nil {
			return "", SourceMap{}, fmt.Errorf("could not read include %s: %w", resolvedPath, err)
		}

		// Strip UTF-8 BOM if present to prevent output corruption
		contentBytes = stripUTF8BOM(contentBytes)

		visited[norm] = true
		if dependencies != nil {
			*dependencies = append(*dependencies, resolvedPath)
		}
		expanded, childMap, err := preprocessASPIncludesWithDepsWithOptions(string(contentBytes), resolvedPath, visited, depth+1, dependencies, options)
		delete(visited, norm)
		if err != nil {
			return "", SourceMap{}, err
		}

		mergeSourceMap(&sourceMap, childMap, currentMergedLine)
		builder.WriteString(expanded)
		currentMergedLine += countLogicalLines(expanded)
		currentSourceLine += countLogicalLines(source[replaceStart:replaceEnd])
		cursor = replaceEnd
	}

	return builder.String(), sourceMap, nil
}

// lineOnlyIncludeDirective reports whether a matched include directive occupies
// an otherwise-empty logical line (whitespace only around it).
func lineOnlyIncludeDirective(s string, start int, end int) bool {
	ls := lineStartIndex(s, start)
	le := lineEndIndex(s, end)
	if !isAllHorizontalWhitespace(s[ls:start]) {
		return false
	}
	if !isAllHorizontalWhitespace(s[end:le]) {
		return false
	}
	return true
}

// lineStartIndex returns the index of the first byte of the logical line that contains pos.
func lineStartIndex(s string, pos int) int {
	if pos <= 0 {
		return 0
	}
	for i := pos - 1; i >= 0; i-- {
		if s[i] == '\n' || s[i] == '\r' {
			return i + 1
		}
	}
	return 0
}

// lineEndIndex returns the index of '\n' or end-of-string for the line that contains pos.
func lineEndIndex(s string, pos int) int {
	if pos >= len(s) {
		return len(s)
	}
	for i := pos; i < len(s); i++ {
		if s[i] == '\n' || s[i] == '\r' {
			return i
		}
	}
	return len(s)
}

// lineEndIndexIncludingNewline returns the line end index and consumes one trailing newline when present.
func lineEndIndexIncludingNewline(s string, pos int) int {
	le := lineEndIndex(s, pos)
	if le >= len(s) {
		return le
	}
	if s[le] == '\r' {
		if le+1 < len(s) && s[le+1] == '\n' {
			return le + 2
		}
		return le + 1
	}
	if s[le] == '\n' {
		return le + 1
	}
	return le
}

// isAllHorizontalWhitespace reports whether the input contains only spaces, tabs, or '\r'.
func isAllHorizontalWhitespace(s string) bool {
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case ' ', '\t', '\r':
			continue
		default:
			return false
		}
	}
	return true
}

// buildIdentitySourceMap maps one source to itself starting at line 1.
func buildIdentitySourceMap(sourceName string) SourceMap {
	m := SourceMap{}
	m.AddBoundary(1, sourceName, 1)
	return m
}

// countLineBreaks counts line separators across CRLF, LF, and CR sequences.
func countLineBreaks(s string) int {
	if s == "" {
		return 0
	}
	count := 0
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '\r':
			count++
			if i+1 < len(s) && s[i+1] == '\n' {
				i++
			}
		case '\n':
			count++
		}
	}
	return count
}

// countLogicalLines returns legacy logical-line count (minimum 1 for non-empty segments).
func countLogicalLines(s string) int {
	if s == "" {
		return 0
	}
	breaks := countLineBreaks(s)
	if strings.HasSuffix(s, "\n") || strings.HasSuffix(s, "\r") {
		return breaks
	}
	return breaks + 1
}

// appendMappedSegment appends one source segment and updates merged/source line cursors.
func appendMappedSegment(builder *strings.Builder, sourceMap *SourceMap, segment string, sourceName string, currentMergedLine *int, currentSourceLine *int) {
	if segment == "" {
		return
	}
	sourceMap.AddBoundary(*currentMergedLine, sourceName, *currentSourceLine)
	builder.WriteString(segment)
	logicalLines := countLogicalLines(segment)
	*currentMergedLine += logicalLines
	*currentSourceLine += logicalLines
}

// mergeSourceMap appends one child source map into a parent map at one merged-line offset.
func mergeSourceMap(target *SourceMap, child SourceMap, mergedLineStart int) {
	if target == nil || mergedLineStart <= 0 || len(child.entries) == 0 {
		return
	}
	for i := 0; i < len(child.entries); i++ {
		entry := child.entries[i]
		target.AddBoundary(mergedLineStart+entry.MergedLineStart-1, entry.OriginalFile, entry.OriginalLine)
	}
}

// resolveIncludePath resolves one ASP include path against current source path context.
func resolveIncludePath(sourceName string, includePath string, isVirtual bool) (string, error) {
	return resolveIncludePathWithOptions(sourceName, includePath, isVirtual, defaultIncludeResolveOptions())
}

func resolveIncludePathWithOptions(sourceName string, includePath string, isVirtual bool, options includeResolveOptions) (string, error) {
	trimmed := strings.TrimSpace(includePath)
	if trimmed == "" {
		return "", fmt.Errorf("empty include path")
	}

	trimmed = strings.ReplaceAll(trimmed, "/", string(filepath.Separator))
	trimmed = strings.ReplaceAll(trimmed, "\\", string(filepath.Separator))

	sourceAbs := filepath.Clean(strings.TrimSpace(sourceName))
	if !filepath.IsAbs(sourceAbs) {
		absSource, err := filepath.Abs(sourceAbs)
		if err != nil {
			return "", err
		}
		sourceAbs = filepath.Clean(absSource)
	}

	sourceDir := filepath.Dir(sourceAbs)

	if isVirtual || strings.HasPrefix(includePath, "/") || strings.HasPrefix(includePath, "\\") {
		// For virtual includes, a leading "/" means "relative to site root",
		// not an absolute filesystem path — skip the absolute-path check.
		siteRoot := normalizeIncludeSiteRoot(options.siteRoot)
		if siteRoot == "" {
			return "", fmt.Errorf("virtual include root is not configured: %s", includePath)
		}
		if filepath.IsAbs(trimmed) {
			trimmed = strings.TrimLeft(trimmed, string(filepath.Separator))
		}
		rel := strings.TrimLeft(trimmed, string(filepath.Separator))
		candidate := filepath.Clean(filepath.Join(siteRoot, rel))
		if !pathInsideRoot(siteRoot, candidate) {
			return "", fmt.Errorf("virtual include escapes site root: %s", includePath)
		}
		resolved, resolveErr := resolveExistingPath(candidate, options.caseInsensitive)
		if resolveErr != nil {
			return "", fmt.Errorf("include not found: %s", includePath)
		}
		return resolved, nil
	}

	// Non-virtual (file) includes must not reference absolute filesystem paths.
	//Attention: This check is done after the virtual-include check to allow virtual includes to use leading slashes on Linux/Unix systems. Absolute paths are still disallowed for non-virtual includes.
	if filepath.IsAbs(trimmed) {
		return "", fmt.Errorf("absolute include file path is not allowed: %s", includePath)
	}
	candidate := filepath.Clean(filepath.Join(sourceDir, trimmed))
	resolved, resolveErr := resolveExistingPath(candidate, options.caseInsensitive)
	if resolveErr != nil {
		return "", fmt.Errorf("include not found: %s", includePath)
	}
	return resolved, nil
}

// clearLastCallTarget resets transient call-target tracking for identifier-based call patching.
func (c *Compiler) clearLastCallTarget() {
	c.lastCallTargetName = ""
	c.lastCallTargetPos = -1
	c.lastCallIsGlobal = false
}

// markLastCallTarget stores the most recent emitted identifier load for a potential immediate call.
func (c *Compiler) markLastCallTarget(name string, op OpCode, emitPos int) {
	c.lastCallTargetName = strings.ToLower(strings.TrimSpace(name))
	c.lastCallTargetPos = emitPos
	c.lastCallIsGlobal = op == OpGetGlobal
}

// registerForwardCallPatch records a call-site load that may target a function declared later in source.
func (c *Compiler) registerForwardCallPatch(name string, emitPos int) {
	trimmed := strings.ToLower(strings.TrimSpace(name))
	if trimmed == "" || emitPos < 0 {
		return
	}
	c.forwardCallPatches[trimmed] = append(c.forwardCallPatches[trimmed], emitPos)
}

// registerForwardConstPatch records one identifier load that may become a constant
// once a later Const declaration is parsed.
func (c *Compiler) registerForwardConstPatch(name string, emitPos int) {
	trimmed := strings.ToLower(strings.TrimSpace(name))
	if trimmed == "" || emitPos < 0 {
		return
	}
	c.forwardConstPatches[trimmed] = append(c.forwardConstPatches[trimmed], emitPos)
}

// patchForwardCallSites rewrites pending OpGetGlobal call-target loads to OpConstant user-sub references.
func (c *Compiler) patchForwardCallSites(name string, userSubConstIdx int) {
	trimmed := strings.ToLower(strings.TrimSpace(name))
	patches, ok := c.forwardCallPatches[trimmed]
	if !ok || len(patches) == 0 {
		return
	}

	for _, pos := range patches {
		if pos < 0 || pos+2 >= len(c.bytecode) {
			continue
		}
		if OpCode(c.bytecode[pos]) != OpGetGlobal {
			continue
		}
		c.bytecode[pos] = byte(OpConstant)
		binary.BigEndian.PutUint16(c.bytecode[pos+1:pos+3], uint16(userSubConstIdx))
	}

	delete(c.forwardCallPatches, trimmed)
}

// patchForwardConstSites rewrites pending OpGetGlobal identifier loads to OpConstant
// using one resolved constant value.
func (c *Compiler) patchForwardConstSites(name string, value Value) {
	trimmed := strings.ToLower(strings.TrimSpace(name))
	patches, ok := c.forwardConstPatches[trimmed]
	if !ok || len(patches) == 0 {
		return
	}

	constIdx := c.addConstant(value)
	for _, pos := range patches {
		if pos < 0 || pos+2 >= len(c.bytecode) {
			continue
		}
		if OpCode(c.bytecode[pos]) != OpGetGlobal {
			continue
		}
		c.bytecode[pos] = byte(OpConstant)
		binary.BigEndian.PutUint16(c.bytecode[pos+1:pos+3], uint16(constIdx))
	}

	delete(c.forwardConstPatches, trimmed)
}

// declareVar marks a variable as declared (for Option Explicit checks).
func (c *Compiler) declareVar(name string) {
	lower := strings.ToLower(name)
	if c.isLocal {
		c.locals.Add(name)
		c.declaredLocals[lower] = true
	} else {
		c.Globals.Add(name)
		c.declaredGlobals[lower] = true
		delete(c.implicitGlobals, lower)
	}
}

// declareConst marks a symbol as constant in the current scope.
func (c *Compiler) declareConst(name string) {
	lower := strings.ToLower(name)
	c.declareVar(name)
	if c.isLocal {
		c.constLocals[lower] = true
		return
	}
	c.constGlobals[lower] = true
}

// resolveVar looks up a variable and performs Option Explicit validation.
func (c *Compiler) resolveVar(name string) (OpCode, int) {
	lower := strings.ToLower(name)

	// 1. Check Locals
	if c.isLocal {
		if globalIdx, isStatic := c.staticLocals[lower]; isStatic {
			return OpGetGlobal, globalIdx
		}
		if idx, exists := c.locals.Get(name); exists {
			return OpGetLocal, idx
		}
	}
	if c.currentClassName != "" && (c.isLocal || c.dynamicMemberResolution) {
		if c.hasClassFieldDeclaration(c.currentClassName, name) ||
			c.hasClassMethodDeclaration(c.currentClassName, name) ||
			c.hasClassZeroArgPropertyGetDeclaration(c.currentClassName, name) {
			memberIdx := c.addConstant(NewString(name))
			return OpGetClassMember, memberIdx
		}
	}

	// Compile-time known global constants should resolve as immediate constants,
	// independent of source declaration order (matching classic VBScript semantics).
	if c.constGlobals[lower] {
		if value, ok := c.constLiteralGlobals[lower]; ok {
			idx := c.addConstant(value)
			return OpConstant, idx
		}
	}

	// 2. Check Globals
	if idx, exists := c.Globals.Get(name); exists {
		isTrueGlobal := idx < c.userGlobalsStart || c.declaredGlobals[lower] || c.constGlobals[lower] || c.implicitGlobals[lower]
		if c.isLocal && !isTrueGlobal {
			if c.optionExplicit {
				panic(c.vbCompileError(vbscript.VariableNotDefined, fmt.Sprintf("Variable not defined: '%s'", name)))
			}
			lidx := c.locals.Add(name)
			return OpGetLocal, lidx
		}
		// Even if not explicitly Dim'ed, if it was used once it exists in SymbolTable.
		// We only error if Option Explicit is on and it wasn't Dim'ed.
		// NOTE: Slots below userGlobalsStart are read-only intrinsics/constants and don't need Dim.
		// Implicit globals (created by ExecuteGlobal at runtime) are also allowed.
		if c.optionExplicit && idx >= c.userGlobalsStart && !c.declaredGlobals[lower] && !c.implicitGlobals[lower] {
			panic(c.vbCompileError(vbscript.VariableNotDefined, fmt.Sprintf("Variable not defined: '%s'", name)))
		}
		return OpGetGlobal, idx
	}

	// 3. Implicit Declaration check
	if c.optionExplicit {
		// Allow implicit global reads for variables that may be created
		// at runtime via ExecuteGlobal. Classic ASP compatibility: the
		// ExecuteGlobal statement can inject global variables/functions
		// that the caller's compiler cannot see at compile time.
		idx := c.Globals.Add(name)
		c.implicitGlobals[lower] = true
		return OpGetGlobal, idx
	}

	if c.isLocal {
		if c.currentClassName != "" {
			idx := c.Globals.Add(name)
			c.implicitGlobals[lower] = true
			return OpGetGlobal, idx
		}
		idx := c.locals.Add(name)
		return OpGetLocal, idx
	}

	idx := c.Globals.Add(name)
	c.implicitGlobals[lower] = true
	return OpGetGlobal, idx
}

// constGlobalValueByIndex resolves one global slot index to a known constant literal value.
func (c *Compiler) constGlobalValueByIndex(idx int) (Value, bool) {
	if c == nil || c.Globals == nil || idx < 0 || idx >= len(c.Globals.names) {
		return Value{}, false
	}
	name := strings.ToLower(strings.TrimSpace(c.Globals.names[idx]))
	if name == "" {
		return Value{}, false
	}
	v, ok := c.constLiteralGlobals[name]
	return v, ok
}

// tryCaptureGlobalConstLiteral stores one compile-time global constant value when the
// parsed Const expression is reducible to a literal or known constant reference.
func (c *Compiler) tryCaptureGlobalConstLiteral(name string, exprStart int, exprEnd int) {
	if c == nil || c.isLocal || c.currentClassName != "" {
		return
	}
	if exprStart < 0 || exprEnd <= exprStart || exprEnd > len(c.bytecode) {
		return
	}

	code := c.bytecode[exprStart:exprEnd]
	if len(code) < 3 {
		return
	}

	resolveSimple := func(op OpCode, idx int) (Value, bool) {
		switch op {
		case OpConstant:
			if idx >= 0 && idx < len(c.constants) {
				return c.constants[idx], true
			}
		case OpGetGlobal:
			return c.constGlobalValueByIndex(idx)
		}
		return Value{}, false
	}

	lower := strings.ToLower(strings.TrimSpace(name))
	if lower == "" {
		return
	}

	if len(code) == 3 {
		op := OpCode(code[0])
		idx := int(binary.BigEndian.Uint16(code[1:3]))
		if val, ok := resolveSimple(op, idx); ok {
			c.constLiteralGlobals[lower] = val
		}
		return
	}

	if len(code) == 4 && OpCode(code[3]) == OpCoerceToValue {
		op := OpCode(code[0])
		idx := int(binary.BigEndian.Uint16(code[1:3]))
		if val, ok := resolveSimple(op, idx); ok {
			c.constLiteralGlobals[lower] = val
		}
	}
}

// resolveSetVar performs lookup for assignment operations.
func (c *Compiler) resolveSetVar(name string) (OpCode, int) {
	lower := strings.ToLower(name)

	if c.isLocal {
		if globalIdx, isStatic := c.staticLocals[lower]; isStatic {
			return OpSetGlobal, globalIdx
		}
		if idx, exists := c.locals.Get(name); exists {
			if c.constLocals[lower] {
				panic(c.vbCompileError(vbscript.IllegalAssignment, fmt.Sprintf("illegal assignment: '%s'", name)))
			}
			return OpSetLocal, idx
		}
	}
	if c.currentClassName != "" && (c.isLocal || c.dynamicMemberResolution) {
		if c.hasClassFieldDeclaration(c.currentClassName, name) || c.hasClassPropertyDeclaration(c.currentClassName, name) {
			memberIdx := c.addConstant(NewString(name))
			return OpSetClassMember, memberIdx
		}
	}

	if c.constGlobals[lower] {
		panic(c.vbCompileError(vbscript.IllegalAssignment, fmt.Sprintf("illegal assignment: '%s'", name)))
	}

	if c.isLocal {
		if gidx, exists := c.Globals.Get(name); exists {
			isTrueGlobal := gidx < c.userGlobalsStart || c.declaredGlobals[lower] || c.constGlobals[lower] || c.implicitGlobals[lower]
			if isTrueGlobal {
				return OpSetGlobal, gidx
			}
		}
		if c.optionExplicit {
			panic(c.vbCompileError(vbscript.VariableNotDefined, fmt.Sprintf("Variable not defined: '%s'", name)))
		}
		idx := c.locals.Add(name)
		return OpSetLocal, idx
	}

	if c.optionExplicit && !c.declaredGlobals[lower] && !c.implicitGlobals[lower] {
		// Allow assignment to implicit globals that may have been created
		// at runtime via ExecuteGlobal. Classic ASP compatibility.
		panic(c.vbCompileError(vbscript.VariableNotDefined, fmt.Sprintf("Variable not defined: '%s'", name)))
	}

	idx := c.Globals.Add(name)
	c.implicitGlobals[lower] = true
	return OpSetGlobal, idx
}

// resolveEraseVar resolves one declared variable slot for an Erase statement.
func (c *Compiler) resolveEraseVar(name string) (OpCode, int) {
	if c.isLocal {
		if idx, exists := c.locals.Get(name); exists {
			return OpEraseLocal, idx
		}
	}
	if c.currentClassName != "" && (c.isLocal || c.dynamicMemberResolution) {
		if c.hasClassFieldDeclaration(c.currentClassName, name) {
			memberIdx := c.addConstant(NewString(name))
			return OpEraseClassMember, memberIdx
		}
	}
	if idx, exists := c.Globals.Get(name); exists {
		return OpEraseGlobal, idx
	}
	panic(c.vbCompileError(vbscript.VariableNotDefined, fmt.Sprintf("Variable not defined: '%s'", name)))
}

// resolveConstInitVar resolves the target slot used only for Const declaration initialization.
func (c *Compiler) resolveConstInitVar(name string) (OpCode, int) {
	if c.isLocal {
		if idx, exists := c.locals.Get(name); exists {
			return OpSetLocal, idx
		}
	}
	if c.currentClassName != "" && (c.isLocal || c.dynamicMemberResolution) && c.hasClassFieldDeclaration(c.currentClassName, name) {
		memberIdx := c.addConstant(NewString(name))
		return OpSetClassMember, memberIdx
	}

	idx := c.Globals.Add(name)
	return OpSetGlobal, idx
}

var vbFastUnaryMathOpcodes = map[string]OpCode{
	"sin":   OpMathSin,
	"cos":   OpMathCos,
	"tan":   OpMathTan,
	"atn":   OpMathAtn,
	"sqr":   OpMathSqr,
	"abs":   OpMathAbs,
	"exp":   OpMathExp,
	"log":   OpMathLog,
	"round": OpMathRound,
	"int":   OpMathInt,
}

// isBuiltinGlobalSlot reports whether one global slot still maps to one builtin name.
func (c *Compiler) isBuiltinGlobalSlot(slot int, builtinLower string) bool {
	if c == nil || c.Globals == nil || slot < 0 || slot >= len(c.Globals.names) {
		return false
	}
	if slot >= c.userGlobalsStart {
		return false
	}
	globalLower := strings.ToLower(strings.TrimSpace(c.Globals.names[slot]))
	if globalLower == "" || globalLower != builtinLower {
		return false
	}
	_, exists := BuiltinIndex[builtinLower]
	return exists
}

// tryEmitFastUnaryMathCall rewrites one builtin call target + OpCall into one direct math opcode.
func (c *Compiler) tryEmitFastUnaryMathCall(callTargetName string, callTargetPos int, argExprStart int, argCount int, callTargetIsGlobal bool) bool {
	if c == nil || !callTargetIsGlobal || argCount != 1 {
		return false
	}
	builtinLower := strings.ToLower(strings.TrimSpace(callTargetName))
	op, exists := vbFastUnaryMathOpcodes[builtinLower]
	if !exists {
		return false
	}
	if callTargetPos < 0 || callTargetPos+3 > len(c.bytecode) || argExprStart < 0 || argExprStart > len(c.bytecode) {
		return false
	}
	if OpCode(c.bytecode[callTargetPos]) != OpGetGlobal {
		return false
	}
	if callTargetPos+3 != argExprStart {
		return false
	}

	globalSlot := int(binary.BigEndian.Uint16(c.bytecode[callTargetPos+1 : callTargetPos+3]))
	if !c.isBuiltinGlobalSlot(globalSlot, builtinLower) {
		return false
	}

	copy(c.bytecode[callTargetPos:], c.bytecode[argExprStart:])
	c.bytecode = c.bytecode[:len(c.bytecode)-(argExprStart-callTargetPos)]
	c.emit(op)
	c.clearLastCallTarget()
	return true
}

func (c *Compiler) emitLine(line int, column int) {
	if line < 0 {
		line = 0
	}
	if column < 0 {
		column = 0
	}
	c.emit(OpLine, line, column)
}

// emitCurrentDebugLocation emits OpLine with current token location for runtime error mapping.
func (c *Compiler) emitCurrentDebugLocation() {
	if c == nil || c.next == nil {
		return
	}

	line := c.next.GetLineNumber()
	column := max(c.next.GetStart()-c.next.GetLineStart(), 0)
	column++

	if line == c.lastDebugLine && column == c.lastDebugColumn {
		return
	}

	c.emitLine(line, column)
	c.lastDebugLine = line
	c.lastDebugColumn = column
}

func (c *Compiler) move() vbscript.Token {
	token := c.next
	c.prevToken = c.lastToken
	c.lastToken = token
	c.next = c.lexer.NextToken()
	c.tokenIndex++
	return token
}

// resetTokenStream reinitializes lexer state and resets the token cursor.
func (c *Compiler) resetTokenStream() {
	if c == nil {
		return
	}

	c.lexer = vbscript.NewLexer(c.sourceCode)
	c.lexer.Mode = c.lexerMode
	if c.lexerMode == vbscript.ModeASP {
		c.lexer.InASPBlock = false
	} else {
		c.lexer.InASPBlock = true
	}

	c.next = nil
	c.tokenIndex = -1
	c.lastDebugLine = -1
	c.lastDebugColumn = -1
	c.prevToken = nil
	c.lastToken = nil
	c.move()
}

// compileDefinitionPreBindingPass pre-binds global Class/Sub/Function declarations.
func (c *Compiler) compileDefinitionPreBindingPass() []definitionTokenBound {
	bounds := make([]definitionTokenBound, 0, 8)

	for !c.matchEof() {
		if !c.isGlobalDefinitionToken(c.next) {
			c.move()
			continue
		}

		start := c.tokenIndex
		c.parseStatement()
		end := max(c.tokenIndex-1, start)
		bounds = append(bounds, definitionTokenBound{start: start, end: end})
	}

	return bounds
}

func isIdentifierLikeTokenForPrescan(token vbscript.Token) (string, bool) {
	switch t := token.(type) {
	case *vbscript.IdentifierToken:
		return strings.TrimSpace(t.Name), true
	case *vbscript.KeywordOrIdentifierToken:
		return strings.TrimSpace(t.Name), true
	case *vbscript.ExtendedIdentifierToken:
		return strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(t.Name, "["), "]")), true
	default:
		return "", false
	}
}

func (c *Compiler) skipTopLevelDefinitionBlock(endKeyword vbscript.Keyword) {
	for !c.matchEof() {
		if c.checkKeyword(vbscript.KeywordEnd) {
			peek := c.peekToken()
			if c.tokenMatchesKeywordOrIdentifier(peek, endKeyword, strings.ToLower(endKeyword.String())) {
				c.move()
				c.move()
				return
			}
		}
		c.move()
	}
}

// prebindTopLevelDimDeclarations scans source-level Dim declarations before
// definition prebinding so local procedure compilation can still resolve true
// globals declared outside procedures.
func (c *Compiler) prebindTopLevelDimDeclarations() {
	if c == nil {
		return
	}

	for !c.matchEof() {
		if c.matchEof() {
			return
		}

		if c.checkKeyword(vbscript.KeywordClass) {
			c.skipTopLevelDefinitionBlock(vbscript.KeywordClass)
			continue
		}
		if c.checkKeyword(vbscript.KeywordSub) {
			c.skipTopLevelDefinitionBlock(vbscript.KeywordSub)
			continue
		}
		if c.checkKeyword(vbscript.KeywordFunction) {
			c.skipTopLevelDefinitionBlock(vbscript.KeywordFunction)
			continue
		}

		if c.checkKeyword(vbscript.KeywordPublic) || c.checkKeyword(vbscript.KeywordPrivate) {
			scopeTok := c.move()
			if c.checkKeyword(vbscript.KeywordClass) {
				c.skipTopLevelDefinitionBlock(vbscript.KeywordClass)
				continue
			}
			if c.checkKeyword(vbscript.KeywordSub) {
				c.skipTopLevelDefinitionBlock(vbscript.KeywordSub)
				continue
			}
			if c.checkKeyword(vbscript.KeywordFunction) {
				c.skipTopLevelDefinitionBlock(vbscript.KeywordFunction)
				continue
			}
			if !c.checkKeyword(vbscript.KeywordDim) {
				_ = scopeTok
				continue
			}
		}

		if c.checkKeyword(vbscript.KeywordDim) {
			c.move()
			for {
				name, ok := isIdentifierLikeTokenForPrescan(c.next)
				if !ok || name == "" {
					break
				}
				c.declareVar(name)
				c.move()

				depth := 0
				for !c.matchEof() {
					if p, ok := c.next.(*vbscript.PunctuationToken); ok {
						switch p.Type {
						case vbscript.PunctLParen:
							depth++
						case vbscript.PunctRParen:
							if depth > 0 {
								depth--
							}
						case vbscript.PunctComma:
							if depth == 0 {
								c.move()
								goto nextDimName
							}
						}
					}
					if depth == 0 {
						switch c.next.(type) {
						case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken, *vbscript.ASPCodeEndToken:
							goto endDimList
						}
					}
					c.move()
				}
				goto endDimList
			nextDimName:
			}
		endDimList:
			continue
		}

		c.move()
	}
}

// preRegisterTypeDeclarations scans the token stream with a cloned lexer for
// top-level Type...End Type blocks and pre-registers the UDT names in
// recordDeclLookup. This must run BEFORE prebindTopLevelDimDeclarations (which
// consumes the token stream) so that UDT names are in recordDeclLookup when
// Class declarations that reference them are compiled during the definition
// pre-binding pass.
func (c *Compiler) preRegisterTypeDeclarations() {
	if c == nil || c.lexer == nil {
		return
	}
	scan := *c.lexer

	for {
		tok := scan.NextToken()
		if tok == nil {
			return
		}
		if _, ok := tok.(*vbscript.EOFToken); ok {
			return
		}

		// Only match "Type" when it is an identifier followed by another
		// identifier (the type name) and then a newline — this is a type
		// declaration, not a member access like "inStream.Type = 1".
		name := ""
		switch t := tok.(type) {
		case *vbscript.IdentifierToken:
			name = t.Name
		case *vbscript.KeywordOrIdentifierToken:
			name = t.Name
		default:
			continue
		}
		if !strings.EqualFold(strings.TrimSpace(name), "type") {
			continue
		}

		// Peek at the next token to see if it's an identifier (type name).
		nextTok := scan.NextToken()
		if nextTok == nil {
			return
		}
		nextName := ""
		switch t := nextTok.(type) {
		case *vbscript.IdentifierToken:
			nextName = t.Name
		case *vbscript.KeywordOrIdentifierToken:
			nextName = t.Name
		default:
			continue
		}
		if nextName == "" {
			continue
		}

		// Peek at the third token. If it's a line terminator or a member access
		// (like "As"), we have a type declaration. If it's "=" or ".", this is
		// an assignment or member access using "Type" as a variable name, not
		// a type declaration.
		thirdTok := scan.NextToken()
		if thirdTok == nil {
			return
		}
		if _, ok := thirdTok.(*vbscript.LineTerminationToken); ok {
			// "Type Point" followed by newline — this is a type declaration.
		} else if _, ok := thirdTok.(*vbscript.CommentToken); ok {
			// "Type Point ' comment" — type declaration with comment.
		} else if kwTok, ok := thirdTok.(*vbscript.KeywordOrIdentifierToken); ok {
			if strings.EqualFold(strings.TrimSpace(kwTok.Name), "as") {
				// "Type Point As ..." — not a type declaration, probably a
				// variable declaration like "Dim x As TypeName".
				continue
			}
		} else if punctTok, ok := thirdTok.(*vbscript.PunctuationToken); ok {
			// If followed by "=" or ".", it's an assignment or member access.
			if punctTok.Type == vbscript.PunctEqual || punctTok.Type == vbscript.PunctDot {
				continue
			}
			// If followed by "(" it could be a function call "Type(...)".
			if punctTok.Type == vbscript.PunctLParen {
				continue
			}
		} else if _, ok := thirdTok.(*vbscript.LineTerminationToken); !ok {
			// If third token is anything unexpected, skip to be safe.
			continue
		}

		lowerTypeName := strings.ToLower(nextName)
		if _, exists := c.recordDeclLookup[lowerTypeName]; exists {
			continue
		}

		c.recordDeclLookup[lowerTypeName] = len(c.recordDecls)
		c.recordDecls = append(c.recordDecls, CompiledRecordDecl{
			Name:    nextName,
			Members: nil, // Populated during actual compilation
		})
	}
}

// isGlobalDefinitionToken reports whether the current token starts a global declaration block.
func (c *Compiler) isGlobalDefinitionToken(token vbscript.Token) bool {
	if c == nil || token == nil || c.isEval {
		return false
	}

	if c.tokenMatchesKeywordOrIdentifier(token, vbscript.KeywordClass, "class") ||
		c.tokenMatchesKeywordOrIdentifier(token, vbscript.KeywordEnum, "enum") ||
		c.tokenMatchesKeywordOrIdentifier(token, vbscript.KeywordSub, "sub") ||
		c.tokenMatchesKeywordOrIdentifier(token, vbscript.KeywordFunction, "function") ||
		c.tokenMatchesKeywordOrIdentifier(token, vbscript.KeywordConst, "const") {
		return true
	}

	if c.tokenMatchesKeywordOrIdentifier(token, vbscript.KeywordPublic, "public") ||
		c.tokenMatchesKeywordOrIdentifier(token, vbscript.KeywordPrivate, "private") {
		peek := c.peekToken()
		return c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordClass, "class") ||
			c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordEnum, "enum") ||
			c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordSub, "sub") ||
			c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordFunction, "function") ||
			c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordConst, "const")
	}

	return false
}

// skipDefinitionBlock advances the token stream until the first token after one definition bound.
func (c *Compiler) skipDefinitionBlock(end int) {
	for !c.matchEof() && c.tokenIndex <= end {
		c.move()
	}
}

// peekToken reads the next lexer token without consuming compiler state.
func (c *Compiler) peekToken() vbscript.Token {
	if c == nil || c.lexer == nil {
		return nil
	}
	lexerCopy := *c.lexer
	return lexerCopy.NextToken()
}

func (c *Compiler) matchEof() bool {
	_, ok := c.next.(*vbscript.EOFToken)
	return ok
}

func (c *Compiler) emit(op OpCode, operands ...int) int {
	pos := len(c.bytecode)
	c.bytecode = append(c.bytecode, byte(op))
	if op == OpCoerceToValue {
		c.lastCoercePos = pos
	} else {
		c.lastCoercePos = -1
	}

	for _, operand := range operands {
		operandWidth := 2
		if usesWideJumpOperand(op) {
			if operand < 0 || uint64(operand) > uint64(math.MaxUint32) {
				panic("Bytecode exceeds 32-bit jump target limit")
			}
			operandWidth = 4
		}
		buf := make([]byte, operandWidth)
		if operandWidth == 4 {
			binary.BigEndian.PutUint32(buf, uint32(operand))
		} else {
			binary.BigEndian.PutUint16(buf, uint16(operand))
		}
		c.bytecode = append(c.bytecode, buf...)
	}

	if op == OpCallMember {
		// Inline cache slot reserved for VM monomorphic call-site caching.
		c.bytecode = append(c.bytecode, 0, 0, 0, 0)
	}

	return pos
}

// emitExt emits one extended opcode sequence using the OpExtPrefix escape.
// Current extended opcodes use 16-bit operands only.
func (c *Compiler) emitExt(op ExtOpCode, operands ...int) int {
	pos := len(c.bytecode)
	c.bytecode = append(c.bytecode, byte(OpExtPrefix), byte(op))
	c.lastCoercePos = -1

	for _, operand := range operands {
		if operand < 0 || uint64(operand) > uint64(math.MaxUint16) {
			panic("Extended opcode operand exceeds 16-bit range")
		}
		buf := make([]byte, 2)
		binary.BigEndian.PutUint16(buf, uint16(operand))
		c.bytecode = append(c.bytecode, buf...)
	}

	return pos
}

func usesWideJumpOperand(op OpCode) bool {
	switch op {
	case OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel:
		fallthrough
	case OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter, OpJSBreak, OpJSContinue, OpJSForInCleanup, OpJSForOfCleanup,
		OpJSJumpIfNullish, OpJSJumpIfNotNullish, OpJSJumpIfNotUndefined,
		OpJumpIfNotEq, OpJumpIfEq, OpJumpIfNotLt, OpJumpIfLte, OpJumpIfNotIs,
		OpJSJumpIfLooseNotEq, OpJSJumpIfLooseEq, OpJSJumpIfStrictNotEq, OpJSJumpIfStrictEq, OpJSJumpIfNotLess, OpJSJumpIfLessEqual:
		return true
	default:
		return false
	}
}

// addConstant appends or reuses a constant value in the compiler's constant pool.
// String constants (VTString) are deduplicated via c.constantMap to avoid exponential
// pool growth from repeated identifiers, property names, and string literals in JScript blocks.
// VTEmpty is never deduplicated because it serves as a mutable placeholder in the VBScript
// compiler (patched later with NewUserSubEx).
func (c *Compiler) addConstant(v Value) int {
	if v.Type == VTString {
		key := v.Str
		if idx, ok := c.constantMap[key]; ok {
			return idx
		}
		c.constants = append(c.constants, v)
		idx := len(c.constants) - 1
		c.constantMap[key] = idx
		return idx
	}
	c.constants = append(c.constants, v)
	return len(c.constants) - 1
}

func (c *Compiler) registerLabel(name string) {
	lower := strings.ToLower(name)
	if _, exists := c.labelMap[lower]; exists {
		panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("Label '%s' already defined", name)))
	}
	pos := len(c.bytecode)
	c.labelMap[lower] = pos

	// Patch forward references
	if patches, ok := c.forwardLabelPatches[lower]; ok {
		for _, patchPos := range patches {
			c.patchJumpTo(patchPos, pos)
		}
		delete(c.forwardLabelPatches, lower)
	}
}

func (c *Compiler) emitGoTo(name string) {
	lower := strings.ToLower(name)
	if pos, ok := c.labelMap[lower]; ok {
		c.emit(OpJump, pos)
	} else {
		// Forward reference
		jumpPos := c.emitJump(OpJump)
		c.forwardLabelPatches[lower] = append(c.forwardLabelPatches[lower], jumpPos)
	}
}

func (c *Compiler) nextIdentifierName() string {
	if c.next == nil {
		return ""
	}
	switch t := c.next.(type) {
	case *vbscript.IdentifierToken:
		return t.Name
	case *vbscript.KeywordOrIdentifierToken:
		return t.Name
	case *vbscript.ExtendedIdentifierToken:
		return strings.TrimSuffix(strings.TrimPrefix(t.Name, "["), "]")
	default:
		return ""
	}
}

func (c *Compiler) Bytecode() []byte {
	return c.bytecode
}

func (c *Compiler) Constants() []Value {
	return c.constants
}

func (c *Compiler) GlobalsCount() int {
	return c.Globals.Count()
}

// IsJSModule reports whether the current compilation targets a pure JavaScript module.
func (c *Compiler) IsJSModule() bool {
	if c == nil {
		return false
	}
	return c.isJSModule
}

// IsASP reports whether the current compilation targets standard ASP (HTML + delimiters).
func (c *Compiler) IsASP() bool {
	if c == nil {
		return false
	}
	return c.lexerMode == vbscript.ModeASP
}

// IsEval reports whether the current compilation targets a dynamic expression.
func (c *Compiler) IsEval() bool {
	if c == nil {
		return false
	}
	return c.isEval
}

// ActiveVBSConstants returns the ordered predefined constant set active in this compilation.
func (c *Compiler) ActiveVBSConstants() []VBSConstant {
	if c == nil {
		return nil
	}
	out := make([]VBSConstant, len(c.activeVBSConstants))
	copy(out, c.activeVBSConstants)
	return out
}

// GlobalVarTypes returns a copy of the VB6 As Type declarations for global variables.
func (c *Compiler) GlobalVarTypes() map[string]ValueType {
	if c == nil {
		return nil
	}
	out := make(map[string]ValueType, len(c.globalVarTypes))
	maps.Copy(out, c.globalVarTypes)
	return out
}

// GlobalRecordTypes returns a copy of the declared UDT/Class names for global variables.
func (c *Compiler) GlobalRecordTypes() map[string]string {
	if c == nil {
		return nil
	}
	out := make(map[string]string, len(c.globalRecordTypes))
	maps.Copy(out, c.globalRecordTypes)
	return out
}

// LocalVarTypes returns a copy of the VB6 As Type declarations for local variables.
func (c *Compiler) LocalVarTypes() map[string]ValueType {
	if c == nil {
		return nil
	}
	out := make(map[string]ValueType, len(c.localVarTypes))
	maps.Copy(out, c.localVarTypes)
	return out
}

// FuncLocalTypes returns the accumulated function local types map.
func (c *Compiler) FuncLocalTypes() map[int]map[string]ValueType {
	if c == nil {
		return nil
	}
	return c.funcLocalTypes
}

// FuncLocalRecordTypes returns the accumulated function local record/class types map.
func (c *Compiler) FuncLocalRecordTypes() map[int]map[string]string {
	if c == nil {
		return nil
	}
	return c.funcLocalRecordTypes
}

// LocalRecordTypes returns a copy of the declared UDT/Class names for local variables.
func (c *Compiler) LocalRecordTypes() map[string]string {
	if c == nil {
		return nil
	}
	out := make(map[string]string, len(c.localRecordTypes))
	maps.Copy(out, c.localRecordTypes)
	return out
}

// SetSourceName attaches the current source file name to compiler-generated errors.
func (c *Compiler) SetSourceName(sourceName string) {
	if c == nil {
		return
	}
	trimmed := strings.TrimSpace(sourceName)
	if trimmed == "" {
		c.sourceName = ""
		return
	}
	if !filepath.IsAbs(trimmed) && !strings.HasPrefix(trimmed, "/") && !strings.HasPrefix(trimmed, "\\") {
		if absSource, err := filepath.Abs(trimmed); err == nil {
			trimmed = absSource
		}
	}
	c.sourceName = filepath.Clean(trimmed)
}

// SetIncludeSiteRoot sets the virtual site root used by SSI include virtual resolution.
func (c *Compiler) SetIncludeSiteRoot(rootDir string) {
	if c == nil {
		return
	}
	c.includeSiteRoot = normalizeIncludeSiteRoot(rootDir)
}

// IncludeSiteRoot returns the normalized virtual site root used for SSI include resolution.
func (c *Compiler) IncludeSiteRoot() string {
	if c == nil {
		return ""
	}
	return c.includeSiteRoot
}

// SetIncludeCaseInsensitiveEnabled toggles case-insensitive filesystem lookup fallback for SSI includes.
func (c *Compiler) SetIncludeCaseInsensitiveEnabled(enabled bool) {
	if c == nil {
		return
	}
	c.includeCaseInsensitive = enabled
}

// IncludeDependencies returns resolved include files discovered during preprocessing.
func (c *Compiler) IncludeDependencies() []string {
	if c == nil || len(c.includeDeps) == 0 {
		return nil
	}
	deps := make([]string, len(c.includeDeps))
	copy(deps, c.includeDeps)
	return deps
}

func (c *Compiler) mapMergedSourceLine(line int) (string, int, bool) {
	if c == nil {
		return "", 0, false
	}
	return c.sourceMap.ResolveLine(line)
}

// SourceMapEntries returns a stable copy of sparse source-map entries for cache/runtime error mapping.
func (c *Compiler) SourceMapEntries() []SourceMapEntry {
	if c == nil {
		return nil
	}
	return c.sourceMap.Entries()
}

func (c *Compiler) lineColumnFromSourceOffset(offset int) (int, int) {
	if c == nil || offset < 0 {
		return 0, 0
	}
	if offset > len(c.sourceCode) {
		offset = len(c.sourceCode)
	}
	line := 1
	lineStart := 0
	for i := 0; i < offset; i++ {
		switch c.sourceCode[i] {
		case '\r':
			line++
			if i+1 < offset && c.sourceCode[i+1] == '\n' {
				i++
			}
			lineStart = i + 1
		case '\n':
			line++
			lineStart = i + 1
		}
	}
	column := max(offset-lineStart, 0)
	return line, column
}

// vbSyntaxError creates a VBScript-compatible syntax error using the current compiler token.
func (c *Compiler) vbSyntaxError(code vbscript.VBSyntaxErrorCode) *vbscript.VBSyntaxError {
	return c.vbSyntaxErrorAtToken(code, c.next)
}

func (c *Compiler) vbSyntaxErrorAtToken(code vbscript.VBSyntaxErrorCode, token vbscript.Token) *vbscript.VBSyntaxError {
	if token == nil {
		err := vbscript.NewVBSyntaxError(code, 0, 0, "", "")
		if c.sourceName != "" {
			err.WithFile(c.sourceName)
		}
		return err
	}

	line, column := c.lineColumnFromSourceOffset(token.GetStart())

	tokenText := c.tokenSourceText(token)
	lineText := c.lineSourceText(token)

	err := vbscript.NewVBSyntaxError(code, line, column, tokenText, lineText)
	if mappedFile, mappedLine, ok := c.mapMergedSourceLine(line); ok {
		if mappedFile != "" {
			err.WithFile(mappedFile)
		}
		if mappedLine > 0 {
			err.Line = mappedLine
		}
	} else if c.sourceName != "" {
		err.WithFile(c.sourceName)
	}
	return err
}

// vbCompileError creates a VBScript-compatible compiler error and preserves the detailed compiler message for ASP consumers.
func (c *Compiler) vbCompileError(code vbscript.VBSyntaxErrorCode, detail string) *vbscript.VBSyntaxError {
	err := c.vbSyntaxErrorAtToken(code, c.next)
	err.WithASPDescription(detail)
	err.Description = detail
	return err
}

// vbCompileErrorAtToken creates a VBScript-compatible compiler error anchored to one explicit token.
func (c *Compiler) vbCompileErrorAtToken(code vbscript.VBSyntaxErrorCode, token vbscript.Token, detail string) *vbscript.VBSyntaxError {
	err := c.vbSyntaxErrorAtToken(code, token)
	err.WithASPDescription(detail)
	err.Description = detail
	return err
}

// normalizeCompileError converts raw compiler panics into VBScript-compatible syntax errors when possible.
func (c *Compiler) normalizeCompileError(err error) error {
	if err == nil {
		return nil
	}

	var jsSyntaxErr *jscript.JSSyntaxError
	if errors.As(err, &jsSyntaxErr) {
		mapped := false
		if c != nil && jsSyntaxErr != nil {
			if file, line, ok := c.mapMergedSourceLine(jsSyntaxErr.Line); ok {
				if file != "" {
					jsSyntaxErr.WithFile(file)
					mapped = true
				}
				if line > 0 {
					jsSyntaxErr.Line = line
					mapped = true
				}
			}
		}
		if !mapped && c.sourceName != "" {
			jsSyntaxErr.WithFile(c.sourceName)
		}
		return jsSyntaxErr
	}

	var syntaxErr *vbscript.VBSyntaxError
	if errors.As(err, &syntaxErr) {
		if syntaxErr != nil && strings.TrimSpace(syntaxErr.File) != "" {
			return syntaxErr
		}
		mapped := false
		if c != nil && syntaxErr != nil {
			if file, line, ok := c.mapMergedSourceLine(syntaxErr.Line); ok {
				if file != "" {
					syntaxErr.WithFile(file)
					mapped = true
				}
				if line > 0 {
					syntaxErr.Line = line
					mapped = true
				}
			}
		}
		if !mapped && c.sourceName != "" {
			syntaxErr.WithFile(c.sourceName)
		}
		return syntaxErr
	}

	message := strings.TrimSpace(err.Error())
	if message == "" {
		return err
	}

	code, ok := c.compileMessageCode(message)
	if !ok {
		return err
	}

	return c.vbCompileError(code, message)
}

// compileMessageCode maps known compiler panic text to the closest VBScript catalog code.
func (c *Compiler) compileMessageCode(message string) (vbscript.VBSyntaxErrorCode, bool) {
	switch {
	case strings.Contains(message, "Variable not defined"):
		return vbscript.VariableNotDefined, true
	case strings.Contains(message, "Unexpected token"):
		return vbscript.SyntaxError, true
	case strings.HasPrefix(message, "Expected ')'"):
		return vbscript.ExpectedRParen, true
	case strings.HasPrefix(message, "Expected array bounds"):
		return vbscript.ExpectedLParen, true
	case strings.HasPrefix(message, "Expected identifier"), strings.HasPrefix(message, "expected directive identifier"):
		return vbscript.ExpectedIdentifier, true
	case strings.HasPrefix(message, "Expected keyword "):
		return c.keywordMessageCode(message), true
	case strings.HasPrefix(message, "expected '='"):
		return vbscript.ExpectedEqual, true
	case strings.HasPrefix(message, "expected directive value"):
		return vbscript.ExpectedLiteral, true
	case strings.Contains(message, "unterminated ASP directive block"):
		return vbscript.SyntaxError, true
	case strings.Contains(message, "invalid ASP code page directive value"):
		return vbscript.InvalidNumber, true
	case strings.Contains(message, "unsupported ASP language directive"), strings.Contains(message, "invalid ASP EnableSessionState directive value"):
		return vbscript.SyntaxError, true
	default:
		return 0, false
	}
}

// keywordMessageCode maps expected-keyword panic messages to the matching VBScript syntax code.
func (c *Compiler) keywordMessageCode(message string) vbscript.VBSyntaxErrorCode {
	keyword := strings.TrimSpace(strings.TrimPrefix(message, "Expected keyword"))
	switch strings.ToLower(keyword) {
	case "if":
		return vbscript.ExpectedIf
	case "to":
		return vbscript.ExpectedTo
	case "end":
		return vbscript.ExpectedEnd
	case "function":
		return vbscript.ExpectedFunction
	case "sub":
		return vbscript.ExpectedSub
	case "then":
		return vbscript.ExpectedThen
	case "wend":
		return vbscript.ExpectedWend
	case "loop":
		return vbscript.ExpectedLoop
	case "next":
		return vbscript.ExpectedNext
	case "case":
		return vbscript.ExpectedCase
	case "select":
		return vbscript.ExpectedSelect
	case "while":
		return vbscript.ExpectedWhileOrUntil
	case "with":
		return vbscript.ExpectedWith
	case "class":
		return vbscript.ExpectedClass
	case "property":
		return vbscript.ExpectedProperty
	default:
		return vbscript.SyntaxError
	}
}

// tokenSourceText returns the original source text for the current token when available.
func (c *Compiler) tokenSourceText(token vbscript.Token) string {
	if c == nil || c.lexer == nil || token == nil {
		return ""
	}

	start := token.GetStart()
	end := max(token.GetEnd(), start)

	return c.sourceSlice(start, end)
}

// lineSourceText returns the full source line that contains the provided token.
func (c *Compiler) lineSourceText(token vbscript.Token) string {
	if c == nil || c.lexer == nil || token == nil {
		return ""
	}

	start := max(token.GetLineStart(), 0)

	codeRunes := []rune(c.lexer.Code)
	if start >= len(codeRunes) {
		return ""
	}

	end := start
	for end < len(codeRunes) {
		if codeRunes[end] == '\n' || codeRunes[end] == '\r' {
			break
		}
		end++
	}

	return string(codeRunes[start:end])
}

// sourceSlice returns a safe rune-aware slice from the compiler source code.
func (c *Compiler) sourceSlice(start int, end int) string {
	if c == nil || c.lexer == nil {
		return ""
	}

	if start < 0 {
		start = 0
	}
	if end < start {
		end = start
	}

	codeRunes := []rune(c.lexer.Code)
	if start >= len(codeRunes) {
		return ""
	}
	if end > len(codeRunes) {
		end = len(codeRunes)
	}

	return string(codeRunes[start:end])
}

// dumpPreprocessedSource writes the fully expanded source code to ./temp/dump_preprocessed_source_<name>.asp
// for debugging and inspection purposes. Executes synchronously to ensure the file is written before returning,
// which is critical for debugging compilation errors that abort mid-parse.
// The file name is derived from sourceName so that nested compilations (Eval, Execute, ExecuteGlobal)
// each produce a distinct file instead of overwriting the parent dump.
// Enabled via SetDumpPreprocessedSourceEnabled or the DUMP_PREPROCESSED_SOURCE environment variable.
func dumpPreprocessedSource(sourceCode, sourceName string) {
	if !dumpPreprocessedSourceEnabled.Load() && os.Getenv("DUMP_PREPROCESSED_SOURCE") == "" {
		return
	}

	dumpDir := resolveConfiguredTempDir()

	// Build a safe file-system name from the source identifier.
	name := strings.TrimSpace(sourceName)
	if name == "" {
		name = "dynamic"
	}
	safe := strings.NewReplacer("\\", "_", "/", "_", ":", "_", " ", "_").Replace(name)
	dumpFile := filepath.Join(dumpDir, "dump_preprocessed_source_"+safe+".asp")

	if err := os.MkdirAll(dumpDir, 0755); err != nil {
		return
	}

	// Synchronous write: guarantees the file exists before Compile() returns, even on panic paths.
	_ = os.WriteFile(dumpFile, []byte(sourceCode), 0644)
}

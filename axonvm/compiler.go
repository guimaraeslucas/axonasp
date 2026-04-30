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
	"math"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"sync/atomic"

	"g3pix.com.br/axonasp/jscript"
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

// SourceLineRef maps one expanded source line back to its original file and line.
type SourceLineRef struct {
	File string
	Line int
}

// LineMap stores one origin entry per expanded line (1-based by index+1).
type LineMap []SourceLineRef

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

func (s *SymbolTable) Count() int {
	return len(s.names)
}

// Compiler handles single-pass compilation of VBScript tokens into VM bytecode.
type Compiler struct {
	lexer               *vbscript.Lexer
	lexerMode           vbscript.LexerMode
	sourceCode          string
	next                vbscript.Token
	tokenIndex          int
	bytecode            []byte
	constants           []Value
	Globals             *SymbolTable
	locals              *SymbolTable     // Current function scope
	declaredGlobals     map[string]bool  // Variables declared via Dim in global scope
	declaredLocals      map[string]bool  // Variables declared via Dim in local scope
	constGlobals        map[string]bool  // Constants declared via Const in global scope
	constLocals         map[string]bool  // Constants declared via Const in local scope
	constLiteralGlobals map[string]Value // Compile-time known global constant values
	isLocal             bool             // True if currently compiling a Sub/Function

	// Compilation Options
	optionExplicit bool // Requires variables to be Dim'ed
	optionCompare  int  // 0: Binary (default), 1: Text
	optionStrict   bool // (Future) Enforces strict typing
	optionInfer    bool // (Future) Allows type inference
	sourceName     string
	lineMap        LineMap
	includeDeps    []string

	forwardCallPatches  map[string][]int
	forwardConstPatches map[string][]int
	lastCallTargetName  string
	lastCallTargetPos   int
	lastCallIsGlobal    bool
	lastCoercePos       int
	lastDebugLine       int
	lastDebugColumn     int
	tempCounter         int
	globalZeroArgFuncs  map[string]bool
	classDecls          []CompiledClassDecl
	classDeclLookup     map[string]int
	ObjectDeclarations  []*vbscript.ASPObjectToken
	currentClassName    string
	currentFunctionName string
	// dynamicMemberResolution enables implicit class-member lookup for dynamic
	// code compiled by Eval/Execute/ExecuteGlobal while keeping global assignment semantics.
	dynamicMemberResolution bool
	loopContexts            []loopContext
	jsLoopContexts          []*jsLoopContext // Loop contexts for JScript
	jsBreakContexts         []*jsBreakContext
	jsStrictMode            bool              // Current JScript strict mode state
	jsFunctionStrictModes   map[int]bool      // Maps function start IP to strict mode
	jsBlockScopeStack       []map[string]bool // Stack of declared block-scoped variables (let/const)
	jsForIterScopes         []jsForIterScope  // Stack of active per-iteration for-let scopes
	// withDepth tracks nesting level of With...End With blocks at compile time.
	// A value > 0 enables the leading-dot '.' statement and expression syntax.
	withDepth          int
	activeVBSConstants []VBSConstant
	// userGlobalsStart is the index of the first user-declared global variable slot.
	// Slots below this index are read-only pre-injected intrinsics, built-ins, or VBScript constants.
	// Only global slots at or above this index are eligible for ByRef argument write-back.
	userGlobalsStart int
	isEval           bool // True if compiling a VBScript expression for Eval()
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
}

type definitionTokenBound struct {
	start int
	end   int
}

// CompiledClassDecl stores bootstrap metadata for one class declaration.
// It is intentionally minimal during staged class implementation.
type CompiledClassDecl struct {
	Name       string
	Fields     []CompiledClassFieldDecl
	Methods    []CompiledClassMethodDecl
	Properties []CompiledClassPropertyDecl
}

// CompiledClassFieldDecl stores one compiled class field metadata entry.
type CompiledClassFieldDecl struct {
	Name     string
	IsPublic bool
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
		lexer:                 lexer,
		lexerMode:             mode,
		sourceCode:            code,
		tokenIndex:            -1,
		bytecode:              make([]byte, 0),
		constants:             make([]Value, 0),
		Globals:               NewSymbolTable(),
		locals:                NewSymbolTable(),
		declaredGlobals:       make(map[string]bool),
		declaredLocals:        make(map[string]bool),
		constGlobals:          make(map[string]bool),
		constLocals:           make(map[string]bool),
		constLiteralGlobals:   make(map[string]Value),
		isLocal:               false,
		optionExplicit:        false,
		optionCompare:         0, // Binary
		forwardCallPatches:    make(map[string][]int),
		forwardConstPatches:   make(map[string][]int),
		lastCallTargetPos:     -1,
		lastCoercePos:         -1,
		lastDebugLine:         -1,
		lastDebugColumn:       -1,
		tempCounter:           0,
		globalZeroArgFuncs:    make(map[string]bool),
		classDecls:            make([]CompiledClassDecl, 0),
		classDeclLookup:       make(map[string]int),
		loopContexts:          make([]loopContext, 0),
		jsStrictMode:          false,
		jsFunctionStrictModes: make(map[int]bool),
		jsBlockScopeStack:     make([]map[string]bool, 0, 32),
		activeVBSConstants:    make([]VBSConstant, 0, len(VBSConstants)),
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

var includeDirectivePattern = regexp.MustCompile(`(?is)<!--\s*#include\s+(virtual|file)\s*=\s*("([^"]+)"|'([^']+)')\s*-->`)

// emptyASPBlockPattern matches empty ASP blocks that carry no executable code.
// This is used as a lightweight preprocessor cleanup step before lexing.

// var emptyASPBlockPattern = regexp.MustCompile(`(?is)<%=\s*%>|<%\s*%>|<%%>`)
var emptyASPBlockPattern = regexp.MustCompile(`<%\s*=?\s*%>`)

// stripEmptyASPBlocks removes empty ASP blocks that should not reach compilation.
func stripEmptyASPBlocks(source string) string {
	if source == "" {
		return source
	}
	return emptyASPBlockPattern.ReplaceAllString(source, "")
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
func preprocessASPIncludesWithDeps(source string, sourceName string, visited map[string]bool, depth int, dependencies *[]string) (string, LineMap, error) {
	if depth > 32 {
		return "", nil, fmt.Errorf("include recursion limit exceeded")
	}

	processed := source
	lineMap := buildIdentityLineMap(source, sourceName)
	for {
		match := includeDirectivePattern.FindStringSubmatchIndex(processed)
		if match == nil {
			break
		}

		replaceStart := match[0]
		replaceEnd := match[1]
		if lineOnlyIncludeDirective(processed, match[0], match[1]) {
			replaceStart = lineStartIndex(processed, match[0])
			replaceEnd = lineEndIndexIncludingNewline(processed, match[1])
		}

		directive := processed[match[0]:match[1]]
		kind := strings.ToLower(processed[match[2]:match[3]])
		pathVal := ""
		if match[6] != -1 && match[7] != -1 {
			pathVal = processed[match[6]:match[7]]
		} else if match[8] != -1 && match[9] != -1 {
			pathVal = processed[match[8]:match[9]]
		}

		resolvedPath, err := resolveIncludePath(sourceName, pathVal, kind == "virtual")
		if err != nil {
			return "", nil, fmt.Errorf("include resolve failed for %s: %w", directive, err)
		}

		norm := strings.ToLower(filepath.Clean(resolvedPath))
		if visited[norm] {
			return "", nil, fmt.Errorf("circular include detected: %s", resolvedPath)
		}

		contentBytes, err := os.ReadFile(resolvedPath)
		if err != nil {
			return "", nil, fmt.Errorf("could not read include %s: %w", resolvedPath, err)
		}

		// Strip UTF-8 BOM if present to prevent output corruption
		contentBytes = stripUTF8BOM(contentBytes)

		visited[norm] = true
		if dependencies != nil {
			*dependencies = append(*dependencies, resolvedPath)
		}
		expanded, childMap, err := preprocessASPIncludesWithDeps(string(contentBytes), resolvedPath, visited, depth+1, dependencies)
		delete(visited, norm)
		if err != nil {
			return "", nil, err
		}

		prefix := processed[:replaceStart]
		suffix := processed[replaceEnd:]
		prefixLines := countLines(prefix)
		directiveLines := countLines(processed[replaceStart:replaceEnd])
		suffixStart := prefixLines + directiveLines

		newMap := make(LineMap, 0, prefixLines+len(childMap)+(len(lineMap)-suffixStart))
		if prefixLines > 0 {
			newMap = append(newMap, lineMap[:prefixLines]...)
		}
		if len(childMap) > 0 {
			newMap = append(newMap, childMap...)
		}
		if suffixStart < len(lineMap) {
			newMap = append(newMap, lineMap[suffixStart:]...)
		}
		lineMap = newMap

		processed = prefix + expanded + suffix
	}

	return processed, lineMap, nil
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
	idx := strings.LastIndexByte(s[:pos], '\n')
	if idx == -1 {
		return 0
	}
	return idx + 1
}

// lineEndIndex returns the index of '\n' or end-of-string for the line that contains pos.
func lineEndIndex(s string, pos int) int {
	if pos >= len(s) {
		return len(s)
	}
	idx := strings.IndexByte(s[pos:], '\n')
	if idx == -1 {
		return len(s)
	}
	return pos + idx
}

// lineEndIndexIncludingNewline returns the line end index and consumes one trailing newline when present.
func lineEndIndexIncludingNewline(s string, pos int) int {
	le := lineEndIndex(s, pos)
	if le < len(s) && s[le] == '\n' {
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

// buildIdentityLineMap maps each line to the same source file and line number.
func buildIdentityLineMap(source string, sourceName string) LineMap {
	total := countLines(source)
	if total <= 0 {
		return nil
	}
	m := make(LineMap, total)
	for i := 0; i < total; i++ {
		m[i] = SourceLineRef{File: sourceName, Line: i + 1}
	}
	return m
}

// countLines returns the number of logical lines using '\n' as separator, minimum 1 for non-empty source.
func countLines(s string) int {
	if s == "" {
		return 0
	}
	return strings.Count(s, "\n") + 1
}

// resolveIncludePath resolves one ASP include path against current source path context.
func resolveIncludePath(sourceName string, includePath string, isVirtual bool) (string, error) {
	trimmed := strings.TrimSpace(includePath)
	if trimmed == "" {
		return "", fmt.Errorf("empty include path")
	}

	trimmed = strings.ReplaceAll(trimmed, "/", string(filepath.Separator))
	trimmed = strings.ReplaceAll(trimmed, "\\", string(filepath.Separator))

	if filepath.IsAbs(trimmed) {
		return filepath.Clean(trimmed), nil
	}

	cwd, err := os.Getwd()
	if err != nil {
		return "", err
	}

	isAbsSource := filepath.IsAbs(sourceName)
	src := filepath.Clean(sourceName)
	sourceAbs := src
	if !isAbsSource {
		relSource := strings.TrimLeft(src, "/\\")
		if strings.HasPrefix(strings.ToLower(filepath.ToSlash(relSource)), "www/") {
			sourceAbs = filepath.Clean(filepath.Join(cwd, relSource))
		} else {
			sourceAbs = filepath.Clean(filepath.Join(cwd, "www", relSource))
		}
	}

	findWebRoot := func(absSource string) string {
		lower := strings.ToLower(filepath.Clean(absSource))
		needle := string(filepath.Separator) + "www" + string(filepath.Separator)
		if pos := strings.LastIndex(lower, needle); pos >= 0 {
			return absSource[:pos+len(needle)-1]
		}
		if strings.HasSuffix(lower, string(filepath.Separator)+"www") {
			return absSource
		}
		return filepath.Clean(filepath.Join(cwd, "www"))
	}

	webRoot := findWebRoot(sourceAbs)
	sourceDir := filepath.Dir(sourceAbs)

	if isVirtual || strings.HasPrefix(includePath, "/") || strings.HasPrefix(includePath, "\\") {
		rel := strings.TrimLeft(trimmed, string(filepath.Separator))
		candidates := make([]string, 0, 6)
		candidates = append(candidates,
			filepath.Clean(filepath.Join(webRoot, rel)),
			filepath.Clean(filepath.Join(cwd, rel)),
		)

		parts := strings.Split(rel, string(filepath.Separator))
		if len(parts) > 1 {
			relWithoutSite := filepath.Join(parts[1:]...)
			candidates = append(candidates,
				filepath.Clean(filepath.Join(webRoot, relWithoutSite)),
				filepath.Clean(filepath.Join(cwd, relWithoutSite)),
			)
		}

		for _, candidate := range candidates {
			if _, statErr := os.Stat(candidate); statErr == nil {
				return candidate, nil
			}
		}

		return candidates[0], nil
	}

	return filepath.Clean(filepath.Join(sourceDir, trimmed)), nil
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
		// Even if not explicitly Dim'ed, if it was used once it exists in SymbolTable.
		// We only error if Option Explicit is on and it wasn't Dim'ed.
		// NOTE: Slots below userGlobalsStart are read-only intrinsics/constants and don't need Dim.
		if c.optionExplicit && idx >= c.userGlobalsStart && !c.declaredGlobals[lower] {
			panic(c.vbCompileError(vbscript.VariableNotDefined, fmt.Sprintf("Variable not defined: '%s'", name)))
		}
		return OpGetGlobal, idx
	}

	// 3. Implicit Declaration check
	if c.optionExplicit {
		panic(c.vbCompileError(vbscript.VariableNotDefined, fmt.Sprintf("Variable not defined: '%s'", name)))
	}

	idx := c.Globals.Add(name)
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

	if c.optionExplicit && !c.declaredGlobals[lower] {
		// Special case: we might allow implicit if we are in local but it's a global
		// but standard VBScript with Option Explicit requires Dim.
		panic(c.vbCompileError(vbscript.VariableNotDefined, fmt.Sprintf("Variable not defined: '%s'", name)))
	}

	idx := c.Globals.Add(name)
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
	column := c.next.GetStart() - c.next.GetLineStart()
	if column < 0 {
		column = 0
	}
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
		end := c.tokenIndex - 1
		if end < start {
			end = start
		}
		bounds = append(bounds, definitionTokenBound{start: start, end: end})
	}

	return bounds
}

// isGlobalDefinitionToken reports whether the current token starts a global declaration block.
func (c *Compiler) isGlobalDefinitionToken(token vbscript.Token) bool {
	if c == nil || token == nil || c.isEval {
		return false
	}

	if c.matchKeywordOrIdentifier(vbscript.KeywordClass, "class") ||
		c.matchKeywordOrIdentifier(vbscript.KeywordSub, "sub") ||
		c.matchKeywordOrIdentifier(vbscript.KeywordFunction, "function") ||
		c.matchKeywordOrIdentifier(vbscript.KeywordConst, "const") {
		return true
	}

	if c.matchKeywordOrIdentifier(vbscript.KeywordPublic, "public") ||
		c.matchKeywordOrIdentifier(vbscript.KeywordPrivate, "private") {
		peek := c.peekToken()
		return c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordClass, "class") ||
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

	return pos
}

func usesWideJumpOperand(op OpCode) bool {
	switch op {
	case OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel:
		fallthrough
	case OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter, OpJSBreak, OpJSContinue, OpJSForInCleanup:
		return true
	default:
		return false
	}
}

func (c *Compiler) addConstant(v Value) int {
	c.constants = append(c.constants, v)
	return len(c.constants) - 1
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

// ActiveVBSConstants returns the ordered predefined constant set active in this compilation.
func (c *Compiler) ActiveVBSConstants() []VBSConstant {
	if c == nil {
		return nil
	}
	out := make([]VBSConstant, len(c.activeVBSConstants))
	copy(out, c.activeVBSConstants)
	return out
}

// SetSourceName attaches the current source file name to compiler-generated errors.
func (c *Compiler) SetSourceName(sourceName string) {
	if c == nil {
		return
	}

	c.sourceName = strings.TrimSpace(sourceName)
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

// vbSyntaxError creates a VBScript-compatible syntax error using the current compiler token.
func (c *Compiler) vbSyntaxError(code vbscript.VBSyntaxErrorCode) *vbscript.VBSyntaxError {
	token := c.next
	if token == nil {
		err := vbscript.NewVBSyntaxError(code, 0, 0, "", "")
		if c.sourceName != "" {
			err.WithFile(c.sourceName)
		}
		return err
	}

	line := token.GetLineNumber()
	column := token.GetStart() - token.GetLineStart()
	if column < 0 {
		column = 0
	}

	tokenText := c.tokenSourceText(token)
	lineText := c.lineSourceText(token)

	err := vbscript.NewVBSyntaxError(code, line, column, tokenText, lineText)
	if c.sourceName != "" {
		err.WithFile(c.sourceName)
	}
	return err
}

// vbCompileError creates a VBScript-compatible compiler error and preserves the detailed compiler message for ASP consumers.
func (c *Compiler) vbCompileError(code vbscript.VBSyntaxErrorCode, detail string) *vbscript.VBSyntaxError {
	err := c.vbSyntaxError(code)
	err.WithASPDescription(detail)
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
			line := jsSyntaxErr.Line
			if line > 0 && line <= len(c.lineMap) {
				ref := c.lineMap[line-1]
				if ref.File != "" {
					jsSyntaxErr.WithFile(ref.File)
					mapped = true
				}
				if ref.Line > 0 {
					jsSyntaxErr.Line = ref.Line
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
		mapped := false
		if c != nil && syntaxErr != nil {
			line := syntaxErr.Line
			if line > 0 && line <= len(c.lineMap) {
				ref := c.lineMap[line-1]
				if ref.File != "" {
					syntaxErr.WithFile(ref.File)
					mapped = true
				}
				if ref.Line > 0 {
					syntaxErr.Line = ref.Line
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
	end := token.GetEnd()
	if end < start {
		end = start
	}

	return c.sourceSlice(start, end)
}

// lineSourceText returns the full source line that contains the provided token.
func (c *Compiler) lineSourceText(token vbscript.Token) string {
	if c == nil || c.lexer == nil || token == nil {
		return ""
	}

	start := token.GetLineStart()
	if start < 0 {
		start = 0
	}

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

	dumpDir := filepath.Join(".", "temp")

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

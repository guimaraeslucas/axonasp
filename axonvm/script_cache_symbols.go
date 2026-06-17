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
	"strconv"
	"strings"
	"sync"

	"github.com/cespare/xxhash/v2"
)

type baseGlobalDictionary struct {
	names      []string
	declared   map[string]struct{}
	constNames map[string]struct{}
}

var (
	baseGlobalsOnce sync.Once
	baseGlobalsData baseGlobalDictionary
)

func getBaseGlobalDictionary() *baseGlobalDictionary {
	baseGlobalsOnce.Do(func() {
		intrinsics := []string{"Response", "Request", "Server", "Session", "Application", "ObjectContext", "Err"}
		events := []string{"OnTransactionCommit", "OnTransactionAbort"}
		total := len(intrinsics) + len(events) + len(BuiltinNames) + len(VBSConstants)
		baseGlobalsData.names = make([]string, 0, total)
		baseGlobalsData.declared = make(map[string]struct{}, total)
		baseGlobalsData.constNames = make(map[string]struct{}, len(VBSConstants))

		appendDeclared := func(name string) {
			trimmed := strings.TrimSpace(name)
			if trimmed == "" {
				return
			}
			baseGlobalsData.names = append(baseGlobalsData.names, trimmed)
			baseGlobalsData.declared[strings.ToLower(trimmed)] = struct{}{}
		}

		for i := range intrinsics {
			appendDeclared(intrinsics[i])
		}
		for i := range events {
			appendDeclared(events[i])
		}
		for i := range BuiltinNames {
			appendDeclared(BuiltinNames[i])
		}
		for i := range VBSConstants {
			name := strings.TrimSpace(VBSConstants[i].Name)
			if name == "" {
				continue
			}
			appendDeclared(name)
			baseGlobalsData.constNames[strings.ToLower(name)] = struct{}{}
		}
	})
	return &baseGlobalsData
}

func computeProgramHash(bytecode []byte, globalCount int, optionCompare int, optionExplicit bool, sourceName string, constants []Value) uint64 {
	h := xxhash.New()
	_, _ = h.Write(bytecode)

	// hash constants
	for _, c := range constants {
		binary.Write(h, binary.LittleEndian, uint8(c.Type))
		binary.Write(h, binary.LittleEndian, c.Num)
		binary.Write(h, binary.LittleEndian, c.Flt)
		switch c.Type {
		case VTString, VTUserSub, VTJSFunctionTemplate, VTJSArrowFunctionTemplate, VTSymbol:
			h.WriteString(c.Str)
		}

		if c.Type == VTUserSub || c.Type == VTJSFunctionTemplate || c.Type == VTJSArrowFunctionTemplate {
			for _, name := range c.Names {
				h.WriteString(name)
			}
		}
	}

	h.WriteString(strings.TrimSpace(sourceName))
	var raw [8]byte
	binary.LittleEndian.PutUint64(raw[:], uint64(globalCount))
	_, _ = h.Write(raw[:])
	binary.LittleEndian.PutUint64(raw[:], uint64(optionCompare))
	_, _ = h.Write(raw[:])
	if optionExplicit {
		raw[0] = 1
	} else {
		raw[0] = 0
	}
	_, _ = h.Write(raw[:1])
	return h.Sum64()
}

func buildCachedProgramFromCompiler(compiler *Compiler) CachedProgram {
	if compiler == nil {
		return CachedProgram{}
	}
	base := getBaseGlobalDictionary()
	allGlobals := compiler.Globals.names
	userStart := compiler.userGlobalsStart
	if userStart < 0 {
		userStart = 0
	}
	if userStart > len(allGlobals) {
		userStart = len(allGlobals)
	}
	baseCount := len(base.names)
	if baseCount > userStart {
		baseCount = userStart
	}
	prelude := cloneStringSlice(allGlobals[baseCount:userStart])
	users := cloneStringSlice(allGlobals[userStart:])

	// Build the VB6 As Type global type declarations list.
	var globalTypeNames []string
	var globalClassNames []string
	if len(compiler.globalVarTypes) > 0 {
		globalTypeNames = make([]string, 0, len(compiler.globalVarTypes))
		for name, t := range compiler.globalVarTypes {
			if t != VTEmpty {
				globalTypeNames = append(globalTypeNames, name+":"+strconv.Itoa(int(t)))
				if t == VTObject || t == VTRecord {
					if className, ok := compiler.globalRecordTypes[name]; ok {
						globalClassNames = append(globalClassNames, name+":"+className)
					}
				}
			}
		}
	}

	// Capture local variable types from function signatures.
	// Since compiler.localVarTypes contains ALL locals from all functions combined,
	// we need to use a different strategy or just store it as is if we can resolve it.
	// Wait! VM.applyLocalVarTypes already handles this by scanning constants.
	// So we just need to provide the maps.

	program := CachedProgram{
		Bytecode:            cloneBytecode(compiler.Bytecode()),
		Constants:           cloneValueSlice(compiler.Constants()),
		GlobalCount:         compiler.GlobalsCount(),
		OptionCompare:       compiler.optionCompare,
		OptionExplicit:      compiler.optionExplicit,
		SourceName:          compiler.sourceName,
		IncludeSiteRoot:     compiler.IncludeSiteRoot(),
		GlobalNames:         cloneStringSlice(allGlobals),
		GlobalPreludeNames:  prelude,
		GlobalPreludeConsts: filterNamesByFlagSet(compiler.constGlobals, prelude),
		UserGlobalNames:     users,
		UserDeclaredGlobals: filterNamesByFlagSet(compiler.declaredGlobals, users),
		UserConstGlobals:    filterNamesByFlagSet(compiler.constGlobals, users),
		GlobalZeroArgFuncs:  sortedTrueKeys(compiler.globalZeroArgFuncs),
		GlobalTypeNames:     globalTypeNames,
		GlobalClassNames:    globalClassNames,
		LocalVarTypes:       compiler.LocalVarTypes(),
		LocalClassTypes:     compiler.LocalRecordTypes(),
		FuncParamDefaults:   cloneFuncParamDefaultsMap(compiler.funcParamDefaults),
		IncludeDependencies: compiler.IncludeDependencies(),
		SourceMapEntries:    compiler.SourceMapEntries(),
		RecordDecls:         cloneRecordDeclSlice(compiler.recordDecls),
		RecordDeclLookup:    cloneIntMap(compiler.recordDeclLookup),
		JSICNodeCount:       compiler.jsICNodeCount,
	}

	if compiler.IsJSModule() {
		program.EngineMode = EngineModeJavaScript
	} else if compiler.IsASP() {
		program.EngineMode = EngineModeDefault
	} else {
		program.EngineMode = EngineModeVBScript
	}

	// Pre-compute lowercased global names for zero-allocation VM resets.
	allLower := make([]string, 0, len(allGlobals))
	for _, name := range allGlobals {
		allLower = append(allLower, strings.ToLower(strings.TrimSpace(name)))
	}
	program.GlobalNamesLower = allLower

	program.ProgramHash = computeProgramHash(
		program.Bytecode,
		program.GlobalCount,
		program.OptionCompare,
		program.OptionExplicit,
		program.SourceName,
		program.Constants,
	)
	return program
}

func cloneFuncParamDefaultsMap(values map[int][]int) map[int][]int {
	if len(values) == 0 {
		return nil
	}
	cloned := make(map[int][]int, len(values))
	for entryPoint, defaults := range values {
		if len(defaults) == 0 {
			cloned[entryPoint] = nil
			continue
		}
		dup := make([]int, len(defaults))
		copy(dup, defaults)
		cloned[entryPoint] = dup
	}
	return cloned
}

func applyProgramGlobalMetadata(vm *VM, program CachedProgram) {
	if vm == nil {
		return
	}
	clear(vm.declaredGlobals)
	clear(vm.constGlobals)

	if len(program.GlobalNames) > 0 {
		vm.globalNames = program.GlobalNames
		if len(program.GlobalNamesLower) > 0 {
			vm.baseGlobalNamesLower = program.GlobalNamesLower
		}
		vm.rebuildGlobalNameIndex()
		for i := range program.DeclaredGlobalNames {
			name := strings.ToLower(strings.TrimSpace(program.DeclaredGlobalNames[i]))
			if name != "" {
				vm.declaredGlobals[name] = true
			}
		}
		for i := range program.ConstGlobalNames {
			name := strings.ToLower(strings.TrimSpace(program.ConstGlobalNames[i]))
			if name != "" {
				vm.constGlobals[name] = true
			}
		}
	} else {
		base := getBaseGlobalDictionary()
		vm.globalNames = vm.globalNames[:0]
		vm.globalNames = append(vm.globalNames, base.names...)
		vm.globalNames = append(vm.globalNames, program.GlobalPreludeNames...)
		vm.globalNames = append(vm.globalNames, program.UserGlobalNames...)
		if len(program.GlobalNamesLower) > 0 {
			vm.baseGlobalNamesLower = program.GlobalNamesLower
		}
		vm.rebuildGlobalNameIndex()

		for key := range base.declared {
			vm.declaredGlobals[key] = true
		}
		for key := range base.constNames {
			vm.constGlobals[key] = true
		}
		for i := range program.GlobalPreludeNames {
			name := strings.ToLower(strings.TrimSpace(program.GlobalPreludeNames[i]))
			if name != "" {
				vm.declaredGlobals[name] = true
			}
		}
		for i := range program.GlobalPreludeConsts {
			name := strings.ToLower(strings.TrimSpace(program.GlobalPreludeConsts[i]))
			if name != "" {
				vm.constGlobals[name] = true
			}
		}
		for i := range program.UserDeclaredGlobals {
			name := strings.ToLower(strings.TrimSpace(program.UserDeclaredGlobals[i]))
			if name != "" {
				vm.declaredGlobals[name] = true
			}
		}
		for i := range program.UserConstGlobals {
			name := strings.ToLower(strings.TrimSpace(program.UserConstGlobals[i]))
			if name != "" {
				vm.constGlobals[name] = true
			}
		}
	}

	clear(vm.globalZeroArgFuncs)
	for i := range program.GlobalZeroArgFuncs {
		name := strings.ToLower(strings.TrimSpace(program.GlobalZeroArgFuncs[i]))
		if name == "" {
			continue
		}
		vm.globalZeroArgFuncs[name] = true
	}

	// Apply VB6 As Type global variable type declarations from cached program.
	if len(program.GlobalTypeNames) > 0 {
		// Ensure globalTypes slice is large enough.
		if len(vm.globalTypes) < len(vm.Globals) {
			vm.globalTypes = make([]ValueType, len(vm.Globals))
		}
		for i := range program.GlobalTypeNames {
			entry := program.GlobalTypeNames[i]
			colonIdx := strings.LastIndex(entry, ":")
			if colonIdx < 1 || colonIdx >= len(entry)-1 {
				continue
			}
			typeName := entry[:colonIdx]
			typeValStr := entry[colonIdx+1:]
			typeVal, err := strconv.Atoi(typeValStr)
			if err != nil || typeVal < 0 || typeVal > 255 {
				continue
			}
			declaredType := ValueType(typeVal)
			if declaredType == VTEmpty {
				continue
			}
			lower := strings.ToLower(typeName)
			// Find the global slot index by matching against vm.globalNames.
			for idx, gname := range vm.globalNames {
				if strings.EqualFold(gname, lower) {
					if idx < len(vm.globalTypes) {
						vm.globalTypes[idx] = declaredType
					}
					if idx < len(vm.Globals) {
						vm.Globals[idx] = vm.zeroValueForType(declaredType)
					}
					break
				}
			}
		}
	}

	// Apply VB6 Class/Interface names for global variables.
	if len(program.GlobalClassNames) > 0 {
		for i := range program.GlobalClassNames {
			entry := program.GlobalClassNames[i]
			colonIdx := strings.LastIndex(entry, ":")
			if colonIdx < 1 || colonIdx >= len(entry)-1 {
				continue
			}
			varName := entry[:colonIdx]
			className := entry[colonIdx+1:]
			lower := strings.ToLower(varName)
			for idx, gname := range vm.globalNames {
				if strings.EqualFold(gname, lower) {
					if idx < len(vm.Globals) {
						vm.Globals[idx].Interface = className
					}
					if vm.globalClassTypes != nil {
						vm.globalClassTypes[uint16(idx)] = className
					}
					break
				}
			}
		}
	}
}

func migrateLegacyCachedProgramSymbols(program *CachedProgram) {
	if program == nil || len(program.GlobalNames) == 0 {
		return
	}
	base := getBaseGlobalDictionary()
	baseCount := len(base.names)
	if baseCount > len(program.GlobalNames) {
		baseCount = len(program.GlobalNames)
	}
	program.GlobalPreludeNames = cloneStringSlice(program.GlobalNames[baseCount:baseCount])
	if len(program.GlobalNames) > baseCount {
		program.UserGlobalNames = cloneStringSlice(program.GlobalNames[baseCount:])
	} else {
		program.UserGlobalNames = nil
	}
	program.GlobalPreludeConsts = nil
	program.UserDeclaredGlobals = diffNamesFromSet(program.DeclaredGlobalNames, base.declared)
	program.UserConstGlobals = diffNamesFromSet(program.ConstGlobalNames, base.constNames)
}

func filterNamesByFlagSet(flags map[string]bool, names []string) []string {
	if len(names) == 0 || len(flags) == 0 {
		return nil
	}
	result := make([]string, 0, len(names))
	for i := range names {
		trimmed := strings.TrimSpace(names[i])
		if trimmed == "" {
			continue
		}
		if flags[strings.ToLower(trimmed)] {
			result = append(result, trimmed)
		}
	}
	if len(result) == 0 {
		return nil
	}
	return result
}

func diffNamesFromSet(values []string, baseline map[string]struct{}) []string {
	if len(values) == 0 {
		return nil
	}
	result := make([]string, 0, len(values))
	for i := range values {
		trimmed := strings.TrimSpace(values[i])
		if trimmed == "" {
			continue
		}
		lower := strings.ToLower(trimmed)
		if _, exists := baseline[lower]; exists {
			continue
		}
		result = append(result, trimmed)
	}
	if len(result) == 0 {
		return nil
	}
	return result
}

func inferCachedProgramZeroArgFuncs(program *CachedProgram) []string {
	_ = program
	return nil
}

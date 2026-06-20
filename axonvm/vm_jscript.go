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
	"encoding/base64"
	_ "encoding/binary"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"math"
	"math/big"
	"math/bits"
	"math/rand"
	"net/url"
	"path/filepath"
	"regexp"
	"slices"
	"sort"
	"strconv"
	"strings"
	"sync/atomic"
	"time"
	"unicode/utf16"
	"unicode/utf8"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/jscript"
	"g3pix.com.br/axonasp/jscript/ftoa"
	"g3pix.com.br/axonasp/vbscript"
	"github.com/dlclark/regexp2"
	"golang.org/x/text/unicode/norm"
)

const jsMaxStringBytes = 8 * 1024 * 1024
const jsMaxStringWorkBytes = 2 * 1024 * 1024
const jsMaxCallStackDepth = 10100
const jsInternalPropPrefix = "__js_"
const jsAccessorGetterPrefix = "__js_getter__"
const jsAccessorSetterPrefix = "__js_setter__"
const jsSymbolPropertyPrefix = "__js_sym__"
const jsRestParamPrefix = "__js_rest__:"
const jsClassConstructorFlag = "__axon_internal__:class_constructor"
const jsStrictModeFlag = "__axon_internal__:strict_mode"
const jsGeneratorFlag = "__axon_internal__:generator"
const jsAsyncFlag = "__axon_internal__:async"
const jsDerivedConstructorFlag = "__axon_internal__:derived_constructor"
const jsModuleExportPrefix = "__js_export__:"
const jsHexUpperDigits = "0123456789ABCDEF"

const (
	jsPropertyKindMethod = 0
	jsPropertyKindGet    = 1
	jsPropertyKindSet    = 2
)

var jsWeakCollectionNextID atomic.Uint64

type jsObjectState struct {
	Extensible     bool
	HiddenWeakData map[uint64]Value
}

type jsPropertyDescriptor struct {
	Value        Value
	Getter       Value
	Setter       Value
	HasValue     bool
	HasGetter    bool
	HasSetter    bool
	Enumerable   bool
	Configurable bool
	Writable     bool
}

type jsEnvFrame struct {
	parentID int64
	bindings map[string]Value
}

type jsArgumentsBinding struct {
	envID        int64
	indexToParam map[string]string
	paramToIndex map[string]string
}

type jsRegExpObject struct {
	pattern   string
	flags     string
	compiled  *regexp2.Regexp
	lastIndex int
}

type jsDefinePropertySpec struct {
	desc            jsPropertyDescriptor
	hasEnumerable   bool
	hasConfigurable bool
	hasWritable     bool
}

type jsFunctionObject struct {
	name       string
	params     []string
	restParam  string
	localCount int
	startIP    int
	endIP      int
	envID      int64
	protoID    int64
	isBound    bool
	boundFn    Value
	boundThis  Value
	boundArgs  []Value
	// isArrow marks this as an ES6 arrow function that captures 'this' lexically.
	// When true, capturedThis is used as the receiver regardless of call site.
	isArrow                 bool
	capturedThis            Value
	isClassConstructor      bool
	isStrict                bool
	isDerived               bool
	homeObjID               int64
	isAsync                 bool
	isGenerator             bool
	capturedBlockScopes     []map[string]Value
	capturedBlockScopeConst []map[string]struct{}
	capturedBlockScopeTDZ   []map[string]struct{}
	hiddenWeakData          map[uint64]Value
}

type jsCallFrame struct {
	returnIP             int
	envID                int64
	savedFP              int
	callLine             int
	callColumn           int
	callFile             string
	fn                   Value
	thisVal              Value
	newTarget            Value
	tryDepth             int
	savedSP              int
	isCtor               bool
	ctorObj              Value
	jsStrictMode         bool
	isSuperCall          bool
	savedBlockScopes     []map[string]Value
	savedBlockScopeConst []map[string]struct{}
	savedBlockScopeTDZ   []map[string]struct{}
	savedBlockScopeDepth int
}

type jsForInEnumerator struct {
	keys  []string
	index int
}

// jsForOfEnumerator holds the collected values for a for...of loop.
// Values are collected once when the loop is first entered and then
// consumed one per iteration to avoid mutating the source during traversal.
type jsForOfEnumerator struct {
	values []Value
	index  int
}

type jsProxyObject struct {
	Target  Value
	Handler Value
	Revoked bool
}

type jsPromiseState uint8

const (
	jsPromisePending jsPromiseState = iota
	jsPromiseFulfilled
	jsPromiseRejected
)

type jsPromiseReaction struct {
	onFulfilled Value
	onRejected  Value
	capability  *jsPromiseCapability
}

type jsPromiseCapability struct {
	promise Value
	resolve Value
	reject  Value
}

type jsPromiseObject struct {
	state     jsPromiseState
	result    Value
	reactions []jsPromiseReaction
	handled   bool // true if a rejection handler was attached
}

func (vm *VM) jsModuleExportKey(name string) string {
	return jsModuleExportPrefix + name
}

func (vm *VM) jsSetModuleExport(name string, value Value) {
	vm.ensureJSRootEnv()
	env := vm.jsEnvItems[vm.jsActiveEnvID]
	if env == nil {
		env = &jsEnvFrame{bindings: make(map[string]Value, 8)}
		vm.jsEnvItems[vm.jsActiveEnvID] = env
	}
	env.bindings[vm.jsModuleExportKey(name)] = value
}

func (vm *VM) jsResolveModulePath(specifier string) (string, error) {
	trimmed := strings.TrimSpace(specifier)
	if trimmed == "" {
		return "", fmt.Errorf("empty module specifier")
	}

	resolved := trimmed
	if !filepath.IsAbs(resolved) {
		baseFile := strings.TrimSpace(vm.sourceName)
		if baseFile == "" {
			baseFile = strings.TrimSpace(vm.baseSourceName)
		}
		if baseFile == "" {
			baseFile = "."
		}
		resolved = filepath.Join(filepath.Dir(baseFile), resolved)
	}
	if filepath.Ext(resolved) == "" {
		resolved += ".js"
	}
	absPath, err := filepath.Abs(resolved)
	if err != nil {
		return "", err
	}
	return absPath, nil
}

func (vm *VM) jsImportModule(specifier string) (*jsEnvFrame, bool) {
	modulePath, err := vm.jsResolveModulePath(specifier)
	if err != nil {
		vm.jsThrowReferenceError("Cannot resolve module '" + specifier + "': " + err.Error())
		return nil, false
	}

	if env, ok := vm.jsModuleInstances[modulePath]; ok && env != nil {
		return env, true
	}

	vm.ensureJSRootEnv()
	rootEnvID := vm.jsActiveEnvID

	moduleEnvID := vm.allocJSID()
	moduleEnv := &jsEnvFrame{parentID: rootEnvID, bindings: make(map[string]Value, 16)}
	vm.jsEnvItems[moduleEnvID] = moduleEnv
	vm.jsModuleInstances[modulePath] = moduleEnv
	vm.jsModuleLoading[modulePath] = struct{}{}
	defer delete(vm.jsModuleLoading, modulePath)

	cache := getExecuteScriptCache()
	// Module imports must use the standard compilation path so export/import
	// semantics match server execution in all modes (CLI/TUI/Eval included).
	program, loadErr := cache.LoadOrCompile(modulePath)
	if loadErr != nil {
		delete(vm.jsModuleInstances, modulePath)
		delete(vm.jsEnvItems, moduleEnvID)
		vm.jsThrowReferenceError("Cannot load module '" + modulePath + "': " + loadErr.Error())
		return nil, false
	}

	startIP := vm.appendExecuteProgram(program.GlobalCount, program.Constants, program.Bytecode)
	child := vm.cloneForExecuteGlobal(startIP)
	child.sourceName = modulePath
	child.baseSourceName = modulePath
	child.jsActiveEnvID = moduleEnvID
	if runErr := child.Run(); runErr != nil {
		delete(vm.jsModuleInstances, modulePath)
		delete(vm.jsEnvItems, moduleEnvID)
		vm.jsThrowReferenceError("Error executing module '" + modulePath + "': " + runErr.Error())
		return nil, false
	}

	vm.syncExecuteGlobalState(child)
	if finalEnv, ok := vm.jsEnvItems[moduleEnvID]; ok && finalEnv != nil {
		vm.jsModuleInstances[modulePath] = finalEnv
		return finalEnv, true
	}
	return moduleEnv, true
}

func (vm *VM) jsIsModuleLoading(specifier string) bool {
	modulePath, err := vm.jsResolveModulePath(specifier)
	if err != nil {
		return false
	}
	_, loading := vm.jsModuleLoading[modulePath]
	return loading
}

func (vm *VM) jsGetModuleNamespace(env *jsEnvFrame) Value {
	if env == nil {
		return Value{Type: VTJSUndefined}
	}
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	props := make(map[string]jsPropertyDescriptor, 8)
	for k, v := range env.bindings {
		if !strings.HasPrefix(k, jsModuleExportPrefix) {
			continue
		}
		exportName := k[len(jsModuleExportPrefix):]
		obj[exportName] = v
		props[exportName] = jsPropertyDescriptor{
			Value:        v,
			HasValue:     true,
			Writable:     false,
			Enumerable:   true,
			Configurable: false,
		}
	}
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = props
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsExportAllFromModule(specifier string) {
	moduleEnv, ok := vm.jsImportModule(specifier)
	if !ok {
		return
	}
	for k, v := range moduleEnv.bindings {
		if !strings.HasPrefix(k, jsModuleExportPrefix) {
			continue
		}
		exportName := k[len(jsModuleExportPrefix):]
		if exportName == "default" {
			continue
		}
		vm.jsSetModuleExport(exportName, v)
	}
}

func (vm *VM) jsIsWeakKey(v Value) bool {
	if v.Type == VTJSObject || v.Type == VTJSFunction {
		return true
	}
	if v.Type == VTSymbol {
		if v.Num < 0 {
			return false // Well-known symbols are not allowed as weak keys
		}
		if _, registered := vm.jsRegisteredSymbolIDs[v.Num]; registered {
			return false // Registered symbols (Symbol.for) are not allowed as weak keys
		}
		return true
	}
	return false
}

func (vm *VM) jsWeakSet(key Value, weakID uint64, val Value) {
	if !vm.jsIsWeakKey(key) {
		vm.jsThrowTypeError("Invalid value used as weak map key")
		return
	}
	if key.Type == VTJSFunction {
		fn := vm.jsFunctionItems[key.Num]
		if fn != nil {
			if fn.hiddenWeakData == nil {
				fn.hiddenWeakData = make(map[uint64]Value)
			}
			fn.hiddenWeakData[weakID] = val
		}
		return
	}
	if key.Type == VTSymbol {
		state := vm.jsGetSymbolState(key.Num)
		if state.HiddenWeakData == nil {
			state.HiddenWeakData = make(map[uint64]Value)
		}
		state.HiddenWeakData[weakID] = val
		vm.jsSymbolStateItems[key.Num] = state
		return
	}
	// VTJSObject
	state := vm.jsGetObjectState(key.Num)
	if state.HiddenWeakData == nil {
		state.HiddenWeakData = make(map[uint64]Value)
	}
	state.HiddenWeakData[weakID] = val
	vm.jsObjectStateItems[key.Num] = state
}

func (vm *VM) jsWeakGet(key Value, weakID uint64) Value {
	if !vm.jsIsWeakKey(key) {
		return Value{Type: VTJSUndefined}
	}
	if key.Type == VTJSFunction {
		fn := vm.jsFunctionItems[key.Num]
		if fn != nil && fn.hiddenWeakData != nil {
			if val, ok := fn.hiddenWeakData[weakID]; ok {
				return val
			}
		}
		return Value{Type: VTJSUndefined}
	}
	if key.Type == VTSymbol {
		state, ok := vm.jsSymbolStateItems[key.Num]
		if ok && state.HiddenWeakData != nil {
			if val, ok := state.HiddenWeakData[weakID]; ok {
				return val
			}
		}
		return Value{Type: VTJSUndefined}
	}
	state, ok := vm.jsObjectStateItems[key.Num]
	if ok && state.HiddenWeakData != nil {
		if val, ok := state.HiddenWeakData[weakID]; ok {
			return val
		}
	}
	return Value{Type: VTJSUndefined}
}

func (vm *VM) jsWeakHas(key Value, weakID uint64) bool {
	if !vm.jsIsWeakKey(key) {
		return false
	}
	if key.Type == VTJSFunction {
		fn := vm.jsFunctionItems[key.Num]
		if fn != nil && fn.hiddenWeakData != nil {
			_, exists := fn.hiddenWeakData[weakID]
			return exists
		}
		return false
	}
	if key.Type == VTSymbol {
		state, ok := vm.jsSymbolStateItems[key.Num]
		if ok && state.HiddenWeakData != nil {
			_, exists := state.HiddenWeakData[weakID]
			return exists
		}
		return false
	}
	state, ok := vm.jsObjectStateItems[key.Num]
	if ok && state.HiddenWeakData != nil {
		_, exists := state.HiddenWeakData[weakID]
		return exists
	}
	return false
}

func (vm *VM) jsWeakDelete(key Value, weakID uint64) bool {
	if !vm.jsIsWeakKey(key) {
		return false
	}
	if key.Type == VTJSFunction {
		fn := vm.jsFunctionItems[key.Num]
		if fn != nil && fn.hiddenWeakData != nil {
			_, exists := fn.hiddenWeakData[weakID]
			if exists {
				delete(fn.hiddenWeakData, weakID)
			}
			return exists
		}
		return false
	}
	if key.Type == VTSymbol {
		state, ok := vm.jsSymbolStateItems[key.Num]
		if ok && state.HiddenWeakData != nil {
			_, exists := state.HiddenWeakData[weakID]
			if exists {
				delete(state.HiddenWeakData, weakID)
				vm.jsSymbolStateItems[key.Num] = state
			}
			return exists
		}
		return false
	}
	state, ok := vm.jsObjectStateItems[key.Num]
	if ok && state.HiddenWeakData != nil {
		_, exists := state.HiddenWeakData[weakID]
		if exists {
			delete(state.HiddenWeakData, weakID)
			vm.jsObjectStateItems[key.Num] = state
		}
		return exists
	}
	return false
}

// jsWeakCollectionID validates a WeakMap/WeakSet receiver and returns its backing ID.
func (vm *VM) jsWeakCollectionID(target Value, typeName string, member string) (uint64, bool) {
	if target.Type != VTJSObject {
		vm.jsThrowTypeError(fmt.Sprintf("Method %s.prototype.%s called on incompatible receiver", typeName, member))
		return 0, false
	}
	actualType := vm.jsObjectStringProperty(target, "__js_type")
	if actualType == "" {
		actualType = vm.jsObjectStringProperty(target, "__js_ctor")
	}
	if actualType != typeName {
		vm.jsThrowTypeError(fmt.Sprintf("Method %s.prototype.%s called on incompatible receiver", typeName, member))
		return 0, false
	}
	weakIDVal, deferred := vm.jsMemberGet(target, "__js_weak_id")
	if deferred {
		return 0, false
	}
	switch weakIDVal.Type {
	case VTInteger:
		return uint64(weakIDVal.Num), true
	case VTDouble:
		return uint64(weakIDVal.Flt), true
	default:
		vm.jsThrowTypeError(fmt.Sprintf("Method %s.prototype.%s called on incompatible receiver", typeName, member))
		return 0, false
	}
}

// jsCallWeakCollectionMethod executes WeakMap/WeakSet prototype methods with receiver checks.
func (vm *VM) jsCallWeakCollectionMethod(target Value, typeName string, member string, args []Value) Value {
	weakID, ok := vm.jsWeakCollectionID(target, typeName, member)
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	switch typeName {
	case "WeakMap":
		switch {
		case strings.EqualFold(member, "get"):
			return vm.jsWeakGet(jsArgOrUndefined(args, 0), weakID)
		case strings.EqualFold(member, "set"):
			vm.jsWeakSet(jsArgOrUndefined(args, 0), weakID, jsArgOrUndefined(args, 1))
			return target
		case strings.EqualFold(member, "has"):
			return NewBool(vm.jsWeakHas(jsArgOrUndefined(args, 0), weakID))
		case strings.EqualFold(member, "delete"):
			return NewBool(vm.jsWeakDelete(jsArgOrUndefined(args, 0), weakID))
		}
	case "WeakSet":
		switch {
		case strings.EqualFold(member, "add"):
			vm.jsWeakSet(jsArgOrUndefined(args, 0), weakID, NewBool(true))
			return target
		case strings.EqualFold(member, "has"):
			return NewBool(vm.jsWeakHas(jsArgOrUndefined(args, 0), weakID))
		case strings.EqualFold(member, "delete"):
			return NewBool(vm.jsWeakDelete(jsArgOrUndefined(args, 0), weakID))
		}
	}
	return Value{Type: VTJSUndefined}
}

// jsCallWeakRefMethod executes WeakRef prototype methods.
func (vm *VM) jsCallWeakRefMethod(target Value, member string, args []Value) Value {
	wr, ok := vm.jsWeakRefItems[target.Num]
	if !ok {
		vm.jsThrowTypeError("Method WeakRef.prototype." + member + " called on incompatible receiver")
		return Value{Type: VTJSUndefined}
	}
	if strings.EqualFold(member, "deref") {
		// Target can be Object, Function or Symbol
		if _, exists := vm.jsObjectItems[wr.targetID]; exists {
			return Value{Type: VTJSObject, Num: wr.targetID}
		}
		if _, exists := vm.jsFunctionItems[wr.targetID]; exists {
			return Value{Type: VTJSFunction, Num: wr.targetID}
		}
		if _, exists := vm.jsSymbolStateItems[wr.targetID]; exists {
			return Value{Type: VTSymbol, Num: wr.targetID}
		}
		// Registered symbols are not allowed as weak keys, but if they were used,
		// they are technically always alive. However, our constructor checks this.
		return Value{Type: VTJSUndefined}
	}
	return Value{Type: VTJSUndefined}
}

// jsCallFinalizationRegistryMethod executes FinalizationRegistry prototype methods.
func (vm *VM) jsCallFinalizationRegistryMethod(target Value, member string, args []Value) Value {
	fr, ok := vm.jsFinalizationRegistryItems[target.Num]
	if !ok {
		vm.jsThrowTypeError("Method FinalizationRegistry.prototype." + member + " called on incompatible receiver")
		return Value{Type: VTJSUndefined}
	}
	switch {
	case strings.EqualFold(member, "register"):
		if len(args) < 2 {
			vm.jsThrowTypeError("FinalizationRegistry.prototype.register requires at least 2 arguments")
			return Value{Type: VTJSUndefined}
		}
		targetObj := args[0]
		if targetObj.Type != VTJSObject && targetObj.Type != VTJSFunction && targetObj.Type != VTSymbol {
			vm.jsThrowTypeError("FinalizationRegistry.prototype.register: target must be an object")
			return Value{Type: VTJSUndefined}
		}
		heldValue := args[1]
		unregisterToken := Value{Type: VTJSUndefined}
		if len(args) > 2 {
			unregisterToken = args[2]
		}
		fr.entries = append(fr.entries, jsFinalizationEntry{
			targetID:        targetObj.Num,
			heldValue:       heldValue,
			unregisterToken: unregisterToken,
		})
		return Value{Type: VTJSUndefined}
	case strings.EqualFold(member, "unregister"):
		if len(args) == 0 {
			return NewBool(false)
		}
		token := args[0]
		found := false
		newEntries := make([]jsFinalizationEntry, 0, len(fr.entries))
		for _, entry := range fr.entries {
			if vm.jsStrictEquals(entry.unregisterToken, token) {
				found = true
				continue
			}
			newEntries = append(newEntries, entry)
		}
		fr.entries = newEntries
		return NewBool(found)
	}
	return Value{Type: VTJSUndefined}
}

// jsCallObjectPrototypeMethod executes Object.prototype methods as callable function objects.
func (vm *VM) jsCallObjectPrototypeMethod(target Value, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "hasOwnProperty"):
		return NewBool(vm.jsObjectHasOwnProperty(target, vm.valueToString(jsArgOrUndefined(args, 0))))
	case strings.EqualFold(member, "propertyIsEnumerable"):
		return NewBool(vm.jsObjectPropertyIsEnumerable(target, vm.valueToString(jsArgOrUndefined(args, 0))))
	case strings.EqualFold(member, "isPrototypeOf"):
		return NewBool(vm.jsObjectIsPrototypeOf(target, jsArgOrUndefined(args, 0)))
	case strings.EqualFold(member, "toString"):
		return NewString(vm.jsObjectToStringTag(target))
	case strings.EqualFold(member, "toLocaleString"):
		return NewString(vm.jsObjectToStringTag(target))
	case strings.EqualFold(member, "valueOf"):
		return target
	}
	return Value{Type: VTJSUndefined}
}

// jsCollectionEntry extracts a [key, value] pair from one Map-like iterable entry.
func (vm *VM) jsCollectionEntry(entry Value) (Value, Value, bool) {
	if entry.Type == VTArray {
		if entry.Arr == nil {
			return Value{Type: VTJSUndefined}, Value{Type: VTJSUndefined}, true
		}
		key := Value{Type: VTJSUndefined}
		val := Value{Type: VTJSUndefined}
		if len(entry.Arr.Values) > 0 {
			key = entry.Arr.Values[0]
		}
		if len(entry.Arr.Values) > 1 {
			val = entry.Arr.Values[1]
		}
		return key, val, true
	}
	if entry.Type != VTJSObject {
		vm.jsThrowTypeError("Map iterable value must be an object")
		return Value{Type: VTJSUndefined}, Value{Type: VTJSUndefined}, false
	}
	key, _ := vm.jsArrayLikeGetIndex(entry, 0)
	val, _ := vm.jsArrayLikeGetIndex(entry, 1)
	return key, val, true
}

// jsCollectionStore validates a Set/Map receiver and returns its backing store.
func (vm *VM) jsCollectionStore(target Value, typeName string, member string) (map[string]Value, bool) {
	if target.Type != VTJSObject {
		vm.jsThrowTypeError(fmt.Sprintf("Method %s.prototype.%s called on incompatible receiver", typeName, member))
		return nil, false
	}
	actualType := vm.jsObjectStringProperty(target, "__js_type")
	if actualType == "" {
		actualType = vm.jsObjectStringProperty(target, "__js_ctor")
	}
	if actualType != typeName {
		vm.jsThrowTypeError(fmt.Sprintf("Method %s.prototype.%s called on incompatible receiver", typeName, member))
		return nil, false
	}
	var store map[string]Value
	var ok bool
	if typeName == "Set" {
		store, ok = vm.jsSetItems[target.Num]
	} else {
		store, ok = vm.jsMapItems[target.Num]
	}
	if !ok {
		vm.jsThrowTypeError(fmt.Sprintf("Method %s.prototype.%s called on incompatible receiver", typeName, member))
		return nil, false
	}
	return store, true
}

// jsCallKeyedCollectionMethod executes Set/Map prototype methods with receiver checks.
func (vm *VM) jsCallKeyedCollectionMethod(target Value, typeName string, member string, args []Value) Value {
	store, ok := vm.jsCollectionStore(target, typeName, member)
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	switch typeName {
	case "Set":
		switch {
		case strings.EqualFold(member, "add"):
			arg := jsArgOrUndefined(args, 0)
			store[vm.jsValueMapKey(arg)] = arg
			return target
		case strings.EqualFold(member, "has"):
			_, exists := store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))]
			return NewBool(exists)
		case strings.EqualFold(member, "delete"):
			key := vm.jsValueMapKey(jsArgOrUndefined(args, 0))
			_, exists := store[key]
			if exists {
				delete(store, key)
			}
			return NewBool(exists)
		case strings.EqualFold(member, "clear"):
			clear(store)
			return Value{Type: VTJSUndefined}
		}
	case "Map":
		switch {
		case strings.EqualFold(member, "set"):
			store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))] = jsArgOrUndefined(args, 1)
			return target
		case strings.EqualFold(member, "get"):
			val, exists := store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))]
			if !exists {
				return Value{Type: VTJSUndefined}
			}
			return val
		case strings.EqualFold(member, "has"):
			_, exists := store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))]
			return NewBool(exists)
		case strings.EqualFold(member, "delete"):
			key := vm.jsValueMapKey(jsArgOrUndefined(args, 0))
			_, exists := store[key]
			if exists {
				delete(store, key)
			}
			return NewBool(exists)
		case strings.EqualFold(member, "clear"):
			clear(store)
			return Value{Type: VTJSUndefined}
		}
	}
	return Value{Type: VTJSUndefined}
}

// jsWeakMapEntry extracts a [key, value] pair from one iterable entry.
func (vm *VM) jsWeakMapEntry(entry Value) (Value, Value, bool) {
	if entry.Type != VTArray && entry.Type != VTJSObject {
		vm.jsThrowTypeError("WeakMap iterable value must be an object")
		return Value{Type: VTJSUndefined}, Value{Type: VTJSUndefined}, false
	}
	return vm.jsCollectionEntry(entry)
}

// jsInitSetFromIterable populates a Set from an ES iterable source.
func (vm *VM) jsInitSetFromIterable(source Value, store map[string]Value) bool {
	itObj := vm.jsGetIterator(source)
	if itObj.Type != VTJSObject {
		return false
	}
	for {
		result, handled := vm.jsCallMember(itObj, "next", nil)
		if !handled || result.Type != VTJSObject {
			vm.jsThrowTypeError("Iterator result is not an object")
			return false
		}
		doneVal, _ := vm.jsMemberGet(result, "done")
		if vm.jsTruthy(doneVal) {
			break
		}
		entry, _ := vm.jsMemberGet(result, "value")
		store[vm.jsValueMapKey(entry)] = entry
	}
	return true
}

// jsInitMapFromIterable populates a Map from an ES iterable source.
func (vm *VM) jsInitMapFromIterable(source Value, store map[string]Value) bool {
	itObj := vm.jsGetIterator(source)
	if itObj.Type != VTJSObject {
		return false
	}
	for {
		result, handled := vm.jsCallMember(itObj, "next", nil)
		if !handled || result.Type != VTJSObject {
			vm.jsThrowTypeError("Iterator result is not an object")
			return false
		}
		doneVal, _ := vm.jsMemberGet(result, "done")
		if vm.jsTruthy(doneVal) {
			break
		}
		entry, _ := vm.jsMemberGet(result, "value")
		key, val, ok := vm.jsCollectionEntry(entry)
		if !ok {
			return false
		}
		store[vm.jsValueMapKey(key)] = val
	}
	return true
}

// jsInitWeakMapFromIterable populates a WeakMap from an ES iterable source.
func (vm *VM) jsInitWeakMapFromIterable(source Value, weakID uint64) bool {
	itObj := vm.jsGetIterator(source)
	if itObj.Type != VTJSObject {
		return false
	}
	for {
		result, handled := vm.jsCallMember(itObj, "next", nil)
		if !handled || result.Type != VTJSObject {
			vm.jsThrowTypeError("Iterator result is not an object")
			return false
		}
		doneVal, _ := vm.jsMemberGet(result, "done")
		if vm.jsTruthy(doneVal) {
			break
		}
		entry, _ := vm.jsMemberGet(result, "value")
		key, val, ok := vm.jsWeakMapEntry(entry)
		if !ok {
			return false
		}
		vm.jsWeakSet(key, weakID, val)
	}
	return true
}

// jsInitWeakSetFromIterable populates a WeakSet from an ES iterable source.
func (vm *VM) jsInitWeakSetFromIterable(source Value, weakID uint64) bool {
	itObj := vm.jsGetIterator(source)
	if itObj.Type != VTJSObject {
		return false
	}
	for {
		result, handled := vm.jsCallMember(itObj, "next", nil)
		if !handled || result.Type != VTJSObject {
			vm.jsThrowTypeError("Iterator result is not an object")
			return false
		}
		doneVal, _ := vm.jsMemberGet(result, "done")
		if vm.jsTruthy(doneVal) {
			break
		}
		entry, _ := vm.jsMemberGet(result, "value")
		vm.jsWeakSet(entry, weakID, NewBool(true))
	}
	return true
}

func (vm *VM) jsEval(args []Value) Value {
	if len(args) == 0 {
		return Value{Type: VTJSUndefined}
	}

	if args[0].Type != VTString {
		return args[0]
	}

	expr := strings.TrimSpace(args[0].String())
	expr = strings.TrimLeft(expr, "\uFEFF")
	if expr == "" {
		return Value{Type: VTJSUndefined}
	}

	compiler := NewASPCompiler("")
	compiler.sourceName = vm.sourceName
	vm.jsPrepareDynamicCompilerIC(compiler)
	compiler.compileJScriptEvalSnippet(expr)

	if len(compiler.bytecode) == 0 {
		return Value{Type: VTJSUndefined}
	}

	vm.jsExtendICStateFromCompiler(compiler)

	startIP := vm.appendExecuteProgram(compiler.GlobalsCount(), compiler.constants, compiler.bytecode)
	if startIP < 0 || startIP >= len(vm.bytecode) {
		return Value{Type: VTJSUndefined}
	}

	child := vm.cloneForExecuteLocal(startIP)
	if err := child.Run(); err != nil {
		vm.syncExecuteGlobalState(child)
		if vmErr, ok := err.(*VMError); ok {
			vm.jsThrowJSError(jscript.JSSyntaxErrorCode(vmErr.Code))
			return Value{Type: VTJSUndefined}
		}
		vm.jsThrowTypeError(err.Error())
		return Value{Type: VTJSUndefined}
	}

	resultValue := Value{Type: VTJSUndefined}
	if child.sp >= 0 {
		resultValue = child.stack[child.sp]
	}

	vm.syncExecuteGlobalState(child)
	return resultValue
}

func defaultJSObjectState() jsObjectState {
	return jsObjectState{Extensible: true}
}

func (vm *VM) jsGetObjectState(objID int64) jsObjectState {
	state, ok := vm.jsObjectStateItems[objID]
	if ok {
		return state
	}
	return defaultJSObjectState()
}

func (vm *VM) jsGetSymbolState(symID int64) jsObjectState {
	state, ok := vm.jsSymbolStateItems[symID]
	if ok {
		return state
	}
	return jsObjectState{Extensible: false}
}

func (vm *VM) jsSetObjectExtensible(objID int64, extensible bool) {
	state := vm.jsGetObjectState(objID)
	state.Extensible = extensible
	vm.jsObjectStateItems[objID] = state
}

func (vm *VM) jsTrackObjectKey(objID int64, key string) {
	if strings.HasPrefix(key, jsInternalPropPrefix) {
		return
	}
	order := vm.jsObjectKeyOrder[objID]
	// Check if already exists to maintain insertion order (only add if new)
	if slices.Contains(order, key) {
		return
	}
	vm.jsObjectKeyOrder[objID] = append(order, key)
}

func (vm *VM) jsUntrackObjectKey(objID int64, key string) {
	order, ok := vm.jsObjectKeyOrder[objID]
	if !ok {
		return
	}
	for i, k := range order {
		if k == key {
			vm.jsObjectKeyOrder[objID] = append(order[:i], order[i+1:]...)
			return
		}
	}
}

func (vm *VM) jsObjectOwnPropertyNames(target Value) []string {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return nil
	}
	objID := target.Num

	var indices []int
	var stringsOrder []string
	seen := make(map[string]struct{})

	// 1. Virtual properties for TypedArrays
	if buf, _, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(target); ok {
		length := byteLength / elemSize
		for i := range length {
			indices = append(indices, i)
			seen[strconv.Itoa(i)] = struct{}{}
		}
		_ = buf
	}

	// 2. Tracked order
	if order, ok := vm.jsObjectKeyOrder[objID]; ok {
		for _, key := range order {
			if strings.HasPrefix(key, jsInternalPropPrefix) || strings.HasPrefix(key, jsSymbolPropertyPrefix) {
				continue
			}
			if _, already := seen[key]; already {
				continue
			}
			if idx, ok := jsParseArrayIndex(key); ok {
				indices = append(indices, idx)
			} else {
				stringsOrder = append(stringsOrder, key)
			}
			seen[key] = struct{}{}
		}
	}

	// 3. Fallback for unindexed properties
	collect := func(key string) {
		if strings.HasPrefix(key, jsInternalPropPrefix) || strings.HasPrefix(key, jsSymbolPropertyPrefix) {
			return
		}
		if _, already := seen[key]; already {
			return
		}
		if idx, ok := jsParseArrayIndex(key); ok {
			indices = append(indices, idx)
		} else {
			stringsOrder = append(stringsOrder, key)
		}
		seen[key] = struct{}{}
	}
	if obj, ok := vm.jsObjectItems[objID]; ok {
		for key := range obj {
			collect(key)
		}
	}
	if props, ok := vm.jsPropertyItems[objID]; ok {
		for key := range props {
			collect(key)
		}
	}

	// 4. Sorting
	sort.Ints(indices)
	if order, ok := vm.jsObjectKeyOrder[objID]; !ok || len(order) < len(stringsOrder) {
		sort.Strings(stringsOrder)
	}

	out := make([]string, 0, len(indices)+len(stringsOrder))
	for _, idx := range indices {
		out = append(out, strconv.Itoa(idx))
	}
	for _, s := range stringsOrder {
		out = append(out, s)
	}

	return out
}
func (vm *VM) jsObjectOwnPropertySymbols(target Value) []Value {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return nil
	}
	ids := make(map[int64]struct{}, 8)
	collect := func(key string) {
		if !strings.HasPrefix(key, jsSymbolPropertyPrefix) {
			return
		}
		numText := strings.TrimPrefix(key, jsSymbolPropertyPrefix)
		if id, err := strconv.ParseInt(numText, 10, 64); err == nil {
			ids[id] = struct{}{}
		}
	}
	if obj, ok := vm.jsObjectItems[target.Num]; ok {
		for key := range obj {
			collect(key)
		}
	}
	if props, ok := vm.jsPropertyItems[target.Num]; ok {
		for key := range props {
			collect(key)
		}
	}
	if len(ids) == 0 {
		return nil
	}
	ordered := make([]int64, 0, len(ids))
	for id := range ids {
		ordered = append(ordered, id)
	}
	slices.Sort(ordered)
	out := make([]Value, len(ordered))
	for i := 0; i < len(ordered); i++ {
		out[i] = Value{Type: VTSymbol, Num: ordered[i]}
	}
	return out
}

func (vm *VM) jsObjectIsExtensible(target Value) bool {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return false
	}
	return vm.jsGetObjectState(target.Num).Extensible
}

func (vm *VM) jsObjectSeal(target Value) Value {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return target
	}
	names := vm.jsObjectOwnPropertyNames(target)
	for i := range names {
		desc, ok := vm.jsGetDescriptor(target.Num, names[i])
		if !ok {
			continue
		}
		desc.Configurable = false
		vm.jsSetDescriptor(target.Num, names[i], desc)
	}
	vm.jsSetObjectExtensible(target.Num, false)
	return target
}

func (vm *VM) jsObjectFreeze(target Value) Value {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return target
	}
	names := vm.jsObjectOwnPropertyNames(target)
	for i := range names {
		desc, ok := vm.jsGetDescriptor(target.Num, names[i])
		if !ok {
			continue
		}
		desc.Configurable = false
		if desc.HasValue {
			desc.Writable = false
		}
		vm.jsSetDescriptor(target.Num, names[i], desc)
	}
	vm.jsSetObjectExtensible(target.Num, false)
	return target
}

func (vm *VM) jsObjectIsSealed(target Value) bool {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return false
	}
	if vm.jsObjectIsExtensible(target) {
		return false
	}
	names := vm.jsObjectOwnPropertyNames(target)
	for i := range names {
		desc, ok := vm.jsGetDescriptor(target.Num, names[i])
		if ok && desc.Configurable {
			return false
		}
	}
	return true
}

func (vm *VM) jsObjectIsFrozen(target Value) bool {
	if !vm.jsObjectIsSealed(target) {
		return false
	}
	names := vm.jsObjectOwnPropertyNames(target)
	for i := range names {
		desc, ok := vm.jsGetDescriptor(target.Num, names[i])
		if ok && desc.HasValue && desc.Writable {
			return false
		}
	}
	return true
}

func (vm *VM) jsEnsurePropertyMap(objID int64) map[string]jsPropertyDescriptor {
	props, ok := vm.jsPropertyItems[objID]
	if ok {
		return props
	}
	props = make(map[string]jsPropertyDescriptor, 8)
	vm.jsPropertyItems[objID] = props
	return props
}

func jsDefaultPropertyDescriptor(v Value) jsPropertyDescriptor {
	return jsPropertyDescriptor{Value: v, HasValue: true, Enumerable: true, Configurable: true, Writable: true}
}

func (vm *VM) jsSetProto(target Value, proto Value) {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return
	}
	id := target.Num
	obj := vm.jsObjectItems[id]
	if obj == nil {
		obj = make(map[string]Value, 2)
		vm.jsObjectItems[id] = obj
	}
	obj["__js_proto"] = proto
	vm.jsInvalidateObjectIC(id)
}

// jsInvalidateObjectIC discards one object's slot/layout metadata.
func (vm *VM) jsInvalidateObjectIC(objID int64) {
	delete(vm.jsObjectShape, objID)
	delete(vm.jsObjectSlots, objID)
	delete(vm.jsObjectSlotIndex, objID)
}

func (vm *VM) jsEnsureObjectICLayout(objID int64) bool {
	obj, ok := vm.jsObjectItems[objID]
	if !ok {
		return false
	}
	keys := make([]string, 0, len(obj))
	for key := range obj {
		if strings.HasPrefix(key, jsInternalPropPrefix) {
			continue
		}
		keys = append(keys, key)
	}
	sort.Strings(keys)

	var shapeSig strings.Builder
	for i := 0; i < len(keys); i++ {
		if i > 0 {
			shapeSig.WriteByte('\x1f')
		}
		shapeSig.WriteString(keys[i])
	}

	shapeID := vm.jsShapeBySignature[shapeSig.String()]
	if shapeID == 0 {
		shapeID = vm.jsNextShapeID
		if shapeID == 0 {
			shapeID = 1
		}
		vm.jsNextShapeID = shapeID + 1
		vm.jsShapeBySignature[shapeSig.String()] = shapeID
		if len(keys) > 0 {
			layout := make([]string, len(keys))
			copy(layout, keys)
			vm.jsShapeSlots[shapeID] = layout
		}
	}

	slots := make([]Value, len(keys))
	indexByName := make(map[string]uint16, len(keys))
	for i := 0; i < len(keys); i++ {
		k := keys[i]
		slots[i] = obj[k]
		indexByName[k] = uint16(i)
	}

	vm.jsObjectShape[objID] = shapeID
	vm.jsObjectSlots[objID] = slots
	vm.jsObjectSlotIndex[objID] = indexByName
	return true
}

func (vm *VM) jsResolveICSlot(target Value, member string) (uint32, uint16, bool) {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return 0, 0, false
	}
	if !vm.jsEnsureObjectICLayout(target.Num) {
		return 0, 0, false
	}
	shapeID := vm.jsObjectShape[target.Num]
	if shapeID == 0 {
		return 0, 0, false
	}
	slot, ok := vm.jsObjectSlotIndex[target.Num][member]
	if !ok {
		return 0, 0, false
	}
	desc, hasDesc := vm.jsGetDescriptor(target.Num, member)
	if hasDesc {
		if desc.HasGetter || desc.HasSetter || !desc.HasValue {
			return 0, 0, false
		}
	}
	return shapeID, slot, true
}

func (vm *VM) jsICMemberGet(target Value, member string, shapeID uint32, slot uint16, flags uint16) (Value, bool) {
	if flags == 0 || shapeID == 0 {
		return Value{Type: VTJSUndefined}, false
	}
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}, false
	}
	if vm.jsObjectShape[target.Num] != shapeID {
		return Value{Type: VTJSUndefined}, false
	}
	layout := vm.jsShapeSlots[shapeID]
	if int(slot) >= len(layout) || layout[slot] != member {
		return Value{Type: VTJSUndefined}, false
	}
	slots := vm.jsObjectSlots[target.Num]
	if int(slot) >= len(slots) {
		return Value{Type: VTJSUndefined}, false
	}
	return slots[slot], true
}

func (vm *VM) jsICMemberSet(target Value, member string, val Value, shapeID uint32, slot uint16, flags uint16) bool {
	if flags == 0 || shapeID == 0 {
		return false
	}
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return false
	}
	if vm.jsObjectShape[target.Num] != shapeID {
		return false
	}
	layout := vm.jsShapeSlots[shapeID]
	if int(slot) >= len(layout) || layout[slot] != member {
		return false
	}
	desc, hasDesc := vm.jsGetDescriptor(target.Num, member)
	if hasDesc {
		if desc.HasSetter || !desc.Writable || !desc.HasValue {
			return false
		}
	}
	obj := vm.jsObjectItems[target.Num]
	if obj == nil {
		return false
	}
	obj[member] = val
	slots := vm.jsObjectSlots[target.Num]
	if int(slot) >= len(slots) {
		return false
	}
	slots[slot] = val
	vm.jsObjectSlots[target.Num] = slots
	if hasDesc {
		desc.Value = val
		desc.HasValue = true
		vm.jsPropertyItems[target.Num][member] = desc
	}
	return true
}

func (vm *VM) jsICPopulate(icNodeID uint16, target Value, member string) {
	if int(icNodeID) >= len(vm.icState) {
		return
	}
	shapeID, slot, ok := vm.jsResolveICSlot(target, member)
	if !ok {
		return
	}
	vm.icState[icNodeID] = InlineCacheSlot{
		ShapeID: shapeID,
		Slot:    slot,
		Flags:   1,
	}
}

// jsPrepareDynamicCompilerIC offsets a fresh compiler's IC node ID counter so
// newly assigned ICNodeIDs do not collide with the VM's existing icState entries.
// After compilation, extendICStateFromCompiler must be called to grow icState.
func (vm *VM) jsPrepareDynamicCompilerIC(c *Compiler) {
	if c == nil || vm == nil {
		return
	}
	c.jsNextICNodeID = uint32(len(vm.icState))
}

// jsExtendICStateFromCompiler grows vm.icState to cover IC nodes allocated by a
// dynamic compiler that was prepared with jsPrepareDynamicCompilerIC.
func (vm *VM) jsExtendICStateFromCompiler(c *Compiler) {
	if c == nil || vm == nil {
		return
	}
	end := int(c.jsNextICNodeID)
	if end > len(vm.icState) {
		extended := make([]InlineCacheSlot, end)
		copy(extended, vm.icState)
		vm.icState = extended
	}
}

// jsClassInherit wires a derived class to its superclass and returns the subclass.
func (vm *VM) jsClassInherit(subclass Value, superclass Value) Value {
	if subclass.Type != VTJSFunction && subclass.Type != VTJSObject {
		vm.jsThrowTypeError("Class constructor cannot be used as a base value")
		return Value{Type: VTJSUndefined}
	}

	if superclass.Type == VTNull {
		proto, deferred := vm.jsMemberGet(subclass, "prototype")
		if !deferred && proto.Type == VTJSObject {
			vm.jsSetProto(proto, NewNull())
		}
		return subclass
	}

	if superclass.Type != VTJSFunction && superclass.Type != VTJSObject {
		vm.jsThrowTypeError("Class extends value is not a constructor or null")
		return Value{Type: VTJSUndefined}
	}

	if superclass.Type == VTJSObject {
		if ctorName := vm.jsObjectStringProperty(superclass, "__js_ctor"); ctorName == "" {
			vm.jsThrowTypeError("Class extends value is not a constructor or null")
			return Value{Type: VTJSUndefined}
		}
	}

	vm.jsSetProto(subclass, superclass)

	derivedProto, deferred := vm.jsMemberGet(subclass, "prototype")
	if deferred || derivedProto.Type != VTJSObject {
		vm.jsThrowTypeError("Class prototype is not available")
		return Value{Type: VTJSUndefined}
	}

	superProto, deferred := vm.jsMemberGet(superclass, "prototype")
	if deferred {
		vm.jsThrowTypeError("Class extends value is not a constructor or null")
		return Value{Type: VTJSUndefined}
	}
	if superProto.Type != VTJSObject && superProto.Type != VTNull {
		vm.jsThrowTypeError("Superclass prototype is not an object or null")
		return Value{Type: VTJSUndefined}
	}

	vm.jsSetProto(derivedProto, superProto)
	return subclass
}

func (vm *VM) jsDefineProperty(target Value, name string, kind int, val Value) {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return
	}
	id := target.Num
	props := vm.jsEnsurePropertyMap(id)
	desc, ok := props[name]
	if !ok {
		desc = jsPropertyDescriptor{
			Enumerable:   false,
			Configurable: true,
			Writable:     true,
		}
	}

	switch kind {
	case jsPropertyKindMethod:
		desc.Value = val
		desc.HasValue = true
		desc.Writable = true
	case jsPropertyKindGet:
		desc.Getter = val
		desc.HasGetter = true
		desc.HasValue = false
		desc.Writable = false
	case jsPropertyKindSet:
		desc.Setter = val
		desc.HasSetter = true
		desc.HasValue = false
		desc.Writable = false
	}

	vm.jsSetDescriptor(id, name, desc)
}

func (vm *VM) jsGetDescriptor(objID int64, key string) (jsPropertyDescriptor, bool) {
	if _, isProxy := vm.jsProxyItems[objID]; isProxy {
		return vm.jsProxyGetOwnPropertyDescriptor(Value{Type: VTJSProxy, Num: objID}, key)
	}
	props, ok := vm.jsPropertyItems[objID]
	if ok {
		d, exists := props[key]
		if exists {
			return d, true
		}
	}
	obj, ok := vm.jsObjectItems[objID]
	if !ok {
		return jsPropertyDescriptor{}, false
	}
	v, exists := obj[key]
	if !exists {
		// Support virtual descriptors for TypedArray indices
		if idx, isIdx := jsParseArrayIndex(key); isIdx {
			if val, handled := vm.jsTypedArrayIndexGet(Value{Type: VTJSObject, Num: objID}, idx); handled {
				return jsDefaultPropertyDescriptor(val), true
			}
		}
		return jsPropertyDescriptor{}, false
	}
	return jsDefaultPropertyDescriptor(v), true
}

func (vm *VM) jsSetDescriptor(objID int64, key string, desc jsPropertyDescriptor) {
	props := vm.jsEnsurePropertyMap(objID)
	_, exists := props[key]
	props[key] = desc
	if !exists {
		vm.jsTrackObjectKey(objID, key)
	}
	if desc.HasValue {
		obj, ok := vm.jsObjectItems[objID]
		if !ok {
			obj = make(map[string]Value, 8)
			vm.jsObjectItems[objID] = obj
		}
		obj[key] = desc.Value
	}
	vm.jsInvalidateObjectIC(objID)
}

func (vm *VM) jsCreatePrototypeObject(owner Value) Value {
	protoID := vm.allocJSID()
	vm.jsObjectItems[protoID] = make(map[string]Value, 2)
	vm.jsObjectKeyOrder[protoID] = make([]string, 0, 2)
	vm.jsPropertyItems[protoID] = make(map[string]jsPropertyDescriptor, 2)
	vm.jsSetDescriptor(protoID, "constructor", jsPropertyDescriptor{
		Value:        owner,
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	return Value{Type: VTJSObject, Num: protoID}
}

func jsCtorNeedsPrototype(ctorName string) bool {
	switch ctorName {
	case "Array", "Object", "String", "Date", "RegExp", "Enumerator", "VBArray", "Set", "Map", "WeakMap", "WeakSet", "Promise",
		"Error", "TypeError", "ReferenceError", "SyntaxError", "RangeError", "EvalError", "URIError",
		"WeakRef", "FinalizationRegistry",
		"ArrayBuffer", "DataView",
		"Int8Array", "Uint8Array", "Uint8ClampedArray",
		"Int16Array", "Uint16Array",
		"Int32Array", "Uint32Array",
		"Float32Array", "Float64Array",
		"BigInt64Array", "BigUint64Array":
		return true
	default:
		return false
	}
}

// jsInstanceOf implements the JScript 'instanceof' operator logic.
func (vm *VM) jsInstanceOf(left Value, right Value) bool {
	if right.Type != VTJSObject && right.Type != VTJSFunction && right.Type != VTJSProxy {
		vm.jsThrowTypeError("instanceof: right-hand side is not an object")
		return false
	}

	// 1. Symbol.hasInstance hook
	hasInstanceKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolHasInstance, 10)
	hook, deferred := vm.jsMemberGet(right, hasInstanceKey)
	if !deferred && vm.jsIsCallable(hook) {
		res := vm.jsCall(hook, right, []Value{left})
		return vm.jsTruthy(res)
	}

	// 2. Default logic: check if right is a function
	if !vm.jsIsCallable(right) {
		vm.jsThrowTypeError("instanceof: right-hand side is not callable")
		return false
	}

	// Prototype chain traversal
	protoVal, _ := vm.jsMemberGet(right, "prototype")
	if protoVal.Type != VTJSObject && protoVal.Type != VTJSFunction && protoVal.Type != VTJSProxy && protoVal.Type != VTArray {
		vm.jsThrowTypeError("instanceof: target prototype is not an object")
		return false
	}

	curr := vm.jsGetPrototypeValue(left)
	for curr.Type == VTJSObject || curr.Type == VTJSFunction || curr.Type == VTJSProxy || curr.Type == VTArray {
		if curr.Type == protoVal.Type && curr.Num == protoVal.Num {
			return true
		}
		curr = vm.jsGetPrototypeValue(curr)
	}

	return false
}

func (vm *VM) jsGetPrototypeValue(v Value) Value {
	if v.Type == VTArray {
		return vm.jsGetIntrinsicPrototype("Array")
	}
	if v.Type == VTJSProxy {
		if proxy, ok := vm.jsProxyItems[v.Num]; ok && proxy != nil && !proxy.Revoked {
			return vm.jsGetPrototypeValue(proxy.Target)
		}
		return Value{Type: VTJSUndefined}
	}
	if v.Type == VTJSObject || v.Type == VTJSFunction || v.Type == VTJSProxy {
		id := v.Num
		if obj, ok := vm.jsObjectItems[id]; ok {
			if proto, exists := obj["__js_proto"]; exists {
				return proto
			}
		}
		if props, ok := vm.jsPropertyItems[id]; ok {
			if desc, exists := props["__js_proto"]; exists {
				if desc.HasValue {
					return desc.Value
				}
			}
		}
		// Fallback to intrinsic prototypes if it's a built-in type object
		if obj, ok := vm.jsObjectItems[id]; ok {
			if t, hasT := obj["__js_type"]; hasT {
				// Don't recurse infinitely, only for non-prototype objects
				if !strings.HasSuffix(t.Str, " Prototype") {
					return vm.jsGetIntrinsicPrototype(t.Str)
				}
			}
		}
		if v.Type == VTJSFunction {
			if proto := vm.jsGetIntrinsicPrototype("Function"); proto.Type == VTJSObject {
				return proto
			}
		}
	}
	if v.Type == VTJSFunction {
		if fn, ok := vm.jsFunctionItems[v.Num]; ok && fn != nil && fn.protoID != 0 {
			return Value{Type: VTJSObject, Num: fn.protoID}
		}
	}
	return Value{Type: VTJSUndefined}
}

func (vm *VM) jsResolveObjectMember(objID int64, member string, visited map[int64]struct{}) (jsPropertyDescriptor, bool) {
	if _, seen := visited[objID]; seen {
		return jsPropertyDescriptor{}, false
	}
	visited[objID] = struct{}{}
	if desc, ok := vm.jsGetDescriptor(objID, member); ok {
		return desc, true
	}
	obj, ok := vm.jsObjectItems[objID]
	if !ok {
		return jsPropertyDescriptor{}, false
	}
	proto, exists := obj["__js_proto"]
	if !exists || (proto.Type != VTJSObject && proto.Type != VTJSFunction) {
		return jsPropertyDescriptor{}, false
	}
	return vm.jsResolveObjectMember(proto.Num, member, visited)
}

func (vm *VM) jsGetIntrinsicPrototype(name string) Value {
	vm.ensureJSRootEnv()
	ctor := vm.jsGetName(name)
	if ctor.Type == VTJSObject || ctor.Type == VTJSFunction {
		if value, deferred := vm.jsMemberGet(ctor, "prototype"); !deferred {
			return value
		}
	}
	return Value{Type: VTJSUndefined}
}

func (vm *VM) jsToDescriptorBoolean(v Value, fallback bool) bool {
	if v.Type == VTJSUndefined {
		return fallback
	}
	return vm.jsTruthy(v)
}

func (vm *VM) jsObjectOwnEnumerableKeys(objID int64) []string {
	names := vm.jsObjectOwnPropertyNames(Value{Type: VTJSObject, Num: objID})
	if len(names) == 0 {
		return nil
	}
	keys := make([]string, 0, len(names))
	for i := range names {
		desc, hasDesc := vm.jsGetDescriptor(objID, names[i])
		if hasDesc && !desc.Enumerable {
			continue
		}
		keys = append(keys, names[i])
	}
	return keys
}

func (vm *VM) jsJSONParse(text string) Value {
	var payload any
	if err := json.Unmarshal([]byte(text), &payload); err != nil {
		return Value{Type: VTJSUndefined}
	}
	return vm.jsFromGoJSON(payload)
}

func (vm *VM) jsFromGoJSON(payload any) Value {
	switch v := payload.(type) {
	case nil:
		return NewNull()
	case bool:
		return NewBool(v)
	case float64:
		return NewDouble(v)
	case string:
		return NewString(v)
	case []any:
		values := make([]Value, len(v))
		for i := range v {
			values[i] = vm.jsFromGoJSON(v[i])
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values))
	case map[string]any:
		objID := vm.allocJSID()
		obj := make(map[string]Value, len(v))
		vm.jsObjectKeyOrder[objID] = make([]string, 0, len(v))
		for key, item := range v {
			val := vm.jsFromGoJSON(item)
			obj[key] = val
			vm.jsTrackObjectKey(objID, key)
		}
		vm.jsObjectItems[objID] = obj
		props := vm.jsEnsurePropertyMap(objID)
		for key, item := range obj {
			props[key] = jsDefaultPropertyDescriptor(item)
		}
		return Value{Type: VTJSObject, Num: objID}
	default:
		return Value{Type: VTJSUndefined}
	}
}

// jsJSONStringify converts a value into JSON text honoring ES5 toJSON hooks.
func (vm *VM) jsJSONStringify(v Value) string {
	return vm.jsJSONStringifyValue(v)
}

// jsJSONStringifyValue converts a value into JSON text honoring ES5 toJSON hooks.
func (vm *VM) jsJSONStringifyValue(v Value) string {
	v = vm.jsJSONStringifyApplyToJSON(v)
	switch v.Type {
	case VTJSUndefined:
		return "null"
	case VTNull, VTEmpty:
		return "null"
	case VTBool:
		if v.Num != 0 {
			return "true"
		}
		return "false"
	case VTInteger:
		return strconv.FormatInt(v.Num, 10)
	case VTDouble:
		if math.IsNaN(v.Flt) || math.IsInf(v.Flt, 0) {
			return "null"
		}
		return strconv.FormatFloat(v.Flt, 'f', -1, 64)
	case VTString:
		encoded, _ := json.Marshal(v.Str)
		return string(encoded)
	case VTArray:
		if v.Arr == nil {
			return "[]"
		}
		var b strings.Builder
		b.WriteByte('[')
		for i := 0; i < len(v.Arr.Values); i++ {
			if i > 0 {
				b.WriteByte(',')
			}
			b.WriteString(vm.jsJSONStringifyValue(v.Arr.Values[i]))
		}
		b.WriteByte(']')
		return b.String()
	case VTJSObject:
		obj, ok := vm.jsObjectItems[v.Num]
		if !ok {
			return "{}"
		}
		keys := vm.jsObjectOwnEnumerableKeys(v.Num)
		var b strings.Builder
		b.WriteByte('{')
		for i := range keys {
			if i > 0 {
				b.WriteByte(',')
			}
			k := keys[i]
			encodedKey, _ := json.Marshal(k)
			b.Write(encodedKey)
			b.WriteByte(':')
			b.WriteString(vm.jsJSONStringifyValue(obj[k]))
		}
		b.WriteByte('}')
		return b.String()
	case VTNativeObject:
		// Serialize native Dictionary objects (e.g. produced by G3JSON.LoadFile/Parse)
		// as JSON objects, recursing into nested dictionaries and arrays.
		if dict, ok := vm.dictionaryItems[v.Num]; ok && dict != nil {
			var b strings.Builder
			b.WriteByte('{')
			for i, k := range dict.keys {
				if i > 0 {
					b.WriteByte(',')
				}
				encodedKey, _ := json.Marshal(k.String())
				b.Write(encodedKey)
				b.WriteByte(':')
				b.WriteString(vm.jsJSONStringifyValue(dict.values[i]))
			}
			b.WriteByte('}')
			return b.String()
		}
		// Non-dictionary native: produce a JSON string (best-effort).
		encoded, _ := json.Marshal(vm.valueToString(v))
		return string(encoded)
	default:
		encoded, _ := json.Marshal(vm.valueToString(v))
		return string(encoded)
	}
}

// jsJSONStringifyApplyToJSON invokes object-level toJSON before serialization.
func (vm *VM) jsJSONStringifyApplyToJSON(v Value) Value {
	if v.Type != VTJSObject {
		return v
	}
	toJSON, deferred := vm.jsMemberGet(v, "toJSON")
	if deferred || toJSON.Type != VTJSFunction {
		return v
	}
	result := vm.jsCall(toJSON, v, nil)
	if toJSON.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
		return Value{Type: VTJSUndefined}
	}
	return result
}

func (vm *VM) ensureJSRootEnv() {
	if vm.jsRootEnvID != 0 {
		return
	}
	rootID := vm.allocJSID()
	vm.jsRootEnvID = rootID
	bindings := make(map[string]Value, 40)
	bindings["Intl"] = vm.jsCreateIntlObject()
	bindings["Math"] = vm.jsCreateMathObject()
	bindings["Date"] = vm.jsCreateIntrinsicObject("", "Date")
	bindings["RegExp"] = vm.jsCreateIntrinsicObject("", "RegExp")
	bindings["Enumerator"] = vm.jsCreateIntrinsicObject("", "Enumerator")
	bindings["VBArray"] = vm.jsCreateIntrinsicObject("", "VBArray")
	bindings["String"] = vm.jsCreateIntrinsicObject("", "String")
	bindings["Array"] = vm.jsCreateIntrinsicObject("", "Array")
	bindings["Object"] = vm.jsCreateIntrinsicObject("", "Object")
	bindings["JSON"] = vm.jsCreateIntrinsicObject("", "JSON")
	bindings["Atomics"] = vm.jsCreateAtomicsObject()
	bindings["Proxy"] = vm.jsCreateProxyObject()
	bindings["Reflect"] = vm.jsCreateReflectObject()
	bindings["Promise"] = vm.jsCreatePromiseObject()
	bindings["Number"] = vm.jsCreateNumberObject()
	bindings["Symbol"] = vm.jsCreateSymbolObject()
	bindings["Error"] = vm.jsCreateIntrinsicObject("", "Error")
	bindings["TypeError"] = vm.jsCreateIntrinsicObject("", "TypeError")
	bindings["ReferenceError"] = vm.jsCreateIntrinsicObject("", "ReferenceError")
	bindings["SyntaxError"] = vm.jsCreateIntrinsicObject("", "SyntaxError")
	bindings["RangeError"] = vm.jsCreateIntrinsicObject("", "RangeError")
	bindings["EvalError"] = vm.jsCreateIntrinsicObject("", "EvalError")
	bindings["URIError"] = vm.jsCreateIntrinsicObject("", "URIError")
	bindings["Set"] = vm.jsCreateIntrinsicObject("", "Set")
	bindings["Map"] = vm.jsCreateIntrinsicObject("", "Map")
	bindings["WeakMap"] = vm.jsCreateIntrinsicObject("", "WeakMap")
	bindings["WeakSet"] = vm.jsCreateIntrinsicObject("", "WeakSet")
	bindings["WeakRef"] = vm.jsCreateIntrinsicObject("", "WeakRef")
	bindings["FinalizationRegistry"] = vm.jsCreateIntrinsicObject("", "FinalizationRegistry")
	// ES6 Binary Data constructors
	bindings["ArrayBuffer"] = vm.jsCreateIntrinsicObject("ArrayBuffer", "ArrayBuffer")
	bindings["SharedArrayBuffer"] = vm.jsCreateIntrinsicObject("SharedArrayBuffer", "SharedArrayBuffer")
	bindings["DataView"] = vm.jsCreateIntrinsicObject("DataView", "DataView")
	bindings["Int8Array"] = vm.jsCreateIntrinsicObject("", "Int8Array")
	bindings["Uint8Array"] = vm.jsCreateIntrinsicObject("", "Uint8Array")
	bindings["Uint8ClampedArray"] = vm.jsCreateIntrinsicObject("", "Uint8ClampedArray")
	bindings["Int16Array"] = vm.jsCreateIntrinsicObject("", "Int16Array")
	bindings["Uint16Array"] = vm.jsCreateIntrinsicObject("", "Uint16Array")
	bindings["Int32Array"] = vm.jsCreateIntrinsicObject("", "Int32Array")
	bindings["Uint32Array"] = vm.jsCreateIntrinsicObject("", "Uint32Array")
	bindings["Float32Array"] = vm.jsCreateIntrinsicObject("", "Float32Array")
	bindings["Float64Array"] = vm.jsCreateIntrinsicObject("", "Float64Array")
	bindings["BigInt64Array"] = vm.jsCreateIntrinsicObject("", "BigInt64Array")
	bindings["BigUint64Array"] = vm.jsCreateIntrinsicObject("", "BigUint64Array")
	bindings["NaN"] = NewDouble(math.NaN())
	bindings["Infinity"] = NewDouble(math.Inf(1))
	bindings["undefined"] = Value{Type: VTJSUndefined}
	bindings["isNaN"] = vm.jsCreateIntrinsicObject("", "isNaN")
	bindings["isFinite"] = vm.jsCreateIntrinsicObject("", "isFinite")
	bindings["parseInt"] = vm.jsCreateIntrinsicObject("", "parseInt")
	bindings["parseFloat"] = vm.jsCreateIntrinsicObject("", "parseFloat")
	bindings["decodeURI"] = vm.jsCreateIntrinsicObject("", "decodeURI")
	bindings["decodeURIComponent"] = vm.jsCreateIntrinsicObject("", "decodeURIComponent")
	bindings["encodeURI"] = vm.jsCreateIntrinsicObject("", "encodeURI")
	bindings["encodeURIComponent"] = vm.jsCreateIntrinsicObject("", "encodeURIComponent")
	if evalIdx, ok := GetBuiltinIndex("Eval"); ok {
		bindings["eval"] = Value{Type: VTBuiltin, Num: int64(evalIdx)}
	}

	// Add Node.js API compatibility globals if enabled via config
	if vm.enableNodeCompatibility() {
		bindings["require"] = vm.jsCreateIntrinsicObject("", "require")
		bindings["process"] = vm.jsCreateProcessObject()
		bindings["Buffer"] = vm.jsCreateBufferConstructor()
		bindings["path"] = vm.jsCreatePathObject()
		bindings["os"] = vm.jsCreateOSObject()
		bindings["fs"] = vm.jsCreateFSObject()
		bindings["crypto"] = vm.jsCreateCryptoObject()
		bindings["http"] = vm.jsCreateHTTPObject("http")
		bindings["https"] = vm.jsCreateHTTPObject("https")
		bindings["querystring"] = vm.jsCreateQueryStringObject()
		urlCtor := vm.jsCreateURLConstructor()
		urlSearchParamsCtor := vm.jsCreateURLSearchParamsConstructor()
		bindings["URL"] = urlCtor
		bindings["URLSearchParams"] = urlSearchParamsCtor
		bindings["url"] = vm.jsCreateURLModuleObject(urlCtor, urlSearchParamsCtor)
		bindings["__axon_stream"] = vm.jsCreateNodeStreamHooksObject()
		// Phase 2: Timing globals
		bindings["setTimeout"] = vm.jsCreateIntrinsicFunction("setTimeout", "SetTimeout")
		bindings["clearTimeout"] = vm.jsCreateIntrinsicFunction("clearTimeout", "ClearTimeout")
		bindings["setInterval"] = vm.jsCreateIntrinsicFunction("setInterval", "SetInterval")
		bindings["clearInterval"] = vm.jsCreateIntrinsicFunction("clearInterval", "ClearInterval")
		bindings["setImmediate"] = vm.jsCreateIntrinsicFunction("setImmediate", "SetImmediate")
		bindings["clearImmediate"] = vm.jsCreateIntrinsicFunction("clearImmediate", "ClearImmediate")

		// Node.js module globals
		if vm.sourceName != "" {
			absPath, _ := filepath.Abs(vm.sourceName)
			bindings["__filename"] = NewString(absPath)
			bindings["__dirname"] = NewString(filepath.Dir(absPath))
		} else {
			bindings["__filename"] = NewString("")
			bindings["__dirname"] = NewString("")
		}
		bindings["global"] = Value{Type: VTJSObject, Num: rootID}
		bindings["globalThis"] = Value{Type: VTJSObject, Num: rootID}
	}

	// Create the root environment frame
	vm.jsEnvItems[rootID] = &jsEnvFrame{parentID: 0, bindings: bindings}
	if vm.jsActiveEnvID == 0 {
		vm.jsActiveEnvID = rootID
	}
	vm.jsThisValue = Value{Type: VTJSUndefined}
	vm.jsPopulatePrototypes(bindings)

	// Add global and globalThis aliases pointing to the root object
	// These must be added to the environment after jsEnvItems is populated
	rootObj := Value{Type: VTJSObject, Num: rootID}
	vm.jsEnvItems[rootID].bindings["global"] = rootObj
	vm.jsEnvItems[rootID].bindings["globalThis"] = rootObj
}

// enableNodeCompatibility checks the viper configuration to determine if Node.js
// API compatibility mode should be enabled for the JScript engine.
func (vm *VM) enableNodeCompatibility() bool {
	// Load the config if available
	cfg := axonconfig.NewViper()
	if cfg == nil {
		return false
	}

	// Check the javascript.enable_node_compatibility setting
	return cfg.GetBool("javascript.enable_node_compatibility")
}

// jsCreateMathObject allocates the global Math object with immutable constants.
func (vm *VM) jsCreateMathObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 10)
	obj["__js_type"] = NewString("Math")
	obj["E"] = NewDouble(math.E)
	obj["PI"] = NewDouble(math.Pi)
	obj["LN2"] = NewDouble(math.Ln2)
	obj["LN10"] = NewDouble(math.Ln10)
	obj["LOG2E"] = NewDouble(math.Log2E)
	obj["LOG10E"] = NewDouble(math.Log10E)
	obj["SQRT2"] = NewDouble(math.Sqrt2)
	obj["SQRT1_2"] = NewDouble(1 / math.Sqrt2)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 10)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateReflectObject allocates the global Reflect namespace object.
func (vm *VM) jsCreateReflectObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("Reflect")
	vm.jsObjectItems[objID] = obj
	props := make(map[string]jsPropertyDescriptor, 8)

	createReflectMethod := func(name string, ctor string, length int) {
		methodID := vm.allocJSID()
		methodObj := make(map[string]Value, 4)
		methodObj["__js_type"] = NewString("Function")
		methodObj["__js_ctor"] = NewString(ctor)
		methodObj["name"] = NewString(name)
		methodObj["length"] = NewInteger(int64(length))
		vm.jsObjectItems[methodID] = methodObj
		vm.jsPropertyItems[methodID] = make(map[string]jsPropertyDescriptor, 4)
		obj[name] = Value{Type: VTJSFunction, Num: methodID}
	}

	createReflectMethod("get", "ReflectGet", 2)
	createReflectMethod("set", "ReflectSet", 3)
	createReflectMethod("apply", "ReflectApply", 3)
	createReflectMethod("construct", "ReflectConstruct", 2)
	createReflectMethod("has", "ReflectHas", 2)
	createReflectMethod("deleteProperty", "ReflectDeleteProperty", 2)
	createReflectMethod("ownKeys", "ReflectOwnKeys", 1)
	createReflectMethod("defineProperty", "ReflectDefineProperty", 3)
	createReflectMethod("getOwnPropertyDescriptor", "ReflectGetOwnPropertyDescriptor", 2)
	createReflectMethod("getPrototypeOf", "ReflectGetPrototypeOf", 1)
	createReflectMethod("isExtensible", "ReflectIsExtensible", 1)
	createReflectMethod("preventExtensions", "ReflectPreventExtensions", 1)
	createReflectMethod("setPrototypeOf", "ReflectSetPrototypeOf", 2)

	vm.jsPropertyItems[objID] = props
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateProxyObject allocates the global Proxy constructor object with static methods.
func (vm *VM) jsCreateProxyObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("Proxy")
	obj["name"] = NewString("Proxy")
	obj["length"] = NewInteger(2)
	vm.jsObjectItems[objID] = obj
	props := make(map[string]jsPropertyDescriptor, 4)

	// Proxy.revocable static method
	revocableID := vm.allocJSID()
	revocableObj := make(map[string]Value, 4)
	revocableObj["__js_type"] = NewString("Function")
	revocableObj["__js_ctor"] = NewString("ProxyRevocable")
	revocableObj["name"] = NewString("revocable")
	revocableObj["length"] = NewInteger(2)
	vm.jsObjectItems[revocableID] = revocableObj
	vm.jsPropertyItems[revocableID] = make(map[string]jsPropertyDescriptor, 4)
	obj["revocable"] = Value{Type: VTJSFunction, Num: revocableID}

	vm.jsPropertyItems[objID] = props
	return Value{Type: VTJSFunction, Num: objID}
}

// jsCreateNumberObject allocates the global Number constructor object with ES6 static methods.
func (vm *VM) jsCreateNumberObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("Number")
	obj["__js_ctor"] = NewString("Number")
	// ES6 static constants
	obj["MAX_SAFE_INTEGER"] = NewInteger(9007199254740991)
	obj["MIN_SAFE_INTEGER"] = NewInteger(-9007199254740991)
	obj["MAX_VALUE"] = NewDouble(math.MaxFloat64)
	obj["MIN_VALUE"] = NewDouble(5e-324)
	obj["POSITIVE_INFINITY"] = NewDouble(math.Inf(1))
	obj["NEGATIVE_INFINITY"] = NewDouble(math.Inf(-1))
	obj["NaN"] = NewDouble(math.NaN())
	obj["EPSILON"] = NewDouble(2.220446049250313e-16)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 10)
	vm.jsSetDescriptor(objID, "MAX_SAFE_INTEGER", jsPropertyDescriptor{Value: obj["MAX_SAFE_INTEGER"], HasValue: true, Enumerable: false, Configurable: false, Writable: false})
	vm.jsSetDescriptor(objID, "MIN_SAFE_INTEGER", jsPropertyDescriptor{Value: obj["MIN_SAFE_INTEGER"], HasValue: true, Enumerable: false, Configurable: false, Writable: false})
	vm.jsSetDescriptor(objID, "EPSILON", jsPropertyDescriptor{Value: obj["EPSILON"], HasValue: true, Enumerable: false, Configurable: false, Writable: false})
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateSymbolObject allocates the global Symbol constructor object and attaches
// well-known symbol properties (Symbol.iterator, Symbol.toStringTag, etc.).
func (vm *VM) jsCreateSymbolObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 10)
	obj["__js_type"] = NewString("Symbol")
	obj["__js_ctor"] = NewString("Symbol")
	// Well-known symbols as non-writable, non-enumerable, non-configurable properties
	obj["iterator"] = jsWellKnownSymbolValue(jsWellKnownSymbolIterator, "Symbol.iterator")
	obj["toStringTag"] = jsWellKnownSymbolValue(jsWellKnownSymbolToStringTag, "Symbol.toStringTag")
	obj["species"] = jsWellKnownSymbolValue(jsWellKnownSymbolSpecies, "Symbol.species")
	obj["hasInstance"] = jsWellKnownSymbolValue(jsWellKnownSymbolHasInstance, "Symbol.hasInstance")
	obj["toPrimitive"] = jsWellKnownSymbolValue(jsWellKnownSymbolToPrimitive, "Symbol.toPrimitive")
	obj["dispose"] = jsWellKnownSymbolValue(jsWellKnownSymbolDispose, "Symbol.dispose")
	obj["asyncDispose"] = jsWellKnownSymbolValue(jsWellKnownSymbolAsyncDispose, "Symbol.asyncDispose")
	obj["unscopables"] = jsWellKnownSymbolValue(jsWellKnownSymbolUnscopables, "Symbol.unscopables")
	obj["matchAll"] = jsWellKnownSymbolValue(jsWellKnownSymbolMatchAll, "Symbol.matchAll")
	obj["isConcatSpreadable"] = jsWellKnownSymbolValue(jsWellKnownSymbolIsConcatSpreadable, "Symbol.isConcatSpreadable")
	vm.jsObjectItems[objID] = obj
	props := make(map[string]jsPropertyDescriptor, 8)
	// Make well-known symbols read-only, non-enumerable, non-configurable
	for _, name := range []string{"iterator", "toStringTag", "species", "hasInstance", "toPrimitive", "dispose", "asyncDispose", "unscopables", "matchAll", "isConcatSpreadable"} {
		props[name] = jsPropertyDescriptor{
			Value: obj[name], HasValue: true,
			Enumerable: false, Configurable: false, Writable: false,
		}
	}
	vm.jsPropertyItems[objID] = props
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsCreatePromiseObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("Promise")
	obj["__js_ctor"] = NewString("Promise")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)

	// Promise.prototype
	proto := vm.jsCreatePrototypeObject(Value{Type: VTJSObject, Num: objID})
	vm.jsSetDescriptor(objID, "prototype", jsPropertyDescriptor{
		Value: proto, HasValue: true, Enumerable: false, Configurable: false, Writable: false,
	})

	// Static methods: Promise.resolve, Promise.reject, Promise.all, Promise.race, Promise.allSettled, Promise.any, Promise.withResolvers
	for _, name := range []string{"resolve", "reject", "all", "race", "allSettled", "any", "withResolvers"} {
		vm.jsSetDescriptor(objID, name, jsPropertyDescriptor{
			Value:        vm.jsCreateIntrinsicFunction("Promise."+name, "PromiseStatic"+strings.Title(name)),
			HasValue:     true,
			Enumerable:   false,
			Configurable: true,
			Writable:     true,
		})
	}

	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsCreateIntrinsicFunction(name string, ctorName string) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString(ctorName)
	obj["name"] = NewString(name)
	vm.jsObjectItems[objID] = obj
	return Value{Type: VTJSFunction, Num: objID}
}

func (vm *VM) jsCreateIntrinsicObject(typeName string, ctorName string) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 2)
	if typeName != "" {
		obj["__js_type"] = NewString(typeName)
	}
	if ctorName != "" {
		obj["__js_ctor"] = NewString(ctorName)
	}
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	if jsCtorNeedsPrototype(ctorName) {
		ctorVal := Value{Type: VTJSObject, Num: objID}
		proto := vm.jsCreatePrototypeObject(ctorVal)
		vm.jsSetDescriptor(objID, "prototype", jsPropertyDescriptor{
			Value:        proto,
			HasValue:     true,
			Enumerable:   false,
			Configurable: false,
			Writable:     false,
		})
	}
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsObjectStringProperty(obj Value, key string) string {
	if obj.Type != VTJSObject && obj.Type != VTJSFunction {
		return ""
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return ""
	}
	v, ok := items[key]
	if !ok || v.Type != VTString {
		return ""
	}
	return v.Str
}

func (vm *VM) allocJSID() int64 {
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	if vm.jsObjectKeyOrder == nil {
		vm.jsObjectKeyOrder = make(map[int64][]string)
	}
	vm.jsObjectKeyOrder[id] = make([]string, 0, 8)
	return id
}

func (vm *VM) jsAllocSymbolID() int64 {
	id := vm.jsNextSymbolID
	vm.jsNextSymbolID++
	if vm.jsNextSymbolID <= 0 {
		vm.jsNextSymbolID = 1
	}
	return id
}

func (vm *VM) jsValueMapKey(v Value) string {
	v = resolveCallable(vm, v)
	switch v.Type {
	case VTJSUndefined:
		return "u"
	case VTNull, VTEmpty:
		return "n"
	case VTBool:
		return "b:" + strconv.FormatInt(v.Num, 10)
	case VTInteger:
		f := float64(v.Num)
		if f == 0 {
			return "num:0"
		}
		return "num:" + strconv.FormatFloat(f, 'g', -1, 64)
	case VTDouble:
		if math.IsNaN(v.Flt) {
			return "num:nan"
		}
		if v.Flt == 0 {
			return "num:0"
		}
		return "num:" + strconv.FormatFloat(v.Flt, 'g', -1, 64)
	case VTString:
		return "s:" + v.Str
	case VTJSBigInt:
		if v.Big == nil {
			return "big:0"
		}
		return "big:" + v.Big.String()
	case VTSymbol:
		return "sym:" + strconv.FormatInt(v.Num, 10)
	case VTArray:
		if v.Arr == nil {
			return "arr:nil"
		}
		return "arr:" + fmt.Sprintf("%p", v.Arr)
	case VTJSObject:
		return "jso:" + strconv.FormatInt(v.Num, 10)
	case VTJSFunction:
		return "jsf:" + strconv.FormatInt(v.Num, 10)
	case VTNativeObject:
		return "nat:" + strconv.FormatInt(v.Num, 10)
	case VTObject:
		return "obj:" + strconv.FormatInt(v.Num, 10)
	case VTDate:
		return "date:" + strconv.FormatInt(v.Num, 10)
	default:
		return "t:" + strconv.FormatInt(int64(v.Type), 10) + ":" + v.String()
	}
}

// jsMapKeyToValue converts an internal Map backing-store key back into a JS value when possible.
func (vm *VM) jsMapKeyToValue(key string) Value {
	if key == "u" {
		return Value{Type: VTJSUndefined}
	}
	if key == "n" {
		return Value{Type: VTNull}
	}
	if strings.HasPrefix(key, "b:") {
		return NewBool(key == "b:1")
	}
	if after, ok := strings.CutPrefix(key, "num:"); ok {
		numText := after
		if numText == "nan" {
			return NewDouble(math.NaN())
		}
		if n, err := strconv.ParseInt(numText, 10, 64); err == nil {
			return NewInteger(n)
		}
		if f, err := strconv.ParseFloat(numText, 64); err == nil {
			return NewDouble(f)
		}
	}
	if after, ok := strings.CutPrefix(key, "s:"); ok {
		return NewString(after)
	}
	if after, ok := strings.CutPrefix(key, "big:"); ok {
		text := after
		if bi, ok := new(big.Int).SetString(text, 10); ok {
			return Value{Type: VTJSBigInt, Big: bi}
		}
	}
	if after, ok := strings.CutPrefix(key, "sym:"); ok {
		if n, err := strconv.ParseInt(after, 10, 64); err == nil {
			return Value{Type: VTSymbol, Num: n}
		}
	}
	return NewString(key)
}

func (vm *VM) jsPropertyKeyFromValue(v Value) string {
	v = resolveCallable(vm, v)
	if v.Type == VTSymbol {
		return jsSymbolPropertyPrefix + strconv.FormatInt(v.Num, 10)
	}
	return vm.valueToString(v)
}

type jsWeakRef struct {
	targetID int64
}

type jsFinalizationRegistry struct {
	callback Value
	entries  []jsFinalizationEntry
}

type jsFinalizationEntry struct {
	targetID        int64
	heldValue       Value
	unregisterToken Value
}

// jsCleanupCollections clears Set/Map backing stores between top-level runs to
// avoid stale growth in long-lived server processes.
func (vm *VM) jsCleanupCollections() {
	for id, setItems := range vm.jsSetItems {
		clear(setItems)
		delete(vm.jsSetItems, id)
	}
	for id, mapItems := range vm.jsMapItems {
		clear(mapItems)
		delete(vm.jsMapItems, id)
	}
	for id := range vm.jsWeakRefItems {
		delete(vm.jsWeakRefItems, id)
	}
	for id, registry := range vm.jsFinalizationRegistryItems {
		clear(registry.entries)
		delete(vm.jsFinalizationRegistryItems, id)
	}
}

func (vm *VM) jsCurrentEnv() *jsEnvFrame {
	vm.ensureJSRootEnv()
	return vm.jsEnvItems[vm.jsActiveEnvID]
}

func (vm *VM) jsDeclareName(name string) {
	env := vm.jsCurrentEnv()
	if env == nil {
		return
	}
	if _, ok := env.bindings[name]; ok {
		return
	}
	env.bindings[name] = Value{Type: VTJSUndefined}
}

// jsGetNameFromEnv reads a variable value from the env frame chain starting at envID.
func (vm *VM) jsGetNameFromEnv(envID int64, name string) Value {
	for id := envID; id != 0; {
		env := vm.jsEnvItems[id]
		if env == nil {
			break
		}
		if val, ok := env.bindings[name]; ok {
			if val.Type == VTArgRef {
				val = vm.stack[int(val.Num)]
			}
			return val
		}
		id = env.parentID
	}
	return Value{Type: VTJSUndefined}
}

// jsSetNameInEnv writes a variable value into the env frame chain starting at envID.
// It searches the chain from envID upward for the first frame that has the binding.
func (vm *VM) jsSetNameInEnv(envID int64, name string, val Value) {
	for id := envID; id != 0; {
		env := vm.jsEnvItems[id]
		if env == nil {
			break
		}
		if _, ok := env.bindings[name]; ok {
			env.bindings[name] = val
			return
		}
		id = env.parentID
	}
	// Not found: declare in the given env frame
	if env, ok := vm.jsEnvItems[envID]; ok {
		env.bindings[name] = val
	}
}

// jsGetBlockScopeValue reads from block scopes (innermost first). Returns the value and
// whether it was found. Raises ReferenceError if the variable is in TDZ (const before init).
func (vm *VM) jsGetBlockScopeValue(name string) (Value, bool) {
	for i := len(vm.jsBlockScopes) - 1; i >= 0; i-- {
		if _, exists := vm.jsBlockScopes[i][name]; exists {
			// Check TDZ for const variables (access before initialization).
			if _, inTDZ := vm.jsBlockScopeTDZ[i][name]; inTDZ {
				vm.jsThrowReferenceError(fmt.Sprintf("Cannot access '%s' before initialization", name))
				return Value{Type: VTJSUndefined}, true
			}
			return vm.jsBlockScopes[i][name], true
		}
	}
	return Value{Type: VTJSUndefined}, false
}

// jsSetBlockScopeValue writes into the innermost block scope that declares the name.
// Returns true if found and set. Raises TypeError for const reassignment.
// Raises ReferenceError if const is in TDZ (accessed before initialization).
func (vm *VM) jsSetBlockScopeValue(name string, val Value) bool {
	for i := len(vm.jsBlockScopes) - 1; i >= 0; i-- {
		if _, exists := vm.jsBlockScopes[i][name]; exists {
			// Check TDZ for const before initialization.
			if _, inTDZ := vm.jsBlockScopeTDZ[i][name]; inTDZ {
				vm.jsThrowReferenceError(fmt.Sprintf("Cannot access '%s' before initialization", name))
				return true
			}
			// Check const immutability: once initialized, const cannot be reassigned.
			if _, isConst := vm.jsBlockScopeConst[i][name]; isConst {
				vm.jsThrowTypeError(fmt.Sprintf("Assignment to constant variable '%s'", name))
				return true
			}
			vm.jsBlockScopes[i][name] = val
			return true
		}
	}
	return false
}

func (vm *VM) jsSetName(name string, val Value) {
	vm.ensureJSRootEnv()
	if vm.jsAssignWithBinding(name, val) {
		return
	}
	// Block scopes take precedence for let/const bindings
	if vm.jsBlockScopeDepth > 0 {
		if vm.jsSetBlockScopeValue(name, val) {
			return
		}
	}
	for envID := vm.jsActiveEnvID; envID != 0; {
		env := vm.jsEnvItems[envID]
		if env == nil {
			break
		}
		if _, ok := env.bindings[name]; ok {
			env.bindings[name] = val
			vm.jsSyncArgumentAliasByParam(envID, name, val)
			return
		}
		envID = env.parentID
	}
	if idx, ok := vm.lookupJSGlobalIndex(name); ok {
		vm.Globals[idx] = val
		return
	}

	// In strict mode, assigning to an undeclared variable is a ReferenceError
	if vm.jsStrictMode {
		vm.jsThrowReferenceError(fmt.Sprintf("%s is not defined", name))
		return
	}

	// Non-strict mode: create variable in root/current environment
	root := vm.jsEnvItems[vm.jsActiveEnvID]
	if root != nil {
		root.bindings[name] = val
		vm.jsSyncArgumentAliasByParam(vm.jsActiveEnvID, name, val)
	}
}

func (vm *VM) jsGetName(name string) Value {
	vm.ensureJSRootEnv()
	if strings.EqualFold(name, "this") {
		if vm.jsThisValue.Type == VTJSUninitialized {
			vm.jsThrowReferenceError("jsGetName: Must call super constructor in derived class before accessing 'this'")
			return Value{Type: VTJSUndefined}
		}
		return vm.jsThisValue
	}
	if strings.EqualFold(name, "eval") {
		if idx, ok := GetBuiltinIndex("Eval"); ok {
			return Value{Type: VTBuiltin, Num: int64(idx)}
		}
	}
	// Block scopes take precedence (innermost let/const declarations)
	if vm.jsBlockScopeDepth > 0 {
		if val, found := vm.jsGetBlockScopeValue(name); found {
			return val
		}
	}
	if value, ok := vm.jsResolveWithBinding(name); ok {
		return value
	}
	for envID := vm.jsActiveEnvID; envID != 0; {
		env := vm.jsEnvItems[envID]
		if env == nil {
			break
		}
		if val, ok := env.bindings[name]; ok {
			if val.Type == VTArgRef {
				val = vm.stack[int(val.Num)]
			}
			return val
		}
		envID = env.parentID
	}
	if vm.jsRootEnvID != 0 {
		if root, ok := vm.jsEnvItems[vm.jsRootEnvID]; ok {
			if val, ok := root.bindings[name]; ok {
				if val.Type == VTArgRef {
					val = vm.stack[int(val.Num)]
				}
				return val
			}
		}
	}
	if idx, ok := vm.lookupJSGlobalIndex(name); ok {
		return vm.Globals[idx]
	}
	return Value{Type: VTJSUndefined}
}

// jsIsUnscopable checks if a property name is marked as unscopable on the target object.
func (vm *VM) jsIsUnscopable(target Value, name string) bool {
	if target.Type != VTJSObject && target.Type != VTArray && target.Type != VTJSProxy {
		return false
	}
	unscopablesKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolUnscopables, 10)
	unscopables, deferred := vm.jsMemberGet(target, unscopablesKey)
	if deferred || (unscopables.Type != VTJSObject && unscopables.Type != VTArray && unscopables.Type != VTJSProxy) {
		return false
	}
	val, _ := vm.jsMemberGet(unscopables, name)
	return vm.asBool(val)
}

func (vm *VM) jsResolveWithBinding(name string) (Value, bool) {
	if len(vm.withStack) == 0 {
		return Value{Type: VTJSUndefined}, false
	}
	for i := len(vm.withStack) - 1; i >= 0; i-- {
		target := vm.withStack[i]
		if !vm.jsHasProperty(target, name) {
			continue
		}
		if vm.jsIsUnscopable(target, name) {
			continue
		}
		value, deferred := vm.jsMemberGet(target, name)
		if deferred {
			return Value{Type: VTJSUndefined}, false
		}
		return value, true
	}
	return Value{Type: VTJSUndefined}, false
}

func (vm *VM) jsAssignWithBinding(name string, val Value) bool {
	if len(vm.withStack) == 0 {
		return false
	}
	for i := len(vm.withStack) - 1; i >= 0; i-- {
		target := vm.withStack[i]
		if !vm.jsHasProperty(target, name) {
			continue
		}
		if vm.jsIsUnscopable(target, name) {
			continue
		}
		vm.jsMemberSet(target, name, val)
		return true
	}
	return false
}

func (vm *VM) jsHasProperty(target Value, name string) bool {
	switch target.Type {
	case VTJSObject, VTJSFunction:
		// Handle global/globalThis object - access properties from root environment bindings
		if target.Num == vm.jsRootEnvID && vm.jsRootEnvID != 0 {
			if rootEnv, ok := vm.jsEnvItems[vm.jsRootEnvID]; ok {
				if _, exists := rootEnv.bindings[name]; exists {
					return true
				}
			}
		}
		_, ok := vm.jsResolveObjectMember(target.Num, name, make(map[int64]struct{}, 4))
		if ok {
			return true
		}
		if obj, hasObj := vm.jsObjectItems[target.Num]; hasObj {
			if _, exists := obj[name]; exists {
				return true
			}
		}
		if obj, hasObj := vm.jsObjectItems[target.Num]; hasObj {
			_, ok = obj[name]
			return ok
		}
		return false
	case VTArray:
		if strings.EqualFold(name, "length") {
			return true
		}
		if target.Arr == nil {
			return false
		}
		idx, err := strconv.Atoi(name)
		if err != nil {
			return false
		}
		adjusted := idx - target.Arr.Lower
		return adjusted >= 0 && adjusted < len(target.Arr.Values)
	case VTString:
		if strings.EqualFold(name, "length") {
			return true
		}
		idx, err := strconv.Atoi(name)
		if err != nil {
			return false
		}
		runes := []rune(target.Str)
		return idx >= 0 && idx < len(runes)
	default:
		return false
	}
}

func (vm *VM) lookupJSGlobalIndex(name string) (int, bool) {
	lowerName := strings.ToLower(name)
	switch lowerName {
	case "response":
		if len(vm.Globals) > 0 {
			return 0, true
		}
	case "request":
		if len(vm.Globals) > 1 {
			return 1, true
		}
	case "server":
		if len(vm.Globals) > 2 {
			return 2, true
		}
	case "session":
		if len(vm.Globals) > 3 {
			return 3, true
		}
	case "application":
		if len(vm.Globals) > 4 {
			return 4, true
		}
	case "objectcontext":
		if len(vm.Globals) > 5 {
			return 5, true
		}
	case "err":
		if len(vm.Globals) > 6 {
			return 6, true
		}
	case "console":
		if len(vm.Globals) > 7 {
			return 7, true
		}
	}

	for i := 0; i < len(vm.globalNames); i++ {
		if vm.globalNames[i] == name {
			return i, true
		}
	}
	if idx, ok := vm.globalNameIndex[lowerName]; ok {
		return idx, true
	}

	// Some execution paths construct a VM without compiler scope metadata
	// (globalNames/globalNameIndex). In that case, resolve builtins by scanning
	// globals for the matching builtin index instead of relying on fixed slots.
	if builtinIdx, ok := BuiltinIndex[lowerName]; ok {
		for i := 0; i < len(vm.Globals); i++ {
			if vm.Globals[i].Type == VTBuiltin && int(vm.Globals[i].Num) == builtinIdx {
				return i, true
			}
		}
	}

	return 0, false
}

func (vm *VM) jsTruthy(v Value) bool {
	switch v.Type {
	case VTJSUndefined, VTNull, VTEmpty:
		return false
	case VTBool:
		return v.Num != 0
	case VTInteger:
		return v.Num != 0
	case VTDouble:
		if math.IsNaN(v.Flt) {
			return false
		}
		return v.Flt != 0
	case VTString:
		return v.Str != ""
	case VTJSBigInt:
		return v.Big != nil && v.Big.Sign() != 0
	default:
		return true
	}
}

func (vm *VM) jsStrictEquals(a Value, b Value) bool {
	if a.Type != b.Type {
		// In JScript, integers and doubles are both "number" type
		if (a.Type == VTInteger || a.Type == VTDouble) && (b.Type == VTInteger || b.Type == VTDouble) {
			return vm.jsToNumber(a).Flt == vm.jsToNumber(b).Flt
		}
		return false
	}
	switch a.Type {
	case VTJSUndefined, VTNull:
		return true
	case VTBool, VTInteger, VTDate, VTNativeObject, VTJSObject, VTJSFunction, VTSymbol:
		return a.Num == b.Num
	case VTDouble:
		return a.Flt == b.Flt
	case VTString:
		return a.Str == b.Str
	case VTJSBigInt:
		return a.Big.Cmp(b.Big) == 0
	default:
		return a.String() == b.String()
	}
}

func (vm *VM) jsIsCallable(v Value) bool {
	switch v.Type {
	case VTJSFunction, VTBuiltin, VTUserSub:
		return true
	case VTJSObject:
		return vm.jsObjectStringProperty(v, "__js_ctor") != ""
	case VTNativeObject:
		return true // Most native objects in our VM are callable (methods/properties)
	case VTJSProxy:
		if proxy, ok := vm.jsProxyItems[v.Num]; ok && !proxy.Revoked {
			return vm.jsIsCallable(proxy.Target)
		}
	}
	return false
}

func (vm *VM) jsIsConstructor(v Value) bool {
	switch v.Type {
	case VTJSFunction:
		if closure, ok := vm.jsFunctionItems[v.Num]; ok && closure != nil {
			return !closure.isGenerator && !closure.isAsync && !closure.isArrow
		}
		return true // Native JScript function (constructor)
	case VTJSObject:
		ctorName := vm.jsObjectStringProperty(v, "__js_ctor")
		return ctorName != "" && ctorName != "Symbol" && ctorName != "isNaN" && ctorName != "isFinite" &&
			ctorName != "parseInt" && ctorName != "parseFloat" && ctorName != "decodeURI" &&
			ctorName != "decodeURIComponent" && ctorName != "encodeURI" && ctorName != "encodeURIComponent"
	case VTJSProxy:
		if proxy, ok := vm.jsProxyItems[v.Num]; ok && !proxy.Revoked {
			return vm.jsIsConstructor(proxy.Target)
		}
	}
	return false
}

func (vm *VM) jsTypeOf(v Value) string {
	switch v.Type {
	case VTJSUndefined:
		return "undefined"
	case VTNull:
		return "object"
	case VTBool:
		return "boolean"
	case VTInteger, VTDouble:
		return "number"
	case VTJSBigInt:
		return "bigint"
	case VTString:
		return "string"
	case VTSymbol:
		return "symbol"
	case VTJSFunction:
		return "function"
	case VTJSProxy:
		if vm.jsIsCallable(v) {
			return "function"
		}
		return "object"
	case VTJSObject:
		if vm.jsIsCallable(v) {
			return "function"
		}
		return "object"
	case VTNativeObject, VTObject, VTArray, VTJSPromise, VTJSGenerator:
		return "object"
	default:
		return "undefined"
	}
}

// jsAddIntegersNoOverflow adds two int64 values and reports if the result is representable as int64.
func jsAddIntegersNoOverflow(a int64, b int64) (int64, bool) {
	sum := a + b
	if (b > 0 && sum < a) || (b < 0 && sum > a) {
		return 0, false
	}
	return sum, true
}

// jsSubtractIntegersNoOverflow subtracts two int64 values and reports if the result is representable as int64.
func jsSubtractIntegersNoOverflow(a int64, b int64) (int64, bool) {
	diff := a - b
	if (b < 0 && diff < a) || (b > 0 && diff > a) {
		return 0, false
	}
	return diff, true
}

// jsIncrementNumberValue increments a numeric value preserving VTInteger whenever possible.
func (vm *VM) jsIncrementNumberValue(v Value) Value {
	if v.Type == VTInteger {
		if next, ok := jsAddIntegersNoOverflow(v.Num, 1); ok {
			return NewInteger(next)
		}
		return NewDouble(float64(v.Num) + 1)
	}
	next := vm.jsToNumber(v)
	next.Flt++
	return next
}

// jsDecrementNumberValue decrements a numeric value preserving VTInteger whenever possible.
func (vm *VM) jsDecrementNumberValue(v Value) Value {
	if v.Type == VTInteger {
		if next, ok := jsSubtractIntegersNoOverflow(v.Num, 1); ok {
			return NewInteger(next)
		}
		return NewDouble(float64(v.Num) - 1)
	}
	next := vm.jsToNumber(v)
	next.Flt--
	return next
}

// jsAddValues implements JScript '+' behavior for string concatenation and numeric addition.
func (vm *VM) jsAddValues(a Value, b Value) Value {
	a = resolveCallable(vm, a)
	b = resolveCallable(vm, b)
	if a.Type == VTString || b.Type == VTString {
		sa := vm.jsConcatString(a)
		sb := vm.jsConcatString(b)
		total := len(sa) + len(sb)
		if !vm.jsEnsureStringSize(total) || !vm.jsChargeStringWork(total) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(sa + sb)
	}
	if a.Type == VTInteger && b.Type == VTInteger {
		if sum, ok := jsAddIntegersNoOverflow(a.Num, b.Num); ok {
			return NewInteger(sum)
		}
		return NewDouble(float64(a.Num) + float64(b.Num))
	}
	return NewDouble(vm.jsToNumber(a).Flt + vm.jsToNumber(b).Flt)
}

// jsConcatString converts one value to the JScript string form used by '+' concatenation.
func (vm *VM) jsConcatString(v Value) string {
	v = resolveCallable(vm, v)
	if arr, ok := vm.jsAsConcatArray(v); ok {
		return vm.jsArrayToString(arr)
	}
	if v.Type == VTJSObject || v.Type == VTJSFunction || v.Type == VTJSProxy {
		return vm.jsToString(v)
	}
	return vm.valueToString(v)
}

// jsSetSpeciesGetter attaches the standard Symbol.species getter to a constructor.
func (vm *VM) jsSetSpeciesGetter(ctor Value) {
	if ctor.Type != VTJSObject && ctor.Type != VTJSFunction {
		return
	}
	speciesKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolSpecies, 10)
	getter := vm.jsCreateNativeFunction("get [Symbol.species]", "SpeciesGetter")
	vm.jsSetDescriptor(ctor.Num, speciesKey, jsPropertyDescriptor{
		Getter:       getter,
		HasGetter:    true,
		Enumerable:   false,
		Configurable: true,
	})
}

// jsArraySpeciesCreate creates a new array object using the species constructor if present.
func (vm *VM) jsArraySpeciesCreate(target Value, length int) Value {
	ctor := vm.jsGetSpeciesConstructor(target, "Array")
	// If it's the standard Array constructor, we can use our optimized VTArray.
	if ctor.Type == VTJSFunction {
		if vm.jsObjectStringProperty(ctor, "__js_ctor") == "Array" {
			return ValueFromVBArray(NewVBArrayFromValues(0, make([]Value, length)))
		}
	}
	// Fallback to calling the constructor
	return vm.jsNew(ctor, []Value{NewInteger(int64(length))})
}

// jsGetSpeciesConstructor implements the Symbol.species lookup for derived objects.
func (vm *VM) jsGetSpeciesConstructor(target Value, defaultConstructor string) Value {
	ctorVal, _ := vm.jsMemberGet(target, "constructor")
	if ctorVal.Type != VTJSObject && ctorVal.Type != VTJSFunction && ctorVal.Type != VTJSProxy {
		return vm.jsGetName(defaultConstructor)
	}
	speciesKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolSpecies, 10)
	species, deferred := vm.jsMemberGet(ctorVal, speciesKey)
	if !deferred && species.Type != VTJSUndefined && species.Type != VTNull {
		return species
	}
	return vm.jsGetName(defaultConstructor)
}

// jsIsConcatSpreadable determines if a value should be flattened by Array.prototype.concat.
func (vm *VM) jsIsConcatSpreadable(v Value) bool {
	if v.Type != VTJSObject && v.Type != VTJSFunction && v.Type != VTArray && v.Type != VTJSProxy {
		return false
	}
	spreadableKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolIsConcatSpreadable, 10)
	spreadable, deferred := vm.jsMemberGet(v, spreadableKey)
	if !deferred && spreadable.Type != VTJSUndefined {
		return vm.jsTruthy(spreadable)
	}
	// Default: only arrays are spreadable
	if v.Type == VTArray {
		return true
	}
	if v.Type == VTJSObject {
		if _, ok := vm.jsObjectItems[v.Num]["__js_vbarray_source"]; ok {
			return true
		}
		return vm.jsObjectStringProperty(v, "__js_type") == "Array"
	}
	return false
}

// jsAsConcatArray resolves supported array-like bridge values to a VTArray source.
func (vm *VM) jsAsConcatArray(v Value) (Value, bool) {
	if v.Type == VTArray && v.Arr != nil {
		return v, true
	}
	if v.Type != VTJSObject {
		return Value{Type: VTJSUndefined}, false
	}
	items, ok := vm.jsObjectItems[v.Num]
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	vbSource, ok := items["__js_vbarray_source"]
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	converted := vm.jsVBArrayToJSArray(vbSource)
	if converted.Type != VTArray || converted.Arr == nil {
		return Value{Type: VTJSUndefined}, false
	}
	return converted, true
}

// jsArrayToString returns the ES3/ES5 Array toString()-style comma-joined representation.
func (vm *VM) jsArrayToString(v Value) string {
	if v.Type != VTArray || v.Arr == nil || len(v.Arr.Values) == 0 {
		return ""
	}
	parts := make([]string, len(v.Arr.Values))
	totalSize := 0
	for i := 0; i < len(v.Arr.Values); i++ {
		item := v.Arr.Values[i]
		if item.Type == VTJSUndefined || item.Type == VTNull {
			parts[i] = ""
		} else {
			parts[i] = vm.jsConcatString(item)
		}
		totalSize += len(parts[i])
		if i > 0 {
			totalSize++
		}
		if !vm.jsEnsureStringSize(totalSize) {
			return ""
		}
	}
	if !vm.jsChargeStringWork(totalSize) {
		return ""
	}
	return strings.Join(parts, ",")
}

// jsEnsureStringSize guards JScript string-producing operations against runaway growth.
func (vm *VM) jsEnsureStringSize(size int) bool {
	if size <= jsMaxStringBytes {
		return true
	}
	vm.raise(vbscript.OutOfStringSpace, fmt.Sprintf("JScript string size exceeded %d bytes", jsMaxStringBytes))
	return false
}

// jsChargeStringWork tracks cumulative JScript string output work per Run() to stop pathological growth loops.
func (vm *VM) jsChargeStringWork(size int) bool {
	if size <= 0 {
		return true
	}
	vm.jsStringWorkBytes += int64(size)
	if vm.jsStringWorkBytes <= jsMaxStringWorkBytes {
		return true
	}
	vm.raise(vbscript.OutOfStringSpace, fmt.Sprintf("JScript cumulative string work exceeded %d bytes", jsMaxStringWorkBytes))
	return false
}

func (vm *VM) jsCreateClosure(template Value) Value {
	if template.Type != VTJSFunctionTemplate && template.Type != VTJSArrowFunctionTemplate {
		return Value{Type: VTJSUndefined}
	}
	id := vm.allocJSID()
	fnVal := Value{Type: VTJSFunction, Num: id}
	proto := vm.jsCreatePrototypeObject(fnVal)
	params := make([]string, 0, len(template.Names))
	restParam := ""
	localCount := 0
	isClassConstructor := false
	isStrict := false
	isDerived := false
	isGenerator := false
	isAsync := false
	for i := 0; i < len(template.Names); i++ {
		name := template.Names[i]
		if after, ok := strings.CutPrefix(name, "__js_local_count__:"); ok {
			if n, err := strconv.Atoi(after); err == nil && n > 0 {
				localCount = n
			}
			continue
		}
		if after, ok := strings.CutPrefix(name, jsRestParamPrefix); ok {
			restParam = after
			continue
		}
		if name == jsClassConstructorFlag {
			isClassConstructor = true
			continue
		}
		if name == jsStrictModeFlag {
			isStrict = true
			continue
		}
		if name == jsGeneratorFlag {
			isGenerator = true
			continue
		}
		if name == jsAsyncFlag {
			isAsync = true
			continue
		}
		if name == jsDerivedConstructorFlag {
			isDerived = true
			continue
		}
		params = append(params, name)
	}
	fnObj := &jsFunctionObject{
		name:               template.Str,
		params:             params,
		restParam:          restParam,
		localCount:         localCount,
		startIP:            int(template.Num),
		endIP:              int(template.Flt),
		envID:              vm.jsActiveEnvID,
		protoID:            proto.Num,
		isClassConstructor: isClassConstructor,
		isStrict:           isStrict,
		isDerived:          isDerived,
		isAsync:            isAsync,
		isGenerator:        isGenerator,
	}
	if vm.jsBlockScopeDepth > 0 {
		activeDepth := min(vm.jsBlockScopeDepth, len(vm.jsBlockScopes))
		if activeDepth > len(vm.jsBlockScopeConst) {
			activeDepth = len(vm.jsBlockScopeConst)
		}
		if activeDepth > len(vm.jsBlockScopeTDZ) {
			activeDepth = len(vm.jsBlockScopeTDZ)
		}
		if activeDepth > 0 {
			fnObj.capturedBlockScopes = append(make([]map[string]Value, 0, activeDepth), vm.jsBlockScopes[:activeDepth]...)
			fnObj.capturedBlockScopeConst = append(make([]map[string]struct{}, 0, activeDepth), vm.jsBlockScopeConst[:activeDepth]...)
			fnObj.capturedBlockScopeTDZ = append(make([]map[string]struct{}, 0, activeDepth), vm.jsBlockScopeTDZ[:activeDepth]...)
		}
	}
	// Arrow functions capture the current 'this' value lexically at creation time.
	if template.Type == VTJSArrowFunctionTemplate {
		fnObj.isArrow = true
		fnObj.capturedThis = vm.jsThisValue
	}
	vm.jsFunctionItems[id] = fnObj
	vm.jsObjectItems[id] = make(map[string]Value, 2)
	vm.jsPropertyItems[id] = make(map[string]jsPropertyDescriptor, 2)
	vm.jsSetDescriptor(id, "prototype", jsPropertyDescriptor{
		Value:        proto,
		HasValue:     true,
		Enumerable:   false,
		Configurable: false,
		Writable:     true,
	})
	return fnVal
}

func (vm *VM) jsPrepareLocalFrame(localCount int, savedSP int) bool {
	if localCount <= 0 {
		vm.fp = savedSP + 1
		vm.sp = savedSP
		return true
	}
	base := savedSP + 1
	end := base + localCount - 1
	if end >= len(vm.stack) {
		vm.jsThrowOutOfStackSpace()
		return false
	}
	for i := base; i <= end; i++ {
		vm.stack[i] = Value{Type: VTJSUndefined}
	}
	vm.fp = base
	vm.sp = end
	return true
}

// jsThrowOutOfStackSpace throws the canonical JScript stack overflow runtime error.
func (vm *VM) jsThrowOutOfStackSpace() {
	vm.jsThrowJSError(jscript.OutOfStackSpace)
}

func (vm *VM) jsBeginFunctionCall(fn Value, thisVal Value, args []Value, ctorObj Value, isCtor bool, newTarget Value, isSuperCall bool) bool {
	// Debug:
	// fmt.Printf("jsBeginFunctionCall: fn=%v, isCtor=%v, isSuper=%v\n", fn, isCtor, isSuperCall)
	closure, ok := vm.jsFunctionItems[fn.Num]
	if !ok || closure == nil {
		return false
	}
	// Arrow functions ignore the caller-supplied 'this' and use the lexically captured one.
	if closure.isArrow {
		thisVal = closure.capturedThis
	}
	if len(vm.jsCallStack) >= jsMaxCallStackDepth {
		vm.jsThrowJSError(jscript.OutOfStackSpace)
		return false
	}
	frame := jsCallFrame{
		returnIP:             vm.ip,
		envID:                vm.jsActiveEnvID,
		savedFP:              vm.fp,
		callLine:             vm.lastLine,
		callColumn:           vm.lastColumn,
		callFile:             vm.sourceName,
		fn:                   fn,
		thisVal:              vm.jsThisValue,
		newTarget:            vm.jsNewTarget,
		tryDepth:             len(vm.jsTryStack),
		savedSP:              vm.sp,
		isCtor:               isCtor,
		ctorObj:              ctorObj,
		jsStrictMode:         vm.jsStrictMode,
		isSuperCall:          isSuperCall,
		savedBlockScopes:     append(make([]map[string]Value, 0, len(vm.jsBlockScopes)), vm.jsBlockScopes...),
		savedBlockScopeConst: append(make([]map[string]struct{}, 0, len(vm.jsBlockScopeConst)), vm.jsBlockScopeConst...),
		savedBlockScopeTDZ:   append(make([]map[string]struct{}, 0, len(vm.jsBlockScopeTDZ)), vm.jsBlockScopeTDZ...),
		savedBlockScopeDepth: vm.jsBlockScopeDepth,
	}
	vm.jsCallStack = append(vm.jsCallStack, frame)
	vm.jsStrictMode = closure.isStrict
	vm.jsNewTarget = newTarget
	vm.jsBlockScopes = append(make([]map[string]Value, 0, len(closure.capturedBlockScopes)), closure.capturedBlockScopes...)
	vm.jsBlockScopeConst = append(make([]map[string]struct{}, 0, len(closure.capturedBlockScopeConst)), closure.capturedBlockScopeConst...)
	vm.jsBlockScopeTDZ = append(make([]map[string]struct{}, 0, len(closure.capturedBlockScopeTDZ)), closure.capturedBlockScopeTDZ...)
	vm.jsBlockScopeDepth = len(vm.jsBlockScopes)
	envID := vm.allocJSID()
	bindings := make(map[string]Value, len(closure.params)+2)
	for i := 0; i < len(closure.params); i++ {
		if i < len(args) {
			bindings[closure.params[i]] = args[i]
		} else {
			bindings[closure.params[i]] = Value{Type: VTJSUndefined}
		}
	}
	if closure.restParam != "" {
		restValues := make([]Value, 0)
		if len(args) > len(closure.params) {
			restValues = append(restValues, args[len(closure.params):]...)
		}
		bindings[closure.restParam] = ValueFromVBArray(NewVBArrayFromValues(0, restValues))
	}
	if _, hasArguments := bindings["arguments"]; !hasArguments {
		argumentsObject := vm.jsCreateArgumentsObject(args, closure.params, envID)
		bindings["arguments"] = argumentsObject
	}
	vm.jsEnvItems[envID] = &jsEnvFrame{parentID: closure.envID, bindings: bindings}
	vm.jsActiveEnvID = envID
	vm.jsThisValue = thisVal
	if !vm.jsPrepareLocalFrame(closure.localCount, frame.savedSP) {
		vm.jsCallStack = vm.jsCallStack[:len(vm.jsCallStack)-1]
		vm.jsActiveEnvID = frame.envID
		vm.jsThisValue = frame.thisVal
		vm.jsNewTarget = frame.newTarget
		vm.jsStrictMode = frame.jsStrictMode
		vm.jsBlockScopes = frame.savedBlockScopes
		vm.jsBlockScopeConst = frame.savedBlockScopeConst
		vm.jsBlockScopeTDZ = frame.savedBlockScopeTDZ
		vm.jsBlockScopeDepth = frame.savedBlockScopeDepth
		vm.fp = frame.savedFP
		vm.sp = frame.savedSP
		return false
	}
	vm.ip = closure.startIP
	return true
}

// jsBeginDirectCall starts one user-defined JScript function call on the active VM without cloning.
func (vm *VM) jsBeginDirectCall(callee Value, thisVal Value, args []Value) bool {
	if callee.Type != VTJSFunction {
		return false
	}
	closure, ok := vm.jsFunctionItems[callee.Num]
	if !ok || closure == nil {
		return false
	}
	if closure.isClassConstructor {
		vm.jsThrowTypeError("Class constructor cannot be invoked without 'new'")
		return false
	}
	if closure.isGenerator {
		return false
	}
	if closure.isAsync {
		return false
	}
	if closure.isBound {
		return false
	}
	if len(vm.jsCallStack) >= jsMaxCallStackDepth {
		vm.jsThrowJSError(jscript.OutOfStackSpace)
		return false
	}
	return vm.jsBeginFunctionCall(callee, thisVal, args, Value{Type: VTJSUndefined}, false, Value{Type: VTJSUndefined}, false)
}

// jsEnsureDirectCallHaltIP returns one VM-local OpHalt trampoline address used
// to stop nested direct callback execution immediately after function return.
func (vm *VM) jsEnsureDirectCallHaltIP() int {
	if vm.jsDirectCallHaltIP >= 0 && vm.jsDirectCallHaltIP < len(vm.bytecode) {
		if OpCode(vm.bytecode[vm.jsDirectCallHaltIP]) == OpHalt {
			return vm.jsDirectCallHaltIP
		}
	}
	vm.jsDirectCallHaltIP = len(vm.bytecode)
	vm.bytecode = append(vm.bytecode, byte(OpHalt))
	return vm.jsDirectCallHaltIP
}

// jsCallDirectNoClone executes one regular user-defined JScript function call
// on the active VM, without cloneForExecuteLocal.
func (vm *VM) jsCallDirectNoClone(callee Value, thisVal Value, args []Value) (Value, bool) {
	if callee.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}, false
	}
	closure, ok := vm.jsFunctionItems[callee.Num]
	if !ok || closure == nil {
		return Value{Type: VTJSUndefined}, false
	}
	if closure.startIP < 0 || closure.endIP <= closure.startIP || closure.endIP > len(vm.bytecode) {
		return Value{Type: VTJSUndefined}, false
	}
	for i := closure.startIP; i < closure.endIP; i++ {
		if OpCode(vm.bytecode[i]) == OpJSThrow {
			return Value{Type: VTJSUndefined}, false
		}
	}

	savedIP := vm.ip
	savedTryStack := append(make([]int, 0, len(vm.jsTryStack)), vm.jsTryStack...)
	savedErrStack := append(make([]Value, 0, len(vm.jsErrStack)), vm.jsErrStack...)
	if !vm.jsBeginDirectCall(callee, thisVal, args) {
		return Value{Type: VTJSUndefined}, false
	}
	frameIdx := len(vm.jsCallStack) - 1
	if frameIdx < 0 {
		vm.ip = savedIP
		return Value{Type: VTJSUndefined}, false
	}
	vm.jsCallStack[frameIdx].tryDepth = 0
	vm.jsTryStack = vm.jsTryStack[:0]
	vm.jsErrStack = vm.jsErrStack[:0]
	frameSavedSP := vm.jsCallStack[frameIdx].savedSP
	vm.jsCallStack[frameIdx].returnIP = vm.jsEnsureDirectCallHaltIP()

	var runErr error
	var runThrow *jsAsyncRejectionError
	func() {
		defer func() {
			if r := recover(); r != nil {
				if are, ok := r.(*jsAsyncRejectionError); ok {
					runThrow = are
					return
				}
				panic(r)
			}
		}()
		runErr = vm.Run()
	}()
	vm.ip = savedIP
	vm.jsTryStack = savedTryStack
	vm.jsErrStack = savedErrStack

	if runThrow != nil {
		vm.jsThrow(runThrow.reason)
		return Value{Type: VTJSUndefined}, true
	}
	if runErr != nil {
		if vmErr, ok := runErr.(*VMError); ok {
			vm.jsThrowJSError(jscript.JSSyntaxErrorCode(vmErr.Code))
			return Value{Type: VTJSUndefined}, true
		}
		vm.jsThrowTypeError(runErr.Error())
		return Value{Type: VTJSUndefined}, true
	}

	if vm.sp > frameSavedSP {
		return vm.pop(), true
	}
	return Value{Type: VTJSUndefined}, true
}

// jsEnvHasCapturedClosures reports whether any closure currently captures envID.
func (vm *VM) jsEnvHasCapturedClosures(envID int64) bool {
	if envID == 0 {
		return false
	}
	for _, fn := range vm.jsFunctionItems {
		if fn != nil && fn.envID == envID {
			return true
		}
	}
	return false
}

// jsReleaseEnvFrame drops one non-captured JScript env frame and its transient arguments object.
func (vm *VM) jsReleaseEnvFrame(envID int64) {
	if envID == 0 || envID == vm.jsRootEnvID {
		return
	}
	if vm.jsEnvHasCapturedClosures(envID) {
		return
	}
	env, ok := vm.jsEnvItems[envID]
	if !ok {
		return
	}
	if env != nil && env.bindings != nil {
		if argsObj, hasArgs := env.bindings["arguments"]; hasArgs && argsObj.Type == VTJSObject {
			delete(vm.jsArgumentsItems, argsObj.Num)
			delete(vm.jsObjectItems, argsObj.Num)
			delete(vm.jsPropertyItems, argsObj.Num)
			delete(vm.jsObjectStateItems, argsObj.Num)
		}
		clear(env.bindings)
	}
	delete(vm.jsEnvItems, envID)
}

// jsRefreshArgumentsObject rewrites one existing arguments object in place.
func (vm *VM) jsRefreshArgumentsObject(objID int64, args []Value, params []string, envID int64) {
	obj, ok := vm.jsObjectItems[objID]
	if !ok {
		obj = make(map[string]Value, len(args)+1)
		vm.jsObjectItems[objID] = obj
	} else {
		clear(obj)
	}
	for i := range args {
		obj[strconv.Itoa(i)] = args[i]
	}
	obj["length"] = NewInteger(int64(len(args)))

	if len(params) > 0 && len(args) > 0 {
		alias, ok := vm.jsArgumentsItems[objID]
		if !ok || alias == nil {
			alias = &jsArgumentsBinding{}
			vm.jsArgumentsItems[objID] = alias
		}
		alias.envID = envID
		if alias.indexToParam == nil {
			alias.indexToParam = make(map[string]string, len(params))
		} else {
			clear(alias.indexToParam)
		}
		if alias.paramToIndex == nil {
			alias.paramToIndex = make(map[string]string, len(params))
		} else {
			clear(alias.paramToIndex)
		}
		max := min(len(args), len(params))
		for i := 0; i < max; i++ {
			key := strconv.Itoa(i)
			paramName := params[i]
			alias.indexToParam[key] = paramName
			if _, exists := alias.paramToIndex[paramName]; !exists {
				alias.paramToIndex[paramName] = key
			}
		}
	} else {
		delete(vm.jsArgumentsItems, objID)
	}
}

// jsTailCallValue replaces the current JScript call frame with one new call target.
func (vm *VM) jsTailCallValue(callee Value, thisVal Value, args []Value) bool {
	if callee.Type != VTJSFunction || len(vm.jsCallStack) == 0 {
		return false
	}

	closure, ok := vm.jsFunctionItems[callee.Num]
	if !ok || closure == nil || closure.isBound {
		return false
	}

	if closure.isArrow {
		thisVal = closure.capturedThis
	}

	frame := &vm.jsCallStack[len(vm.jsCallStack)-1]

	canReuseEnv := vm.jsActiveEnvID != 0 && !vm.jsEnvHasCapturedClosures(vm.jsActiveEnvID)
	envID := vm.jsActiveEnvID
	var bindings map[string]Value
	reusedArgs := Value{Type: VTJSUndefined}
	if canReuseEnv {
		if env, ok := vm.jsEnvItems[envID]; ok && env != nil {
			if env.bindings == nil {
				env.bindings = make(map[string]Value, len(closure.params)+2)
			}
			bindings = env.bindings
			if existingArgs, exists := bindings["arguments"]; exists {
				reusedArgs = existingArgs
			}
			clear(bindings)
			env.parentID = closure.envID
		} else {
			canReuseEnv = false
		}
	}
	if !canReuseEnv {
		vm.jsReleaseEnvFrame(vm.jsActiveEnvID)
		envID = vm.allocJSID()
		bindings = make(map[string]Value, len(closure.params)+2)
		vm.jsEnvItems[envID] = &jsEnvFrame{parentID: closure.envID, bindings: bindings}
	}

	for i := 0; i < len(closure.params); i++ {
		if i < len(args) {
			bindings[closure.params[i]] = args[i]
		} else {
			bindings[closure.params[i]] = Value{Type: VTJSUndefined}
		}
	}
	if closure.restParam != "" {
		restValues := make([]Value, 0)
		if len(args) > len(closure.params) {
			restValues = append(restValues, args[len(closure.params):]...)
		}
		bindings[closure.restParam] = ValueFromVBArray(NewVBArrayFromValues(0, restValues))
	}
	if _, hasArguments := bindings["arguments"]; !hasArguments {
		if canReuseEnv {
			if reusedArgs.Type == VTJSObject {
				vm.jsRefreshArgumentsObject(reusedArgs.Num, args, closure.params, envID)
				bindings["arguments"] = reusedArgs
			} else {
				bindings["arguments"] = vm.jsCreateArgumentsObject(args, closure.params, envID)
			}
		} else {
			bindings["arguments"] = vm.jsCreateArgumentsObject(args, closure.params, envID)
		}
	}

	if len(vm.jsTryStack) > frame.tryDepth {
		vm.jsTryStack = vm.jsTryStack[:frame.tryDepth]
	}

	vm.jsActiveEnvID = envID
	vm.jsThisValue = thisVal
	vm.jsBlockScopes = append(make([]map[string]Value, 0, len(closure.capturedBlockScopes)), closure.capturedBlockScopes...)
	vm.jsBlockScopeConst = append(make([]map[string]struct{}, 0, len(closure.capturedBlockScopeConst)), closure.capturedBlockScopeConst...)
	vm.jsBlockScopeTDZ = append(make([]map[string]struct{}, 0, len(closure.capturedBlockScopeTDZ)), closure.capturedBlockScopeTDZ...)
	vm.jsBlockScopeDepth = len(vm.jsBlockScopes)
	vm.ip = closure.startIP
	if !vm.jsPrepareLocalFrame(closure.localCount, frame.savedSP) {
		vm.jsActiveEnvID = frame.envID
		vm.jsThisValue = frame.thisVal
		vm.jsNewTarget = frame.newTarget
		vm.jsStrictMode = frame.jsStrictMode
		vm.jsBlockScopes = frame.savedBlockScopes
		vm.jsBlockScopeConst = frame.savedBlockScopeConst
		vm.jsBlockScopeTDZ = frame.savedBlockScopeTDZ
		vm.jsBlockScopeDepth = frame.savedBlockScopeDepth
		vm.fp = frame.savedFP
		vm.sp = frame.savedSP
		vm.ip = frame.returnIP
		return false
	}
	return true
}

func (vm *VM) jsCreateArgumentsObject(args []Value, params []string, envID int64) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, len(args)+1)
	for i := range args {
		obj[strconv.Itoa(i)] = args[i]
	}
	obj["length"] = NewInteger(int64(len(args)))
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, len(args)+1)
	if len(params) > 0 && len(args) > 0 {
		alias := &jsArgumentsBinding{
			envID:        envID,
			indexToParam: make(map[string]string, len(params)),
			paramToIndex: make(map[string]string, len(params)),
		}
		max := min(len(args), len(params))
		for i := 0; i < max; i++ {
			key := strconv.Itoa(i)
			paramName := params[i]
			alias.indexToParam[key] = paramName
			if _, exists := alias.paramToIndex[paramName]; !exists {
				alias.paramToIndex[paramName] = key
			}
		}
		if vm.jsArgumentsItems == nil {
			vm.jsArgumentsItems = make(map[int64]*jsArgumentsBinding, 8)
		}
		vm.jsArgumentsItems[objID] = alias
	}
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsSyncArgumentAliasByParam(envID int64, name string, value Value) {
	if len(vm.jsArgumentsItems) == 0 {
		return
	}
	for objID, alias := range vm.jsArgumentsItems {
		if alias == nil || alias.envID != envID {
			continue
		}
		idxStr, ok := alias.paramToIndex[name]
		if !ok {
			continue
		}
		if obj, exists := vm.jsObjectItems[objID]; exists {
			obj[idxStr] = value
		}
	}
}
func (vm *VM) jsSetAliasedArgumentValue(objID int64, key string, value Value) bool {
	alias, ok := vm.jsArgumentsItems[objID]
	if !ok || alias == nil {
		return false
	}
	paramName, ok := alias.indexToParam[key]
	if !ok {
		return false
	}
	env := vm.jsEnvItems[alias.envID]
	if env == nil {
		return false
	}
	env.bindings[paramName] = value
	if alias.envID == vm.jsActiveEnvID {
		if slot, convErr := strconv.Atoi(key); convErr == nil && slot >= 0 {
			idx := vm.fp + slot
			if idx >= 0 && idx < len(vm.stack) {
				vm.stack[idx] = value
			}
		}
	}
	if obj, exists := vm.jsObjectItems[objID]; exists {
		obj[key] = value
	}
	return true
}

// jsSyncAliasedLocalSlot mirrors a parameter local-slot write to the active env.
// This keeps non-strict arguments aliasing coherent when params are lowered to local slots.
func (vm *VM) jsSyncAliasedLocalSlot(slot int, value Value) {
	if slot < 0 || len(vm.jsCallStack) == 0 {
		return
	}
	frame := vm.jsCallStack[len(vm.jsCallStack)-1]
	closure, ok := vm.jsFunctionItems[frame.fn.Num]
	if !ok || closure == nil {
		return
	}
	if slot >= len(closure.params) {
		return
	}
	envID := vm.jsActiveEnvID
	if envID == 0 {
		return
	}
	env := vm.jsEnvItems[envID]
	if env == nil {
		return
	}
	name := closure.params[slot]
	env.bindings[name] = value
	vm.jsSyncArgumentAliasByParam(envID, name, value)
}

// jsGetAliasedArgumentValue reads a live argument value from the env frame
// for an arguments binding, so that mutations of named params are visible via arguments[n].
func (vm *VM) jsGetAliasedArgumentValue(objID int64, key string) (Value, bool) {
	alias, ok := vm.jsArgumentsItems[objID]
	if !ok || alias == nil {
		return Value{Type: VTJSUndefined}, false
	}
	paramName, ok := alias.indexToParam[key]
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	env := vm.jsEnvItems[alias.envID]
	if env == nil {
		return Value{Type: VTJSUndefined}, false
	}
	val, exists := env.bindings[paramName]
	if !exists {
		return Value{Type: VTJSUndefined}, false
	}
	return val, true
}

func (vm *VM) jsExtractApplyArgs(argArray Value) []Value {
	switch argArray.Type {
	case VTJSUndefined, VTNull:
		return nil
	case VTArray:
		if argArray.Arr == nil || len(argArray.Arr.Values) == 0 {
			return nil
		}
		return argArray.Arr.Values
	case VTJSObject:
		obj, ok := vm.jsObjectItems[argArray.Num]
		if !ok || len(obj) == 0 {
			return nil
		}
		lengthVal, hasLength := obj["length"]
		if !hasLength {
			return nil
		}
		lengthNum := int(vm.jsToNumber(lengthVal).Flt)
		if lengthNum <= 0 {
			return nil
		}
		out := make([]Value, lengthNum)
		for i := range lengthNum {
			key := strconv.Itoa(i)
			if v, exists := obj[key]; exists {
				out[i] = v
			} else {
				out[i] = Value{Type: VTJSUndefined}
			}
		}
		return out
	default:
		return nil
	}
}

func (vm *VM) jsArrayFlat(vals []Value, depth int) []Value {
	if depth < 1 {
		return append([]Value(nil), vals...)
	}
	out := make([]Value, 0, len(vals))
	for _, v := range vals {
		if v.Type == VTArray && v.Arr != nil {
			flattened := vm.jsArrayFlat(v.Arr.Values, depth-1)
			out = append(out, flattened...)
		} else {
			out = append(out, v)
		}
	}
	return out
}

func (vm *VM) jsBindFunction(fn Value, thisArg Value, boundArgs []Value) Value {
	if fn.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}
	}
	id := vm.allocJSID()
	bound := &jsFunctionObject{
		name:      "bound",
		isBound:   true,
		boundFn:   fn,
		boundThis: thisArg,
	}
	if len(boundArgs) > 0 {
		bound.boundArgs = append([]Value(nil), boundArgs...)
	}
	proto := vm.jsCreatePrototypeObject(Value{Type: VTJSFunction, Num: id})
	bound.protoID = proto.Num
	vm.jsFunctionItems[id] = bound
	vm.jsObjectItems[id] = make(map[string]Value, 2)
	vm.jsPropertyItems[id] = make(map[string]jsPropertyDescriptor, 2)
	vm.jsSetDescriptor(id, "prototype", jsPropertyDescriptor{
		Value:        proto,
		HasValue:     true,
		Enumerable:   false,
		Configurable: false,
		Writable:     false,
	})
	return Value{Type: VTJSFunction, Num: id}
}

func jsClampIndex(idx int, length int) int {
	if idx < 0 {
		return 0
	}
	if idx > length {
		return length
	}
	return idx
}

func jsNormalizeRelativeIndex(idx int, length int) int {
	if idx < 0 {
		idx = length + idx
	}
	if idx < 0 {
		return 0
	}
	if idx > length {
		return length
	}
	return idx
}

func jsParseHexByte(hi byte, lo byte) (byte, bool) {
	var high byte
	switch {
	case hi >= '0' && hi <= '9':
		high = hi - '0'
	case hi >= 'a' && hi <= 'f':
		high = hi - 'a' + 10
	case hi >= 'A' && hi <= 'F':
		high = hi - 'A' + 10
	default:
		return 0, false
	}

	var low byte
	switch {
	case lo >= '0' && lo <= '9':
		low = lo - '0'
	case lo >= 'a' && lo <= 'f':
		low = lo - 'a' + 10
	case lo >= 'A' && lo <= 'F':
		low = lo - 'A' + 10
	default:
		return 0, false
	}

	return (high << 4) | low, true
}

func jsIsEncodeURIComponentSafeByte(ch byte) bool {
	if (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9') {
		return true
	}
	switch ch {
	case '-', '_', '.', '!', '~', '*', '\'', '(', ')':
		return true
	default:
		return false
	}
}

func jsIsEncodeURISafeByte(ch byte) bool {
	if jsIsEncodeURIComponentSafeByte(ch) {
		return true
	}
	switch ch {
	case ';', ',', '/', '?', ':', '@', '&', '=', '+', '$', '#':
		return true
	default:
		return false
	}
}

func jsIsDecodeURIReservedByte(ch byte) bool {
	switch ch {
	case ';', ',', '/', '?', ':', '@', '&', '=', '+', '$', '#':
		return true
	default:
		return false
	}
}

func jsEncodeURIValue(input string, component bool) string {
	if input == "" {
		return ""
	}

	var out strings.Builder
	out.Grow(len(input))
	for _, r := range input {
		if r < utf8.RuneSelf {
			ch := byte(r)
			if component {
				if jsIsEncodeURIComponentSafeByte(ch) {
					out.WriteByte(ch)
					continue
				}
			} else {
				if jsIsEncodeURISafeByte(ch) {
					out.WriteByte(ch)
					continue
				}
			}
		}

		var buf [utf8.UTFMax]byte
		n := utf8.EncodeRune(buf[:], r)
		for i := range n {
			b := buf[i]
			out.WriteByte('%')
			out.WriteByte(jsHexUpperDigits[b>>4])
			out.WriteByte(jsHexUpperDigits[b&0x0F])
		}
	}

	return out.String()
}

func jsDecodeURIComponentValue(input string) (string, error) {
	if input == "" {
		return "", nil
	}
	return url.PathUnescape(input)
}

func jsDecodeURIValue(input string) (string, error) {
	if input == "" {
		return "", nil
	}

	var out strings.Builder
	out.Grow(len(input))
	chunkStart := 0
	flushChunk := func(end int) error {
		if end <= chunkStart {
			return nil
		}
		decoded, err := url.PathUnescape(input[chunkStart:end])
		if err != nil {
			return err
		}
		out.WriteString(decoded)
		return nil
	}

	for i := 0; i < len(input); {
		if input[i] != '%' {
			i++
			continue
		}
		if i+2 >= len(input) {
			i++
			continue
		}
		decodedByte, ok := jsParseHexByte(input[i+1], input[i+2])
		if !ok {
			i++
			continue
		}
		if jsIsDecodeURIReservedByte(decodedByte) {
			if err := flushChunk(i); err != nil {
				return "", err
			}
			out.WriteByte('%')
			out.WriteByte(input[i+1])
			out.WriteByte(input[i+2])
			i += 3
			chunkStart = i
			continue
		}
		i += 3
	}

	if err := flushChunk(len(input)); err != nil {
		return "", err
	}
	return out.String(), nil
}

func (vm *VM) jsParseIntES5(args []Value) Value {
	if len(args) == 0 {
		return NewDouble(math.NaN())
	}
	s := strings.TrimSpace(vm.valueToString(args[0]))
	if s == "" {
		return NewDouble(math.NaN())
	}
	sign := 1.0
	if s[0] == '+' || s[0] == '-' {
		if s[0] == '-' {
			sign = -1.0
		}
		s = s[1:]
	}
	if s == "" {
		return NewDouble(math.NaN())
	}
	radix := 0
	if len(args) > 1 {
		r := int(vm.jsToNumber(args[1]).Flt)
		if r != 0 {
			if r < 2 || r > 36 {
				return NewDouble(math.NaN())
			}
			radix = r
		}
	}
	if radix == 0 {
		if len(s) >= 2 && (s[0:2] == "0x" || s[0:2] == "0X") {
			radix = 16
			s = s[2:]
		} else {
			radix = 10
		}
	} else if radix == 16 && len(s) >= 2 && (s[0:2] == "0x" || s[0:2] == "0X") {
		s = s[2:]
	}
	if s == "" {
		return NewDouble(math.NaN())
	}
	value := 0.0
	digits := 0
	for i := 0; i < len(s); i++ {
		ch := s[i]
		d := -1
		switch {
		case ch >= '0' && ch <= '9':
			d = int(ch - '0')
		case ch >= 'a' && ch <= 'z':
			d = int(ch-'a') + 10
		case ch >= 'A' && ch <= 'Z':
			d = int(ch-'A') + 10
		}
		if d < 0 || d >= radix {
			break
		}
		value = value*float64(radix) + float64(d)
		digits++
	}
	if digits == 0 {
		return NewDouble(math.NaN())
	}
	return NewDouble(sign * value)
}

// jsFunctionExpectedLength returns the ES5 Function.length value.
func (vm *VM) jsFunctionExpectedLength(fn Value) int {
	if fn.Type != VTJSFunction {
		return 0
	}
	closure, ok := vm.jsFunctionItems[fn.Num]
	if !ok || closure == nil {
		return 0
	}
	if closure.isBound {
		base := vm.jsFunctionExpectedLength(closure.boundFn)
		remaining := base - len(closure.boundArgs)
		if remaining < 0 {
			return 0
		}
		return remaining
	}
	return len(closure.params)
}

// jsObjectToStringTag returns the canonical Object.prototype.toString tag.
func (vm *VM) jsObjectToStringTag(v Value) string {
	switch v.Type {
	case VTJSUndefined:
		return "[object Undefined]"
	case VTNull:
		return "[object Null]"
	case VTArray:
		return "[object Array]"
	case VTDate:
		return "[object Date]"
	case VTJSFunction:
		return "[object Function]"
	case VTString:
		return "[object String]"
	case VTSymbol:
		return "[object Symbol]"
	case VTBool:
		return "[object Boolean]"
	case VTInteger, VTDouble:
		return "[object Number]"
	case VTJSObject:
		tag := vm.jsObjectStringProperty(v, "__js_type")
		if tag == "" {
			tag = vm.jsObjectStringProperty(v, "__js_ctor")
		}
		switch tag {
		case "Array", "Date", "Function", "RegExp", "Math", "JSON", "Enumerator", "VBArray", "String", "Number", "Boolean", "Object", "Set", "Map", "WeakMap", "WeakSet",
			"ArrayBuffer", "DataView",
			"Int8Array", "Uint8Array", "Uint8ClampedArray",
			"Int16Array", "Uint16Array",
			"Int32Array", "Uint32Array",
			"Float32Array", "Float64Array",
			"BigInt64Array", "BigUint64Array":
			return "[object " + tag + "]"
		default:
			if tag != "" {
				return "[object " + tag + "]"
			}
			return "[object Object]"
		}
	default:
		return "[object Object]"
	}
}

// jsArrayLikeLength resolves the effective ES5 array-like length from an object.
func (vm *VM) jsArrayLikeLength(target Value) (int, bool, bool) {
	switch target.Type {
	case VTArray:
		if target.Arr == nil {
			return 0, true, false
		}
		return len(target.Arr.Values), true, false
	case VTJSObject:
		lengthVal, deferred := vm.jsMemberGet(target, "length")
		if deferred {
			return 0, false, true
		}
		if lengthVal.Type == VTJSUndefined {
			arrayProto := vm.jsGetIntrinsicPrototype("Array")
			if arrayProto.Type == VTJSObject && vm.jsObjectIsPrototypeOf(arrayProto, target) {
				return 0, true, false
			}
			return 0, false, false
		}
		n := max(int(vm.jsToNumber(lengthVal).Flt), 0)
		return n, true, false
	default:
		return 0, false, false
	}
}

// jsArrayLikeHasIndex reports whether an array-like value has an element index.
func (vm *VM) jsArrayLikeHasIndex(target Value, idx int) bool {
	if idx < 0 {
		return false
	}
	if target.Type == VTArray {
		if target.Arr == nil {
			return false
		}
		return idx < len(target.Arr.Values)
	}
	if target.Type != VTJSObject {
		return false
	}
	return vm.jsHasProperty(target, strconv.Itoa(idx))
}

// jsArrayLikeGetIndex resolves one element from an array-like value.
func (vm *VM) jsArrayLikeGetIndex(target Value, idx int) (Value, bool) {
	if idx < 0 {
		return Value{Type: VTJSUndefined}, false
	}
	if target.Type == VTArray {
		if target.Arr == nil || idx >= len(target.Arr.Values) {
			return Value{Type: VTJSUndefined}, false
		}
		return target.Arr.Values[idx], true
	}
	if target.Type != VTJSObject {
		return Value{Type: VTJSUndefined}, false
	}
	v, deferred := vm.jsMemberGet(target, strconv.Itoa(idx))
	if deferred {
		return Value{Type: VTJSUndefined}, false
	}
	if v.Type == VTJSUndefined && !vm.jsArrayLikeHasIndex(target, idx) {
		return Value{Type: VTJSUndefined}, false
	}
	return v, true
}

func (vm *VM) jsParseFloatES5(args []Value) Value {
	if len(args) == 0 {
		return NewDouble(math.NaN())
	}
	s := strings.TrimSpace(vm.valueToString(args[0]))
	if s == "" {
		return NewDouble(math.NaN())
	}
	if strings.HasPrefix(s, "Infinity") || strings.HasPrefix(s, "+Infinity") {
		return NewDouble(math.Inf(1))
	}
	if strings.HasPrefix(s, "-Infinity") {
		return NewDouble(math.Inf(-1))
	}
	i := 0
	if i < len(s) && (s[i] == '+' || s[i] == '-') {
		i++
	}
	intStart := i
	for i < len(s) && s[i] >= '0' && s[i] <= '9' {
		i++
	}
	hasInt := i > intStart
	if i < len(s) && s[i] == '.' {
		i++
		fracStart := i
		for i < len(s) && s[i] >= '0' && s[i] <= '9' {
			i++
		}
		hasInt = hasInt || (i > fracStart)
	}
	if !hasInt {
		return NewDouble(math.NaN())
	}
	end := i
	if i < len(s) && (s[i] == 'e' || s[i] == 'E') {
		expPos := i
		i++
		if i < len(s) && (s[i] == '+' || s[i] == '-') {
			i++
		}
		expStart := i
		for i < len(s) && s[i] >= '0' && s[i] <= '9' {
			i++
		}
		if i > expStart {
			end = i
		} else {
			end = expPos
		}
	}
	parsed, err := strconv.ParseFloat(s[:end], 64)
	if err != nil {
		return NewDouble(math.NaN())
	}
	return NewDouble(parsed)
}

func (vm *VM) jsObjectHasOwnProperty(target Value, key string) bool {
	switch target.Type {
	case VTJSObject, VTJSFunction:
		if _, ok := vm.jsGetDescriptor(target.Num, key); ok {
			return true
		}
		obj, ok := vm.jsObjectItems[target.Num]
		if !ok {
			return false
		}
		_, exists := obj[key]
		return exists
	case VTArray:
		if strings.EqualFold(key, "length") {
			return true
		}
		if target.Arr == nil {
			return false
		}
		i, err := strconv.Atoi(key)
		if err != nil {
			return false
		}
		idx := i - target.Arr.Lower
		return idx >= 0 && idx < len(target.Arr.Values)
	case VTString:
		if strings.EqualFold(key, "length") {
			return true
		}
		i, err := strconv.Atoi(key)
		if err != nil {
			return false
		}
		r := []rune(target.Str)
		return i >= 0 && i < len(r)
	default:
		return false
	}
}

func (vm *VM) jsObjectPropertyIsEnumerable(target Value, key string) bool {
	switch target.Type {
	case VTJSObject, VTJSFunction:
		d, ok := vm.jsGetDescriptor(target.Num, key)
		if !ok {
			return false
		}
		return d.Enumerable
	case VTArray, VTString:
		return vm.jsObjectHasOwnProperty(target, key)
	default:
		return false
	}
}

func (vm *VM) jsObjectIsPrototypeOf(proto Value, candidate Value) bool {
	if proto.Type != VTJSObject && proto.Type != VTJSFunction {
		return false
	}
	seen := make(map[int64]struct{}, 4)
	current := vm.jsGetPrototypeValue(candidate)
	for current.Type == VTJSObject {
		if current.Num == proto.Num {
			return true
		}
		if _, ok := seen[current.Num]; ok {
			break
		}
		seen[current.Num] = struct{}{}
		obj, ok := vm.jsObjectItems[current.Num]
		if !ok {
			break
		}
		next, ok := obj["__js_proto"]
		if !ok {
			break
		}
		current = next
	}
	return false
}

func (vm *VM) jsReturn(retVal Value) {
	if len(vm.jsCallStack) == 0 {
		vm.push(retVal)
		return
	}
	currentEnvID := vm.jsActiveEnvID
	frame := vm.jsCallStack[len(vm.jsCallStack)-1]
	vm.jsCallStack = vm.jsCallStack[:len(vm.jsCallStack)-1]
	vm.jsReleaseEnvFrame(currentEnvID)
	if len(vm.jsTryStack) > frame.tryDepth {
		vm.jsTryStack = vm.jsTryStack[:frame.tryDepth]
	}
	vm.jsActiveEnvID = frame.envID
	vm.jsThisValue = frame.thisVal
	vm.jsNewTarget = frame.newTarget
	vm.jsStrictMode = frame.jsStrictMode
	vm.jsBlockScopes = frame.savedBlockScopes
	vm.jsBlockScopeConst = frame.savedBlockScopeConst
	vm.jsBlockScopeTDZ = frame.savedBlockScopeTDZ
	vm.jsBlockScopeDepth = frame.savedBlockScopeDepth
	vm.ip = frame.returnIP
	vm.fp = frame.savedFP
	vm.sp = frame.savedSP
	if frame.isSuperCall {
		// This was a super() call in a constructor. Assign the result to 'this'.
		newThis := retVal
		if retVal.Type != VTJSObject && retVal.Type != VTJSFunction && retVal.Type != VTObject && retVal.Type != VTNativeObject && retVal.Type != VTArray {
			newThis = frame.ctorObj
		}
		vm.jsThisValue = newThis
		if len(vm.jsCallStack) > 0 {
			vm.jsCallStack[len(vm.jsCallStack)-1].thisVal = newThis
			if vm.jsCallStack[len(vm.jsCallStack)-1].ctorObj.Type == VTJSUninitialized {
				vm.jsCallStack[len(vm.jsCallStack)-1].ctorObj = newThis
			}
		}
		vm.push(newThis)
		return
	}

	if frame.isCtor {
		switch retVal.Type {
		case VTJSObject, VTJSFunction, VTObject, VTNativeObject, VTArray:
			vm.push(retVal)
			return
		default:
			vm.push(frame.ctorObj)
			return
		}
	}
	vm.push(retVal)
}

func (vm *VM) jsPropertyKeyToValue(key string) Value {
	if after, ok := strings.CutPrefix(key, jsSymbolPropertyPrefix); ok {
		idStr := after
		if id, err := strconv.ParseInt(idStr, 10, 64); err == nil {
			return Value{Type: VTSymbol, Num: id}
		}
	}
	return NewString(key)
}

// jsProxyGet handles the [[Get]] internal method for a Proxy, enforcing all
// ECMAScript invariants: a non-configurable non-writable data property must
// return exactly its stored value; a non-configurable accessor with no getter
// must return undefined.
func (vm *VM) jsProxyGet(proxy Value, member string, receiver Value) (Value, bool) {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return Value{Type: VTJSUndefined}, false
	}

	// 1. Get the 'get' trap from handler
	trap, _ := vm.jsMemberGet(pObj.Handler, "get")

	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		// 2. No trap — forward to target directly
		return vm.jsMemberGet(pObj.Target, member)
	}

	// 3. Invoke the trap: trap(target, property, receiver)
	args := []Value{pObj.Target, vm.jsPropertyKeyToValue(member), receiver}
	result := vm.jsCall(trap, pObj.Handler, args)

	// 4. Invariant check (§10.5.8 step 7)
	if pObj.Target.Type == VTJSObject || pObj.Target.Type == VTJSFunction {
		if targetDesc, hasDesc := vm.jsGetDescriptor(pObj.Target.Num, member); hasDesc {
			if !targetDesc.Configurable {
				if jsDescriptorIsData(targetDesc) && !targetDesc.Writable {
					// Must return the exact same value.
					if !vm.jsStrictEquals(result, targetDesc.Value) {
						vm.jsThrowJSError(jscript.ProxyGetTrapInvariantViolation)
						return Value{Type: VTJSUndefined}, false
					}
				} else if (targetDesc.HasGetter || targetDesc.HasSetter) && !targetDesc.HasGetter {
					// Accessor with no getter — result must be undefined.
					if result.Type != VTJSUndefined {
						vm.jsThrowJSError(jscript.ProxyGetTrapInvariantViolation)
						return Value{Type: VTJSUndefined}, false
					}
				}
			}
		}
	}
	return result, false
}

// jsProxySet handles the [[Set]] internal method for a Proxy, enforcing the
// invariant that a non-configurable, non-writable data property cannot have
// its value changed via a trap that returns true.
func (vm *VM) jsProxySet(proxy Value, member string, val Value, receiver Value) {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return
	}

	// 1. Get the 'set' trap
	trap, _ := vm.jsMemberGet(pObj.Handler, "set")

	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		// 2. No trap — forward to target directly
		vm.jsMemberSet(pObj.Target, member, val)
		return
	}

	// 3. Invoke the trap: trap(target, property, value, receiver)
	args := []Value{pObj.Target, vm.jsPropertyKeyToValue(member), val, receiver}
	result := vm.jsCall(trap, pObj.Handler, args)

	// 4. Falsy result check in strict mode
	if !vm.jsTruthy(result) {
		if vm.jsStrictMode {
			vm.jsThrowTypeError(fmt.Sprintf("'set' on proxy: trap returned falsy for property '%s'", member))
		}
		return
	}

	// 5. Invariant check (§10.5.9 step 8): trap returned true
	// A non-configurable, non-writable data property cannot be set to a different value.
	if pObj.Target.Type == VTJSObject || pObj.Target.Type == VTJSFunction {
		if targetDesc, hasDesc := vm.jsGetDescriptor(pObj.Target.Num, member); hasDesc && !targetDesc.Configurable {
			if jsDescriptorIsData(targetDesc) && !targetDesc.Writable {
				if !vm.jsStrictEquals(val, targetDesc.Value) {
					vm.jsThrowJSError(jscript.ProxySetTrapInvariantViolation)
					return
				}
			} else if targetDesc.HasSetter && !targetDesc.HasSetter {
				// Non-configurable accessor with no setter
				vm.jsThrowJSError(jscript.ProxySetTrapInvariantViolation)
				return
			}
		}
	}
}

// jsProxyHas handles the [[HasProperty]] internal method for a Proxy, enforcing
// the invariant that a non-configurable own property (or any own property of a
// non-extensible target) cannot be reported as absent.
func (vm *VM) jsProxyHas(proxy Value, member string) bool {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return false
	}

	trap, _ := vm.jsMemberGet(pObj.Handler, "has")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		return vm.jsHas(pObj.Target, member)
	}

	args := []Value{pObj.Target, vm.jsPropertyKeyToValue(member)}
	result := vm.jsCall(trap, pObj.Handler, args)
	trapResult := vm.jsTruthy(result)

	// Invariant check (§10.5.7 step 8): trap returned false
	if pObj.Target.Type == VTJSObject || pObj.Target.Type == VTJSFunction {
		targetExtensible := vm.jsObjectIsExtensible(pObj.Target)
		if targetDesc, hasDesc := vm.jsGetDescriptor(pObj.Target.Num, member); hasDesc {
			if code, violated := jscript.ValidateProxyHasMissingPropertyInvariant(hasDesc, targetDesc.Configurable, targetExtensible, trapResult); violated {
				vm.jsThrowJSError(code)
				return false
			}
		}
	}
	return trapResult
}

// jsProxyDelete handles the [[Delete]] internal method for a Proxy, enforcing
// the invariant that a non-configurable property cannot be reported as deleted.
func (vm *VM) jsProxyDelete(proxy Value, member string) bool {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return false
	}

	trap, _ := vm.jsMemberGet(pObj.Handler, "deleteProperty")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		return vm.jsMemberDelete(pObj.Target, member)
	}

	args := []Value{pObj.Target, vm.jsPropertyKeyToValue(member)}
	result := vm.jsCall(trap, pObj.Handler, args)
	success := vm.jsTruthy(result)

	if !success {
		if vm.jsStrictMode {
			vm.jsThrowTypeError(fmt.Sprintf("'deleteProperty' on proxy: trap returned falsy for property '%s'", member))
		}
		return false
	}

	// Invariant check (§10.5.10 step 8): trap returned true
	// A non-configurable property cannot be deleted.
	if pObj.Target.Type == VTJSObject || pObj.Target.Type == VTJSFunction {
		if targetDesc, hasDesc := vm.jsGetDescriptor(pObj.Target.Num, member); hasDesc && !targetDesc.Configurable {
			vm.jsThrowJSError(jscript.ProxyDeletePropertyTrapInvariantViolation)
			return false
		}
		// Non-extensible target: cannot delete an existing own property
		if !vm.jsObjectIsExtensible(pObj.Target) {
			if _, hasDesc := vm.jsGetDescriptor(pObj.Target.Num, member); hasDesc {
				vm.jsThrowJSError(jscript.ProxyDeletePropertyTrapInvariantViolation)
				return false
			}
		}
	}
	return true
}

// jsProxyOwnKeys handles the [[OwnPropertyKeys]] internal method for a Proxy,
// enforcing that the result includes every non-configurable own property key and,
// for non-extensible targets, exactly matches the target's own key set.
func (vm *VM) jsProxyOwnKeys(proxy Value) []string {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return nil
	}

	trap, _ := vm.jsMemberGet(pObj.Handler, "ownKeys")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		return vm.jsObjectOwnKeys(pObj.Target)
	}

	args := []Value{pObj.Target}
	result := vm.jsCall(trap, pObj.Handler, args)

	var keys []string
	if result.Type == VTArray && result.Arr != nil {
		keys = make([]string, len(result.Arr.Values))
		for i := 0; i < len(result.Arr.Values); i++ {
			keys[i] = vm.valueToString(result.Arr.Values[i])
		}
	} else if result.Type == VTJSObject || result.Type == VTJSFunction {
		lengthVal, _ := vm.jsMemberGet(result, "length")
		length := int(vm.jsToNumber(lengthVal).Flt)
		keys = make([]string, length)
		for i := range length {
			val := vm.jsIndexGet(result, NewInteger(int64(i)))
			keys[i] = vm.valueToString(val)
		}
	} else {
		vm.jsThrowTypeError("Proxy 'ownKeys' trap must return an array")
		return nil
	}

	// Invariant checks (§10.5.11 steps 14-26).
	if pObj.Target.Type == VTJSObject || pObj.Target.Type == VTJSFunction {
		// Build a set from the trap result for O(1) lookups.
		resultSet := make(map[string]struct{}, len(keys))
		for _, k := range keys {
			resultSet[k] = struct{}{}
		}

		targetKeys := vm.jsObjectOwnKeys(pObj.Target)

		// All non-configurable own property keys must be present in the result.
		for _, tk := range targetKeys {
			if desc, hasDesc := vm.jsGetDescriptor(pObj.Target.Num, tk); hasDesc && !desc.Configurable {
				if _, inResult := resultSet[tk]; !inResult {
					vm.jsThrowJSError(jscript.ProxyOwnKeysTrapInvariantViolation)
					return nil
				}
			}
		}

		// If target is non-extensible, the result must not introduce new keys.
		if !vm.jsObjectIsExtensible(pObj.Target) {
			targetSet := make(map[string]struct{}, len(targetKeys))
			for _, tk := range targetKeys {
				targetSet[tk] = struct{}{}
			}
			for _, k := range keys {
				if _, inTarget := targetSet[k]; !inTarget {
					vm.jsThrowJSError(jscript.ProxyOwnKeysTrapInvariantViolation)
					return nil
				}
			}
			// All target keys must also appear in the result.
			for _, tk := range targetKeys {
				if _, inResult := resultSet[tk]; !inResult {
					vm.jsThrowJSError(jscript.ProxyOwnKeysTrapInvariantViolation)
					return nil
				}
			}
		}
	}
	return keys
}

// jsProxyDefineProperty handles the [[DefineOwnProperty]] internal method for a
// Proxy. When the trap returns true the result is validated against the target's
// extensibility and existing non-configurable descriptors.
func (vm *VM) jsProxyDefineProperty(proxy Value, member string, attributes Value) bool {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return false
	}
	trap, _ := vm.jsMemberGet(pObj.Handler, "defineProperty")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		objID := pObj.Target.Num
		current, currentExists := vm.jsGetDescriptor(objID, member)
		if !currentExists && !vm.jsObjectIsExtensible(pObj.Target) {
			return false
		}
		spec := vm.jsReadDefinePropertySpec(attributes)
		if !vm.jsValidateDefinePropertyTransition(current, currentExists, spec) {
			return false
		}
		finalDesc := vm.jsApplyDefinePropertySpec(current, currentExists, spec)
		vm.jsSetDescriptor(objID, member, finalDesc)
		return true
	}
	args := []Value{pObj.Target, NewString(member), attributes}
	result := vm.jsCall(trap, pObj.Handler, args)
	trapResult := vm.asBool(result)

	if !trapResult {
		return false
	}

	// Invariant checks (§10.5.6 steps 19-27): trap returned true.
	if pObj.Target.Type == VTJSObject || pObj.Target.Type == VTJSFunction {
		spec := vm.jsReadDefinePropertySpec(attributes)
		current, currentExists := vm.jsGetDescriptor(pObj.Target.Num, member)

		// Cannot add a new property to a non-extensible target.
		if !currentExists && !vm.jsObjectIsExtensible(pObj.Target) {
			vm.jsThrowJSError(jscript.ProxyDefinePropertyTrapInvariantViolation)
			return false
		}

		// Cannot make an existing non-configurable property configurable.
		if currentExists && !current.Configurable {
			if spec.hasConfigurable && spec.desc.Configurable {
				vm.jsThrowJSError(jscript.ProxyDefinePropertyTrapInvariantViolation)
				return false
			}
			// Cannot make a non-configurable non-writable property writable.
			if jsDescriptorIsData(current) && !current.Writable && spec.hasWritable && spec.desc.Writable {
				vm.jsThrowJSError(jscript.ProxyDefinePropertyTrapInvariantViolation)
				return false
			}
		}
	}
	return true
}

// jsProxyGetOwnPropertyDescriptor handles [[GetOwnProperty]] for a Proxy and
// enforces the invariants: the trap cannot hide a non-configurable own property
// of the target, cannot report a non-configurable property as configurable, and
// cannot omit an existing own property of a non-extensible target.
func (vm *VM) jsProxyGetOwnPropertyDescriptor(proxy Value, member string) (jsPropertyDescriptor, bool) {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return jsPropertyDescriptor{}, false
	}
	trap, _ := vm.jsMemberGet(pObj.Handler, "getOwnPropertyDescriptor")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		return vm.jsGetDescriptor(pObj.Target.Num, member)
	}
	args := []Value{pObj.Target, NewString(member)}
	result := vm.jsCall(trap, pObj.Handler, args)

	// Trap returned undefined — check invariant.
	if result.Type == VTJSUndefined {
		if pObj.Target.Type == VTJSObject || pObj.Target.Type == VTJSFunction {
			if targetDesc, hasDesc := vm.jsGetDescriptor(pObj.Target.Num, member); hasDesc {
				// Cannot hide a non-configurable own property.
				if !targetDesc.Configurable {
					vm.jsThrowJSError(jscript.ProxyGetOwnPropertyDescriptorTrapInvariantViolation)
					return jsPropertyDescriptor{}, false
				}
				// Cannot omit an existing own property of a non-extensible target.
				if !vm.jsObjectIsExtensible(pObj.Target) {
					vm.jsThrowJSError(jscript.ProxyGetOwnPropertyDescriptorTrapInvariantViolation)
					return jsPropertyDescriptor{}, false
				}
			}
		}
		return jsPropertyDescriptor{}, false
	}

	if result.Type != VTJSObject && result.Type != VTJSFunction {
		vm.jsThrowTypeError("Proxy 'getOwnPropertyDescriptor' trap must return an object or undefined")
		return jsPropertyDescriptor{}, false
	}

	spec := vm.jsReadDefinePropertySpec(result)
	trapDesc := spec.desc

	// Invariant checks (§10.5.5): trap returned a descriptor object.
	if pObj.Target.Type == VTJSObject || pObj.Target.Type == VTJSFunction {
		targetDesc, hasTargetDesc := vm.jsGetDescriptor(pObj.Target.Num, member)

		// Cannot report a configurable descriptor for a non-configurable target property.
		if hasTargetDesc && !targetDesc.Configurable && trapDesc.Configurable {
			vm.jsThrowJSError(jscript.ProxyGetOwnPropertyDescriptorTrapInvariantViolation)
			return jsPropertyDescriptor{}, false
		}

		// Cannot report a non-configurable non-writable data descriptor as writable
		// when the target has it as non-writable.
		if hasTargetDesc && !targetDesc.Configurable && jsDescriptorIsData(targetDesc) && !targetDesc.Writable {
			if jsDescriptorIsData(trapDesc) && trapDesc.Writable {
				vm.jsThrowJSError(jscript.ProxyGetOwnPropertyDescriptorTrapInvariantViolation)
				return jsPropertyDescriptor{}, false
			}
		}

		// For a non-extensible target, cannot report a property that does not exist on target.
		if !vm.jsObjectIsExtensible(pObj.Target) && !hasTargetDesc {
			vm.jsThrowJSError(jscript.ProxyGetOwnPropertyDescriptorTrapInvariantViolation)
			return jsPropertyDescriptor{}, false
		}
	}

	return trapDesc, true
}

// jsProxyGetPrototypeOf handles [[GetPrototypeOf]] for a Proxy. When the target
// is non-extensible the trap must return the same prototype as the target.
func (vm *VM) jsProxyGetPrototypeOf(proxy Value) Value {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return Value{Type: VTJSUndefined}
	}
	trap, _ := vm.jsMemberGet(pObj.Handler, "getPrototypeOf")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		return vm.jsGetPrototypeValue(pObj.Target)
	}
	args := []Value{pObj.Target}
	result := vm.jsCall(trap, pObj.Handler, args)
	if result.Type != VTJSObject && result.Type != VTJSFunction && result.Type != VTNull && result.Type != VTArray && result.Type != VTJSProxy {
		vm.jsThrowTypeError("Proxy 'getPrototypeOf' trap must return an object or null")
		return Value{Type: VTJSUndefined}
	}
	// Invariant check (§10.5.1 step 8): if target is non-extensible the returned
	// prototype must be identical to the target's actual prototype.
	targetExtensible := vm.jsObjectIsExtensible(pObj.Target)
	if !targetExtensible {
		targetProto := vm.jsGetPrototypeValue(pObj.Target)
		if code, violated := jscript.ValidateProxyGetPrototypeOfInvariant(targetExtensible, vm.jsStrictEquals(result, targetProto)); violated {
			vm.jsThrowJSError(code)
			return Value{Type: VTJSUndefined}
		}
	}
	return result
}

func (vm *VM) jsProxyIsExtensible(proxy Value) bool {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return false
	}
	trap, _ := vm.jsMemberGet(pObj.Handler, "isExtensible")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		return vm.jsObjectIsExtensible(pObj.Target)
	}
	args := []Value{pObj.Target}
	result := vm.jsCall(trap, pObj.Handler, args)
	booleanResult := vm.asBool(result)
	if booleanResult != vm.jsObjectIsExtensible(pObj.Target) {
		vm.jsThrowTypeError("Proxy 'isExtensible' trap result must match the target's extensibility")
		return false
	}
	return booleanResult
}

// jsProxyPreventExtensions handles [[PreventExtensions]] for a Proxy. The spec
// invariant states: if the trap returns true the target must be non-extensible
// already; silently marking the target is not permitted.
func (vm *VM) jsProxyPreventExtensions(proxy Value) bool {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return false
	}
	trap, _ := vm.jsMemberGet(pObj.Handler, "preventExtensions")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		vm.jsSetObjectExtensible(pObj.Target.Num, false)
		return true
	}
	args := []Value{pObj.Target}
	result := vm.jsCall(trap, pObj.Handler, args)
	booleanResult := vm.asBool(result)
	// Invariant check (§10.5.4 step 8): trap returned true but target is
	// still extensible — this is a spec violation.
	if booleanResult && vm.jsObjectIsExtensible(pObj.Target) {
		vm.jsThrowJSError(jscript.ProxyPreventExtensionsTrapInvariantViolation)
		return false
	}
	return booleanResult
}

// jsProxySetPrototypeOf handles [[SetPrototypeOf]] for a Proxy. When the target
// is non-extensible the trap must not attempt to change the prototype.
func (vm *VM) jsProxySetPrototypeOf(proxy Value, proto Value) bool {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return false
	}
	trap, _ := vm.jsMemberGet(pObj.Handler, "setPrototypeOf")
	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		return vm.jsSetPrototype(pObj.Target, proto)
	}
	args := []Value{pObj.Target, proto}
	result := vm.jsCall(trap, pObj.Handler, args)
	trapResult := vm.asBool(result)

	// Invariant check (§10.5.2 step 9): if the trap returned true and the target
	// is non-extensible, the new prototype must be the same as the current one.
	if trapResult && !vm.jsObjectIsExtensible(pObj.Target) {
		currentProto := vm.jsGetPrototypeValue(pObj.Target)
		if !vm.jsStrictEquals(proto, currentProto) {
			vm.jsThrowJSError(jscript.ProxySetPrototypeOfTrapInvariantViolation)
			return false
		}
	}
	return trapResult
}

func (vm *VM) jsSetPrototype(target Value, proto Value) bool {
	if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray {
		return false
	}
	if proto.Type != VTJSObject && proto.Type != VTJSFunction && proto.Type != VTArray && proto.Type != VTNull && proto.Type != VTJSProxy {
		return false
	}
	if !vm.jsObjectIsExtensible(target) {
		currentProto := vm.jsGetPrototypeValue(target)
		if currentProto.Type != proto.Type || currentProto.Num != proto.Num {
			return false
		}
	}
	curr := proto
	for curr.Type == VTJSObject || curr.Type == VTJSFunction || curr.Type == VTArray {
		if curr.Num == target.Num {
			return false
		}
		curr = vm.jsGetPrototypeValue(curr)
	}
	obj, ok := vm.jsObjectItems[target.Num]
	if !ok {
		obj = make(map[string]Value, 8)
		vm.jsObjectItems[target.Num] = obj
	}
	obj["__js_proto"] = proto
	vm.jsInvalidateObjectIC(target.Num)
	return true
}

func (vm *VM) jsHas(target Value, key string) bool {
	switch target.Type {
	case VTJSProxy:
		return vm.jsProxyHas(target, key)
	case VTJSObject, VTJSFunction:
		_, exists := vm.jsResolveObjectMember(target.Num, key, make(map[int64]struct{}, 4))
		return exists
	case VTNativeObject:
		// Check if property exists on native object
		return vm.dispatchMemberGet(target, key).Type != VTJSUndefined
	case VTString:
		if strings.EqualFold(key, "length") {
			return true
		}
		if idx, err := strconv.Atoi(key); err == nil && idx >= 0 && int64(idx) < jsStringLength(target.Str) {
			return true
		}
	case VTArray:
		if target.Arr != nil {
			if strings.EqualFold(key, "length") {
				return true
			}
			if idx, err := strconv.Atoi(key); err == nil && idx >= 0 && int64(idx) < int64(len(target.Arr.Values)) {
				return true
			}
		}
	}
	return false
}

func (vm *VM) jsObjectOwnKeys(obj Value) []string {
	names := vm.jsObjectOwnPropertyNames(obj)
	symbols := vm.jsObjectOwnPropertySymbols(obj)
	keys := make([]string, 0, len(names)+len(symbols))
	keys = append(keys, names...)
	for _, s := range symbols {
		keys = append(keys, jsSymbolPropertyPrefix+strconv.FormatInt(s.Num, 10))
	}
	return keys
}

func (vm *VM) jsMemberGet(target Value, member string) (Value, bool) {
	if target.Type == VTJSUninitialized {
		vm.jsThrowReferenceError("Must call super constructor in derived class before accessing 'this'")
		return Value{Type: VTJSUndefined}, false
	}
	if target.Type == VTJSUndefined {
		vm.jsThrowTypeError(fmt.Sprintf("Cannot read property '%s' of undefined", member))
		return Value{Type: VTJSUndefined}, false
	}
	if target.Type == VTNull {
		vm.jsThrowTypeError(fmt.Sprintf("Cannot read property '%s' of null", member))
		return Value{Type: VTJSUndefined}, false
	}
	switch target.Type {
	case VTJSProxy:
		return vm.jsProxyGet(target, member, target)
	case VTNativeObject:
		return vm.dispatchMemberGet(target, member), false
	case VTString:
		if strings.EqualFold(member, "length") {
			return NewInteger(jsStringLength(target.Str)), false
		}
		proto := vm.jsGetIntrinsicPrototype("String")
		if proto.Type == VTJSObject {
			desc, hasDesc := vm.jsResolveObjectMember(proto.Num, member, make(map[int64]struct{}, 4))
			if hasDesc {
				if desc.HasGetter {
					result := vm.jsCall(desc.Getter, target, nil)
					if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
						return Value{Type: VTJSUndefined}, true
					}
					return result, false
				}
				if desc.HasValue {
					return desc.Value, false
				}
			}
		}
		return Value{Type: VTJSUndefined}, false
	case VTArray:
		if target.Arr != nil {
			if strings.EqualFold(member, "length") {
				return NewInteger(int64(len(target.Arr.Values))), false
			}
			if target.Arr.JSProps != nil {
				if val, ok := target.Arr.JSProps[member]; ok {
					return val, false
				}
			}
		}
		proto := vm.jsGetIntrinsicPrototype("Array")
		if proto.Type == VTJSObject {
			desc, hasDesc := vm.jsResolveObjectMember(proto.Num, member, make(map[int64]struct{}, 4))
			if hasDesc {
				if desc.HasGetter {
					result := vm.jsCall(desc.Getter, target, nil)
					if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
						return Value{Type: VTJSUndefined}, true
					}
					return result, false
				}
				if desc.HasValue {
					return desc.Value, false
				}
			}
		}
		return Value{Type: VTJSUndefined}, false
	case VTDate:
		proto := vm.jsGetIntrinsicPrototype("Date")
		if proto.Type == VTJSObject {
			desc, hasDesc := vm.jsResolveObjectMember(proto.Num, member, make(map[int64]struct{}, 4))
			if hasDesc {
				if desc.HasGetter {
					result := vm.jsCall(desc.Getter, target, nil)
					if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
						return Value{Type: VTJSUndefined}, true
					}
					return result, false
				}
				if desc.HasValue {
					return desc.Value, false
				}
			}
		}
		return Value{Type: VTJSUndefined}, false
	case VTJSObject:
		// Handle global/globalThis object - access properties from root environment bindings first
		if target.Num == vm.jsRootEnvID && vm.jsRootEnvID != 0 {
			if rootEnv, ok := vm.jsEnvItems[vm.jsRootEnvID]; ok {
				if val, exists := rootEnv.bindings[member]; exists {
					return val, false
				}
			}
		}
		if vm.jsObjectStringProperty(target, "__js_type") == "Array" {
			if idx, ok := jsParseArrayIndex(member); ok {
				if slots, hasSlots := vm.jsObjectSlots[target.Num]; hasSlots {
					if idx >= 0 && idx < len(slots) {
						return slots[idx], false
					}
				}
			}
		}
		// Handle Buffer instance index access (e.g., buffer[0])
		if vm.jsObjectStringProperty(target, "__js_type") == "Buffer" {
			if idx, ok := jsParseArrayIndex(member); ok {
				bufItem, exists := vm.jsBufferItems[target.Num]
				if exists && idx >= 0 && idx < len(bufItem.data) {
					return NewInteger(int64(bufItem.data[idx])), false
				}
			}
		}
		if val, handled := vm.jsHandleNodeURLMemberGet(target, member); handled {
			return val, false
		}
		if value, ok := vm.jsGetAliasedArgumentValue(target.Num, member); ok {
			return value, false
		}
		if vm.jsObjectStringProperty(target, "__js_type") == "RegExp" {
			switch {
			case strings.EqualFold(member, "flags"):
				return NewString(vm.jsRegExpGetFlags(target)), false
			case strings.EqualFold(member, "source"):
				return vm.jsObjectItems[target.Num]["pattern"], false
			case strings.EqualFold(member, "global"):
				return NewBool(strings.Contains(vm.jsObjectStringProperty(target, "flags"), "g")), false
			case strings.EqualFold(member, "ignoreCase"):
				return NewBool(strings.Contains(vm.jsObjectStringProperty(target, "flags"), "i")), false
			case strings.EqualFold(member, "multiline"):
				return NewBool(strings.Contains(vm.jsObjectStringProperty(target, "flags"), "m")), false
			case strings.EqualFold(member, "sticky"):
				return NewBool(strings.Contains(vm.jsObjectStringProperty(target, "flags"), "y")), false
			case strings.EqualFold(member, "unicode"):
				return NewBool(strings.Contains(vm.jsObjectStringProperty(target, "flags"), "u")), false
			case strings.EqualFold(member, "dotAll"):
				return NewBool(strings.Contains(vm.jsObjectStringProperty(target, "flags"), "s")), false
			}
		}
		if strings.EqualFold(member, "size") {
			if targetType := vm.jsObjectStringProperty(target, "__js_type"); targetType == "Set" || targetType == "Map" {
				if store, ok := vm.jsCollectionStore(target, targetType, "size"); ok {
					return NewInteger(int64(len(store))), false
				}
			}
		}
		// ArrayBuffer / SharedArrayBuffer property get (byteLength)
		if backing, isBuf := vm.jsArrayBuffers[target.Num]; isBuf {
			if strings.EqualFold(member, "byteLength") {
				return NewInteger(int64(len(backing))), false
			}
		} else if backing, isSBuf := vm.jsSharedArrayBuffers[target.Num]; isSBuf {
			if strings.EqualFold(member, "byteLength") {
				return NewInteger(int64(len(backing))), false
			}
		}
		// Typed array / DataView property get
		if v, handled := vm.jsTypedArrayMemberGet(target, member); handled {
			return v, false
		}
		desc, hasDesc := vm.jsResolveObjectMember(target.Num, member, make(map[int64]struct{}, 4))
		if hasDesc {
			if desc.HasGetter {
				result := vm.jsCall(desc.Getter, target, nil)
				if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				return result, false
			}
			if desc.HasValue {
				return desc.Value, false
			}
		}
		if obj, ok := vm.jsObjectItems[target.Num]; ok {
			if val, exists := obj[member]; exists {
				return val, false
			}
		}
		return Value{Type: VTJSUndefined}, false
	case VTJSPromise:
		if strings.EqualFold(member, "then") {
			return vm.jsCreateIntrinsicFunction("Promise.prototype.then", "PromisePrototypeThen"), false
		}
		if strings.EqualFold(member, "catch") {
			return vm.jsCreateIntrinsicFunction("Promise.prototype.catch", "PromisePrototypeCatch"), false
		}
		if strings.EqualFold(member, "finally") {
			return vm.jsCreateIntrinsicFunction("Promise.prototype.finally", "PromisePrototypeFinally"), false
		}
		return Value{Type: VTJSUndefined}, false
	case VTJSFunction:
		if strings.EqualFold(member, "length") {
			return NewInteger(int64(vm.jsFunctionExpectedLength(target))), false
		}
		desc, hasDesc := vm.jsResolveObjectMember(target.Num, member, make(map[int64]struct{}, 4))
		if hasDesc {
			if desc.HasGetter {
				result := vm.jsCall(desc.Getter, target, nil)
				if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				return result, false
			}
			if desc.HasValue {
				return desc.Value, false
			}
		}
		if obj, ok := vm.jsObjectItems[target.Num]; ok {
			if val, exists := obj[member]; exists {
				return val, false
			}
		}
		return Value{Type: VTJSUndefined}, false
	default:
		return Value{Type: VTJSUndefined}, false
	}
}

// jsPrepareMemberCallee resolves one member access into a callable callee and its receiver.
func (vm *VM) jsPrepareMemberCallee(target Value, member string) (Value, Value, bool, bool) {
	callee, deferred := vm.jsMemberGet(target, member)
	if deferred {
		return Value{Type: VTJSUndefined}, Value{Type: VTJSUndefined}, false, true
	}
	if !vm.jsIsCallable(callee) {
		return Value{Type: VTJSUndefined}, Value{Type: VTJSUndefined}, false, false
	}
	return callee, target, true, false
}

func (vm *VM) jsCallMember(target Value, member string, args []Value) (Value, bool) {
	if target.Type == VTJSFunction {
		switch {
		case strings.EqualFold(member, "call"):
			callThis := Value{Type: VTJSUndefined}
			callArgs := args
			if len(args) > 0 {
				callThis = args[0]
				callArgs = args[1:]
			}
			return vm.jsCall(target, callThis, callArgs), true
		case strings.EqualFold(member, "apply"):
			applyThis := Value{Type: VTJSUndefined}
			if len(args) > 0 {
				applyThis = args[0]
			}
			applyArgs := vm.jsExtractApplyArgs(jsArgOrUndefined(args, 1))
			return vm.jsCall(target, applyThis, applyArgs), true
		case strings.EqualFold(member, "bind"):
			bindThis := jsArgOrUndefined(args, 0)
			boundArgs := []Value(nil)
			if len(args) > 1 {
				boundArgs = args[1:]
			}
			return vm.jsBindFunction(target, bindThis, boundArgs), true
		}
		// Handle Buffer constructor static methods
		if vm.jsObjectStringProperty(target, "__js_ctor") == "Buffer" {
			if result, handled := vm.jsCallBufferMethod(member, args); handled {
				return result, true
			}
		}
	}

	if target.Type == VTJSObject {
		class := vm.jsObjectStringProperty(target, "__js_type")
		switch class {
		case "fs":
			if result, handled := vm.jsCallFSMethod(member, args); handled {
				return result, true
			}
		case "fs.promises":
			if result, handled := vm.jsCallFSPromisesMethod(member, args); handled {
				return result, true
			}
		case "crypto":
			if result, handled := vm.jsCallCryptoMethod(member, args); handled {
				return result, true
			}
		case "http":
			if result, handled := vm.jsCallHTTPMethod("http", member, args); handled {
				return result, true
			}
		case "https":
			if result, handled := vm.jsCallHTTPMethod("https", member, args); handled {
				return result, true
			}
		case "fs.Stats":
			if result, handled := vm.jsCallFSStatsMethod(target, member); handled {
				return result, true
			}
		case "crypto.Hash", "crypto.Hmac":
			if result, handled := vm.jsCallCryptoHashMethod(target, member, args); handled {
				return result, true
			}
		case "http.IncomingMessage":
			if result, handled := vm.jsCallHTTPResponseMethod(target, member, args); handled {
				return result, true
			}
		case "path":
			if result, handled := vm.jsCallPathMethod(member, args); handled {
				return result, true
			}
		case "os":
			if result, handled := vm.jsCallOSMethod(member, args); handled {
				return result, true
			}
		case "querystring":
			if result, handled := vm.jsCallQueryStringMethod(member, args); handled {
				return result, true
			}
		case "url":
			if result, handled := vm.jsCallURLModuleMethod(member, args); handled {
				return result, true
			}
		case "process":
			// Node.js process object methods
			if result, handled := vm.jsCallProcessMethod(member, args); handled {
				return result, true
			}
		case "__axon_stream":
			if result, handled := vm.jsCallNodeStreamHookMethod(member, args); handled {
				return result, true
			}
		case "Buffer":
			// Node.js Buffer constructor methods (static) or instance methods
			if target.Num == vm.nextDynamicNativeID || vm.jsObjectStringProperty(target, "__js_ctor") == "Buffer" {
				// Constructor methods (static)
				if result, handled := vm.jsCallBufferMethod(member, args); handled {
					return result, true
				}
			} else {
				// Instance methods
				if result, handled := vm.jsCallBufferInstanceMethod(target, member, args); handled {
					return result, true
				}
			}
		case "URL":
			if result, handled := vm.jsCallURLInstanceMethod(target, member, args); handled {
				return result, true
			}
		case "URLSearchParams":
			if result, handled := vm.jsCallURLSearchParamsMethod(target, member, args); handled {
				return result, true
			}
		case "Timeout":
			if result, handled := vm.jsCallTimeoutMethod(target, member, args); handled {
				return result, true
			}
		case "Array Iterator":
			if strings.EqualFold(member, "next") {
				return vm.jsArrayIteratorNext(target), true
			}
		case "String Iterator":
			if strings.EqualFold(member, "next") {
				return vm.jsStringIteratorNext(target), true
			}
		case "RegExp String Iterator":
			if strings.EqualFold(member, "next") {
				return vm.jsRegExpStringIteratorNext(target), true
			}
		case "Atomics":
			return vm.jsAtomicsCall(member, args)
		case "Uint8Array":
			switch {
			case strings.EqualFold(member, "toBase64"):
				buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(target)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
				data := buf[byteOffset : byteOffset+(byteLength/elemSize)]
				return NewString(base64.StdEncoding.EncodeToString(data)), true
			case strings.EqualFold(member, "toHex"):
				buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(target)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
				data := buf[byteOffset : byteOffset+(byteLength/elemSize)]
				return NewString(hex.EncodeToString(data)), true
			}
		}
	}

	if member == "slice" || member == "forEach" || member == "map" || member == "filter" || member == "at" || member == "findLast" || member == "toSorted" || member == "with" || member == "toReversed" || member == "toSpliced" || member == "flat" || member == "flatMap" ||
		strings.EqualFold(member, "slice") || strings.EqualFold(member, "forEach") || strings.EqualFold(member, "map") || strings.EqualFold(member, "filter") || strings.EqualFold(member, "at") || strings.EqualFold(member, "findLast") || strings.EqualFold(member, "toSorted") || strings.EqualFold(member, "with") || strings.EqualFold(member, "toReversed") || strings.EqualFold(member, "toSpliced") || strings.EqualFold(member, "flat") || strings.EqualFold(member, "flatMap") {
		length, isArrayLike, deferred := vm.jsArrayLikeLength(target)
		if deferred {
			return Value{Type: VTJSUndefined}, true
		}
		if isArrayLike {
			switch {
			case strings.EqualFold(member, "at"):
				idx := 0
				if len(args) > 0 {
					idx = int(vm.jsToNumber(args[0]).Flt)
				}
				if idx < 0 {
					idx = length + idx
				}
				if idx < 0 || idx >= length {
					return Value{Type: VTJSUndefined}, true
				}
				if v, ok := vm.jsArrayLikeGetIndex(target, idx); ok {
					return v, true
				}
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "slice"):
				start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
				end := length
				if len(args) > 1 && args[1].Type != VTJSUndefined {
					end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[1]).Flt), length)
				}
				if end < start {
					end = start
				}
				out := make([]Value, end-start)
				for i := start; i < end; i++ {
					if v, ok := vm.jsArrayLikeGetIndex(target, i); ok {
						out[i-start] = v
					} else {
						out[i-start] = Value{Type: VTJSUndefined}
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
			case strings.EqualFold(member, "forEach"):
				callback := jsArgOrUndefined(args, 0)
				if callback.Type != VTJSFunction {
					return Value{Type: VTJSUndefined}, true
				}
				thisArg := jsArgOrUndefined(args, 1)
				for i := range length {
					if !vm.jsArrayLikeHasIndex(target, i) {
						continue
					}
					item, _ := vm.jsArrayLikeGetIndex(target, i)
					_ = vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
				}
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "map"):
				callback := jsArgOrUndefined(args, 0)
				if callback.Type != VTJSFunction {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				thisArg := jsArgOrUndefined(args, 1)
				A := vm.jsArraySpeciesCreate(target, length)
				for i := range length {
					if !vm.jsArrayLikeHasIndex(target, i) {
						continue
					}
					item, _ := vm.jsArrayLikeGetIndex(target, i)
					result := vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
					vm.jsIndexSet(A, NewInteger(int64(i)), result)
				}
				return A, true
			case strings.EqualFold(member, "toReversed"):
				out := make([]Value, length)
				for i := range length {
					if v, ok := vm.jsArrayLikeGetIndex(target, length-1-i); ok {
						out[i] = v
					} else {
						out[i] = Value{Type: VTJSUndefined}
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
			case strings.EqualFold(member, "toSorted"):
				out := make([]Value, length)
				for i := range length {
					if v, ok := vm.jsArrayLikeGetIndex(target, i); ok {
						out[i] = v
					} else {
						out[i] = Value{Type: VTJSUndefined}
					}
				}
				compareFn := jsArgOrUndefined(args, 0)
				sort.Slice(out, func(i int, j int) bool {
					a := out[i]
					b := out[j]
					if compareFn.Type == VTJSFunction {
						res := vm.jsCall(compareFn, Value{Type: VTJSUndefined}, []Value{a, b})
						if res.Type == VTJSUndefined {
							return vm.valueToString(a) < vm.valueToString(b)
						}
						return vm.jsToNumber(res).Flt < 0
					}
					return vm.valueToString(a) < vm.valueToString(b)
				})
				return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
			case strings.EqualFold(member, "with"):
				index := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
				value := jsArgOrUndefined(args, 1)
				if index < 0 {
					index = length + index
				}
				if index < 0 || index >= length {
					vm.jsThrowRangeError("Invalid index for with()")
					return Value{Type: VTJSUndefined}, true
				}
				out := make([]Value, length)
				for i := range length {
					if i == index {
						out[i] = value
					} else if v, ok := vm.jsArrayLikeGetIndex(target, i); ok {
						out[i] = v
					} else {
						out[i] = Value{Type: VTJSUndefined}
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
			case strings.EqualFold(member, "toSpliced"):
				start := jsClampIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
				deleteCount := 0
				if len(args) == 0 {
					start = 0
				} else if len(args) == 1 {
					deleteCount = length - start
				} else {
					deleteCount = max(int(vm.jsToNumber(args[1]).Flt), 0)
					if start+deleteCount > length {
						deleteCount = length - start
					}
				}
				insertItems := []Value{}
				if len(args) > 2 {
					insertItems = args[2:]
				}

				newLen := length - deleteCount + len(insertItems)
				out := make([]Value, newLen)
				outIdx := 0
				for i := 0; i < start; i++ {
					if v, ok := vm.jsArrayLikeGetIndex(target, i); ok {
						out[outIdx] = v
					} else {
						out[outIdx] = Value{Type: VTJSUndefined}
					}
					outIdx++
				}
				for _, v := range insertItems {
					out[outIdx] = v
					outIdx++
				}
				for i := start + deleteCount; i < length; i++ {
					if v, ok := vm.jsArrayLikeGetIndex(target, i); ok {
						out[outIdx] = v
					} else {
						out[outIdx] = Value{Type: VTJSUndefined}
					}
					outIdx++
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
			case strings.EqualFold(member, "flat"):
				depth := 1
				if len(args) > 0 && args[0].Type != VTJSUndefined {
					depth = int(vm.jsToNumber(args[0]).Flt)
				}
				var flatten func(v Value, currentDepth int, out *[]Value)
				flatten = func(v Value, currentDepth int, out *[]Value) {
					vLen, isArrLike, _ := vm.jsArrayLikeLength(v)
					if isArrLike && currentDepth < depth {
						for i := range vLen {
							if item, ok := vm.jsArrayLikeGetIndex(v, i); ok {
								flatten(item, currentDepth+1, out)
							}
						}
					} else {
						*out = append(*out, v)
					}
				}
				out := make([]Value, 0, length)
				for i := range length {
					if item, ok := vm.jsArrayLikeGetIndex(target, i); ok {
						flatten(item, 0, &out)
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
			case strings.EqualFold(member, "flatMap"):
				callback := jsArgOrUndefined(args, 0)
				if callback.Type != VTJSFunction {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				thisArg := jsArgOrUndefined(args, 1)
				out := make([]Value, 0, length)
				for i := range length {
					if !vm.jsArrayLikeHasIndex(target, i) {
						continue
					}
					item, _ := vm.jsArrayLikeGetIndex(target, i)
					result := vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
					resLen, isArrLike, _ := vm.jsArrayLikeLength(result)
					if isArrLike {
						for j := range resLen {
							if subItem, ok := vm.jsArrayLikeGetIndex(result, j); ok {
								out = append(out, subItem)
							}
						}
					} else {
						out = append(out, result)
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
			case strings.EqualFold(member, "findLast"):
				callback := jsArgOrUndefined(args, 0)
				if callback.Type != VTJSFunction {
					return Value{Type: VTJSUndefined}, true
				}
				thisArg := jsArgOrUndefined(args, 1)
				for i := length - 1; i >= 0; i-- {
					if !vm.jsArrayLikeHasIndex(target, i) {
						continue
					}
					item, _ := vm.jsArrayLikeGetIndex(target, i)
					result := vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
					if vm.jsTruthy(result) {
						return item, true
					}
				}
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "filter"):
				callback := jsArgOrUndefined(args, 0)
				if callback.Type != VTJSFunction {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				thisArg := jsArgOrUndefined(args, 1)
				filtered := make([]Value, 0, length)
				for i := range length {
					if !vm.jsArrayLikeHasIndex(target, i) {
						continue
					}
					item, _ := vm.jsArrayLikeGetIndex(target, i)
					result := vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
					if vm.jsTruthy(result) {
						filtered = append(filtered, item)
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, filtered)), true
			}
		}
	}

	switch target.Type {
	case VTInteger, VTDouble:
		number := vm.jsToNumber(target).Flt
		switch {
		case strings.EqualFold(member, "toString"):
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				radix := int(vm.jsToNumber(args[0]).Flt)
				if radix < 2 || radix > 36 {
					return Value{Type: VTJSUndefined}, true
				}
				text := vm.jsNumberToStringRadix(number, radix)
				if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
					return Value{Type: VTJSUndefined}, true
				}
				return NewString(text), true
			}
			text := vm.jsNumberToString(number)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		case strings.EqualFold(member, "toLocaleString"):
			text := vm.jsNumberToString(number)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		case strings.EqualFold(member, "valueOf"):
			return NewDouble(number), true
		case strings.EqualFold(member, "toFixed"):
			digits := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
			if digits < 0 || digits > 20 {
				return Value{Type: VTJSUndefined}, true
			}
			text := vm.jsNumberToFixed(number, digits)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		case strings.EqualFold(member, "toExponential"):
			digits := -1
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				digits = int(vm.jsToNumber(args[0]).Flt)
				if digits < 0 || digits > 20 {
					return Value{Type: VTJSUndefined}, true
				}
			}
			text := vm.jsNumberToExponential(number, digits)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		case strings.EqualFold(member, "toPrecision"):
			if len(args) == 0 || args[0].Type == VTJSUndefined {
				text := vm.jsNumberToString(number)
				if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
					return Value{Type: VTJSUndefined}, true
				}
				return NewString(text), true
			}
			precision := int(vm.jsToNumber(args[0]).Flt)
			if precision < 1 || precision > 21 {
				return Value{Type: VTJSUndefined}, true
			}
			text := vm.jsNumberToPrecision(number, precision)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		}
	case VTString:
		text := target.Str
		runes := []rune(text)
		switch {
		case strings.EqualFold(member, "charAt"):
			idx := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
			if idx < 0 || idx >= len(runes) {
				return NewString(""), true
			}
			ch := string(runes[idx])
			if !vm.jsEnsureStringSize(len(ch)) || !vm.jsChargeStringWork(len(ch)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(ch), true
		case strings.EqualFold(member, "at"):
			length := len(runes)
			idx := 0
			if len(args) > 0 {
				idx = int(vm.jsToNumber(args[0]).Flt)
			}
			if idx < 0 {
				idx = length + idx
			}
			if idx < 0 || idx >= length {
				return Value{Type: VTJSUndefined}, true
			}
			ch := string(runes[idx])
			if !vm.jsEnsureStringSize(len(ch)) || !vm.jsChargeStringWork(len(ch)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(ch), true
		case strings.EqualFold(member, "charCodeAt"):
			idx := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
			if idx < 0 || idx >= len(runes) {
				return NewDouble(math.NaN()), true
			}
			return NewInteger(int64(runes[idx])), true
		case strings.EqualFold(member, "codePointAt"):
			position := vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt
			if math.IsNaN(position) {
				position = 0
			}
			if math.IsInf(position, 0) {
				return Value{Type: VTJSUndefined}, true
			}
			position = math.Trunc(position)
			if position < 0 {
				return Value{Type: VTJSUndefined}, true
			}
			units := utf16.Encode(runes)
			maxInt := float64(int(^uint(0) >> 1))
			if position > maxInt {
				return Value{Type: VTJSUndefined}, true
			}
			idx := int(position)
			if idx >= len(units) {
				return Value{Type: VTJSUndefined}, true
			}
			first := units[idx]
			if first >= 0xD800 && first <= 0xDBFF && idx+1 < len(units) {
				second := units[idx+1]
				if second >= 0xDC00 && second <= 0xDFFF {
					r := utf16.DecodeRune(rune(first), rune(second))
					return NewInteger(int64(r)), true
				}
			}
			return NewInteger(int64(first)), true
		case strings.EqualFold(member, "substring"):
			start := jsClampIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), len(runes))
			end := len(runes)
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				end = jsClampIndex(int(vm.jsToNumber(args[1]).Flt), len(runes))
			}
			if start > end {
				start, end = end, start
			}
			out := string(runes[start:end])
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "substr"):
			start := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
			if start < 0 {
				start = max(len(runes)+start, 0)
			}
			if start > len(runes) {
				start = len(runes)
			}
			length := len(runes) - start
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				length = max(int(vm.jsToNumber(args[1]).Flt), 0)
			}
			end := min(start+length, len(runes))
			out := string(runes[start:end])
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "slice"):
			start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), len(runes))
			end := len(runes)
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[1]).Flt), len(runes))
			}
			if end < start {
				end = start
			}
			out := string(runes[start:end])
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "concat"):
			total := len(text)
			parts := make([]string, 0, len(args)+1)
			parts = append(parts, text)
			for i := range args {
				part := vm.valueToString(args[i])
				total += len(part)
				parts = append(parts, part)
			}
			if !vm.jsEnsureStringSize(total) || !vm.jsChargeStringWork(total) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(strings.Join(parts, "")), true
		case strings.EqualFold(member, "match"):
			if len(args) == 0 {
				return NewNull(), true
			}
			reVal := args[0]
			if reVal.Type != VTJSObject || vm.jsObjectStringProperty(reVal, "__js_type") != "RegExp" {
				reVal = vm.jsNew(vm.jsCreateIntrinsicObject("", "RegExp"), []Value{reVal})
			}
			flags := vm.jsObjectStringProperty(reVal, "flags")
			if !strings.Contains(flags, "g") {
				return vm.jsRegExpExec(reVal, text), true
			}
			// Global match returns all matches in an array
			vm.jsMemberSet(reVal, "lastIndex", NewInteger(0))
			var matches []Value
			for {
				res := vm.jsRegExpExec(reVal, text)
				if res.Type == VTNull {
					break
				}
				matches = append(matches, vm.jsObjectSlots[res.Num][0])
				// Ensure progress if empty match
				lastIdxVal, _ := vm.jsMemberGet(reVal, "lastIndex")
				currIdxVal, _ := vm.jsMemberGet(res, "index")
				if int(vm.jsToNumber(lastIdxVal).Flt) == int(vm.jsToNumber(currIdxVal).Flt) {
					curr := int(vm.jsToNumber(lastIdxVal).Flt)
					vm.jsMemberSet(reVal, "lastIndex", NewInteger(int64(curr+1)))
				}
			}
			if len(matches) == 0 {
				return NewNull(), true
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, matches)), true
		case strings.EqualFold(member, "search"):
			if len(args) == 0 {
				return NewInteger(-1), true
			}
			reVal := args[0]
			if reVal.Type != VTJSObject || vm.jsObjectStringProperty(reVal, "__js_type") != "RegExp" {
				reVal = vm.jsNew(vm.jsCreateIntrinsicObject("", "RegExp"), []Value{reVal})
			}
			// search ignores global flag and lastIndex
			vm.jsMemberSet(reVal, "lastIndex", NewInteger(0))
			res := vm.jsRegExpExec(reVal, text)
			if res.Type == VTNull {
				return NewInteger(-1), true
			}
			idxVal, _ := vm.jsMemberGet(res, "index")
			return idxVal, true
		case strings.EqualFold(member, "toLowerCase"):
			out := strings.ToLower(text)
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "toUpperCase"):
			out := strings.ToUpper(text)
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "localeCompare"):
			other := vm.valueToString(jsArgOrUndefined(args, 0))
			cmp := strings.Compare(text, other)
			if cmp < 0 {
				return NewInteger(-1), true
			}
			if cmp > 0 {
				return NewInteger(1), true
			}
			return NewInteger(0), true
		case strings.EqualFold(member, "indexOf"):
			if len(args) == 0 {
				return NewInteger(-1), true
			}
			needle := vm.valueToString(args[0])
			start := 0
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
			}
			if start < 0 {
				start = 0
			}
			if start > len(text) {
				return NewInteger(-1), true
			}
			idx := strings.Index(text[start:], needle)
			if idx < 0 {
				return NewInteger(-1), true
			}
			return NewInteger(int64(start + idx)), true
		case strings.EqualFold(member, "split"):
			if len(args) == 0 || args[0].Type == VTJSUndefined {
				return ValueFromVBArray(NewVBArrayFromValues(0, []Value{NewString(text)})), true
			}
			sepVal := args[0]
			limit := -1
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				limit = max(int(vm.jsToNumber(args[1]).Flt), 0)
			}

			if sepVal.Type == VTJSObject && vm.jsObjectStringProperty(sepVal, "__js_type") == "RegExp" {
				re, err := vm.jsGetCompiledRegExp(sepVal.Num)
				if err != nil {
					return ValueFromVBArray(NewVBArrayFromValues(0, []Value{NewString(text)})), true
				}
				var values []Value
				lastByteIdx := 0
				m, err := re.FindStringMatch(text)
				for m != nil && (limit < 0 || len(values) < limit) {
					startByte := vm.jsRuneToByteOffset(text, m.Index)
					lengthByte := vm.jsRuneToByteOffset(text[startByte:], m.Length)
					values = append(values, NewString(text[lastByteIdx:startByte]))

					groups := m.Groups()
					for i := 1; i < len(groups); i++ {
						if limit >= 0 && len(values) >= limit {
							break
						}
						g := groups[i]
						if g.Capture.Length < 0 {
							values = append(values, Value{Type: VTJSUndefined})
						} else {
							values = append(values, NewString(g.String()))
						}
					}

					lastByteIdx = startByte + lengthByte
					if lengthByte == 0 {
						// Ensure progress for zero-length match
						if lastByteIdx < len(text) {
							// We need to advance by one rune
							_, size := utf8.DecodeRuneInString(text[lastByteIdx:])
							lastByteIdx += size
						} else {
							break
						}
					}
					m, err = re.FindNextMatch(m)
				}
				if limit < 0 || len(values) < limit {
					if lastByteIdx <= len(text) {
						values = append(values, NewString(text[lastByteIdx:]))
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
			}

			sep := vm.valueToString(sepVal)
			var pieces []string
			if sep == "" {
				pieces = make([]string, 0, len(text))
				for _, r := range text {
					pieces = append(pieces, string(r))
				}
			} else {
				pieces = strings.Split(text, sep)
			}
			if limit >= 0 && len(pieces) > limit {
				pieces = pieces[:limit]
			}
			values := make([]Value, len(pieces))
			for i := range pieces {
				values[i] = NewString(pieces[i])
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
		case strings.EqualFold(member, "replace"):
			if len(args) == 0 {
				return NewString(text), true
			}
			replacementArg := jsArgOrUndefined(args, 1)
			return vm.jsStringReplace(text, args[0], replacementArg, false), true
		case strings.EqualFold(member, "replaceAll"):
			if len(args) == 0 {
				return NewString(text), true
			}
			replacementArg := jsArgOrUndefined(args, 1)
			return vm.jsStringReplace(text, args[0], replacementArg, true), true
		case strings.EqualFold(member, "trim"):
			trimmed := strings.TrimSpace(text)
			if !vm.jsEnsureStringSize(len(trimmed)) || !vm.jsChargeStringWork(len(trimmed)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(trimmed), true
		case strings.EqualFold(member, "normalize"):
			form := "NFC"
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				form = vm.valueToString(args[0])
			}
			var normalized string
			switch form {
			case "NFC":
				normalized = norm.NFC.String(text)
			case "NFD":
				normalized = norm.NFD.String(text)
			case "NFKC":
				normalized = norm.NFKC.String(text)
			case "NFKD":
				normalized = norm.NFKD.String(text)
			default:
				vm.jsThrowRangeError("The normalization form should be one of NFC, NFD, NFKC, NFKD")
				return Value{Type: VTJSUndefined}, true
			}
			if !vm.jsEnsureStringSize(len(normalized)) || !vm.jsChargeStringWork(len(normalized)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(normalized), true
		// ES6 String.prototype additions
		case strings.EqualFold(member, "includes"):
			if len(args) == 0 {
				return NewBool(false), true
			}
			if args[0].Type == VTJSObject && vm.jsObjectStringProperty(args[0], "__js_type") == "RegExp" {
				vm.jsThrowTypeError("First argument to String.prototype.includes must not be a RegExp")
				return Value{Type: VTJSUndefined}, true
			}
			needle := vm.valueToString(args[0])
			start := 0
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
			}
			if start < 0 {
				start = 0
			}
			if start > len(text) {
				return NewBool(false), true
			}
			return NewBool(strings.Contains(text[start:], needle)), true
		case strings.EqualFold(member, "startsWith"):
			if len(args) == 0 {
				return NewBool(false), true
			}
			needle := vm.valueToString(args[0])
			start := 0
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
			}
			if start < 0 {
				start = 0
			}
			if start > len(text) {
				return NewBool(false), true
			}
			return NewBool(strings.HasPrefix(text[start:], needle)), true
		case strings.EqualFold(member, "endsWith"):
			if len(args) == 0 {
				return NewBool(false), true
			}
			needle := vm.valueToString(args[0])
			end := len(text)
			if len(args) > 1 {
				e := int(vm.jsToNumber(args[1]).Flt)
				if e < end {
					end = e
				}
			}
			if end < 0 {
				end = 0
			}
			substr := text
			if end < len(text) {
				substr = text[:end]
			}
			return NewBool(strings.HasSuffix(substr, needle)), true
		case strings.EqualFold(member, "repeat"):
			if len(args) == 0 {
				return NewString(""), true
			}
			count := int(vm.jsToNumber(args[0]).Flt)
			if count < 0 {
				vm.jsThrowTypeError("Invalid count value for String.prototype.repeat")
				return Value{Type: VTJSUndefined}, true
			}
			if count == 0 || text == "" {
				return NewString(""), true
			}
			totalSize := len(text) * count
			if !vm.jsEnsureStringSize(totalSize) || !vm.jsChargeStringWork(totalSize) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(strings.Repeat(text, count)), true
		case strings.EqualFold(member, "padStart"):
			targetLen := 0
			if len(args) > 0 {
				targetLen = int(vm.jsToNumber(args[0]).Flt)
			}
			padStr := " "
			if len(args) > 1 {
				padStr = vm.valueToString(args[1])
			}
			if targetLen <= len(text) || padStr == "" {
				return NewString(text), true
			}
			needed := targetLen - len(text)
			padded := strings.Repeat(padStr, (needed/len(padStr))+1)[:needed] + text
			if !vm.jsEnsureStringSize(len(padded)) || !vm.jsChargeStringWork(len(padded)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(padded), true
		case strings.EqualFold(member, "padEnd"):
			targetLen := 0
			if len(args) > 0 {
				targetLen = int(vm.jsToNumber(args[0]).Flt)
			}
			padStr := " "
			if len(args) > 1 {
				padStr = vm.valueToString(args[1])
			}
			if targetLen <= len(text) || padStr == "" {
				return NewString(text), true
			}
			needed := targetLen - len(text)
			padded := text + strings.Repeat(padStr, (needed/len(padStr))+1)[:needed]
			if !vm.jsEnsureStringSize(len(padded)) || !vm.jsChargeStringWork(len(padded)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(padded), true
		}
	case VTArray:
		if target.Arr == nil {
			return Value{Type: VTJSUndefined}, true
		}
		switch {
		case strings.EqualFold(member, "keys"):
			return vm.jsCreateArrayIterator(target, 1), true
		case strings.EqualFold(member, "entries"):
			return vm.jsCreateArrayIterator(target, 2), true
		case strings.EqualFold(member, "values"):
			return vm.jsCreateArrayIterator(target, 0), true
		case strings.EqualFold(member, "fill"):
			fillValue := jsArgOrUndefined(args, 0)
			length := len(target.Arr.Values)
			start := 0
			if len(args) > 1 {
				start = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[1]).Flt), length)
			}
			end := length
			if len(args) > 2 && args[2].Type != VTJSUndefined {
				end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[2]).Flt), length)
			}
			if end < start {
				end = start
			}
			for i := start; i < end; i++ {
				target.Arr.Values[i] = fillValue
			}
			return target, true
		case strings.EqualFold(member, "at"):
			length := len(target.Arr.Values)
			idx := 0
			if len(args) > 0 {
				idx = int(vm.jsToNumber(args[0]).Flt)
			}
			if idx < 0 {
				idx = length + idx
			}
			if idx < 0 || idx >= length {
				return Value{Type: VTJSUndefined}, true
			}
			return target.Arr.Values[idx], true
		case strings.EqualFold(member, "flat"):
			depth := 1
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				num := vm.jsToNumber(args[0])
				if math.IsInf(num.Flt, 1) {
					depth = 1000 // effectively infinity
				} else {
					depth = int(num.Flt)
				}
			}
			if depth < 0 {
				depth = 0
			}
			flattened := vm.jsArrayFlat(target.Arr.Values, depth)
			return ValueFromVBArray(NewVBArrayFromValues(0, flattened)), true
		case strings.EqualFold(member, "flatMap"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				vm.jsThrowTypeError("flatMap callback must be a function")
				return Value{Type: VTJSUndefined}, true
			}
			thisArg := jsArgOrUndefined(args, 1)
			length := len(target.Arr.Values)
			out := make([]Value, 0, length)
			for i := range length {
				item := target.Arr.Values[i]
				res := vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
				if res.Type == VTArray && res.Arr != nil {
					out = append(out, res.Arr.Values...)
				} else {
					out = append(out, res)
				}
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
		case strings.EqualFold(member, "toSorted"):
			length := len(target.Arr.Values)
			newVals := make([]Value, length)
			copy(newVals, target.Arr.Values)
			compareFn := jsArgOrUndefined(args, 0)
			sort.Slice(newVals, func(i, j int) bool {
				a := newVals[i]
				b := newVals[j]
				if compareFn.Type == VTJSFunction {
					res := vm.jsCall(compareFn, Value{Type: VTJSUndefined}, []Value{a, b})
					return vm.jsToNumber(res).Flt < 0
				}
				return vm.valueToString(a) < vm.valueToString(b)
			})
			return ValueFromVBArray(NewVBArrayFromValues(0, newVals)), true
		case strings.EqualFold(member, "toReversed"):
			length := len(target.Arr.Values)
			newVals := make([]Value, length)
			for i := range length {
				newVals[i] = target.Arr.Values[length-1-i]
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, newVals)), true
		case strings.EqualFold(member, "toSpliced"):
			length := len(target.Arr.Values)
			start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
			deleteCount := length - start
			if len(args) > 1 {
				deleteCount = max(int(vm.jsToNumber(args[1]).Flt), 0)
				if deleteCount > length-start {
					deleteCount = length - start
				}
			}
			insertItems := []Value(nil)
			if len(args) > 2 {
				insertItems = args[2:]
			}
			newVals := make([]Value, 0, length-deleteCount+len(insertItems))
			newVals = append(newVals, target.Arr.Values[:start]...)
			newVals = append(newVals, insertItems...)
			newVals = append(newVals, target.Arr.Values[start+deleteCount:]...)
			return ValueFromVBArray(NewVBArrayFromValues(0, newVals)), true
		case strings.EqualFold(member, "copyWithin"):
			length := len(target.Arr.Values)
			to := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
			from := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 1)).Flt), length)
			end := length
			if len(args) > 2 && args[2].Type != VTJSUndefined {
				end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[2]).Flt), length)
			}
			if end < from {
				end = from
			}
			count := min(end-from, length-to)
			if count > 0 {
				copy(target.Arr.Values[to:to+count], target.Arr.Values[from:from+count])
			}
			return target, true
		case strings.EqualFold(member, "slice"):
			length := len(target.Arr.Values)
			start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
			end := length
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[1]).Flt), length)
			}
			if end < start {
				end = start
			}
			outLen := end - start
			A := vm.jsArraySpeciesCreate(target, outLen)
			for i := range outLen {
				vm.jsIndexSet(A, NewInteger(int64(i)), target.Arr.Values[start+i])
			}
			return A, true
		case strings.EqualFold(member, "splice"):
			length := len(target.Arr.Values)
			start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
			deleteCount := length - start
			if len(args) > 1 {
				deleteCount = max(int(vm.jsToNumber(args[1]).Flt), 0)
				if deleteCount > length-start {
					deleteCount = length - start
				}
			}
			removed := make([]Value, deleteCount)
			copy(removed, target.Arr.Values[start:start+deleteCount])
			insertItems := args[2:]
			newVals := make([]Value, 0, length-deleteCount+len(insertItems))
			newVals = append(newVals, target.Arr.Values[:start]...)
			newVals = append(newVals, insertItems...)
			newVals = append(newVals, target.Arr.Values[start+deleteCount:]...)
			target.Arr.Values = newVals
			return ValueFromVBArray(NewVBArrayFromValues(0, removed)), true
		case strings.EqualFold(member, "shift"):
			if len(target.Arr.Values) == 0 {
				return Value{Type: VTJSUndefined}, true
			}
			first := target.Arr.Values[0]
			target.Arr.Values = target.Arr.Values[1:]
			return first, true
		case strings.EqualFold(member, "unshift"):
			if len(args) == 0 {
				return NewInteger(int64(len(target.Arr.Values))), true
			}
			newVals := make([]Value, 0, len(args)+len(target.Arr.Values))
			newVals = append(newVals, args...)
			newVals = append(newVals, target.Arr.Values...)
			target.Arr.Values = newVals
			return NewInteger(int64(len(target.Arr.Values))), true
		case strings.EqualFold(member, "reverse"):
			for i, j := 0, len(target.Arr.Values)-1; i < j; i, j = i+1, j-1 {
				target.Arr.Values[i], target.Arr.Values[j] = target.Arr.Values[j], target.Arr.Values[i]
			}
			return target, true
		case strings.EqualFold(member, "sort"):
			compareFn := jsArgOrUndefined(args, 0)
			sort.Slice(target.Arr.Values, func(i int, j int) bool {
				a := target.Arr.Values[i]
				b := target.Arr.Values[j]
				if compareFn.Type == VTJSFunction {
					res := vm.jsCall(compareFn, Value{Type: VTJSUndefined}, []Value{a, b})
					if res.Type == VTJSUndefined {
						return vm.valueToString(a) < vm.valueToString(b)
					}
					return vm.jsToNumber(res).Flt < 0
				}
				return vm.valueToString(a) < vm.valueToString(b)
			})
			return target, true
		case strings.EqualFold(member, "concat"):
			out := make([]Value, 0, len(target.Arr.Values)+len(args))
			out = append(out, target.Arr.Values...)
			for i := range args {
				arg := args[i]
				if vm.jsIsConcatSpreadable(arg) {
					// Extract values from array or array-like
					var values []Value
					if arrLike, ok := vm.jsAsConcatArray(arg); ok && arrLike.Arr != nil {
						values = arrLike.Arr.Values
					} else if arg.Type == VTArray && arg.Arr != nil {
						values = arg.Arr.Values
					} else {
						// It could be a bridge object or a JS object marked as spreadable
						lenVal, _ := vm.jsMemberGet(arg, "length")
						length := max(int(vm.jsToNumber(lenVal).Flt), 0)
						values = make([]Value, length)
						for j := 0; j < length; j++ {
							if vm.jsArrayLikeHasIndex(arg, j) {
								item, _ := vm.jsArrayLikeGetIndex(arg, j)
								values[j] = item
							} else {
								values[j] = Value{Type: VTJSUndefined}
							}
						}
					}
					out = append(out, values...)
				} else {
					out = append(out, arg)
				}
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
		case strings.EqualFold(member, "indexOf"):
			if len(target.Arr.Values) == 0 {
				return NewInteger(-1), true
			}
			needle := jsArgOrUndefined(args, 0)
			start := 0
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
			}
			if start < 0 {
				start = max(len(target.Arr.Values)+start, 0)
			}
			for i := start; i < len(target.Arr.Values); i++ {
				if vm.jsStrictEquals(target.Arr.Values[i], needle) {
					return NewInteger(int64(i)), true
				}
			}
			return NewInteger(-1), true
		case strings.EqualFold(member, "find"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return Value{Type: VTJSUndefined}, true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				callResult := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if vm.jsTruthy(callResult) {
					return target.Arr.Values[i], true
				}
			}
			return Value{Type: VTJSUndefined}, true
		case strings.EqualFold(member, "findIndex"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return NewInteger(-1), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				callResult := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if vm.jsTruthy(callResult) {
					return NewInteger(int64(i)), true
				}
			}
			return NewInteger(-1), true
		case strings.EqualFold(member, "lastIndexOf"):
			if len(target.Arr.Values) == 0 {
				return NewInteger(-1), true
			}
			needle := jsArgOrUndefined(args, 0)
			start := len(target.Arr.Values) - 1
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
				if start >= len(target.Arr.Values) {
					start = len(target.Arr.Values) - 1
				}
			}
			if start < 0 {
				return NewInteger(-1), true
			}
			for i := start; i >= 0; i-- {
				if vm.jsStrictEquals(target.Arr.Values[i], needle) {
					return NewInteger(int64(i)), true
				}
			}
			return NewInteger(-1), true
		case strings.EqualFold(member, "push"):
			target.Arr.Values = append(target.Arr.Values, args...)
			return NewInteger(int64(len(target.Arr.Values))), true
		case strings.EqualFold(member, "__spreadPush"):
			if len(args) == 0 {
				return NewInteger(int64(len(target.Arr.Values))), true
			}
			source := args[0]
			if source.Type == VTJSUndefined || source.Type == VTNull {
				vm.jsThrowTypeError("Spread source cannot be null or undefined")
				return Value{Type: VTJSUndefined}, true
			}
			length, hasLength, deferred := vm.jsArrayLikeLength(source)
			if deferred {
				return Value{Type: VTJSUndefined}, true
			}
			if !hasLength {
				vm.jsThrowTypeError("Spread source must be array-like")
				return Value{Type: VTJSUndefined}, true
			}
			if length <= 0 {
				return NewInteger(int64(len(target.Arr.Values))), true
			}
			start := len(target.Arr.Values)
			target.Arr.Values = append(target.Arr.Values, make([]Value, length)...)
			for i := range length {
				if v, ok := vm.jsArrayLikeGetIndex(source, i); ok {
					target.Arr.Values[start+i] = v
				} else {
					target.Arr.Values[start+i] = Value{Type: VTJSUndefined}
				}
			}
			return NewInteger(int64(len(target.Arr.Values))), true
		case strings.EqualFold(member, "pop"):
			if len(target.Arr.Values) == 0 {
				return Value{Type: VTJSUndefined}, true
			}
			last := target.Arr.Values[len(target.Arr.Values)-1]
			target.Arr.Values = target.Arr.Values[:len(target.Arr.Values)-1]
			return last, true
		case strings.EqualFold(member, "join"):
			sep := ","
			if len(args) > 0 {
				sep = vm.valueToString(args[0])
			}
			if len(target.Arr.Values) == 0 {
				return NewString(""), true
			}
			parts := make([]string, len(target.Arr.Values))
			totalSize := 0
			for i := range target.Arr.Values {
				parts[i] = vm.jsConcatString(target.Arr.Values[i])
				totalSize += len(parts[i])
				if i > 0 {
					totalSize += len(sep)
				}
				if !vm.jsEnsureStringSize(totalSize) {
					return Value{Type: VTJSUndefined}, true
				}
			}
			if !vm.jsChargeStringWork(totalSize) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(strings.Join(parts, sep)), true
		case strings.EqualFold(member, "forEach"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return Value{Type: VTJSUndefined}, true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				_ = vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
			}
			return Value{Type: VTJSUndefined}, true
		case strings.EqualFold(member, "every"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return NewBool(false), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if !vm.jsTruthy(result) {
					return NewBool(false), true
				}
			}
			return NewBool(true), true
		case strings.EqualFold(member, "some"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return NewBool(false), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if vm.jsTruthy(result) {
					return NewBool(true), true
				}
			}
			return NewBool(false), true
		case strings.EqualFold(member, "map"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			length := len(target.Arr.Values)
			A := vm.jsArraySpeciesCreate(target, length)
			for i := range length {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				vm.jsIndexSet(A, NewInteger(int64(i)), result)
			}
			return A, true
		case strings.EqualFold(member, "filter"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			filtered := make([]Value, 0, len(target.Arr.Values))
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if vm.jsTruthy(result) {
					filtered = append(filtered, target.Arr.Values[i])
				}
			}
			A := vm.jsArraySpeciesCreate(target, len(filtered))
			for i := 0; i < len(filtered); i++ {
				vm.jsIndexSet(A, NewInteger(int64(i)), filtered[i])
			}
			return A, true
		case strings.EqualFold(member, "reduce"):
			if len(target.Arr.Values) == 0 && len(args) < 2 {
				return Value{Type: VTJSUndefined}, true
			}
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return Value{Type: VTJSUndefined}, true
			}
			acc := Value{Type: VTJSUndefined}
			start := 0
			if len(args) > 1 {
				acc = args[1]
			} else {
				acc = target.Arr.Values[0]
				start = 1
			}
			for i := start; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, Value{Type: VTJSUndefined}, []Value{acc, target.Arr.Values[i], NewInteger(int64(i)), target})
				acc = result
			}
			return acc, true
		case strings.EqualFold(member, "reduceRight"):
			if len(target.Arr.Values) == 0 && len(args) < 2 {
				return Value{Type: VTJSUndefined}, true
			}
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return Value{Type: VTJSUndefined}, true
			}
			acc := Value{Type: VTJSUndefined}
			start := len(target.Arr.Values) - 1
			if len(args) > 1 {
				acc = args[1]
			} else {
				acc = target.Arr.Values[start]
				start--
			}
			for i := start; i >= 0; i-- {
				result := vm.jsCall(callback, Value{Type: VTJSUndefined}, []Value{acc, target.Arr.Values[i], NewInteger(int64(i)), target})
				acc = result
			}
			return acc, true
		}
	case VTDate:
		switch {
		case strings.EqualFold(member, "getFullYear"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Year())), true
		case strings.EqualFold(member, "getMonth"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(int(t.Month()) - 1)), true
		case strings.EqualFold(member, "getDate"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Day())), true
		case strings.EqualFold(member, "getDay"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Weekday())), true
		case strings.EqualFold(member, "getHours"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Hour())), true
		case strings.EqualFold(member, "getMinutes"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Minute())), true
		case strings.EqualFold(member, "getSeconds"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Second())), true
		case strings.EqualFold(member, "getMilliseconds"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Nanosecond() / int(time.Millisecond))), true
		case strings.EqualFold(member, "getTimezoneOffset"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			_, offsetSeconds := t.Zone()
			return NewInteger(int64(-(offsetSeconds / 60))), true
		case strings.EqualFold(member, "getTime"):
			t := valueToTimeInLocale(vm, target)
			return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
		case strings.EqualFold(member, "now"):
			return NewInteger(time.Now().UnixNano() / int64(time.Millisecond)), true
		case strings.EqualFold(member, "toString"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewString(t.Format("Mon Jan 02 2006 15:04:05 GMT-0700")), true
		case strings.EqualFold(member, "toLocaleString"):
			return NewString(vm.dateToLocalizedString(target)), true
		case strings.EqualFold(member, "toUTCString"):
			t := valueToTimeInLocale(vm, target).UTC()
			return NewString(t.Format(time.RFC1123)), true
		case strings.EqualFold(member, "toISOString"):
			t := valueToTimeInLocale(vm, target).UTC()
			return NewString(t.Format(time.RFC3339)), true
		case strings.EqualFold(member, "toJSON"):
			result, _ := vm.jsCallMember(target, "toISOString", nil)
			return result, true
		case strings.EqualFold(member, "valueOf"):
			t := valueToTimeInLocale(vm, target)
			return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
		}
	case VTJSPromise:
		switch {
		case strings.EqualFold(member, "then"):
			return vm.jsPromiseThen(target, args), true
		case strings.EqualFold(member, "catch"):
			return vm.jsPromiseCatch(target, args), true
		case strings.EqualFold(member, "finally"):
			return vm.jsPromiseFinally(target, args), true
		}
	case VTJSObject:
		switch {
		case strings.EqualFold(member, "hasOwnProperty"):
			return NewBool(vm.jsObjectHasOwnProperty(target, vm.valueToString(jsArgOrUndefined(args, 0)))), true
		case strings.EqualFold(member, "propertyIsEnumerable"):
			return NewBool(vm.jsObjectPropertyIsEnumerable(target, vm.valueToString(jsArgOrUndefined(args, 0)))), true
		case strings.EqualFold(member, "isPrototypeOf"):
			return NewBool(vm.jsObjectIsPrototypeOf(target, jsArgOrUndefined(args, 0))), true
		case strings.EqualFold(member, "toString"):
			return NewString(vm.jsObjectToStringTag(target)), true
		case strings.EqualFold(member, "toLocaleString"):
			return NewString(vm.jsObjectToStringTag(target)), true
		case strings.EqualFold(member, "valueOf"):
			return target, true
		}
		objType := vm.jsObjectStringProperty(target, "__js_type")
		if objType == "" {
			objType = vm.jsObjectStringProperty(target, "__js_ctor")
		}
		switch objType {
		case "Date":
			switch {
			case strings.EqualFold(member, "now"):
				return NewInteger(time.Now().UnixNano() / int64(time.Millisecond)), true
			case strings.EqualFold(member, "parse"):
				if len(args) == 0 {
					return NewDouble(math.NaN()), true
				}
				t := valueToTimeInLocale(vm, args[0])
				if t.IsZero() {
					return NewDouble(math.NaN()), true
				}
				return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
			case strings.EqualFold(member, "UTC"):
				year := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
				month := int(vm.jsToNumber(jsArgOrUndefined(args, 1)).Flt)
				day := 1
				hour := 0
				minute := 0
				second := 0
				millisecond := 0
				if len(args) > 2 {
					day = int(vm.jsToNumber(args[2]).Flt)
				}
				if len(args) > 3 {
					hour = int(vm.jsToNumber(args[3]).Flt)
				}
				if len(args) > 4 {
					minute = int(vm.jsToNumber(args[4]).Flt)
				}
				if len(args) > 5 {
					second = int(vm.jsToNumber(args[5]).Flt)
				}
				if len(args) > 6 {
					millisecond = int(vm.jsToNumber(args[6]).Flt)
				}
				t := time.Date(year, time.Month(month+1), day, hour, minute, second, millisecond*int(time.Millisecond), time.UTC)
				return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
			}
		case "Array":
			switch {
			case strings.EqualFold(member, "isArray"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				return NewBool(args[0].Type == VTArray), true
			case strings.EqualFold(member, "of"):
				if len(args) == 0 {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				vals := make([]Value, len(args))
				copy(vals, args)
				return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
			case strings.EqualFold(member, "from"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Array.from requires an array-like source")
					return Value{Type: VTJSUndefined}, true
				}
				source := args[0]
				switch source.Type {
				case VTArray:
					if source.Arr == nil {
						return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
					}
					vals := make([]Value, len(source.Arr.Values))
					copy(vals, source.Arr.Values)
					return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
				case VTString:
					r := []rune(source.Str)
					vals := make([]Value, len(r))
					for i := range r {
						vals[i] = NewString(string(r[i]))
					}
					return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
				default:
					length, hasLength, deferred := vm.jsArrayLikeLength(source)
					if deferred {
						return Value{Type: VTJSUndefined}, true
					}
					if !hasLength {
						vm.jsThrowTypeError("Array.from source is not array-like")
						return Value{Type: VTJSUndefined}, true
					}
					vals := make([]Value, length)
					for i := range length {
						if v, ok := vm.jsArrayLikeGetIndex(source, i); ok {
							vals[i] = v
						} else {
							vals[i] = Value{Type: VTJSUndefined}
						}
					}
					return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
				}
			}
		case "String":
			switch {
			case strings.EqualFold(member, "fromCodePoint"):
				// String.fromCodePoint(...codePoints)
				// Converts one or more code points to a string
				if len(args) == 0 {
					return NewString(""), true
				}
				var runes []rune
				for i := range args {
					codePointVal := vm.jsToNumber(args[i])
					codePoint := int64(codePointVal.Flt)
					// Convert to integer by truncating towards zero
					if math.IsNaN(codePointVal.Flt) {
						codePoint = 0
					} else if codePointVal.Flt < 0 || codePointVal.Flt > 0x10FFFF {
						// Out of range
						vm.jsThrowRangeError(fmt.Sprintf("Invalid code point %d", codePoint))
						return Value{Type: VTJSUndefined}, true
					} else if int64(codePointVal.Flt) != codePoint {
						// Not an integer value
						vm.jsThrowRangeError(fmt.Sprintf("Invalid code point %v", codePointVal.Flt))
						return Value{Type: VTJSUndefined}, true
					}
					runes = append(runes, rune(codePoint))
				}
				result := string(runes)
				if !vm.jsEnsureStringSize(len(result)) || !vm.jsChargeStringWork(len(result)) {
					return Value{Type: VTJSUndefined}, true
				}
				return NewString(result), true
			}
		case "Object":
			switch {
			case strings.EqualFold(member, "is"):
				if len(args) < 2 {
					return NewBool(false), true
				}
				a := args[0]
				b := args[1]
				if (a.Type == VTInteger || a.Type == VTDouble) && (b.Type == VTInteger || b.Type == VTDouble) {
					aNum := vm.jsToNumber(a).Flt
					bNum := vm.jsToNumber(b).Flt
					if math.IsNaN(aNum) && math.IsNaN(bNum) {
						return NewBool(true), true
					}
					if aNum == 0 && bNum == 0 {
						return NewBool(math.Signbit(aNum) == math.Signbit(bNum)), true
					}
					return NewBool(aNum == bNum), true
				}
				return NewBool(vm.jsStrictEquals(a, b)), true
			case strings.EqualFold(member, "setPrototypeOf"):
				if len(args) < 2 {
					vm.jsThrowTypeError("Object.setPrototypeOf requires an object and a prototype")
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.setPrototypeOf called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type != VTJSObject && args[0].Type != VTJSFunction && args[0].Type != VTJSProxy && args[0].Type != VTArray {
					return args[0], true
				}
				// Use ReflectSetPrototypeOf logic
				success := vm.dispatchJSIntrinsicCall(Value{}, "ReflectSetPrototypeOf", args)
				if !vm.asBool(success) {
					vm.jsThrowTypeError("Object.setPrototypeOf failed")
				}
				return args[0], true
			case strings.EqualFold(member, "create"):
				objID := vm.allocJSID()
				obj := make(map[string]Value, 8)
				vm.jsObjectItems[objID] = obj
				vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
				created := Value{Type: VTJSObject, Num: objID}
				if len(args) > 0 {
					vm.jsSetProto(created, args[0])
				}
				if len(args) > 1 {
					_, _ = vm.jsCallMember(target, "defineProperties", []Value{created, args[1]})
				}
				return created, true
			case strings.EqualFold(member, "getPrototypeOf"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.getPrototypeOf called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type != VTJSObject && args[0].Type != VTJSFunction && args[0].Type != VTJSProxy && args[0].Type != VTArray {
					// ES6+: primitive values return their prototype
					if proto := vm.jsGetIntrinsicPrototype(vm.jsPrimitiveTypeName(args[0])); proto.Type == VTJSObject {
						return proto, true
					}
					return NewNull(), true
				}
				if args[0].Type == VTJSProxy {
					return vm.jsProxyGetPrototypeOf(args[0]), true
				}
				return vm.jsGetPrototypeValue(args[0]), true
			case strings.EqualFold(member, "fromEntries"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.fromEntries requires an iterable")
					return Value{Type: VTJSUndefined}, true
				}
				iterable := args[0]
				objID := vm.allocJSID()
				obj := make(map[string]Value, 8)
				vm.jsObjectItems[objID] = obj
				vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
				if iterable.Type == VTArray && iterable.Arr != nil {
					for _, entry := range iterable.Arr.Values {
						if entry.Type == VTArray && entry.Arr != nil && len(entry.Arr.Values) >= 2 {
							key := vm.valueToString(entry.Arr.Values[0])
							val := entry.Arr.Values[1]
							obj[key] = val
							vm.jsSetDescriptor(objID, key, jsDefaultPropertyDescriptor(val))
						}
					}
				}
				return Value{Type: VTJSObject, Num: objID}, true
			case strings.EqualFold(member, "keys"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.keys called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				keys := make([]string, 0)
				source := args[0]
				switch source.Type {
				case VTJSObject, VTJSFunction:
					keys = vm.jsObjectOwnEnumerableKeys(source.Num)
				case VTJSProxy:
					rawKeys := vm.jsProxyOwnKeys(source)
					keys = make([]string, 0, len(rawKeys))
					for _, k := range rawKeys {
						desc, ok := vm.jsGetDescriptor(source.Num, k)
						if ok && desc.Enumerable {
							keys = append(keys, k)
						}
					}
				case VTArray:
					if source.Arr != nil {
						keys = make([]string, len(source.Arr.Values))
						for i := 0; i < len(source.Arr.Values); i++ {
							keys[i] = strconv.Itoa(i)
						}
					}
				case VTString:
					runes := []rune(source.Str)
					keys = make([]string, len(runes))
					for i := range runes {
						keys[i] = strconv.Itoa(i)
					}
				}
				values := make([]Value, len(keys))
				for i := 0; i < len(keys); i++ {
					values[i] = NewString(keys[i])
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
			case strings.EqualFold(member, "values"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.values called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				source := args[0]
				values := make([]Value, 0)
				switch source.Type {
				case VTJSObject, VTJSFunction:
					keys := vm.jsObjectOwnEnumerableKeys(source.Num)
					values = make([]Value, len(keys))
					for i := range keys {
						v, deferred := vm.jsMemberGet(source, keys[i])
						if deferred {
							return Value{Type: VTJSUndefined}, true
						}
						values[i] = v
					}
				case VTArray:
					if source.Arr != nil {
						values = make([]Value, len(source.Arr.Values))
						copy(values, source.Arr.Values)
					}
				case VTString:
					runes := []rune(source.Str)
					values = make([]Value, len(runes))
					for i := range runes {
						values[i] = NewString(string(runes[i]))
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
			case strings.EqualFold(member, "entries"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.entries called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				source := args[0]
				keys := make([]string, 0)
				vals := make([]Value, 0)
				switch source.Type {
				case VTJSObject, VTJSFunction:
					keys = vm.jsObjectOwnEnumerableKeys(source.Num)
					vals = make([]Value, len(keys))
					for i := 0; i < len(keys); i++ {
						v, deferred := vm.jsMemberGet(source, keys[i])
						if deferred {
							return Value{Type: VTJSUndefined}, true
						}
						vals[i] = v
					}
				case VTArray:
					if source.Arr != nil {
						keys = make([]string, len(source.Arr.Values))
						vals = make([]Value, len(source.Arr.Values))
						for i := 0; i < len(source.Arr.Values); i++ {
							keys[i] = strconv.Itoa(i)
							vals[i] = source.Arr.Values[i]
						}
					}
				case VTString:
					runes := []rune(source.Str)
					keys = make([]string, len(runes))
					vals = make([]Value, len(runes))
					for i := range runes {
						keys[i] = strconv.Itoa(i)
						vals[i] = NewString(string(runes[i]))
					}
				}
				entries := make([]Value, len(keys))
				for i := 0; i < len(keys); i++ {
					entries[i] = ValueFromVBArray(NewVBArrayFromValues(0, []Value{NewString(keys[i]), vals[i]}))
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, entries)), true
			case strings.EqualFold(member, "hasOwn"):
				if len(args) < 2 {
					return NewBool(false), true
				}
				obj := args[0]
				if obj.Type == VTNull || obj.Type == VTJSUndefined {
					vm.jsThrowTypeError("Cannot convert undefined or null to object")
					return Value{Type: VTJSUndefined}, true
				}
				key := vm.jsPropertyKeyFromValue(args[1])
				return NewBool(vm.jsObjectHasOwnProperty(obj, key)), true
			case strings.EqualFold(member, "fromEntries"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.fromEntries requires an iterable")
					return Value{Type: VTJSUndefined}, true
				}
				objID := vm.allocJSID()
				objMap := make(map[string]Value, 8)
				vm.jsObjectItems[objID] = objMap
				vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
				if proto := vm.jsGetIntrinsicPrototype("Object"); proto.Type == VTJSObject {
					objMap["__js_proto"] = proto
				}
				outObj := Value{Type: VTJSObject, Num: objID}
				iterable := args[0]
				iterator := vm.jsGetIterator(iterable)
				if iterator.Type == VTJSUndefined {
					vm.jsThrowTypeError("Object.fromEntries requires an iterable")
					return Value{Type: VTJSUndefined}, true
				}
				for {
					nextVal := vm.jsIteratorNextValue(iterator)
					if nextVal.Type == VTJSUndefined {
						break
					}
					len, isArr, _ := vm.jsArrayLikeLength(nextVal)
					if !isArr || len < 2 {
						continue
					}
					keyVal, _ := vm.jsArrayLikeGetIndex(nextVal, 0)
					val, _ := vm.jsArrayLikeGetIndex(nextVal, 1)
					keyStr := vm.jsPropertyKeyFromValue(keyVal)
					vm.jsMemberSet(outObj, keyStr, val)
				}
				return outObj, true
			case strings.EqualFold(member, "getOwnPropertySymbols"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.getOwnPropertySymbols called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type != VTJSObject && args[0].Type != VTJSFunction {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				symbols := vm.jsObjectOwnPropertySymbols(args[0])
				return ValueFromVBArray(NewVBArrayFromValues(0, symbols)), true
			case strings.EqualFold(member, "assign"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.assign target cannot be null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				targetObj := args[0]
				if targetObj.Type != VTJSObject && targetObj.Type != VTJSFunction && targetObj.Type != VTArray {
					vm.jsThrowTypeError("Object.assign target must be an object")
					return Value{Type: VTJSUndefined}, true
				}
				for si := 1; si < len(args); si++ {
					source := args[si]
					if source.Type == VTJSUndefined || source.Type == VTNull {
						continue
					}
					switch source.Type {
					case VTJSObject, VTJSFunction:
						keys := vm.jsObjectOwnEnumerableKeys(source.Num)
						for i := range keys {
							v, deferred := vm.jsMemberGet(source, keys[i])
							if deferred {
								return Value{Type: VTJSUndefined}, true
							}
							if targetObj.Type == VTArray {
								if idx, err := strconv.Atoi(keys[i]); err == nil {
									vm.jsIndexSet(targetObj, NewInteger(int64(idx)), v)
								}
							} else {
								vm.jsMemberSet(targetObj, keys[i], v)
							}
						}
					case VTArray:
						if source.Arr == nil {
							continue
						}
						for i := 0; i < len(source.Arr.Values); i++ {
							key := strconv.Itoa(i)
							if targetObj.Type == VTArray {
								vm.jsIndexSet(targetObj, NewInteger(int64(i)), source.Arr.Values[i])
							} else {
								vm.jsMemberSet(targetObj, key, source.Arr.Values[i])
							}
						}
					case VTString:
						runes := []rune(source.Str)
						for i := range runes {
							v := NewString(string(runes[i]))
							if targetObj.Type == VTArray {
								vm.jsIndexSet(targetObj, NewInteger(int64(i)), v)
							} else {
								vm.jsMemberSet(targetObj, strconv.Itoa(i), v)
							}
						}
					}
				}
				return targetObj, true
			case strings.EqualFold(member, "getOwnPropertyNames"):
				if len(args) == 0 {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				var names []string
				if args[0].Type == VTJSProxy {
					names = vm.jsProxyOwnKeys(args[0])
				} else {
					names = vm.jsObjectOwnPropertyNames(args[0])
				}
				values := make([]Value, len(names))
				for i := 0; i < len(names); i++ {
					values[i] = NewString(names[i])
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
			case strings.EqualFold(member, "defineProperty"):
				if len(args) < 1 || (args[0].Type != VTJSObject && args[0].Type != VTJSFunction && args[0].Type != VTJSProxy && args[0].Type != VTArray) {
					vm.jsThrowTypeError("Object.defineProperty called on non-object")
					return Value{Type: VTJSUndefined}, true
				}
				if len(args) < 3 {
					return args[0], true
				}
				// Use ReflectDefineProperty logic
				success := vm.dispatchJSIntrinsicCall(Value{}, "ReflectDefineProperty", args)
				if !vm.asBool(success) && args[0].Type == VTJSProxy {
					vm.jsThrowTypeError("Object.defineProperty failed")
				}
				return args[0], true
			case strings.EqualFold(member, "defineProperties"):
				if len(args) < 2 || (args[0].Type != VTJSObject && args[0].Type != VTJSFunction) || args[1].Type != VTJSObject {
					return jsArgOrUndefined(args, 0), true
				}
				descObj := vm.jsObjectItems[args[1].Num]
				for key, oneDesc := range descObj {
					if strings.HasPrefix(key, jsInternalPropPrefix) || oneDesc.Type != VTJSObject {
						continue
					}
					_, _ = vm.jsCallMember(target, "defineProperty", []Value{args[0], NewString(key), oneDesc})
				}
				return args[0], true
			case strings.EqualFold(member, "getOwnPropertyDescriptor"):
				if len(args) < 1 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.getOwnPropertyDescriptor called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type != VTJSObject && args[0].Type != VTJSFunction && args[0].Type != VTJSProxy && args[0].Type != VTArray {
					return Value{Type: VTJSUndefined}, true
				}
				if len(args) < 2 {
					return Value{Type: VTJSUndefined}, true
				}
				name := vm.jsPropertyKeyFromValue(args[1])
				desc, ok := vm.jsGetDescriptor(args[0].Num, name)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
				return vm.jsBuildDescriptorObject(desc), true
			case strings.EqualFold(member, "getOwnPropertyDescriptors"):
				if len(args) < 1 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.getOwnPropertyDescriptors called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type != VTJSObject && args[0].Type != VTJSFunction {
					return Value{Type: VTJSUndefined}, true
				}
				resultID := vm.allocJSID()
				resultObj := make(map[string]Value, 8)
				vm.jsObjectItems[resultID] = resultObj
				vm.jsPropertyItems[resultID] = make(map[string]jsPropertyDescriptor, 8)
				names := vm.jsObjectOwnPropertyNames(args[0])
				for i := range names {
					desc, ok := vm.jsGetDescriptor(args[0].Num, names[i])
					if !ok {
						continue
					}
					descObj := vm.jsBuildDescriptorObject(desc)
					resultObj[names[i]] = descObj
					vm.jsSetDescriptor(resultID, names[i], jsDefaultPropertyDescriptor(descObj))
				}
				return Value{Type: VTJSObject, Num: resultID}, true
			case strings.EqualFold(member, "preventExtensions"):
				if len(args) == 0 {
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type == VTJSObject || args[0].Type == VTJSFunction || args[0].Type == VTJSProxy || args[0].Type == VTArray {
					// Use ReflectPreventExtensions logic via dispatchJSIntrinsicCall
					success := vm.dispatchJSIntrinsicCall(Value{}, "ReflectPreventExtensions", args)
					_ = success
				}
				return args[0], true
			case strings.EqualFold(member, "isExtensible"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				if args[0].Type != VTJSObject && args[0].Type != VTJSFunction && args[0].Type != VTJSProxy && args[0].Type != VTArray {
					return NewBool(false), true
				}
				// Use ReflectIsExtensible logic via dispatchJSIntrinsicCall
				return vm.dispatchJSIntrinsicCall(Value{}, "ReflectIsExtensible", args), true
			case strings.EqualFold(member, "seal"):
				if len(args) == 0 {
					return Value{Type: VTJSUndefined}, true
				}
				return vm.jsObjectSeal(args[0]), true
			case strings.EqualFold(member, "isSealed"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				return NewBool(vm.jsObjectIsSealed(args[0])), true
			case strings.EqualFold(member, "freeze"):
				if len(args) == 0 {
					return Value{Type: VTJSUndefined}, true
				}
				return vm.jsObjectFreeze(args[0]), true
			case strings.EqualFold(member, "isFrozen"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				return NewBool(vm.jsObjectIsFrozen(args[0])), true
			}
		case "JSON":
			switch {
			case strings.EqualFold(member, "parse"):
				if len(args) == 0 {
					return Value{Type: VTJSUndefined}, true
				}
				return vm.jsJSONParse(vm.valueToString(args[0])), true
			case strings.EqualFold(member, "stringify"):
				if len(args) == 0 {
					return NewString("null"), true
				}
				jsonText := vm.jsJSONStringify(args[0])
				if !vm.jsEnsureStringSize(len(jsonText)) || !vm.jsChargeStringWork(len(jsonText)) {
					return Value{Type: VTJSUndefined}, true
				}
				return NewString(jsonText), true
			}
		case "Math":
			switch {
			case strings.EqualFold(member, "abs"):
				if len(args) == 0 {
					return NewDouble(math.NaN()), true
				}
				return NewDouble(math.Abs(vm.jsToNumber(args[0]).Flt)), true
			case strings.EqualFold(member, "sin"):
				return NewDouble(math.Sin(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "cos"):
				return NewDouble(math.Cos(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "tan"):
				return NewDouble(math.Tan(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "asin"):
				return NewDouble(math.Asin(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "acos"):
				return NewDouble(math.Acos(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "atan"):
				return NewDouble(math.Atan(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "acosh"):
				return NewDouble(math.Acosh(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "asinh"):
				return NewDouble(math.Asinh(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "atanh"):
				return NewDouble(math.Atanh(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "atan2"):
				y := vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt
				x := vm.jsToNumber(jsArgOrUndefined(args, 1)).Flt
				return NewDouble(math.Atan2(y, x)), true
			case strings.EqualFold(member, "ceil"):
				return NewDouble(math.Ceil(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "floor"):
				return NewDouble(math.Floor(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "trunc"):
				return NewDouble(math.Trunc(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "sign"):
				n := vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt
				if math.IsNaN(n) {
					return NewDouble(math.NaN()), true
				}
				if n > 0 {
					return NewInteger(1), true
				}
				if n < 0 {
					return NewInteger(-1), true
				}
				return NewInteger(0), true
			case strings.EqualFold(member, "cbrt"):
				return NewDouble(math.Cbrt(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "expm1"):
				return NewDouble(math.Expm1(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "log1p"):
				return NewDouble(math.Log1p(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "log10"):
				return NewDouble(math.Log10(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "log2"):
				return NewDouble(math.Log2(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "round"):
				return NewDouble(math.Round(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "sqrt"):
				return NewDouble(math.Sqrt(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "fround"):
				return NewDouble(float64(float32(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt))), true
			case strings.EqualFold(member, "clz32"):
				value := vm.jsToUint32Exact(jsArgOrUndefined(args, 0))
				return NewInteger(int64(bits.LeadingZeros32(value))), true
			case strings.EqualFold(member, "imul"):
				a := vm.jsToUint32Exact(jsArgOrUndefined(args, 0))
				b := vm.jsToUint32Exact(jsArgOrUndefined(args, 1))
				return NewInteger(int64(int32(a * b))), true
			case strings.EqualFold(member, "hypot"):
				if len(args) == 0 {
					return NewInteger(0), true
				}
				hyp := 0.0
				for i := range args {
					n := vm.jsToNumber(args[i]).Flt
					hyp = math.Hypot(hyp, n)
				}
				return NewDouble(hyp), true
			case strings.EqualFold(member, "pow"):
				base := vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt
				exp := vm.jsToNumber(jsArgOrUndefined(args, 1)).Flt
				return NewDouble(math.Pow(base, exp)), true
			case strings.EqualFold(member, "max"):
				if len(args) == 0 {
					return NewDouble(math.Inf(-1)), true
				}
				maxVal := vm.jsToNumber(args[0]).Flt
				for i := 1; i < len(args); i++ {
					n := vm.jsToNumber(args[i]).Flt
					if n > maxVal || math.IsNaN(n) {
						maxVal = n
					}
				}
				return NewDouble(maxVal), true
			case strings.EqualFold(member, "min"):
				if len(args) == 0 {
					return NewDouble(math.Inf(1)), true
				}
				minVal := vm.jsToNumber(args[0]).Flt
				for i := 1; i < len(args); i++ {
					n := vm.jsToNumber(args[i]).Flt
					if n < minVal || math.IsNaN(n) {
						minVal = n
					}
				}
				return NewDouble(minVal), true
			case strings.EqualFold(member, "random"):
				return NewDouble(rand.Float64()), true
			}
		case "Set":
			switch {
			case strings.EqualFold(member, "add"):
				if store, ok := vm.jsSetItems[target.Num]; ok {
					arg := jsArgOrUndefined(args, 0)
					store[vm.jsValueMapKey(arg)] = arg
				}
				return target, true
			case strings.EqualFold(member, "has"):
				store, ok := vm.jsSetItems[target.Num]
				if !ok {
					return NewBool(false), true
				}
				_, exists := store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))]
				return NewBool(exists), true
			case strings.EqualFold(member, "delete"):
				store, ok := vm.jsSetItems[target.Num]
				if !ok {
					return NewBool(false), true
				}
				key := vm.jsValueMapKey(jsArgOrUndefined(args, 0))
				_, exists := store[key]
				if exists {
					delete(store, key)
				}
				return NewBool(exists), true
			case strings.EqualFold(member, "clear"):
				if store, ok := vm.jsSetItems[target.Num]; ok {
					clear(store)
				}
				return Value{Type: VTJSUndefined}, true
			}
		case "Map":
			switch {
			case strings.EqualFold(member, "set"):
				if store, ok := vm.jsMapItems[target.Num]; ok {
					store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))] = jsArgOrUndefined(args, 1)
				}
				return target, true
			case strings.EqualFold(member, "has"):
				store, ok := vm.jsMapItems[target.Num]
				if !ok {
					return NewBool(false), true
				}
				_, exists := store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))]
				return NewBool(exists), true
			case strings.EqualFold(member, "delete"):
				store, ok := vm.jsMapItems[target.Num]
				if !ok {
					return NewBool(false), true
				}
				key := vm.jsValueMapKey(jsArgOrUndefined(args, 0))
				_, exists := store[key]
				if exists {
					delete(store, key)
				}
				return NewBool(exists), true
			case strings.EqualFold(member, "clear"):
				if store, ok := vm.jsMapItems[target.Num]; ok {
					clear(store)
				}
				return Value{Type: VTJSUndefined}, true
			}
		case "WeakMap":
			return vm.jsCallWeakCollectionMethod(target, "WeakMap", member, args), true
		case "WeakSet":
			return vm.jsCallWeakCollectionMethod(target, "WeakSet", member, args), true
		case "Number":
			// ES6 Number static methods.
			switch {
			case strings.EqualFold(member, "isInteger"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				v := args[0]
				switch v.Type {
				case VTInteger:
					return NewBool(true), true
				case VTDouble:
					return NewBool(!math.IsNaN(v.Flt) && !math.IsInf(v.Flt, 0) && math.Trunc(v.Flt) == v.Flt), true
				default:
					return NewBool(false), true
				}
			case strings.EqualFold(member, "isNaN"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				v := args[0]
				// ES6 Number.isNaN does NOT coerce; only actual NaN returns true.
				if v.Type == VTDouble {
					return NewBool(math.IsNaN(v.Flt)), true
				}
				return NewBool(false), true
			case strings.EqualFold(member, "isFinite"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				v := args[0]
				// ES6 Number.isFinite does NOT coerce; only actual numbers are checked.
				switch v.Type {
				case VTInteger:
					return NewBool(true), true
				case VTDouble:
					return NewBool(!math.IsNaN(v.Flt) && !math.IsInf(v.Flt, 0)), true
				default:
					return NewBool(false), true
				}
			case strings.EqualFold(member, "isSafeInteger"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				v := args[0]
				switch v.Type {
				case VTInteger:
					return NewBool(v.Num >= -9007199254740991 && v.Num <= 9007199254740991), true
				case VTDouble:
					return NewBool(!math.IsNaN(v.Flt) && !math.IsInf(v.Flt, 0) && math.Trunc(v.Flt) == v.Flt && v.Flt >= -9007199254740991 && v.Flt <= 9007199254740991), true
				default:
					return NewBool(false), true
				}
			case strings.EqualFold(member, "parseInt"):
				return vm.jsParseIntES5(args), true
			case strings.EqualFold(member, "parseFloat"):
				return vm.jsParseFloatES5(args), true
			}
		case "RegExp":
			if strings.EqualFold(member, "test") {
				needle := ""
				if len(args) > 0 {
					needle = vm.valueToString(args[0])
				}
				res := vm.jsRegExpExec(target, needle)
				return NewBool(res.Type == VTJSObject), true
			}
			if strings.EqualFold(member, "exec") {
				needle := ""
				if len(args) > 0 {
					needle = vm.valueToString(args[0])
				}
				return vm.jsRegExpExec(target, needle), true
			}
		case "Enumerator":
			switch {
			case strings.EqualFold(member, "atEnd"):
				return NewBool(vm.jsEnumeratorAtEnd(target)), true
			case strings.EqualFold(member, "moveNext"):
				vm.jsEnumeratorMoveNext(target)
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "moveFirst"):
				vm.jsEnumeratorMoveFirst(target)
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "item"):
				return vm.jsEnumeratorItem(target), true
			}
		case "VBArray":
			switch {
			case strings.EqualFold(member, "dimensions"):
				return NewInteger(int64(vm.jsVBArrayDimensions(target))), true
			case strings.EqualFold(member, "lbound"):
				dim := 1
				if len(args) > 0 {
					dim = int(vm.jsToNumber(args[0]).Flt)
				}
				lower, _, ok := arrayBounds(vm.jsVBArraySource(target), dim)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
				return NewInteger(int64(lower)), true
			case strings.EqualFold(member, "ubound"):
				dim := 1
				if len(args) > 0 {
					dim = int(vm.jsToNumber(args[0]).Flt)
				}
				_, upper, ok := arrayBounds(vm.jsVBArraySource(target), dim)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
				return NewInteger(int64(upper)), true
			case strings.EqualFold(member, "toArray"):
				return vm.jsVBArrayToJSArray(vm.jsVBArraySource(target)), true
			case strings.EqualFold(member, "getItem"):
				return vm.jsVBArrayGetItem(target, args), true
			}
		case "Symbol":
			// Symbol constructor static methods
			switch {
			case strings.EqualFold(member, "for"):
				key := ""
				if len(args) > 0 {
					key = vm.valueToString(args[0])
				}
				return vm.jsSymbolFor(key), true
			case strings.EqualFold(member, "keyFor"):
				return vm.jsSymbolKeyFor(jsArgOrUndefined(args, 0)), true
			}
		case "ArrayBuffer", "SharedArrayBuffer":
			// ArrayBuffer.isView(arg) static method (called on the ctor object)
			if strings.EqualFold(member, "isView") {
				if vm.jsObjectStringProperty(target, "__js_ctor") == "ArrayBuffer" {
					if len(args) == 0 {
						return NewBool(false), true
					}
					if args[0].Type != VTJSObject {
						return NewBool(false), true
					}
					t := vm.jsObjectStringProperty(args[0], "__js_type")
					return NewBool(jsIsTypedArrayType(t)), true
				}
			}
			// ArrayBuffer / SharedArrayBuffer instance method: slice
			if strings.EqualFold(member, "slice") {
				return vm.jsArrayBufferSlice(target, args), true
			}
		default:
			// Instance methods on typed arrays and DataView
			if jsIsTypedArrayType(objType) {
				switch {
				case strings.EqualFold(member, "set"):
					return vm.jsTypedArraySet(target, args), true
				case objType == "Uint8Array" && strings.EqualFold(member, "fromBase64"):
					if len(args) == 0 || args[0].Type == VTJSUndefined {
						vm.jsThrowTypeError("fromBase64 requires a string")
						return Value{Type: VTJSUndefined}, true
					}
					data, err := base64.StdEncoding.DecodeString(vm.valueToString(args[0]))
					if err != nil {
						vm.jsThrowTypeError("Invalid base64 string")
						return Value{Type: VTJSUndefined}, true
					}
					bufObj := vm.jsNewArrayBuffer(len(data))
					copy(vm.jsArrayBuffers[bufObj.Num], data)
					return vm.jsNewTypedArray("Uint8Array", []Value{bufObj}), true
				case objType == "Uint8Array" && strings.EqualFold(member, "fromHex"):
					if len(args) == 0 || args[0].Type == VTJSUndefined {
						vm.jsThrowTypeError("fromHex requires a string")
						return Value{Type: VTJSUndefined}, true
					}
					data, err := hex.DecodeString(vm.valueToString(args[0]))
					if err != nil {
						vm.jsThrowTypeError("Invalid hex string")
						return Value{Type: VTJSUndefined}, true
					}
					bufObj := vm.jsNewArrayBuffer(len(data))
					copy(vm.jsArrayBuffers[bufObj.Num], data)
					return vm.jsNewTypedArray("Uint8Array", []Value{bufObj}), true
				case objType == "Uint8Array" && strings.EqualFold(member, "toBase64"):
					buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(target)
					if !ok {
						return Value{Type: VTJSUndefined}, true
					}
					data := buf[byteOffset : byteOffset+(byteLength/elemSize)]
					return NewString(base64.StdEncoding.EncodeToString(data)), true
				case objType == "Uint8Array" && strings.EqualFold(member, "toHex"):
					buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(target)
					if !ok {
						return Value{Type: VTJSUndefined}, true
					}
					data := buf[byteOffset : byteOffset+(byteLength/elemSize)]
					return NewString(hex.EncodeToString(data)), true
				case strings.EqualFold(member, "subarray"):
					return vm.jsTypedArraySubarray(target, args), true
				case strings.EqualFold(member, "fill"):
					return vm.jsTypedArrayFill(target, args), true
				case strings.EqualFold(member, "slice"):
					// Returns a new typed array copy of the range
					return vm.jsTypedArraySubarray(target, args), true
				case strings.EqualFold(member, "forEach"):
					length := 0
					if items, ok := vm.jsObjectItems[target.Num]; ok {
						if lv, ok2 := items["__js_byte_length"]; ok2 {
							elemSz := jsTypedArrayElementSize(objType)
							if elemSz > 0 {
								length = int(lv.Num) / elemSz
							}
						}
					}
					callback := jsArgOrUndefined(args, 0)
					if callback.Type == VTJSFunction {
						buf, byteOffset, _, elemSize, ok := vm.jsGetTypedArrayInfo(target)
						if ok {
							for i := 0; i < length; i++ {
								v := jsReadTypedArrayElement(objType, elemSize, buf, byteOffset, i)
								_ = vm.jsCall(callback, Value{Type: VTJSUndefined}, []Value{v, NewInteger(int64(i)), target})
							}
						}
					}
					return Value{Type: VTJSUndefined}, true
				case strings.EqualFold(member, "indexOf"):
					target2 := jsArgOrUndefined(args, 0)
					fromIndex := 0
					if len(args) > 1 {
						fromIndex = int(vm.jsToNumber(args[1]).Flt)
					}
					buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(target)
					if !ok {
						return NewInteger(-1), true
					}
					length := byteLength / elemSize
					searchVal := vm.jsToNumber(target2).Flt
					for i := fromIndex; i < length; i++ {
						v := jsReadTypedArrayElement(objType, elemSize, buf, byteOffset, i)
						if vm.jsToNumber(v).Flt == searchVal {
							return NewInteger(int64(i)), true
						}
					}
					return NewInteger(-1), true
				default:
					// DataView method calls
					if objType == "DataView" {
						if v, handled := vm.jsDataViewCallMember(target, member, args); handled {
							return v, true
						}
					}
				}
			}
		}
	}

	// General fallback: dispatch based on target type
	switch target.Type {
	case VTNativeObject:
		return vm.dispatchNativeCall(target.Num, member, args), true
	case VTJSObject, VTJSFunction, VTArray, VTString, VTDate:
		callee, deferred := vm.jsMemberGet(target, member)
		if deferred {
			return Value{Type: VTJSUndefined}, true
		}
		if callee.Type != VTJSUndefined {
			return vm.jsCall(callee, target, args), true
		}
	}

	return Value{Type: VTJSUndefined}, false
}

func (vm *VM) jsCreateErrorObject(name string, msg string) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("Error")
	obj["__js_ctor"] = NewString(name)
	obj["name"] = NewString(name)
	obj["message"] = NewString(msg)
	if proto := vm.jsGetIntrinsicPrototype(name); proto.Type == VTJSObject {
		obj["__js_proto"] = proto
	} else if proto := vm.jsGetIntrinsicPrototype("Error"); proto.Type == VTJSObject {
		obj["__js_proto"] = proto
	}
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// jsNumberToString formats a numeric primitive using JScript-compatible defaults.
func (vm *VM) jsNumberToString(n float64) string {
	if math.IsNaN(n) {
		return "NaN"
	}
	if math.IsInf(n, 1) {
		return "Infinity"
	}
	if math.IsInf(n, -1) {
		return "-Infinity"
	}
	if n == 0 {
		return "0"
	}
	return vm.jsNormalizeExponent(strconv.FormatFloat(n, 'g', -1, 64))
}

// jsNumberToStringRadix formats integral numbers in the requested radix.
func (vm *VM) jsNumberToStringRadix(n float64, radix int) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return vm.jsNumberToString(n)
	}
	if n == 0 {
		return "0"
	}
	truncated := math.Trunc(n)
	if truncated != n {
		return vm.jsNumberToString(n)
	}
	sign := ""
	if truncated < 0 {
		sign = "-"
		truncated = -truncated
	}
	const digits = "0123456789abcdefghijklmnopqrstuvwxyz"
	value := uint64(truncated)
	buf := [65]byte{}
	idx := len(buf)
	for value > 0 {
		idx--
		buf[idx] = digits[value%uint64(radix)]
		value /= uint64(radix)
	}
	if idx == len(buf) {
		return "0"
	}
	return sign + string(buf[idx:])
}

// jsNumberToFixed implements Number.prototype.toFixed with bounded precision.
func (vm *VM) jsNumberToFixed(n float64, digits int) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return vm.jsNumberToString(n)
	}
	if n >= 1e21 || n <= -1e21 {
		return vm.jsNumberToString(n)
	}
	if n == 0 {
		n = 0
	}
	return string(ftoa.FToStr(n, ftoa.ModeFixed, digits, nil))
}

// jsNumberToExponential implements Number.prototype.toExponential.
func (vm *VM) jsNumberToExponential(n float64, digits int) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return vm.jsNumberToString(n)
	}
	if n == 0 {
		n = 0
	}
	if digits < 0 {
		return vm.jsNormalizeExponent(strconv.FormatFloat(n, 'e', -1, 64))
	}
	return vm.jsNormalizeExponent(strconv.FormatFloat(n, 'e', digits, 64))
}

// jsNumberToPrecision implements Number.prototype.toPrecision.
func (vm *VM) jsNumberToPrecision(n float64, precision int) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return vm.jsNumberToString(n)
	}
	if n == 0 {
		n = 0
	}
	return vm.jsNormalizeExponent(strconv.FormatFloat(n, 'g', precision, 64))
}

// jsNormalizeExponent removes zero-padding in exponent suffix to match JS formatting.
func (vm *VM) jsNormalizeExponent(text string) string {
	ePos := strings.IndexAny(text, "eE")
	if ePos < 0 || ePos+2 >= len(text) {
		return text
	}
	base := text[:ePos]
	exp := text[ePos+1:]
	sign := exp[0:1]
	digits := strings.TrimLeft(exp[1:], "0")
	if digits == "" {
		digits = "0"
	}
	return base + "e" + sign + digits
}

// jsArgOrUndefined returns args[idx] or undefined for missing arguments.
func jsArgOrUndefined(args []Value, idx int) Value {
	if idx >= 0 && idx < len(args) {
		return args[idx]
	}
	return Value{Type: VTJSUndefined}
}

// jsStringReplace implements String.prototype.replace and replaceAll with size guards.
func (vm *VM) jsStringReplace(source string, patternArg Value, replacementArg Value, replaceAll bool) Value {
	if patternArg.Type == VTJSObject {
		objType := vm.jsObjectStringProperty(patternArg, "__js_type")
		if objType == "RegExp" {
			pattern := vm.jsObjectStringProperty(patternArg, "pattern")
			flags := vm.jsObjectStringProperty(patternArg, "flags")
			return vm.jsStringReplaceRegex(source, pattern, flags, replacementArg, replaceAll)
		}
	}

	search := vm.valueToString(patternArg)
	useCallback := replacementArg.Type == VTJSFunction
	replacement := vm.valueToString(replacementArg)
	if search == "" {
		if replaceAll {
			var b strings.Builder
			for i := 0; i <= len(source); i++ {
				repl := replacement
				if useCallback {
					cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, []Value{NewString(""), NewInteger(int64(i)), NewString(source)})
					repl = vm.valueToString(cb)
				}
				b.WriteString(repl)
				if i < len(source) {
					b.WriteByte(source[i])
				}
			}
			out := b.String()
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}
			}
			return NewString(out)
		}
		repl := replacement
		if useCallback {
			cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, []Value{NewString(""), NewInteger(0), NewString(source)})
			repl = vm.valueToString(cb)
		}
		out := repl + source
		if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(out)
	}

	if !replaceAll {
		idx := strings.Index(source, search)
		if idx < 0 {
			if !vm.jsEnsureStringSize(len(source)) || !vm.jsChargeStringWork(len(source)) {
				return Value{Type: VTJSUndefined}
			}
			return NewString(source)
		}
		repl := replacement
		if useCallback {
			cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, []Value{NewString(search), NewInteger(int64(idx)), NewString(source)})
			repl = vm.valueToString(cb)
		}
		out := source[:idx] + repl + source[idx+len(search):]
		if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(out)
	}

	var b strings.Builder
	start := 0
	for {
		idx := strings.Index(source[start:], search)
		if idx < 0 {
			b.WriteString(source[start:])
			break
		}
		idx += start
		b.WriteString(source[start:idx])
		repl := replacement
		if useCallback {
			cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, []Value{NewString(search), NewInteger(int64(idx)), NewString(source)})
			repl = vm.valueToString(cb)
		}
		b.WriteString(repl)
		start = idx + len(search)
	}
	out := b.String()
	if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
		return Value{Type: VTJSUndefined}
	}
	return NewString(out)
}

// jsStringReplaceRegex applies one or all RegExp matches to the source string.
func (vm *VM) jsStringReplaceRegex(source string, pattern string, flags string, replacementArg Value, replaceAll bool) Value {
	re, err := vm.jsCompileRegExp(pattern, flags)
	if err != nil {
		return NewString(source)
	}
	replacement := vm.valueToString(replacementArg)
	useCallback := replacementArg.Type == VTJSFunction
	flagsLower := strings.ToLower(flags)
	useAll := replaceAll || strings.Contains(flagsLower, "g")

	var b strings.Builder
	lastByteIdx := 0
	lastUTF16Idx := int64(0)

	m, err := re.FindStringMatch(source)
	if err != nil || m == nil {
		return NewString(source)
	}

	for m != nil {
		startByte := vm.jsRuneToByteOffset(source, m.Index)
		lengthByte := vm.jsRuneToByteOffset(source[startByte:], m.Length)
		endByte := startByte + lengthByte

		// Append portion before match
		b.WriteString(source[lastByteIdx:startByte])

		// Calculate UTF-16 index of match start
		matchStartUTF16 := lastUTF16Idx
		for _, r := range source[lastByteIdx:startByte] {
			if r >= 0x10000 {
				matchStartUTF16 += 2
			} else {
				matchStartUTF16 += 1
			}
		}

		repl := ""
		if useCallback {
			groups := m.Groups()
			callbackArgs := make([]Value, 0, len(groups)+2)
			callbackArgs = append(callbackArgs, NewString(m.String()))
			for i := 1; i < len(groups); i++ {
				g := groups[i]
				if g.Capture.Length < 0 {
					callbackArgs = append(callbackArgs, Value{Type: VTJSUndefined})
				} else {
					callbackArgs = append(callbackArgs, NewString(g.String()))
				}
			}
			callbackArgs = append(callbackArgs, NewInteger(matchStartUTF16), NewString(source))
			// ES2018: if named groups are present, an extra argument with groups object is passed
			hasNamedGroups := false
			for i, g := range groups {
				if g.Name != "" && g.Name != strconv.Itoa(i) {
					hasNamedGroups = true
					break
				}
			}
			if hasNamedGroups {
				groupsObjID := vm.allocJSID()
				groupsObj := make(map[string]Value)
				for i, g := range groups {
					if g.Name != "" && g.Name != strconv.Itoa(i) {
						if g.Capture.Length < 0 {
							groupsObj[g.Name] = Value{Type: VTJSUndefined}
						} else {
							groupsObj[g.Name] = NewString(g.String())
						}
					}
				}
				vm.jsObjectItems[groupsObjID] = groupsObj
				callbackArgs = append(callbackArgs, Value{Type: VTJSObject, Num: groupsObjID})
			}

			cb, handled := vm.jsCallDirectNoClone(replacementArg, Value{Type: VTJSUndefined}, callbackArgs)
			if !handled {
				cb = vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, callbackArgs)
			}
			repl = vm.valueToString(cb)
		} else {
			// String replacement with $ tokens
			repl = vm.jsExpandReplaceTokens(m, source, replacement, matchStartUTF16)
		}

		b.WriteString(repl)

		// Update last indices
		lastByteIdx = endByte
		lastUTF16Idx = matchStartUTF16
		for _, r := range source[startByte:endByte] {
			if r >= 0x10000 {
				lastUTF16Idx += 2
			} else {
				lastUTF16Idx += 1
			}
		}

		if !useAll {
			break
		}
		m, err = re.FindNextMatch(m)
		if err != nil {
			break
		}
	}

	b.WriteString(source[lastByteIdx:])
	out := b.String()
	if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
		return Value{Type: VTJSUndefined}
	}
	return NewString(out)
}

func (vm *VM) jsExpandReplaceTokens(m *regexp2.Match, source, replacement string, matchStartUTF16 int64) string {
	var b strings.Builder
	groups := m.Groups()
	startByte := vm.jsRuneToByteOffset(source, m.Index)
	lengthByte := vm.jsRuneToByteOffset(source[startByte:], m.Length)
	endByte := startByte + lengthByte

	for i := 0; i < len(replacement); i++ {
		if replacement[i] == '$' && i+1 < len(replacement) {
			next := replacement[i+1]
			switch next {
			case '$':
				b.WriteByte('$')
				i++
			case '&':
				b.WriteString(m.String())
				i++
			case '`':
				b.WriteString(source[:startByte])
				i++
			case '\'':
				b.WriteString(source[endByte:])
				i++
			case '0', '1', '2', '3', '4', '5', '6', '7', '8', '9':
				// Handle $n and $nn
				num := int(next - '0')
				i++
				if i+1 < len(replacement) && replacement[i+1] >= '0' && replacement[i+1] <= '9' {
					potentialNum := num*10 + int(replacement[i+1]-'0')
					if potentialNum < len(groups) {
						num = potentialNum
						i++
					}
				}
				if num > 0 && num < len(groups) {
					g := groups[num]
					if g.Capture.Length >= 0 {
						b.WriteString(g.String())
					}
				} else {
					// If not a valid group, just literal $n
					b.WriteByte('$')
					b.WriteByte(byte('0' + num)) // This is slightly wrong for $nn, but good enough for now
				}
			case '<':
				// Named groups: $<name>
				endIdx := strings.IndexByte(replacement[i+1:], '>')
				if endIdx > 0 {
					name := replacement[i+2 : i+1+endIdx]
					g := m.GroupByName(name)
					if g != nil && g.Capture.Length >= 0 {
						b.WriteString(g.String())
					}
					i += endIdx + 1
				} else {
					b.WriteByte('$')
				}
			default:
				b.WriteByte('$')
			}
		} else {
			b.WriteByte(replacement[i])
		}
	}
	return b.String()
}

// jsEnumeratorSource returns the wrapped collection for a JScript Enumerator.
func (vm *VM) jsEnumeratorSource(obj Value) Value {
	if obj.Type != VTJSObject {
		return Value{Type: VTJSUndefined}
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	source, ok := items["__js_enum_source"]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	return source
}

// jsEnumeratorIndex returns the current Enumerator cursor index.
func (vm *VM) jsEnumeratorIndex(obj Value) int {
	if obj.Type != VTJSObject {
		return 0
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return 0
	}
	v, ok := items["__js_enum_index"]
	if !ok {
		return 0
	}
	return int(vm.jsToNumber(v).Flt)
}

// jsEnumeratorSetIndex writes the Enumerator cursor index.
func (vm *VM) jsEnumeratorSetIndex(obj Value, idx int) {
	if obj.Type != VTJSObject {
		return
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return
	}
	items["__js_enum_index"] = NewInteger(int64(idx))
}

// jsEnumeratorItemCount returns the number of available items in an Enumerator.
func (vm *VM) jsEnumeratorItemCount(obj Value) int {
	source := vm.jsEnumeratorSource(obj)
	switch source.Type {
	case VTArray:
		if source.Arr == nil {
			return 0
		}
		return len(source.Arr.Values)
	case VTJSObject:
		items, ok := vm.jsObjectItems[obj.Num]
		if !ok {
			return 0
		}
		keyList, ok := items["__js_enum_keys"]
		if !ok || keyList.Type != VTArray || keyList.Arr == nil {
			return 0
		}
		return len(keyList.Arr.Values)
	case VTNativeObject:
		if dict, ok := vm.dictionaryItems[source.Num]; ok && dict != nil {
			return len(dict.keys)
		}
		countVal := vm.dispatchMemberGet(source, "Count")
		return int(vm.jsToNumber(countVal).Flt)
	default:
		return 0
	}
}

// jsNativeCollectionItemByKeyIndex resolves one collection value from Key(index) lookups.
func (vm *VM) jsNativeCollectionItemByKeyIndex(source Value, idx int) Value {
	if idx < 0 {
		return Value{Type: VTJSUndefined}
	}
	keyMethod := vm.dispatchMemberGet(source, "Key")
	if keyMethod.Type != VTNativeObject {
		return Value{Type: VTJSUndefined}
	}
	lookup := func(pos int) Value {
		key := vm.dispatchNativeCall(keyMethod.Num, "", []Value{NewInteger(int64(pos))})
		if key.Type == VTJSUndefined || key.Type == VTEmpty {
			return Value{Type: VTJSUndefined}
		}
		keyStr := vm.valueToString(key)
		if keyStr == "" {
			return Value{Type: VTJSUndefined}
		}
		return vm.dispatchNativeCall(source.Num, "", []Value{NewString(keyStr)})
	}
	value := lookup(idx + 1)
	if value.Type != VTJSUndefined && value.Type != VTEmpty {
		return value
	}
	return lookup(idx)
}

// jsEnumeratorAtEnd reports whether the Enumerator cursor reached the end.
func (vm *VM) jsEnumeratorAtEnd(obj Value) bool {
	return vm.jsEnumeratorIndex(obj) >= vm.jsEnumeratorItemCount(obj)
}

// jsEnumeratorMoveNext advances the Enumerator cursor by one.
func (vm *VM) jsEnumeratorMoveNext(obj Value) {
	vm.jsEnumeratorSetIndex(obj, vm.jsEnumeratorIndex(obj)+1)
}

// jsEnumeratorMoveFirst resets the Enumerator cursor to the first item.
func (vm *VM) jsEnumeratorMoveFirst(obj Value) {
	vm.jsEnumeratorSetIndex(obj, 0)
}

// jsEnumeratorItem returns the current item in the wrapped collection.
func (vm *VM) jsEnumeratorItem(obj Value) Value {
	source := vm.jsEnumeratorSource(obj)
	idx := vm.jsEnumeratorIndex(obj)
	switch source.Type {
	case VTArray:
		if source.Arr == nil || idx < 0 || idx >= len(source.Arr.Values) {
			return Value{Type: VTJSUndefined}
		}
		return source.Arr.Values[idx]
	case VTJSObject:
		items, ok := vm.jsObjectItems[obj.Num]
		if !ok {
			return Value{Type: VTJSUndefined}
		}
		keyList, ok := items["__js_enum_keys"]
		if !ok || keyList.Type != VTArray || keyList.Arr == nil || idx < 0 || idx >= len(keyList.Arr.Values) {
			return Value{Type: VTJSUndefined}
		}
		key := keyList.Arr.Values[idx].Str
		objMap, ok := vm.jsObjectItems[source.Num]
		if !ok {
			return Value{Type: VTJSUndefined}
		}
		if v, exists := objMap[key]; exists {
			return v
		}
		return Value{Type: VTJSUndefined}
	case VTNativeObject:
		if idx < 0 {
			return Value{Type: VTJSUndefined}
		}
		if dict, ok := vm.dictionaryItems[source.Num]; ok && dict != nil {
			if idx >= len(dict.keys) {
				return Value{Type: VTJSUndefined}
			}
			return dict.keys[idx]
		}
		zeroBased := vm.dispatchNativeCall(source.Num, "Item", []Value{NewInteger(int64(idx))})
		if zeroBased.Type != VTJSUndefined && zeroBased.Type != VTEmpty {
			return zeroBased
		}
		oneBased := vm.dispatchNativeCall(source.Num, "Item", []Value{NewInteger(int64(idx + 1))})
		if oneBased.Type != VTJSUndefined && oneBased.Type != VTEmpty {
			return oneBased
		}
		return vm.jsNativeCollectionItemByKeyIndex(source, idx)
	default:
		return Value{Type: VTJSUndefined}
	}
}

// jsVBArrayDepth returns the maximum nested array depth starting at one VBArray node.
func (vm *VM) jsVBArrayDepth(arr *VBArray) int {
	if arr == nil {
		return 0
	}
	maxChildDepth := 0
	for i := 0; i < len(arr.Values); i++ {
		child, ok := toVBArray(arr.Values[i])
		if !ok {
			continue
		}
		depth := vm.jsVBArrayDepth(child)
		if depth > maxChildDepth {
			maxChildDepth = depth
		}
	}
	return 1 + maxChildDepth
}

// jsVBArraySource extracts the wrapped VBArray value from a JScript VBArray wrapper object.
func (vm *VM) jsVBArraySource(obj Value) Value {
	if obj.Type != VTJSObject {
		return Value{Type: VTJSUndefined}
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	source, ok := items["__js_vbarray_source"]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	return source
}

// jsVBArrayDimensions returns the number of dimensions for the wrapped VBArray.
func (vm *VM) jsVBArrayDimensions(obj Value) int {
	value := vm.jsVBArraySource(obj)
	arr, ok := toVBArray(value)
	if !ok {
		return 0
	}
	return vm.jsVBArrayDepth(arr)
}

// jsVBArrayToJSArray converts a VBArray value into a zero-based array value.
func (vm *VM) jsVBArrayToJSArray(value Value) Value {
	arr, ok := toVBArray(value)
	if !ok || arr == nil {
		return Value{Type: VTJSUndefined}
	}
	converted := make([]Value, len(arr.Values))
	for i := 0; i < len(arr.Values); i++ {
		if child, childOK := toVBArray(arr.Values[i]); childOK {
			converted[i] = vm.jsVBArrayToJSArray(ValueFromVBArray(child))
		} else {
			converted[i] = arr.Values[i]
		}
	}
	return ValueFromVBArray(NewVBArrayFromValues(0, converted))
}

// jsVBArrayGetItem fetches one element from VBArray using one or more indexes.
func (vm *VM) jsVBArrayGetItem(obj Value, args []Value) Value {
	current := vm.jsVBArraySource(obj)
	if current.Type != VTArray || current.Arr == nil {
		return Value{Type: VTJSUndefined}
	}
	if len(args) == 0 {
		return Value{Type: VTJSUndefined}
	}
	for i := range args {
		if current.Type != VTArray || current.Arr == nil {
			return Value{Type: VTJSUndefined}
		}
		idx := int(vm.jsToNumber(args[i]).Flt)
		offset := idx - current.Arr.Lower
		if offset < 0 || offset >= len(current.Arr.Values) {
			return Value{Type: VTJSUndefined}
		}
		current = current.Arr.Values[offset]
	}
	return current
}

func (vm *VM) jsMemberSet(target Value, member string, val Value) {
	if target.Type == VTJSUninitialized {
		vm.jsThrowReferenceError("Must call super constructor in derived class before accessing 'this'")
		return
	}
	switch target.Type {
	case VTJSProxy:
		vm.jsProxySet(target, member, val, target)
		return
	case VTNativeObject:
		vm.dispatchMemberSet(target.Num, member, val)
	case VTArray:
		if target.Arr == nil {
			return
		}
		if target.Arr.JSProps == nil {
			target.Arr.JSProps = make(map[string]Value, 4)
		}
		target.Arr.JSProps[member] = val
	case VTJSObject, VTJSFunction:
		var targetID int64
		if target.Type == VTJSObject || target.Type == VTJSFunction {
			targetID = target.Num
		}
		if vm.jsHandleNodeURLMemberSet(target, member, val) {
			return
		}
		if vm.jsSetAliasedArgumentValue(targetID, member, val) {
			return
		}
		// Typed array: handle numeric-index writes via string key.
		if vm.jsTypedArrayMemberSet(target, member, val) {
			return
		}
		if after, ok := strings.CutPrefix(member, jsAccessorGetterPrefix); ok {
			name := after
			desc, _ := vm.jsGetDescriptor(targetID, name)
			desc.Getter = val
			desc.HasGetter = true
			desc.HasValue = false
			desc.Writable = false
			desc.Enumerable = true
			desc.Configurable = true
			vm.jsSetDescriptor(targetID, name, desc)
			return
		}
		if after, ok := strings.CutPrefix(member, jsAccessorSetterPrefix); ok {
			name := after
			desc, _ := vm.jsGetDescriptor(targetID, name)
			desc.Setter = val
			desc.HasSetter = true
			desc.HasValue = false
			desc.Writable = false
			desc.Enumerable = true
			desc.Configurable = true
			vm.jsSetDescriptor(targetID, name, desc)
			return
		}

		desc, hasDesc := vm.jsGetDescriptor(targetID, member)
		if hasDesc {
			if desc.HasSetter {
				vm.jsCall(desc.Setter, target, []Value{val})
				return
			}
			if !desc.Writable {
				// In strict mode, writing to a non-writable property is a TypeError.
				if vm.jsStrictMode {
					vm.jsThrowTypeError(fmt.Sprintf("Cannot assign to read-only property '%s'", member))
				}
				return
			}
			desc.Value = val
			desc.HasValue = true
			vm.jsSetDescriptor(targetID, member, desc)
			return
		}

		if proto := vm.jsGetPrototypeValue(target); proto.Type == VTJSObject {
			if inherited, ok := vm.jsResolveObjectMember(proto.Num, member, make(map[int64]struct{}, 4)); ok {
				if inherited.HasSetter {
					vm.jsCall(inherited.Setter, target, []Value{val})
					return
				}
				if !inherited.Writable {
					if vm.jsStrictMode {
						vm.jsThrowTypeError(fmt.Sprintf("Cannot assign to read-only property '%s'", member))
					}
					return
				}
			}
		}

		obj, ok := vm.jsObjectItems[targetID]
		if !ok {
			obj = make(map[string]Value, 8)
			vm.jsObjectItems[targetID] = obj
		}
		if !vm.jsObjectIsExtensible(target) {
			// In strict mode, adding a property to a non-extensible object is a TypeError.
			if vm.jsStrictMode {
				vm.jsThrowTypeError(fmt.Sprintf("Cannot add property '%s', object is not extensible", member))
			}
			return
		}
		obj[member] = val
		vm.jsSetDescriptor(targetID, member, jsDefaultPropertyDescriptor(val))
	}
}

func jsDescriptorIsAccessor(desc jsPropertyDescriptor) bool {
	return desc.HasGetter || desc.HasSetter
}

func jsDescriptorIsData(desc jsPropertyDescriptor) bool {
	return desc.HasValue
}

func (vm *VM) jsReadDefinePropertySpec(descVal Value) jsDefinePropertySpec {
	spec := jsDefinePropertySpec{desc: jsPropertyDescriptor{Enumerable: false, Configurable: false, Writable: false}}
	if descVal.Type != VTJSObject && descVal.Type != VTJSFunction && descVal.Type != VTJSProxy && descVal.Type != VTArray {
		return spec
	}
	if v, deferred := vm.jsMemberGet(descVal, "value"); !deferred && v.Type != VTJSUndefined {
		spec.desc.Value = v
		spec.desc.HasValue = true
	}
	if v, deferred := vm.jsMemberGet(descVal, "get"); !deferred && v.Type != VTJSUndefined {
		spec.desc.Getter = v
		spec.desc.HasGetter = true
	}
	if v, deferred := vm.jsMemberGet(descVal, "set"); !deferred && v.Type != VTJSUndefined {
		spec.desc.Setter = v
		spec.desc.HasSetter = true
	}
	if v, deferred := vm.jsMemberGet(descVal, "enumerable"); !deferred && v.Type != VTJSUndefined {
		spec.hasEnumerable = true
		spec.desc.Enumerable = vm.jsToDescriptorBoolean(v, false)
	}
	if v, deferred := vm.jsMemberGet(descVal, "configurable"); !deferred && v.Type != VTJSUndefined {
		spec.hasConfigurable = true
		spec.desc.Configurable = vm.jsToDescriptorBoolean(v, false)
	}
	if v, deferred := vm.jsMemberGet(descVal, "writable"); !deferred && v.Type != VTJSUndefined {
		spec.hasWritable = true
		spec.desc.Writable = vm.jsToDescriptorBoolean(v, false)
	}
	return spec
}

func (vm *VM) jsBuildDescriptorObject(desc jsPropertyDescriptor) Value {
	descObjID := vm.allocJSID()
	descObj := make(map[string]Value, 8)
	if desc.HasValue {
		descObj["value"] = desc.Value
		descObj["writable"] = NewBool(desc.Writable)
	}
	if desc.HasGetter {
		descObj["get"] = desc.Getter
	} else if desc.HasSetter {
		descObj["get"] = Value{Type: VTJSUndefined}
	}
	if desc.HasSetter {
		descObj["set"] = desc.Setter
	} else if desc.HasGetter {
		descObj["set"] = Value{Type: VTJSUndefined}
	}
	descObj["enumerable"] = NewBool(desc.Enumerable)
	descObj["configurable"] = NewBool(desc.Configurable)
	vm.jsObjectItems[descObjID] = descObj
	vm.jsPropertyItems[descObjID] = make(map[string]jsPropertyDescriptor, 8)
	for k, v := range descObj {
		vm.jsSetDescriptor(descObjID, k, jsDefaultPropertyDescriptor(v))
	}
	return Value{Type: VTJSObject, Num: descObjID}
}

func (vm *VM) jsValidateDefinePropertyTransition(current jsPropertyDescriptor, currentExists bool, spec jsDefinePropertySpec) bool {
	if spec.desc.HasValue && (spec.desc.HasGetter || spec.desc.HasSetter) {
		return false
	}
	if spec.hasWritable && (spec.desc.HasGetter || spec.desc.HasSetter) {
		return false
	}
	if !currentExists {
		return true
	}
	if current.Configurable {
		return true
	}
	if spec.hasConfigurable && spec.desc.Configurable {
		return false
	}
	if spec.hasEnumerable && spec.desc.Enumerable != current.Enumerable {
		return false
	}
	isCurrentAccessor := jsDescriptorIsAccessor(current)
	isSpecAccessor := spec.desc.HasGetter || spec.desc.HasSetter
	isSpecData := spec.desc.HasValue || spec.hasWritable
	if isSpecAccessor && jsDescriptorIsData(current) {
		return false
	}
	if isSpecData && isCurrentAccessor {
		return false
	}
	if isCurrentAccessor {
		if spec.desc.HasGetter && !vm.jsStrictEquals(spec.desc.Getter, current.Getter) {
			return false
		}
		if spec.desc.HasSetter && !vm.jsStrictEquals(spec.desc.Setter, current.Setter) {
			return false
		}
		return true
	}
	if !current.Writable {
		if spec.hasWritable && spec.desc.Writable {
			return false
		}
		if spec.desc.HasValue && !vm.jsStrictEquals(spec.desc.Value, current.Value) {
			return false
		}
	}
	return true
}

func (vm *VM) jsApplyDefinePropertySpec(current jsPropertyDescriptor, currentExists bool, spec jsDefinePropertySpec) jsPropertyDescriptor {
	result := current
	if !currentExists {
		result = jsPropertyDescriptor{Enumerable: false, Configurable: false, Writable: false}
	}
	isSpecAccessor := spec.desc.HasGetter || spec.desc.HasSetter
	isSpecData := spec.desc.HasValue || spec.hasWritable
	if isSpecAccessor {
		result.HasValue = false
		result.Writable = false
		result.Value = Value{Type: VTJSUndefined}
		if spec.desc.HasGetter {
			result.Getter = spec.desc.Getter
			result.HasGetter = true
		}
		if spec.desc.HasSetter {
			result.Setter = spec.desc.Setter
			result.HasSetter = true
		}
	} else if isSpecData {
		if jsDescriptorIsAccessor(result) {
			result.Getter = Value{Type: VTJSUndefined}
			result.Setter = Value{Type: VTJSUndefined}
			result.HasGetter = false
			result.HasSetter = false
		}
		if !currentExists {
			result.HasValue = true
			result.Value = Value{Type: VTJSUndefined}
		}
		if spec.desc.HasValue {
			result.HasValue = true
			result.Value = spec.desc.Value
		}
		if spec.hasWritable {
			result.Writable = spec.desc.Writable
		}
	}
	if spec.hasEnumerable {
		result.Enumerable = spec.desc.Enumerable
	}
	if spec.hasConfigurable {
		result.Configurable = spec.desc.Configurable
	}
	if !currentExists && !isSpecAccessor {
		if !result.HasValue {
			result.HasValue = true
			result.Value = Value{Type: VTJSUndefined}
		}
	}
	return result
}

func (vm *VM) jsEnumerateForInKeys(source Value) []string {
	if source.Type == VTJSProxy {
		rawKeys := vm.jsProxyOwnKeys(source)
		var keys []string
		for _, k := range rawKeys {
			desc, ok := vm.jsGetDescriptor(source.Num, k)
			if ok && desc.Enumerable {
				keys = append(keys, k)
			}
		}
		return keys
	}

	if source.Type == VTJSObject || source.Type == VTJSFunction {
		seen := make(map[string]struct{})
		var allKeys []string

		curr := source
		for curr.Type == VTJSObject || curr.Type == VTJSFunction {
			names := vm.jsObjectOwnPropertyNames(curr)
			for _, k := range names {
				if _, alreadySeen := seen[k]; !alreadySeen {
					desc, hasDesc := vm.jsGetDescriptor(curr.Num, k)
					if hasDesc && desc.Enumerable {
						allKeys = append(allKeys, k)
						seen[k] = struct{}{}
					} else if !hasDesc {
						// For virtual properties (like TypedArray indices) without descriptors,
						// assume they are enumerable if they showed up in OwnPropertyNames.
						allKeys = append(allKeys, k)
						seen[k] = struct{}{}
					}
				}
			}
			curr = vm.jsGetPrototypeValue(curr)
			if curr.Type == VTNull || curr.Type == VTJSUndefined {
				break
			}
		}
		return allKeys
	}

	if source.Type == VTArray && source.Arr != nil && len(source.Arr.Values) > 0 {
		keys := make([]string, 0, len(source.Arr.Values))
		for i := 0; i < len(source.Arr.Values); i++ {
			keys = append(keys, strconv.Itoa(source.Arr.Lower+i))
		}
		return keys
	}

	return nil
}

// jsEnumerateForOfValues collects the iterable values from a source for for...of traversal.
// Arrays yield their elements, Strings yield each character, JS Sets yield their members,
// JS Maps yield [key, value] pair arrays, and JS Arrays (VTJSObject with __js_class Array)
// yield indexed elements. All other types produce an empty slice.
func (vm *VM) jsEnumerateForOfValues(source Value) []Value {
	// Fast path: VBScript/internal Array — yield elements directly.
	if source.Type == VTArray && source.Arr != nil {
		out := make([]Value, len(source.Arr.Values))
		copy(out, source.Arr.Values)
		return out
	}

	// String — yield each character as a single-character string.
	if source.Type == VTString {
		runes := []rune(source.Str)
		out := make([]Value, len(runes))
		for i, r := range runes {
			out[i] = NewString(string(r))
		}
		return out
	}

	// JS Object — inspect the runtime class to decide how to iterate.
	if source.Type == VTJSObject {
		class := vm.jsObjectStringProperty(source, "__js_class")
		if class == "" {
			class = vm.jsObjectStringProperty(source, "__js_type")
		}
		if class == "" {
			class = vm.jsObjectStringProperty(source, "__js_ctor")
		}
		switch class {
		case "Array":
			// JS Array: iterate by numeric index up to .length.
			lengthVal, _ := vm.jsMemberGet(source, "length")
			n := int(vm.jsToNumber(lengthVal).Flt)
			out := make([]Value, 0, n)
			for i := range n {
				out = append(out, vm.jsIndexGet(source, NewInteger(int64(i))))
			}
			return out
		case "Set":
			// Set: iterate over unique members stored in jsSetItems.
			if setMap, ok := vm.jsSetItems[source.Num]; ok {
				out := make([]Value, 0, len(setMap))
				for _, v := range setMap {
					out = append(out, v)
				}
				return out
			}
		case "Map":
			// Map: iterate over [key, value] pairs stored in jsMapItems.
			if mapData, ok := vm.jsMapItems[source.Num]; ok {
				out := make([]Value, 0, len(mapData))
				for k, v := range mapData {
					out = append(out, ValueFromVBArray(NewVBArrayFromValues(0, []Value{vm.jsMapKeyToValue(k), v})))
				}
				return out
			}
		case "Array Iterator", "String Iterator":
			out := make([]Value, 0)
			for {
				result, handled := vm.jsCallMember(source, "next", nil)
				if !handled || result.Type != VTJSObject {
					break
				}
				doneVal, _ := vm.jsMemberGet(result, "done")
				if vm.jsTruthy(doneVal) {
					break
				}
				val, _ := vm.jsMemberGet(result, "value")
				out = append(out, val)
			}
			return out
		default:
			// Typed arrays are iterable
			if jsIsTypedArrayType(class) {
				if vals := vm.jsTypedArrayValues(source); vals != nil {
					return vals
				}
			}
		}
	}

	// Fallback: ES6 Iteration Protocol
	itKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolIterator, 10)
	itFn, _ := vm.jsMemberGet(source, itKey)
	if itFn.Type == VTJSFunction || itFn.Type == VTJSObject {
		itObj := vm.jsCall(itFn, source, nil)
		if itObj.Type == VTJSObject {
			out := make([]Value, 0)
			for {
				result, handled := vm.jsCallMember(itObj, "next", nil)
				if !handled || result.Type != VTJSObject {
					break
				}
				doneVal, _ := vm.jsMemberGet(result, "done")
				if vm.jsTruthy(doneVal) {
					break
				}
				val, _ := vm.jsMemberGet(result, "value")
				out = append(out, val)
			}
			return out
		}
	}

	return nil
}

func (vm *VM) jsSuperCall(args []Value) Value {
	if len(vm.jsCallStack) == 0 {
		vm.jsThrowReferenceError("super() call outside of function")
		return Value{Type: VTJSUndefined}
	}
	currentFrameIdx := len(vm.jsCallStack) - 1
	currentFrame := &vm.jsCallStack[currentFrameIdx]
	currentFn := currentFrame.fn
	if currentFn.Type != VTJSFunction {
		vm.jsThrowReferenceError("super() call outside of class constructor")
		return Value{Type: VTJSUndefined}
	}

	// 1. Get the super constructor
	superCtor := vm.jsGetPrototypeValue(currentFn)
	if superCtor.Type != VTJSFunction && superCtor.Type != VTJSObject {
		vm.jsThrowTypeError("Super constructor is not a constructor")
		return Value{Type: VTJSUndefined}
	}

	// 2. Call the super constructor
	// In ES6, super() call in derived constructor uses new.target
	newTarget := vm.jsNewTarget

	// Call the constructor as a super call
	return vm.jsConstruct(superCtor, args, newTarget, true)
}

func (vm *VM) jsProxyApply(proxy Value, thisVal Value, args []Value) Value {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return Value{Type: VTJSUndefined}
	}

	// 1. Get the 'apply' trap
	trap, _ := vm.jsMemberGet(pObj.Handler, "apply")

	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		// 2. Forward to target
		return vm.jsCall(pObj.Target, thisVal, args)
	}

	// 3. Invoke the trap: trap(target, thisArgument, argumentsList)
	// Create argumentsList array
	argList := ValueFromVBArray(NewVBArrayFromValues(0, args))
	trapArgs := []Value{pObj.Target, thisVal, argList}
	return vm.jsCall(trap, pObj.Handler, trapArgs)
}

func (vm *VM) jsProxyConstruct(proxy Value, args []Value, newTarget Value) Value {
	pObj, ok := vm.jsProxyItems[proxy.Num]
	if !ok || pObj.Revoked {
		vm.jsThrowJSError(jscript.ProxyTrapResultRevoked)
		return Value{Type: VTJSUndefined}
	}

	// 1. Get the 'construct' trap
	trap, _ := vm.jsMemberGet(pObj.Handler, "construct")

	if trap.Type == VTJSUndefined || trap.Type == VTNull {
		// 2. Forward to target
		return vm.jsConstruct(pObj.Target, args, newTarget, false)
	}

	// 3. Invoke the trap: trap(target, argumentsList, newTarget)
	argList := ValueFromVBArray(NewVBArrayFromValues(0, args))
	trapArgs := []Value{pObj.Target, argList, newTarget}
	result := vm.jsCall(trap, pObj.Handler, trapArgs)

	// 4. Result must be an object
	isObject := func(v Value) bool {
		switch v.Type {
		case VTJSObject, VTJSFunction, VTJSPromise, VTJSGenerator, VTJSProxy, VTNativeObject, VTArray, VTObject:
			return true
		}
		return false
	}
	if !isObject(result) {
		vm.jsThrowJSError(jscript.ProxyTrapReturnedInvalidValue)
		return Value{Type: VTJSUndefined}
	}
	return result
}

func (vm *VM) jsPrimitiveTypeName(v Value) string {
	switch v.Type {
	case VTBool:
		return "Boolean"
	case VTInteger, VTDouble:
		return "Number"
	case VTString:
		return "String"
	case VTSymbol:
		return "Symbol"
	case VTJSBigInt:
		return "BigInt"
	default:
		return "Object"
	}
}

func (vm *VM) dispatchJSIntrinsicCall(thisVal Value, ctorName string, args []Value) Value {
	// Internal helper to reuse native JScript function logic without creating fake callee objects.
	// This mimics the switch block in jsCall.
	switch ctorName {
	case "ReflectGet":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.get requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		key := vm.jsPropertyKeyFromValue(jsArgOrUndefined(args, 1))
		val, _ := vm.jsMemberGet(target, key)
		return val
	case "ReflectSet":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.set requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		key := vm.jsPropertyKeyFromValue(jsArgOrUndefined(args, 1))
		val := jsArgOrUndefined(args, 2)
		savedStrict := vm.jsStrictMode
		vm.jsStrictMode = false
		vm.jsMemberSet(target, key, val)
		vm.jsStrictMode = savedStrict
		return NewBool(true)
	case "ReflectHas":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.has requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		key := vm.jsPropertyKeyFromValue(jsArgOrUndefined(args, 1))
		return NewBool(vm.jsHas(target, key))
	case "ReflectDefineProperty":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.defineProperty requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		name := vm.jsPropertyKeyFromValue(jsArgOrUndefined(args, 1))
		attributes := jsArgOrUndefined(args, 2)
		if target.Type == VTJSProxy {
			return NewBool(vm.jsProxyDefineProperty(target, name, attributes))
		}
		objID := target.Num
		current, currentExists := vm.jsGetDescriptor(objID, name)
		if !currentExists && !vm.jsObjectIsExtensible(target) {
			return NewBool(false)
		}
		spec := vm.jsReadDefinePropertySpec(attributes)
		if !vm.jsValidateDefinePropertyTransition(current, currentExists, spec) {
			return NewBool(false)
		}
		finalDesc := vm.jsApplyDefinePropertySpec(current, currentExists, spec)
		vm.jsSetDescriptor(objID, name, finalDesc)
		return NewBool(true)
	case "ReflectGetOwnPropertyDescriptor":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.getOwnPropertyDescriptor requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		name := vm.jsPropertyKeyFromValue(jsArgOrUndefined(args, 1))
		var desc jsPropertyDescriptor
		var ok bool
		if target.Type == VTJSProxy {
			desc, ok = vm.jsProxyGetOwnPropertyDescriptor(target, name)
		} else {
			desc, ok = vm.jsGetDescriptor(target.Num, name)
		}
		if !ok {
			return Value{Type: VTJSUndefined}
		}
		return vm.jsBuildDescriptorObject(desc)
	case "ReflectGetPrototypeOf":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.getPrototypeOf requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		if target.Type == VTJSProxy {
			return vm.jsProxyGetPrototypeOf(target)
		}
		return vm.jsGetPrototypeValue(target)
	case "ReflectIsExtensible":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.isExtensible requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		if target.Type == VTJSProxy {
			return NewBool(vm.jsProxyIsExtensible(target))
		}
		return NewBool(vm.jsObjectIsExtensible(target))
	case "ReflectPreventExtensions":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.preventExtensions requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		if target.Type == VTJSProxy {
			return NewBool(vm.jsProxyPreventExtensions(target))
		}
		vm.jsSetObjectExtensible(target.Num, false)
		return NewBool(true)
	case "ReflectDeleteProperty":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.deleteProperty requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		key := vm.jsPropertyKeyFromValue(jsArgOrUndefined(args, 1))
		success := false
		if target.Type == VTJSProxy {
			success = vm.jsProxyDelete(target, key)
		} else {
			success = vm.jsMemberDelete(target, key)
		}
		return NewBool(success)
	case "ReflectOwnKeys":
		if len(args) < 1 {
			vm.jsThrowTypeError("Reflect.ownKeys requires at least 1 argument")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		var keys []string
		if target.Type == VTJSProxy {
			keys = vm.jsProxyOwnKeys(target)
		} else {
			keys = vm.jsObjectOwnKeys(target)
		}
		values := make([]Value, len(keys))
		for i := 0; i < len(keys); i++ {
			values[i] = NewString(keys[i])
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values))
	case "ReflectSetPrototypeOf":
		if len(args) < 2 {
			vm.jsThrowTypeError("Reflect.setPrototypeOf requires at least 2 arguments")
			return Value{Type: VTJSUndefined}
		}
		target := args[0]
		if target.Type != VTJSObject && target.Type != VTJSFunction && target.Type != VTArray && target.Type != VTJSProxy {
			vm.jsThrowJSError(jscript.ReflectArgumentNotObject)
			return Value{Type: VTJSUndefined}
		}
		proto := args[1]
		if proto.Type != VTJSObject && proto.Type != VTJSFunction && proto.Type != VTArray && proto.Type != VTNull && proto.Type != VTJSProxy {
			vm.jsThrowTypeError("Reflect.setPrototypeOf prototype must be an object or null")
			return Value{Type: VTJSUndefined}
		}
		if target.Type == VTJSProxy {
			return NewBool(vm.jsProxySetPrototypeOf(target, proto))
		}
		return NewBool(vm.jsSetPrototype(target, proto))
	case "ConsoleLog":
		consoleDispatch(vm, "log", args)
		return Value{Type: VTJSUndefined}
	case "ConsoleWarn":
		consoleDispatch(vm, "warn", args)
		return Value{Type: VTJSUndefined}
	case "ConsoleError":
		consoleDispatch(vm, "error", args)
		return Value{Type: VTJSUndefined}
	case "ConsoleInfo":
		consoleDispatch(vm, "info", args)
		return Value{Type: VTJSUndefined}
	case "ConsoleDebug":
		consoleDispatch(vm, "debug", args)
		return Value{Type: VTJSUndefined}
	case "ConsoleTrace":
		consoleDispatch(vm, "trace", args)
		return Value{Type: VTJSUndefined}
	case "ConsoleClear":
		consoleDispatch(vm, "clear", args)
		return Value{Type: VTJSUndefined}
	case "ProcessCwd":
		res, _ := vm.jsCallProcessMethod("cwd", args)
		return res
	case "ProcessExit":
		res, _ := vm.jsCallProcessMethod("exit", args)
		return res
	case "ProcessNextTick":
		res, _ := vm.jsCallProcessMethod("nextTick", args)
		return res
	case "OSHostname":
		res, _ := vm.jsCallOSMethod("hostname", args)
		return res
	case "OSType":
		res, _ := vm.jsCallOSMethod("type", args)
		return res
	case "OSPlatform":
		res, _ := vm.jsCallOSMethod("platform", args)
		return res
	case "OSArch":
		res, _ := vm.jsCallOSMethod("arch", args)
		return res
	case "OSRelease":
		res, _ := vm.jsCallOSMethod("release", args)
		return res
	case "OSUptime":
		res, _ := vm.jsCallOSMethod("uptime", args)
		return res
	case "OSFreemem":
		res, _ := vm.jsCallOSMethod("freemem", args)
		return res
	case "OSTotalmem":
		res, _ := vm.jsCallOSMethod("totalmem", args)
		return res
	case "OSCpus":
		res, _ := vm.jsCallOSMethod("cpus", args)
		return res
	case "PathJoin":
		res, _ := vm.jsCallPathMethod("join", args)
		return res
	case "PathResolve":
		res, _ := vm.jsCallPathMethod("resolve", args)
		return res
	case "PathBasename":
		res, _ := vm.jsCallPathMethod("basename", args)
		return res
	case "PathDirname":
		res, _ := vm.jsCallPathMethod("dirname", args)
		return res
	case "PathExtname":
		res, _ := vm.jsCallPathMethod("extname", args)
		return res
	case "PathNormalize":
		res, _ := vm.jsCallPathMethod("normalize", args)
		return res
	// fs module methods
	case "FSReadFile":
		res, _ := vm.jsCallFSMethod("readFile", args)
		return res
	case "FSReadFileSync":
		res, _ := vm.jsCallFSMethod("readFileSync", args)
		return res
	case "FSWriteFileSync":
		res, _ := vm.jsCallFSMethod("writeFileSync", args)
		return res
	case "FSExistsSync":
		res, _ := vm.jsCallFSMethod("existsSync", args)
		return res
	case "FSStatSync":
		res, _ := vm.jsCallFSMethod("statSync", args)
		return res
	// fs.promises module methods
	case "FSPromisesReadFile":
		res, _ := vm.jsCallFSPromisesMethod("readFile", args)
		return res
	// crypto module methods
	case "CryptoCreateHash":
		res, _ := vm.jsCallCryptoMethod("createHash", args)
		return res
	case "CryptoCreateHmac":
		res, _ := vm.jsCallCryptoMethod("createHmac", args)
		return res
	case "CryptoRandomBytes":
		res, _ := vm.jsCallCryptoMethod("randomBytes", args)
		return res
	case "HTTPCreateServer":
		res, _ := vm.jsCallHTTPMethod("http", "createServer", args)
		return res
	case "HTTPRequest":
		res, _ := vm.jsCallHTTPMethod("http", "request", args)
		return res
	case "HTTPGet":
		res, _ := vm.jsCallHTTPMethod("http", "get", args)
		return res
	case "HTTPServerListen":
		res, _ := vm.jsCallHTTPServerMethod(thisVal, "listen", args)
		return res
	case "QSParse":
		res, _ := vm.jsCallQueryStringMethod("parse", args)
		return res
	case "QSStringify":
		res, _ := vm.jsCallQueryStringMethod("stringify", args)
		return res
	case "SetTimeout":
		return vm.jsSetTimeout(args)
	case "ClearTimeout":
		return vm.jsClearTimeout(args)
	case "SetInterval":
		return vm.jsSetInterval(args)
	case "ClearInterval":
		return vm.jsClearInterval(args)
	case "SetImmediate":
		return vm.jsSetImmediate(args)
	case "ClearImmediate":
		return vm.jsClearImmediate(args)
	case "AtomicsAdd", "AtomicsSub", "AtomicsAnd", "AtomicsOr", "AtomicsXor", "AtomicsLoad", "AtomicsStore", "AtomicsExchange", "AtomicsCompareExchange", "AtomicsIsLockFree":
		methodName := ctorName[7:] // Strip "Atomics"
		res, _ := vm.jsAtomicsCall(methodName, args)
		return res
	}
	return Value{Type: VTJSUndefined}
}

func (vm *VM) jsCall(callee Value, thisVal Value, args []Value) Value {
	if !vm.jsIsCallable(callee) {
		vm.jsThrowTypeError(vm.valueToString(callee) + " is not a function")
		return Value{Type: VTJSUndefined}
	}
	switch callee.Type {
	case VTJSProxy:
		return vm.jsProxyApply(callee, thisVal, args)
	case VTJSFunction:
		if closure, ok := vm.jsFunctionItems[callee.Num]; ok && closure != nil {
			if closure.isClassConstructor {
				vm.jsThrowTypeError("Class constructor cannot be invoked without 'new'")
				return Value{Type: VTJSUndefined}
			}
			if closure.isGenerator {
				return vm.jsCreateGeneratorObject(callee, thisVal, args)
			}
			if closure.isAsync {
				return vm.jsAsyncCall(callee, thisVal, args)
			}
			if closure.isBound {
				callArgs := closure.boundArgs
				if len(args) > 0 {
					merged := make([]Value, 0, len(callArgs)+len(args))
					merged = append(merged, callArgs...)
					merged = append(merged, args...)
					callArgs = merged
				}
				return vm.jsCall(closure.boundFn, closure.boundThis, callArgs)
			}
			if len(vm.jsCallStack) >= jsMaxCallStackDepth {
				vm.jsThrowJSError(jscript.OutOfStackSpace)
				return Value{Type: VTJSUndefined}
			}
			child := vm.cloneForExecuteLocal(len(vm.bytecode))
			if child.jsBeginFunctionCall(callee, thisVal, args, Value{Type: VTJSUndefined}, false, Value{Type: VTJSUndefined}, false) {
				var childErr error
				var childThrow *jsAsyncRejectionError
				func() {
					defer func() {
						if r := recover(); r != nil {
							if are, ok := r.(*jsAsyncRejectionError); ok {
								childThrow = are
								return
							}
							panic(r)
						}
					}()
					childErr = child.Run()
				}()
				if childThrow != nil {
					vm.syncExecuteGlobalState(child)
					vm.jsThrow(childThrow.reason)
					return Value{Type: VTJSUndefined}
				}
				if childErr != nil {
					vm.syncExecuteGlobalState(child)
					if vmErr, ok := childErr.(*VMError); ok {
						vm.jsThrowJSError(jscript.JSSyntaxErrorCode(vmErr.Code))
						return Value{Type: VTJSUndefined}
					}
					vm.jsThrowTypeError(childErr.Error())
					return Value{Type: VTJSUndefined}
				}
				result := Value{Type: VTJSUndefined}
				if child.sp >= 0 {
					result = child.stack[child.sp]
				}
				vm.syncExecuteGlobalState(child)
				return result
			}
			return Value{Type: VTJSUndefined}
		}
		// Native JScript functions (without closures)
		ctorName := vm.jsObjectStringProperty(callee, "__js_ctor")
		switch ctorName {
		case "Proxy":
			vm.jsThrowTypeError("Constructor Proxy requires 'new'")
			return Value{Type: VTJSUndefined}
		case "ProxyRevocable":
			if len(args) < 2 {
				vm.jsThrowTypeError("Proxy.revocable requires 2 arguments: target and handler")
				return Value{Type: VTJSUndefined}
			}
			target := args[0]
			handler := args[1]
			isObject := func(v Value) bool {
				switch v.Type {
				case VTJSObject, VTJSFunction, VTJSPromise, VTJSGenerator, VTJSProxy, VTNativeObject, VTArray, VTObject:
					return true
				}
				return false
			}
			if !isObject(target) || !isObject(handler) {
				vm.jsThrowJSError(jscript.ProxyTargetOrHandlerNotObject)
				return Value{Type: VTJSUndefined}
			}
			proxyVal := vm.jsCreateProxy(args)
			if proxyVal.Type != VTJSProxy {
				return Value{Type: VTJSUndefined}
			}
			// Create revoke function
			revokeID := vm.allocJSID()
			revokeObj := make(map[string]Value, 4)
			revokeObj["__js_type"] = NewString("Function")
			revokeObj["__js_ctor"] = NewString("ProxyRevoke")
			revokeObj["__js_proxy_id"] = NewInteger(proxyVal.Num)
			revokeObj["name"] = NewString("")
			revokeObj["length"] = NewInteger(0)
			vm.jsObjectItems[revokeID] = revokeObj
			vm.jsPropertyItems[revokeID] = make(map[string]jsPropertyDescriptor, 4)
			revokeFn := Value{Type: VTJSFunction, Num: revokeID}

			// Return {proxy, revoke}
			resultID := vm.allocJSID()
			resultObj := make(map[string]Value, 2)
			resultObj["proxy"] = proxyVal
			resultObj["revoke"] = revokeFn
			vm.jsObjectItems[resultID] = resultObj
			vm.jsPropertyItems[resultID] = make(map[string]jsPropertyDescriptor, 4)
			return Value{Type: VTJSObject, Num: resultID}
		case "ProxyRevoke":
			proxyID := vm.jsObjectItems[callee.Num]["__js_proxy_id"].Num
			if p, ok := vm.jsProxyItems[proxyID]; ok {
				p.Revoked = true
			}
			return Value{Type: VTJSUndefined}
		case "ReflectGet", "ReflectSet", "ReflectHas", "ReflectDeleteProperty", "ReflectOwnKeys", "ReflectDefineProperty", "ReflectGetOwnPropertyDescriptor", "ReflectGetPrototypeOf", "ReflectIsExtensible", "ReflectPreventExtensions", "ReflectSetPrototypeOf",
			"AtomicsAdd", "AtomicsSub", "AtomicsAnd", "AtomicsOr", "AtomicsXor", "AtomicsLoad", "AtomicsStore", "AtomicsExchange", "AtomicsCompareExchange", "AtomicsIsLockFree",
			"ConsoleLog", "ConsoleWarn", "ConsoleError", "ConsoleInfo", "ConsoleDebug", "ConsoleTrace", "ConsoleClear",
			"ProcessCwd", "ProcessExit", "ProcessNextTick",
			"OSHostname", "OSType", "OSPlatform", "OSArch", "OSRelease", "OSUptime", "OSFreemem", "OSTotalmem", "OSCpus",
			"PathJoin", "PathResolve", "PathBasename", "PathDirname", "PathExtname", "PathNormalize",
			"FSReadFile", "FSReadFileSync", "FSWriteFileSync", "FSExistsSync", "FSStatSync",
			"FSPromisesReadFile",
			"CryptoCreateHash", "CryptoCreateHmac", "CryptoRandomBytes",
			"HTTPCreateServer", "HTTPRequest", "HTTPGet", "HTTPServerListen",
			"QSParse", "QSStringify", "QSEscape", "QSUnescape",
			"SetTimeout", "ClearTimeout", "SetInterval", "ClearInterval", "SetImmediate", "ClearImmediate",
			"ObjectToString", "ObjectToLocaleString", "ObjectValueOf":
			return vm.dispatchJSIntrinsicCall(thisVal, ctorName, args)
		case "ReflectApply":
			if len(args) < 1 {
				vm.jsThrowTypeError("Reflect.apply requires at least 1 argument")
				return Value{Type: VTJSUndefined}
			}
			target := args[0]
			if !vm.jsIsCallable(target) {
				vm.jsThrowTypeError("Reflect.apply target is not callable")
				return Value{Type: VTJSUndefined}
			}
			thisArg := jsArgOrUndefined(args, 1)
			applyArgs := vm.jsExtractApplyArgs(jsArgOrUndefined(args, 2))
			return vm.jsCall(target, thisArg, applyArgs)
		case "ReflectConstruct":
			if len(args) < 1 {
				vm.jsThrowTypeError("Reflect.construct requires at least 1 argument")
				return Value{Type: VTJSUndefined}
			}
			target := args[0]
			if !vm.jsIsConstructor(target) {
				vm.jsThrowTypeError("Reflect.construct target is not a constructor")
				return Value{Type: VTJSUndefined}
			}
			applyArgs := vm.jsExtractApplyArgs(jsArgOrUndefined(args, 1))
			newTarget := target
			if len(args) > 2 {
				newTarget = args[2]
				if !vm.jsIsConstructor(newTarget) {
					vm.jsThrowTypeError("Reflect.construct newTarget is not a constructor")
					return Value{Type: VTJSUndefined}
				}
			}
			return vm.jsConstruct(target, applyArgs, newTarget, false)
		case "SpeciesGetter":
			return thisVal
		case "ArrayValues":
			return vm.jsCreateArrayIterator(thisVal, 0)
		case "ArrayKeys":
			return vm.jsCreateArrayIterator(thisVal, 1)
		case "ArrayEntries":
			return vm.jsCreateArrayIterator(thisVal, 2)
		case "StringIteratorFactory":
			return vm.jsCreateStringIterator(vm.valueToString(thisVal))
		case "RegExpStringIteratorIterator":
			return thisVal
		case "StringPrototypeMatchAll":
			if len(args) == 0 {
				reCtor := vm.jsGetName("RegExp")
				rx := vm.jsNew(reCtor, []Value{Value{Type: VTJSUndefined}})
				return vm.jsRegExpMatchAll(rx, vm.valueToString(thisVal))
			}
			regexp := args[0]
			if regexp.Type != VTNull && regexp.Type != VTJSUndefined {
				isRegExp := false
				if regexp.Type == VTJSObject {
					if vm.jsObjectStringProperty(regexp, "__js_type") == "RegExp" {
						isRegExp = true
					}
				}
				if isRegExp {
					flags := vm.jsObjectStringProperty(regexp, "flags")
					if !strings.Contains(flags, "g") {
						vm.jsThrowTypeError("String.prototype.matchAll called with a non-global RegExp argument")
						return Value{Type: VTJSUndefined}
					}
				}
			}
			reCtor := vm.jsGetName("RegExp")
			rx := vm.jsNew(reCtor, []Value{regexp, NewString("g")})
			return vm.jsRegExpMatchAll(rx, vm.valueToString(thisVal))
		case "RegExpPrototypeMatchAll":
			if thisVal.Type != VTJSObject && thisVal.Type != VTJSProxy {
				vm.jsThrowTypeError("RegExp.prototype[Symbol.matchAll] called on non-object")
				return Value{Type: VTJSUndefined}
			}
			input := ""
			if len(args) > 0 {
				input = vm.valueToString(args[0])
			}
			return vm.jsRegExpMatchAll(thisVal, input)
		case "PromiseResolve":
			promise := vm.jsObjectItems[callee.Num]["__js_promise"]
			resolution := jsArgOrUndefined(args, 0)
			vm.jsResolvePromise(promise, resolution)
			return Value{Type: VTJSUndefined}
		case "PromiseReject":
			promise := vm.jsObjectItems[callee.Num]["__js_promise"]
			reason := jsArgOrUndefined(args, 0)
			vm.jsRejectPromise(promise, reason)
			return Value{Type: VTJSUndefined}
		case "PromiseStaticResolve":
			return vm.jsPromiseStaticResolve(args)
		case "PromiseStaticReject":
			return vm.jsPromiseStaticReject(args)
		case "PromiseStaticAll":
			return vm.jsPromiseStaticAll(args)
		case "PromiseStaticRace":
			return vm.jsPromiseStaticRace(args)
		case "PromiseStaticAllSettled":
			return vm.jsPromiseStaticAllSettled(args)
		case "PromiseStaticAny":
			return vm.jsPromiseStaticAny(args)
		case "PromiseStaticWithResolvers":
			return vm.jsPromiseStaticWithResolvers(args)
		case "PromiseAllResolver":
			vm.jsHandlePromiseAllResolver(callee, args)
			return Value{Type: VTJSUndefined}
		case "PromiseAllSettledHandler":
			vm.jsHandlePromiseAllSettledHandler(callee, args)
			return Value{Type: VTJSUndefined}
		case "PromiseAnyRejecter":
			vm.jsHandlePromiseAnyRejecter(callee, args)
			return Value{Type: VTJSUndefined}
		case "PromiseFinallyHandler":
			return vm.jsHandlePromiseFinallyHandler(callee, args)
		case "PromiseConstantHandler":
			return vm.jsHandlePromiseConstantHandler(callee, args)
		case "PromisePrototypeThen":
			return vm.jsPromiseThen(thisVal, args)
		case "PromisePrototypeCatch":
			return vm.jsPromiseCatch(thisVal, args)
		case "PromisePrototypeFinally":
			return vm.jsPromiseFinally(thisVal, args)
		case "IntlDateTimeFormatFormat":
			return vm.jsIntlDateTimeFormatFormat(callee, thisVal, args)
		case "IntlDateTimeFormatFormatToParts":
			return vm.jsIntlDateTimeFormatFormatToParts(callee, thisVal, args)
		case "IntlNumberFormatFormat":
			return vm.jsIntlNumberFormatFormat(callee, thisVal, args)
		case "IntlNumberFormatFormatToParts":
			return vm.jsIntlNumberFormatFormatToParts(callee, thisVal, args)
		case "IntlCollatorCompare":
			return vm.jsIntlCollatorCompare(callee, thisVal, args)
		case "IntlPluralRulesSelect":
			return vm.jsIntlPluralRulesSelect(callee, thisVal, args)
		case "IntlRelativeTimeFormatFormat":
			return vm.jsIntlRelativeTimeFormatFormat(callee, thisVal, args)
		case "IntlRelativeTimeFormatFormatToParts":
			return vm.jsIntlRelativeTimeFormatFormatToParts(callee, thisVal, args)
		case "ObjectPrototype":
			return vm.jsCallObjectPrototypeMethod(thisVal, vm.jsObjectStringProperty(callee, "name"), args)
		case "Set":
			return vm.jsCallKeyedCollectionMethod(thisVal, "Set", vm.jsObjectStringProperty(callee, "name"), args)
		case "Map":
			return vm.jsCallKeyedCollectionMethod(thisVal, "Map", vm.jsObjectStringProperty(callee, "name"), args)
		case "WeakMap":
			return vm.jsCallWeakCollectionMethod(thisVal, "WeakMap", vm.jsObjectStringProperty(callee, "name"), args)
		case "WeakSet":
			return vm.jsCallWeakCollectionMethod(thisVal, "WeakSet", vm.jsObjectStringProperty(callee, "name"), args)
		case "WeakRef":
			return vm.jsCallWeakRefMethod(thisVal, vm.jsObjectStringProperty(callee, "name"), args)
		case "FinalizationRegistry":
			return vm.jsCallFinalizationRegistryMethod(thisVal, vm.jsObjectStringProperty(callee, "name"), args)
		case "GeneratorPrototypeNext":
			return vm.jsHandleGeneratorNext(thisVal, args)
		case "GeneratorPrototypeThrow":
			return vm.jsHandleGeneratorThrow(thisVal, args)
		case "GeneratorPrototypeReturn":
			return vm.jsHandleGeneratorReturn(thisVal, args)
		}
		return Value{Type: VTJSUndefined}
	case VTNativeObject:
		return vm.dispatchNativeCall(callee.Num, "", args)
	case VTJSObject:
		ctorName := vm.jsObjectStringProperty(callee, "__js_ctor")
		switch ctorName {
		case "Number":
			if len(args) == 0 {
				return NewDouble(0)
			}
			return vm.jsToNumber(args[0])
		case "String":
			if len(args) == 0 {
				return NewString("")
			}
			return NewString(vm.jsToString(args[0]))
		case "Boolean":
			if len(args) == 0 {
				return NewBool(false)
			}
			return NewBool(vm.jsTruthy(args[0]))
		case "Error", "TypeError", "ReferenceError", "SyntaxError", "RangeError", "EvalError", "URIError":
			msg := ""
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				msg = vm.jsToString(args[0])
			}
			return vm.jsCreateErrorObject(ctorName, msg)
		case "Symbol":
			desc := ""
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				desc = vm.valueToString(args[0])
			}
			return Value{Type: VTSymbol, Num: vm.jsAllocSymbolID(), Str: desc}
		case "Proxy":
			vm.jsThrowTypeError("Constructor Proxy requires 'new'")
			return Value{Type: VTJSUndefined}
		case "Set", "Map", "WeakMap", "WeakSet":
			vm.jsThrowTypeError(fmt.Sprintf("Constructor %s requires 'new'", ctorName))
			return Value{Type: VTJSUndefined}
		case "URL", "URLSearchParams":
			vm.jsEnsureURLConstructorWithNew(ctorName)
			return Value{Type: VTJSUndefined}
		case "ArrayBuffer", "DataView",
			"Int8Array", "Uint8Array", "Uint8ClampedArray",
			"Int16Array", "Uint16Array",
			"Int32Array", "Uint32Array",
			"Float32Array", "Float64Array",
			"BigInt64Array", "BigUint64Array":
			vm.jsThrowTypeError(fmt.Sprintf("Constructor %s requires 'new'", ctorName))
			return Value{Type: VTJSUndefined}
		case "isNaN":
			if len(args) == 0 {
				return NewBool(true)
			}
			return NewBool(math.IsNaN(vm.jsToNumber(args[0]).Flt))
		case "isFinite":
			if len(args) == 0 {
				return NewBool(false)
			}
			num := vm.jsToNumber(args[0]).Flt
			return NewBool(!math.IsNaN(num) && !math.IsInf(num, 0))
		case "parseInt":
			return vm.jsParseIntES5(args)
		case "parseFloat":
			return vm.jsParseFloatES5(args)
		case "require":
			return vm.jsRequire(args)
		case "decodeURI":
			decoded, err := jsDecodeURIValue(vm.valueToString(jsArgOrUndefined(args, 0)))
			if err != nil {
				vm.jsThrowTypeError("URI malformed")
				return Value{Type: VTJSUndefined}
			}
			return NewString(decoded)
		case "decodeURIComponent":
			decoded, err := jsDecodeURIComponentValue(vm.valueToString(jsArgOrUndefined(args, 0)))
			if err != nil {
				vm.jsThrowTypeError("URI malformed")
				return Value{Type: VTJSUndefined}
			}
			return NewString(decoded)
		case "encodeURI":
			return NewString(jsEncodeURIValue(vm.valueToString(jsArgOrUndefined(args, 0)), false))
		case "encodeURIComponent":
			return NewString(jsEncodeURIValue(vm.valueToString(jsArgOrUndefined(args, 0)), true))
		}
		return Value{Type: VTJSUndefined}
	case VTBuiltin:
		idx := int(callee.Num)
		if idx < 0 || idx >= len(BuiltinRegistry) {
			return Value{Type: VTJSUndefined}
		}
		if idx < len(BuiltinNames) && strings.EqualFold(BuiltinNames[idx], "Eval") {
			return vm.jsEval(args)
		}
		result, err := BuiltinRegistry[idx](vm, args)
		if err != nil {
			vm.raise(vbscript.InternalError, err.Error())
			return Value{Type: VTJSUndefined}
		}
		return result
	case VTObject:
		defaultProperty, ok := vm.resolveRuntimeClassPropertyGet(callee, "__default__", len(args), true)
		if ok {
			if vm.beginUserSubCall(defaultProperty, args, false, callee.Num) {
				return Value{Type: VTJSUndefined}
			}
		}
		return Value{Type: VTJSUndefined}
	default:
		return Value{Type: VTJSUndefined}
	}
}

func (vm *VM) jsThrow(v Value) {
	if len(vm.jsTryStack) == 0 {
		panic(&jsAsyncRejectionError{reason: v})
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, v)
	vm.ip = target
}

// jsRaiseRuntimeError raises an uncaught JavaScript runtime error in ASP-compatible shape.
func (vm *VM) jsRaiseRuntimeError(code jscript.JSSyntaxErrorCode, msg string) {
	description := strings.TrimSpace(msg)
	if description == "" {
		description = code.String()
	}

	vbCode := vbscript.VBSyntaxErrorCode(code)
	vme := &VMError{
		Code:           vbCode,
		Line:           vm.lastLine,
		Column:         vm.lastColumn,
		Msg:            description,
		ASPCode:        int(code),
		ASPDescription: description,
		Category:       "JavaScript runtime",
		Description:    description,
		Number:         jscript.HRESULTFromJScriptCode(code),
		Source:         "JavaScript runtime error",
	}

	vm.errSetFromVMError(vme)

	if vm.onResumeNext || vm.executeGlobalResumeGuard {
		vm.lastError = vme
		vm.skipToNextStmt = true
		return
	}

	panic(vme)
}

// jsThrowError throws a standard JScript Error that can be caught by a JS try/catch.
func (vm *VM) jsThrowError(msg string) {
	if len(vm.jsTryStack) == 0 {
		vm.jsRaiseRuntimeError(jscript.InternalError, msg)
		return
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, vm.jsCreateErrorObject("Error", msg))
	vm.ip = target
}

// jsThrowTypeError throws a JScript TypeError that can be caught by a JS try/catch.
// If no active catch handler exists, raises a VBScript TypeMismatch error instead.
func (vm *VM) jsThrowTypeError(msg string) {
	if len(vm.jsTryStack) == 0 {
		vm.jsRaiseRuntimeError(jscript.TypeMismatch, msg)
		return
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, vm.jsCreateErrorObject("TypeError", msg))
	vm.ip = target
}

// jsThrowJSError throws a specific JScript error code.
func (vm *VM) jsThrowJSError(code jscript.JSSyntaxErrorCode) {
	msg := code.String()
	if len(vm.jsTryStack) == 0 {
		vm.jsRaiseRuntimeError(code, msg)
		return
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, vm.jsCreateErrorObject("TypeError", msg)) // Most ES6+ errors are TypeErrors
	vm.ip = target
}

// jsThrowReferenceError throws a JScript ReferenceError that can be caught by a JS try/catch.
// If no active catch handler exists, raises a VBScript VariableNotDefined error instead.
func (vm *VM) jsThrowReferenceError(msg string) {
	if len(vm.jsTryStack) == 0 {
		vm.jsRaiseRuntimeError(jscript.UndefinedIdentifier, msg)
		return
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, vm.jsCreateErrorObject("ReferenceError", msg))
	vm.ip = target
}

// jsToNumber converts a Value to a numeric value (VTDouble) following JScript semantics.
// jsToPrimitive implements the ECMAScript ToPrimitive abstract operation.
func (vm *VM) jsToPrimitive(v Value, hint string) Value {
	if v.Type != VTJSObject && v.Type != VTJSFunction && v.Type != VTJSProxy && v.Type != VTArray {
		return v
	}
	if hint == "string" && v.Type == VTJSObject && vm.jsObjectStringProperty(v, "__js_type") == "Error" {
		name := vm.jsObjectStringProperty(v, "name")
		msg := vm.jsObjectStringProperty(v, "message")
		if name == "" {
			name = vm.jsObjectStringProperty(v, "__js_ctor")
		}
		if name == "" {
			name = "Error"
		}
		if msg == "" {
			return NewString(name)
		}
		return NewString(name + ": " + msg)
	}

	// 1. Check for @@toPrimitive
	toPrimitive := jsWellKnownSymbolValue(jsWellKnownSymbolToPrimitive, "Symbol.toPrimitive")
	if method, ok := vm.jsResolveObjectMember(v.Num, vm.jsPropertyKeyFromValue(toPrimitive), make(map[int64]struct{}, 4)); ok {
		if method.HasValue && vm.jsIsCallable(method.Value) {
			res := vm.jsCall(method.Value, v, []Value{NewString(hint)})
			if res.Type != VTJSObject && res.Type != VTJSFunction && res.Type != VTJSProxy && res.Type != VTArray {
				return res
			}
			vm.jsThrowTypeError("Symbol.toPrimitive returned an object")
			return Value{Type: VTJSUndefined}
		}
	}

	// 2. Default OrdinaryToPrimitive
	if hint == "string" {
		// toString then valueOf
		if res, ok := vm.jsCallOrdinaryToPrimitive(v, "toString"); ok {
			return res
		}
		if res, ok := vm.jsCallOrdinaryToPrimitive(v, "valueOf"); ok {
			return res
		}
	} else {
		// valueOf then toString
		if res, ok := vm.jsCallOrdinaryToPrimitive(v, "valueOf"); ok {
			return res
		}
		if res, ok := vm.jsCallOrdinaryToPrimitive(v, "toString"); ok {
			return res
		}
	}

	vm.jsThrowTypeError("Cannot convert object to primitive value")
	return Value{Type: VTJSUndefined}
}

func (vm *VM) jsCallOrdinaryToPrimitive(v Value, methodName string) (Value, bool) {
	method, deferred := vm.jsMemberGet(v, methodName)
	if !deferred && vm.jsIsCallable(method) {
		res := vm.jsCall(method, v, nil)
		if res.Type != VTJSObject && res.Type != VTJSFunction && res.Type != VTJSProxy && res.Type != VTArray {
			return res, true
		}
	}
	return Value{}, false
}

// jsToObject implements the ECMAScript ToObject abstract operation.
func (vm *VM) jsToObject(v Value) Value {
	switch v.Type {
	case VTJSUndefined, VTNull, VTEmpty:
		vm.jsThrowTypeError("Cannot convert undefined or null to object")
		return Value{Type: VTJSUndefined}
	case VTJSObject, VTJSFunction, VTJSProxy, VTArray:
		return v
	case VTBool:
		return vm.jsNew(vm.jsGetName("Boolean"), []Value{v})
	case VTInteger, VTDouble:
		return vm.jsNew(vm.jsGetName("Number"), []Value{v})
	case VTString:
		return vm.jsNew(vm.jsGetName("String"), []Value{v})
	case VTSymbol:
		// Symbols are handled differently, they don't have a direct 'new' constructor in ES6
		// but we can wrap them in an object for ToObject.
		objID := vm.allocJSID()
		obj := make(map[string]Value, 4)
		obj["__js_type"] = NewString("Symbol")
		obj["__js_primitive_value"] = v
		if proto := vm.jsGetIntrinsicPrototype("Symbol"); proto.Type == VTJSObject {
			obj["__js_proto"] = proto
		}
		vm.jsObjectItems[objID] = obj
		vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
		return Value{Type: VTJSObject, Num: objID}
	default:
		return Value{Type: VTJSUndefined}
	}
}

func (vm *VM) jsToString(v Value) string {
	if v.Type == VTArgRef {
		v = vm.stack[int(v.Num)]
	}
	switch v.Type {
	case VTJSUndefined:
		return "undefined"
	case VTNull, VTEmpty:
		return "null"
	case VTBool:
		if v.Num != 0 {
			return "true"
		}
		return "false"
	case VTInteger:
		return strconv.FormatInt(v.Num, 10)
	case VTDouble:
		if math.IsNaN(v.Flt) {
			return "NaN"
		}
		if math.IsInf(v.Flt, 1) {
			return "Infinity"
		}
		if math.IsInf(v.Flt, -1) {
			return "-Infinity"
		}
		if v.Flt == 0 {
			return "0"
		}
		return strconv.FormatFloat(v.Flt, 'g', -1, 64)
	case VTString:
		return v.Str
	case VTJSBigInt:
		if v.Big == nil {
			return "0"
		}
		return v.Big.String()
	case VTSymbol:
		vm.jsThrowTypeError("Cannot convert a Symbol value to a string")
		return ""
	case VTJSObject, VTJSFunction, VTJSProxy, VTArray:
		prim := vm.jsToPrimitive(v, "string")
		return vm.jsToString(prim)
	default:
		return vm.valueToString(v)
	}
}

func (vm *VM) jsToNumber(v Value) Value {
	switch v.Type {
	case VTJSBigInt:
		vm.jsThrowTypeError("Cannot convert a BigInt value to a number")
		return NewDouble(math.NaN())
	case VTJSUndefined:
		return NewDouble(math.NaN())
	case VTNull, VTEmpty:
		return NewDouble(0)
	case VTInteger:
		return Value{Type: VTDouble, Flt: float64(v.Num)}
	case VTDouble:
		return v
	case VTString:
		trimmed := strings.TrimSpace(v.Str)
		if trimmed == "" {
			return NewDouble(0)
		}
		parsed, err := strconv.ParseFloat(trimmed, 64)
		if err != nil {
			return NewDouble(math.NaN())
		}
		return NewDouble(parsed)
	case VTBool:
		if v.Num != 0 {
			return NewDouble(1)
		}
		return NewDouble(0)
	case VTDate:
		return NewDouble(float64(v.Num) / float64(time.Millisecond))
	case VTJSObject, VTJSFunction, VTJSProxy, VTArray:
		prim := vm.jsToPrimitive(v, "number")
		return vm.jsToNumber(prim)
	default:
		return NewDouble(math.NaN())
	}
}

// jsToInt32 converts a value to a 32-bit signed integer for bitwise operations.
func (vm *VM) jsToInteger(v Value) float64 {
	number := vm.jsToNumber(v).Flt
	if math.IsNaN(number) {
		return 0
	}
	if number == 0 || math.IsInf(number, 0) {
		return number
	}
	res := math.Trunc(number)
	if res == 0 && number < 0 {
		return math.Copysign(0, -1) // maintain signed zero
	}
	return res
}

func (vm *VM) jsToInt32(v Value) int32 {
	num := vm.jsToInteger(NewDouble(vm.jsToNumber(v).Flt)) // Simplified but usually enough
	return int32(num)
}

// jsToUint32 converts a value to a 32-bit unsigned integer for bitwise operations.
func (vm *VM) jsToUint32(v Value) uint32 {
	num := vm.jsToNumber(v).Flt
	return uint32(int32(num))
}

// jsToUint32Exact converts a value to uint32 using ECMAScript ToUint32 semantics.
func (vm *VM) jsToUint32Exact(v Value) uint32 {
	num := vm.jsToNumber(v).Flt
	if math.IsNaN(num) || math.IsInf(num, 0) || num == 0 {
		return 0
	}
	truncated := math.Trunc(num)
	mod := math.Mod(truncated, 4294967296.0)
	if mod < 0 {
		mod += 4294967296.0
	}
	return uint32(mod)
}

// jsAdd implements JScript '+' operator (string concatenation or numeric addition).
func (vm *VM) jsAdd(a Value, b Value) Value {
	a = resolveCallable(vm, a)
	b = resolveCallable(vm, b)
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		res := new(big.Int).Add(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		if a.Type == VTString || b.Type == VTString {
			// Fall through to string concatenation
		} else {
			vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
			return Value{Type: VTJSUndefined}
		}
	}
	return vm.jsAddValues(a, b)
}

// jsSubtract implements JScript '-' operator (numeric subtraction).
func (vm *VM) jsSubtract(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		res := new(big.Int).Sub(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	if a.Type == VTInteger && b.Type == VTInteger {
		if diff, ok := jsSubtractIntegersNoOverflow(a.Num, b.Num); ok {
			return NewInteger(diff)
		}
		return NewDouble(float64(a.Num) - float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewDouble(aNum - bNum)
}

// jsMultiply implements JScript '*' operator.
func (vm *VM) jsMultiply(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		res := new(big.Int).Mul(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewDouble(aNum * bNum)
}

// jsDivide implements JScript '/' operator.
func (vm *VM) jsDivide(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		if b.Big.Sign() == 0 {
			vm.jsThrowTypeError("Division by zero")
			return Value{Type: VTJSUndefined}
		}
		res := new(big.Int).Div(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	if math.IsNaN(aNum) || math.IsNaN(bNum) {
		return NewDouble(math.NaN())
	}
	if bNum == 0 {
		if aNum == 0 {
			return NewDouble(math.NaN())
		}
		if aNum > 0 {
			return NewDouble(math.Inf(1))
		}
		return NewDouble(math.Inf(-1))
	}
	return NewDouble(aNum / bNum)
}

// jsModulo implements JScript '%' operator.
func (vm *VM) jsModulo(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		if b.Big.Sign() == 0 {
			vm.jsThrowTypeError("Division by zero")
			return Value{Type: VTJSUndefined}
		}
		res := new(big.Int).Mod(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	if bNum == 0 || math.IsNaN(aNum) || math.IsNaN(bNum) || math.IsInf(aNum, 0) || math.IsInf(bNum, 0) {
		return NewDouble(math.NaN()) // NaN for modulo by zero
	}
	return NewDouble(math.Mod(aNum, bNum))
}

// jsExponent implements JScript '**' operator.
func (vm *VM) jsExponent(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		if b.Big.Sign() < 0 {
			vm.jsThrowTypeError("Exponent must be positive")
			return Value{Type: VTJSUndefined}
		}
		if !b.Big.IsUint64() {
			vm.jsThrowTypeError("Exponent too large")
			return Value{Type: VTJSUndefined}
		}
		res := new(big.Int).Exp(a.Big, b.Big, nil)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewDouble(math.Pow(aNum, bNum))
}

// jsNegate implements JScript unary '-' operator.
func (vm *VM) jsNegate(v Value) Value {
	if v.Type == VTJSBigInt {
		res := new(big.Int).Neg(v.Big)
		return NewBigInt(res)
	}
	num := vm.jsToNumber(v).Flt
	return NewDouble(-num)
}

// jsBitwiseAnd implements JScript '&' operator.
func (vm *VM) jsBitwiseAnd(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToInt32(b)
	return NewInteger(int64(aInt & bInt))
}

// jsBitwiseOr implements JScript '|' operator.
func (vm *VM) jsBitwiseOr(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToInt32(b)
	return NewInteger(int64(aInt | bInt))
}

// jsBitwiseXor implements JScript '^' operator.
func (vm *VM) jsBitwiseXor(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToInt32(b)
	return NewInteger(int64(aInt ^ bInt))
}

// jsBitwiseNot implements JScript '~' operator (bitwise NOT).
func (vm *VM) jsBitwiseNot(v Value) Value {
	vInt := vm.jsToInt32(v)
	return NewInteger(int64(^vInt))
}

// jsLeftShift implements JScript '<<' operator.
func (vm *VM) jsLeftShift(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToUint32(b) & 0x1f // Only use lower 5 bits
	return NewInteger(int64(aInt << bInt))
}

// jsRightShift implements JScript '>>' operator (sign-extending right shift).
func (vm *VM) jsRightShift(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToUint32(b) & 0x1f // Only use lower 5 bits
	return NewInteger(int64(aInt >> bInt))
}

// jsUnsignedRightShift implements JScript '>>>' operator (zero-filling right shift).
func (vm *VM) jsUnsignedRightShift(a Value, b Value) Value {
	aUint := vm.jsToUint32(a)
	bInt := vm.jsToUint32(b) & 0x1f // Only use lower 5 bits
	return NewInteger(int64(aUint >> bInt))
}

// jsLess implements JScript '<' operator.
func (vm *VM) jsLess(a Value, b Value) Value {
	// Fast path: both integers — avoid jsToNumber call and int→float conversion.
	if a.Type == VTInteger && b.Type == VTInteger {
		return NewBool(a.Num < b.Num)
	}
	// Fast path: both doubles.
	if a.Type == VTDouble && b.Type == VTDouble {
		return NewBool(a.Flt < b.Flt)
	}
	// Fast path: integer vs double mixed.
	if a.Type == VTInteger && b.Type == VTDouble {
		return NewBool(float64(a.Num) < b.Flt)
	}
	if a.Type == VTDouble && b.Type == VTInteger {
		return NewBool(a.Flt < float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum < bNum)
}

// jsGreater implements JScript '>' operator.
func (vm *VM) jsGreater(a Value, b Value) Value {
	// Fast path: both integers.
	if a.Type == VTInteger && b.Type == VTInteger {
		return NewBool(a.Num > b.Num)
	}
	// Fast path: both doubles.
	if a.Type == VTDouble && b.Type == VTDouble {
		return NewBool(a.Flt > b.Flt)
	}
	// Fast path: integer vs double mixed.
	if a.Type == VTInteger && b.Type == VTDouble {
		return NewBool(float64(a.Num) > b.Flt)
	}
	if a.Type == VTDouble && b.Type == VTInteger {
		return NewBool(a.Flt > float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum > bNum)
}

// jsLessEqual implements JScript '<=' operator.
func (vm *VM) jsLessEqual(a Value, b Value) Value {
	// Fast path: both integers.
	if a.Type == VTInteger && b.Type == VTInteger {
		return NewBool(a.Num <= b.Num)
	}
	// Fast path: both doubles.
	if a.Type == VTDouble && b.Type == VTDouble {
		return NewBool(a.Flt <= b.Flt)
	}
	// Fast path: integer vs double mixed.
	if a.Type == VTInteger && b.Type == VTDouble {
		return NewBool(float64(a.Num) <= b.Flt)
	}
	if a.Type == VTDouble && b.Type == VTInteger {
		return NewBool(a.Flt <= float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum <= bNum)
}

// jsGreaterEqual implements JScript '>=' operator.
func (vm *VM) jsGreaterEqual(a Value, b Value) Value {
	// Fast path: both integers.
	if a.Type == VTInteger && b.Type == VTInteger {
		return NewBool(a.Num >= b.Num)
	}
	// Fast path: both doubles.
	if a.Type == VTDouble && b.Type == VTDouble {
		return NewBool(a.Flt >= b.Flt)
	}
	// Fast path: integer vs double mixed.
	if a.Type == VTInteger && b.Type == VTDouble {
		return NewBool(float64(a.Num) >= b.Flt)
	}
	if a.Type == VTDouble && b.Type == VTInteger {
		return NewBool(a.Flt >= float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum >= bNum)
}

// jsLooseEqual implements JScript '==' (loose equality) operator.
func (vm *VM) jsLooseEqual(a Value, b Value) Value {
	a = resolveCallable(vm, a)
	b = resolveCallable(vm, b)

	if vm.jsStrictEquals(a, b) {
		return NewBool(true)
	}

	// 1. null == undefined (and VTEmpty, which is VBScript's "Empty", is treated
	//    as a JS-side absence-of-value when returned from native Go methods).
	aNullish := a.Type == VTNull || a.Type == VTJSUndefined || a.Type == VTEmpty
	bNullish := b.Type == VTNull || b.Type == VTJSUndefined || b.Type == VTEmpty
	if aNullish && bNullish {
		return NewBool(true)
	}
	if aNullish || bNullish {
		return NewBool(false)
	}

	// 2. Coerce types
	// ES5 11.9.3

	// String and Number
	if (a.Type == VTString && (b.Type == VTInteger || b.Type == VTDouble)) ||
		((a.Type == VTInteger || a.Type == VTDouble) && b.Type == VTString) {
		return vm.jsLooseEqual(vm.jsToNumber(a), vm.jsToNumber(b))
	}

	// Boolean as Number
	if a.Type == VTBool {
		return vm.jsLooseEqual(vm.jsToNumber(a), b)
	}
	if b.Type == VTBool {
		return vm.jsLooseEqual(a, vm.jsToNumber(b))
	}

	// Object and (String or Number or Symbol)
	isObject := func(v Value) bool {
		return v.Type == VTJSObject || v.Type == VTJSFunction || v.Type == VTJSProxy || v.Type == VTArray
	}
	isPrimitive := func(v Value) bool {
		return v.Type == VTString || v.Type == VTInteger || v.Type == VTDouble || v.Type == VTSymbol
	}

	if isObject(a) && isPrimitive(b) {
		return vm.jsLooseEqual(vm.jsToPrimitive(a, ""), b)
	}
	if isPrimitive(a) && isObject(b) {
		return vm.jsLooseEqual(a, vm.jsToPrimitive(b, ""))
	}

	// BigInt cases (ES2020)
	if (a.Type == VTJSBigInt && isPrimitive(b) && b.Type != VTSymbol) ||
		(isPrimitive(a) && a.Type != VTSymbol && b.Type == VTJSBigInt) {
		// Simplified BigInt comparison
		return NewBool(vm.jsToString(a) == vm.jsToString(b))
	}

	return NewBool(false)
}

// jsLooseNotEqual implements JScript '!=' (loose inequality) operator.
func (vm *VM) jsLooseNotEqual(a Value, b Value) Value {
	eq := vm.jsLooseEqual(a, b)
	return NewBool(eq.Num == 0)
}

// jsLogicalAnd implements JScript '&&' operator.
func (vm *VM) jsLogicalAnd(a Value, b Value) Value {
	if !vm.jsTruthy(a) {
		return a
	}
	return b
}

// jsLogicalOr implements JScript '||' operator.
func (vm *VM) jsLogicalOr(a Value, b Value) Value {
	if vm.jsTruthy(a) {
		return a
	}
	return b
}

// jsLogicalNot implements JScript '!' operator.
func (vm *VM) jsLogicalNot(v Value) Value {
	return NewBool(!vm.jsTruthy(v))
}

// jsCoalesce implements JScript '??' operator.
func (vm *VM) jsCoalesce(a Value, b Value) Value {
	if a.Type == VTNull || a.Type == VTJSUndefined {
		return b
	}
	return a
}

// jsMemberIndexGet implements member[index] access.
func (vm *VM) jsMemberIndexGet(obj Value, index Value, memberName string) Value {
	container, deferred := vm.jsMemberGet(obj, memberName)
	if deferred {
		return Value{Type: VTJSUndefined}
	}
	return vm.jsIndexGet(container, index)
}

// jsMemberIndexSet implements member[index] = value assignment.
func (vm *VM) jsMemberIndexSet(obj Value, index Value, value Value, memberName string) {
	container, deferred := vm.jsMemberGet(obj, memberName)
	if deferred {
		return
	}
	vm.jsIndexSet(container, index, value)
}

func (vm *VM) jsUpdateMember(target Value, member string, delta float64, post bool) Value {
	current, deferred := vm.jsMemberGet(target, member)
	if deferred {
		return Value{Type: VTJSUndefined}
	}
	next := vm.jsToNumber(current)
	next.Flt += delta
	vm.jsMemberSet(target, member, next)
	if post {
		return current
	}
	return next
}

func (vm *VM) jsUpdateIndex(target Value, index Value, delta float64, post bool) Value {
	current := vm.jsIndexGet(target, index)
	next := vm.jsToNumber(current)
	next.Flt += delta
	vm.jsIndexSet(target, index, next)
	if post {
		return current
	}
	return next
}

func (vm *VM) jsNew(constructor Value, args []Value) Value {
	return vm.jsConstruct(constructor, args, constructor, false)
}

func (vm *VM) jsCreateProxy(args []Value) Value {
	if len(args) < 2 {
		vm.jsThrowTypeError("Proxy requires 2 arguments: target and handler")
		return Value{Type: VTJSUndefined}
	}
	target := args[0]
	handler := args[1]
	isObject := func(v Value) bool {
		switch v.Type {
		case VTJSObject, VTJSFunction, VTJSPromise, VTJSGenerator, VTJSProxy, VTNativeObject, VTArray, VTObject:
			return true
		}
		return false
	}
	if !isObject(target) || !isObject(handler) {
		vm.jsThrowJSError(jscript.ProxyTargetOrHandlerNotObject)
		return Value{Type: VTJSUndefined}
	}
	proxyID := vm.allocJSID()
	vm.jsProxyItems[proxyID] = &jsProxyObject{
		Target:  target,
		Handler: handler,
	}
	return Value{Type: VTJSProxy, Num: proxyID}
}

func (vm *VM) jsConstruct(constructor Value, args []Value, newTarget Value, isSuper bool) Value {
	if constructor.Type == VTJSProxy {
		if !vm.jsIsConstructor(constructor) {
			vm.jsThrowTypeError("Proxy target is not a constructor")
			return Value{Type: VTJSUndefined}
		}
		return vm.jsProxyConstruct(constructor, args, newTarget)
	}
	if constructor.Type == VTJSFunction {
		closure := vm.jsFunctionItems[constructor.Num]
		if closure != nil {
			if closure.isDerived {
				// Derived class constructor: 'this' is uninitialized until super() is called.
				if vm.jsBeginFunctionCall(constructor, Value{Type: VTJSUninitialized}, args, Value{Type: VTJSUninitialized}, true, newTarget, isSuper) {
					return Value{Type: VTJSUndefined}
				}
				return Value{Type: VTJSUninitialized}
			}

			instanceID := vm.allocJSID()
			vm.jsObjectItems[instanceID] = make(map[string]Value, 8)
			vm.jsPropertyItems[instanceID] = make(map[string]jsPropertyDescriptor, 8)
			instance := Value{Type: VTJSObject, Num: instanceID}
			// Use prototype from newTarget (which might be the subclass if this is a base constructor call via super())
			proto, deferred := vm.jsMemberGet(newTarget, "prototype")
			if !deferred && proto.Type == VTJSObject {
				vm.jsObjectItems[instanceID]["__js_proto"] = proto
			} else {
				fallback := vm.jsGetIntrinsicPrototype("Object")
				if fallback.Type == VTJSObject {
					vm.jsObjectItems[instanceID]["__js_proto"] = fallback
				}
			}
			if vm.jsBeginFunctionCall(constructor, instance, args, instance, true, newTarget, isSuper) {
				return Value{Type: VTJSUndefined}
			}
			return instance
		}
	}

	if constructor.Type == VTJSObject || constructor.Type == VTJSFunction {
		ctorName := vm.jsObjectStringProperty(constructor, "__js_ctor")
		switch ctorName {
		case "Array":
			if len(args) == 0 {
				return ValueFromVBArray(NewVBArrayFromValues(0, nil))
			}
			if len(args) == 1 && (args[0].Type == VTInteger || args[0].Type == VTDouble) {
				length := max(int(vm.jsToNumber(args[0]).Flt), 0)
				return ValueFromVBArray(NewVBArrayFromValues(0, make([]Value, length)))
			}
			vals := make([]Value, len(args))
			copy(vals, args)
			return ValueFromVBArray(NewVBArrayFromValues(0, vals))
		case "Symbol":
			vm.jsThrowTypeError("Symbol is not a constructor")
			return Value{Type: VTJSUndefined}
		case "Error", "TypeError", "ReferenceError", "SyntaxError", "RangeError", "EvalError", "URIError":
			msg := ""
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				msg = vm.jsToString(args[0])
			}
			return vm.jsCreateErrorObject(ctorName, msg)
		case "Proxy":
			return vm.jsCreateProxy(args)
		case "Date":
			if len(args) == 0 {
				return NewDate(time.Now().In(builtinCurrentLocation(vm)))
			}
			if len(args) == 1 {
				if args[0].Type == VTString {
					parsed := valueToTimeInLocale(vm, args[0])
					if parsed.IsZero() {
						return NewDate(time.Time{})
					}
					return NewDate(parsed)
				}
				millis := int64(vm.jsToNumber(args[0]).Flt)
				return NewDate(time.Unix(0, millis*int64(time.Millisecond)).In(builtinCurrentLocation(vm)))
			}
			year := int(vm.jsToNumber(args[0]).Flt)
			month := 0
			day := 1
			hour := 0
			minute := 0
			second := 0
			millisecond := 0
			if len(args) > 1 {
				month = int(vm.jsToNumber(args[1]).Flt)
			}
			if len(args) > 2 {
				day = int(vm.jsToNumber(args[2]).Flt)
			}
			if len(args) > 3 {
				hour = int(vm.jsToNumber(args[3]).Flt)
			}
			if len(args) > 4 {
				minute = int(vm.jsToNumber(args[4]).Flt)
			}
			if len(args) > 5 {
				second = int(vm.jsToNumber(args[5]).Flt)
			}
			if len(args) > 6 {
				millisecond = int(vm.jsToNumber(args[6]).Flt)
			}
			loc := builtinCurrentLocation(vm)
			t := time.Date(year, time.Month(month+1), day, hour, minute, second, millisecond*int(time.Millisecond), loc)
			return NewDate(t)
		case "URL":
			return vm.jsConstructURL(args)
		case "URLSearchParams":
			return vm.jsConstructURLSearchParams(args)
		case "IntlDateTimeFormat":
			return vm.jsIntlCreateDateTimeFormat(args)
		case "IntlNumberFormat":
			return vm.jsIntlCreateNumberFormat(args)
		case "IntlCollator":
			return vm.jsIntlCreateCollator(args)
		case "IntlPluralRules":
			return vm.jsIntlCreatePluralRules(args)
		case "IntlRelativeTimeFormat":
			return vm.jsIntlCreateRelativeTimeFormat(args)
		case "RegExp":
			pattern := ""
			flags := ""
			if len(args) > 0 {
				if args[0].Type == VTJSObject && vm.jsObjectStringProperty(args[0], "__js_type") == "RegExp" {
					pattern = vm.jsObjectStringProperty(args[0], "pattern")
					if len(args) > 1 && args[1].Type != VTJSUndefined {
						flags = vm.valueToString(args[1])
					} else {
						flags = vm.jsObjectStringProperty(args[0], "flags")
					}
				} else {
					pattern = vm.valueToString(args[0])
					if len(args) > 1 {
						flags = vm.valueToString(args[1])
					}
				}
			}
			objID := vm.allocJSID()
			obj := make(map[string]Value, 8)
			obj["__js_type"] = NewString("RegExp")
			obj["pattern"] = NewString(pattern)
			obj["flags"] = NewString(flags)
			obj["lastIndex"] = NewInteger(0)
			if proto := vm.jsGetIntrinsicPrototype("RegExp"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
			return Value{Type: VTJSObject, Num: objID}
		case "Enumerator":
			source := Value{Type: VTJSUndefined}
			if len(args) > 0 {
				source = args[0]
			}
			objID := vm.allocJSID()
			obj := make(map[string]Value, 4)
			obj["__js_type"] = NewString("Enumerator")
			obj["__js_enum_source"] = source
			obj["__js_enum_index"] = NewInteger(0)
			if source.Type == VTJSObject {
				keys := vm.jsEnumerateForInKeys(source)
				keyVals := make([]Value, len(keys))
				for i := range keys {
					keyVals[i] = NewString(keys[i])
				}
				obj["__js_enum_keys"] = ValueFromVBArray(NewVBArrayFromValues(0, keyVals))
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
			return Value{Type: VTJSObject, Num: objID}
		case "VBArray":
			source := Value{Type: VTJSUndefined}
			if len(args) > 0 {
				source = args[0]
			}
			objID := vm.allocJSID()
			obj := make(map[string]Value, 2)
			obj["__js_type"] = NewString("VBArray")
			obj["__js_vbarray_source"] = source
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 3)
			if proto := vm.jsGetIntrinsicPrototype("Array"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			return Value{Type: VTJSObject, Num: objID}
		case "Set":
			objID := vm.allocJSID()
			obj := make(map[string]Value, 2)
			obj["__js_type"] = NewString("Set")
			obj["__js_ctor"] = NewString("Set")
			if proto := vm.jsGetIntrinsicPrototype("Set"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
			vm.jsSetItems[objID] = make(map[string]Value, 8)
			setObj := Value{Type: VTJSObject, Num: objID}
			if len(args) > 0 && args[0].Type != VTJSUndefined && args[0].Type != VTNull {
				if !vm.jsInitSetFromIterable(args[0], vm.jsSetItems[objID]) {
					return Value{Type: VTJSUndefined}
				}
			}
			return setObj
		case "Map":
			objID := vm.allocJSID()
			obj := make(map[string]Value, 2)
			obj["__js_type"] = NewString("Map")
			obj["__js_ctor"] = NewString("Map")
			if proto := vm.jsGetIntrinsicPrototype("Map"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
			vm.jsMapItems[objID] = make(map[string]Value, 8)
			mapObj := Value{Type: VTJSObject, Num: objID}
			if len(args) > 0 && args[0].Type != VTJSUndefined && args[0].Type != VTNull {
				if !vm.jsInitMapFromIterable(args[0], vm.jsMapItems[objID]) {
					return Value{Type: VTJSUndefined}
				}
			}
			return mapObj
		case "WeakMap":
			objID := vm.allocJSID()
			obj := make(map[string]Value, 4)
			obj["__js_type"] = NewString("WeakMap")
			obj["__js_ctor"] = NewString("WeakMap")
			weakID := jsWeakCollectionNextID.Add(1)
			obj["__js_weak_id"] = NewDouble(float64(weakID))
			if proto := vm.jsGetIntrinsicPrototype("WeakMap"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
			weakMapObj := Value{Type: VTJSObject, Num: objID}
			if len(args) > 0 && args[0].Type != VTJSUndefined && args[0].Type != VTNull {
				if !vm.jsInitWeakMapFromIterable(args[0], weakID) {
					return Value{Type: VTJSUndefined}
				}
			}
			return weakMapObj
		case "WeakSet":
			objID := vm.allocJSID()
			obj := make(map[string]Value, 4)
			obj["__js_type"] = NewString("WeakSet")
			obj["__js_ctor"] = NewString("WeakSet")
			weakID := jsWeakCollectionNextID.Add(1)
			obj["__js_weak_id"] = NewDouble(float64(weakID))
			if proto := vm.jsGetIntrinsicPrototype("WeakSet"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
			weakSetObj := Value{Type: VTJSObject, Num: objID}
			if len(args) > 0 && args[0].Type != VTJSUndefined && args[0].Type != VTNull {
				if !vm.jsInitWeakSetFromIterable(args[0], weakID) {
					return Value{Type: VTJSUndefined}
				}
			}
			return weakSetObj
		case "WeakRef":
			if len(args) == 0 || (args[0].Type != VTJSObject && args[0].Type != VTJSFunction && args[0].Type != VTSymbol) {
				vm.jsThrowTypeError("WeakRef target must be an object")
				return Value{Type: VTJSUndefined}
			}
			objID := vm.allocJSID()
			obj := make(map[string]Value, 2)
			obj["__js_type"] = NewString("WeakRef")
			obj["__js_ctor"] = NewString("WeakRef")
			if proto := vm.jsGetIntrinsicPrototype("WeakRef"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 2)
			vm.jsWeakRefItems[objID] = &jsWeakRef{targetID: args[0].Num}
			return Value{Type: VTJSObject, Num: objID}
		case "FinalizationRegistry":
			if len(args) == 0 || args[0].Type != VTJSFunction {
				vm.jsThrowTypeError("FinalizationRegistry callback must be a function")
				return Value{Type: VTJSUndefined}
			}
			objID := vm.allocJSID()
			obj := make(map[string]Value, 2)
			obj["__js_type"] = NewString("FinalizationRegistry")
			obj["__js_ctor"] = NewString("FinalizationRegistry")
			if proto := vm.jsGetIntrinsicPrototype("FinalizationRegistry"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 2)
			vm.jsFinalizationRegistryItems[objID] = &jsFinalizationRegistry{
				callback: args[0],
			}
			return Value{Type: VTJSObject, Num: objID}
		case "Promise":
			return vm.jsNewPromise(args)
		case "ArrayBuffer", "SharedArrayBuffer":
			byteLength := 0
			if len(args) > 0 {
				byteLength = int(vm.jsToNumber(args[0]).Flt)
			}
			if ctorName == "SharedArrayBuffer" {
				return vm.jsNewSharedArrayBuffer(byteLength)
			}
			return vm.jsNewArrayBuffer(byteLength)
		case "DataView":
			return vm.jsNewDataView(args)
		case "Int8Array", "Uint8Array", "Uint8ClampedArray",
			"Int16Array", "Uint16Array",
			"Int32Array", "Uint32Array",
			"Float32Array", "Float64Array",
			"BigInt64Array", "BigUint64Array":
			return vm.jsNewTypedArray(ctorName, args)
		}
	}

	objID := vm.allocJSID()
	vm.jsObjectItems[objID] = make(map[string]Value, 8)
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsMemberDelete implements the 'delete' operator for object properties.
func (vm *VM) jsMemberDelete(obj Value, member string) bool {
	if obj.Type == VTJSObject {
		desc, hasDesc := vm.jsGetDescriptor(obj.Num, member)
		if hasDesc && !desc.Configurable {
			return false
		}
		if jsObj, ok := vm.jsObjectItems[obj.Num]; ok {
			delete(jsObj, member)
			if props, ok := vm.jsPropertyItems[obj.Num]; ok {
				delete(props, member)
			}
			vm.jsUntrackObjectKey(obj.Num, member)
			return true
		}
	}
	return false
}

// jsIndexGet implements array[index] access.
func (vm *VM) jsIndexGet(arr Value, index Value) Value {
	if index.Type == VTSymbol {
		key := vm.jsPropertyKeyFromValue(index)
		val, _ := vm.jsMemberGet(arr, key)
		return val
	}
	switch arr.Type {
	case VTJSProxy:
		key := vm.jsPropertyKeyFromValue(index)
		val, _ := vm.jsMemberGet(arr, key)
		return val
	case VTArray:
		if arr.Arr == nil {
			return Value{Type: VTJSUndefined}
		}
		indexNum := int(vm.jsToNumber(index).Flt)
		adjustedIndex := indexNum - arr.Arr.Lower
		if adjustedIndex < 0 || adjustedIndex >= len(arr.Arr.Values) {
			// Fallback to property/prototype lookup (e.g. Symbol.iterator)
			key := vm.jsPropertyKeyFromValue(index)
			val, _ := vm.jsMemberGet(arr, key)
			return val
		}
		return arr.Arr.Values[adjustedIndex]
	case VTJSObject:
		// First try typed array index access.
		idxInt := int(vm.jsToNumber(index).Flt)
		if v, handled := vm.jsTypedArrayIndexGet(arr, idxInt); handled {
			return v
		}
		// General Array slots
		if vm.jsObjectStringProperty(arr, "__js_type") == "Array" {
			if slots, ok := vm.jsObjectSlots[arr.Num]; ok {
				if idxInt >= 0 && idxInt < len(slots) {
					return slots[idxInt]
				}
			}
		}
		key := vm.jsPropertyKeyFromValue(index)
		if v, _ := vm.jsMemberGet(arr, key); v.Type != VTJSUndefined {
			return v
		}
		return Value{Type: VTJSUndefined}
	case VTString:
		runes := []rune(arr.Str)
		idx := int(vm.jsToNumber(index).Flt)
		if idx < 0 || idx >= len(runes) {
			// Fallback to property/prototype lookup
			key := vm.jsPropertyKeyFromValue(index)
			val, _ := vm.jsMemberGet(arr, key)
			return val
		}
		ch := string(runes[idx])
		if !vm.jsEnsureStringSize(len(ch)) || !vm.jsChargeStringWork(len(ch)) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(ch)
	default:
		return Value{Type: VTJSUndefined}
	}
}

// jsIndexSet implements array[index] = value assignment.
func (vm *VM) jsIndexSet(arr Value, index Value, value Value) {
	switch arr.Type {
	case VTJSProxy:
		key := vm.jsPropertyKeyFromValue(index)
		vm.jsMemberSet(arr, key, value)
	case VTArray:
		if arr.Arr == nil {
			return
		}
		if index.Type == VTSymbol {
			key := vm.jsPropertyKeyFromValue(index)
			vm.jsMemberSet(arr, key, value)
			return
		}
		indexNum, isIdx := jsParseArrayIndex(vm.valueToString(index))
		if isIdx {
			adjustedIndex := indexNum - arr.Arr.Lower
			if adjustedIndex >= 0 && adjustedIndex < len(arr.Arr.Values) {
				arr.Arr.Values[adjustedIndex] = value
				return
			}
		}
		// Fallback to member set for non-numeric keys or out-of-bounds indices (JScript arrays are objects)
		key := vm.jsPropertyKeyFromValue(index)
		vm.jsMemberSet(arr, key, value)
	case VTJSObject:
		// Try typed array index set first.
		idxInt := int(vm.jsToNumber(index).Flt)
		if vm.jsTypedArrayIndexSet(arr, idxInt, value) {
			return
		}
		key := vm.jsPropertyKeyFromValue(index)
		vm.jsMemberSet(arr, key, value)
	}
}

var jsRegExpUnicodeEscapeRegex = regexp.MustCompile(`\\u\{([0-9a-fA-F]+)\}`)

func jsNormalizeUnicodePropertyName(name string) string {
	name = strings.ToLower(name)
	name = strings.ReplaceAll(name, "_", "")
	name = strings.ReplaceAll(name, "-", "")
	name = strings.ReplaceAll(name, " ", "")
	return name
}

func jsCanonicalUnicodeProperty(name string) (string, bool) {
	switch jsNormalizeUnicodePropertyName(name) {
	case "l", "letter", "alphabetic":
		return "L", true
	case "lu", "uppercase":
		return "Lu", true
	case "ll", "lowercase":
		return "Ll", true
	case "m", "mark":
		return "M", true
	case "n", "number", "numeric":
		return "N", true
	case "p", "punctuation":
		return "P", true
	case "s", "symbol":
		return "S", true
	case "z", "separator", "space":
		return "Z", true
	case "zs":
		return "Zs", true
	case "c", "other":
		return "C", true
	default:
		return "", false
	}
}

func jsTranslateUnicodePropertyEscapes(pattern string) string {
	var builder strings.Builder
	changed := false

	for i := 0; i < len(pattern); {
		if pattern[i] != '\\' || i+3 >= len(pattern) {
			if changed {
				builder.WriteByte(pattern[i])
			}
			i++
			continue
		}

		kind := pattern[i+1]
		if (kind != 'p' && kind != 'P') || pattern[i+2] != '{' {
			if changed {
				builder.WriteByte(pattern[i])
			}
			i++
			continue
		}

		backslashCount := 0
		for j := i - 1; j >= 0 && pattern[j] == '\\'; j-- {
			backslashCount++
		}
		if backslashCount%2 == 1 {
			if changed {
				builder.WriteByte(pattern[i])
			}
			i++
			continue
		}

		closeRel := strings.IndexByte(pattern[i+3:], '}')
		if closeRel < 0 {
			if changed {
				builder.WriteByte(pattern[i])
			}
			i++
			continue
		}

		close := i + 3 + closeRel
		propertyName := pattern[i+3 : close]
		canonical, ok := jsCanonicalUnicodeProperty(propertyName)
		if !ok {
			if changed {
				builder.WriteByte(pattern[i])
			}
			i++
			continue
		}

		if !changed {
			builder.Grow(len(pattern) + 16)
			builder.WriteString(pattern[:i])
			changed = true
		}

		if kind == 'P' {
			builder.WriteString(`\P{`)
		} else {
			builder.WriteString(`\p{`)
		}
		builder.WriteString(canonical)
		builder.WriteByte('}')
		i = close + 1
	}

	if !changed {
		return pattern
	}
	return builder.String()
}

func (vm *VM) jsCompileRegExp(pattern string, flags string) (*regexp2.Regexp, error) {
	var options regexp2.RegexOptions

	flagsLower := strings.ToLower(flags)
	if strings.Contains(flagsLower, "i") {
		options |= regexp2.IgnoreCase
	}
	if strings.Contains(flagsLower, "m") {
		options |= regexp2.Multiline
	}
	if strings.Contains(flagsLower, "s") {
		options |= regexp2.Singleline
	}
	if strings.Contains(flagsLower, "u") {
		options |= regexp2.Unicode
		// Translate JS \u{...} to regexp2 \x{...}
		pattern = jsRegExpUnicodeEscapeRegex.ReplaceAllString(pattern, `\x{$1}`)
		// Normalize JS Unicode property aliases to regex category shorthands.
		pattern = jsTranslateUnicodePropertyEscapes(pattern)
	}

	return regexp2.Compile(pattern, options)
}

func (vm *VM) jsGetCompiledRegExp(objID int64) (*regexp2.Regexp, error) {
	if reObj, ok := vm.jsRegExpItems[objID]; ok && reObj.compiled != nil {
		return reObj.compiled, nil
	}

	pattern := vm.jsObjectStringProperty(Value{Type: VTJSObject, Num: objID}, "pattern")
	flags := vm.jsObjectStringProperty(Value{Type: VTJSObject, Num: objID}, "flags")

	re, err := vm.jsCompileRegExp(pattern, flags)
	if err != nil {
		return nil, err
	}

	vm.jsRegExpItems[objID] = &jsRegExpObject{
		pattern:  pattern,
		flags:    flags,
		compiled: re,
	}
	return re, nil
}

func (vm *VM) jsRegExpMatchAll(reVal Value, input string) Value {
	flags := vm.jsRegExpGetFlags(reVal)
	isGlobal := strings.Contains(flags, "g")
	isUnicode := strings.Contains(flags, "u")
	return vm.jsCreateRegExpStringIterator(reVal, input, isGlobal, isUnicode)
}

func (vm *VM) jsRegExpExec(reVal Value, input string) Value {
	objID := reVal.Num
	re, err := vm.jsGetCompiledRegExp(objID)
	if err != nil {
		return NewNull()
	}

	flags := vm.jsObjectStringProperty(reVal, "flags")
	isGlobal := strings.Contains(flags, "g")
	isSticky := strings.Contains(flags, "y")

	lastIndexVal, _ := vm.jsMemberGet(reVal, "lastIndex")
	lastIndex := max(int(vm.jsToNumber(lastIndexVal).Flt), 0)

	if !isGlobal && !isSticky {
		lastIndex = 0
	}

	// JScript lastIndex is in UTF-16 code units.
	// We need to convert it to byte offset for regexp2.
	runeOffset := 0
	if lastIndex > 0 {
		utf16Count := 0
		for utf16Count < lastIndex && runeOffset < len(input) {
			r, size := utf8.DecodeRuneInString(input[runeOffset:])
			if r >= 0x10000 {
				utf16Count += 2
			} else {
				utf16Count += 1
			}
			runeOffset += size
		}
		if utf16Count < lastIndex {
			// lastIndex is beyond string length
			vm.jsMemberSet(reVal, "lastIndex", NewInteger(0))
			return NewNull()
		}
	}

	// regexp2.FindStringMatchStartingAt takes RUNE index, not byte offset.
	// But we need to convert runeOffset (byte offset) to rune count.
	runeCount := 0
	for i := 0; i < runeOffset; {
		_, size := utf8.DecodeRuneInString(input[i:])
		runeCount++
		i += size
	}

	m, err := re.FindStringMatchStartingAt(input, runeCount)
	if err != nil || m == nil {
		if isGlobal || isSticky {
			vm.jsMemberSet(reVal, "lastIndex", NewInteger(0))
		}
		return NewNull()
	}

	// If sticky, match MUST start exactly at lastIndex (runeCount)
	if isSticky && m.Index != runeCount {
		vm.jsMemberSet(reVal, "lastIndex", NewInteger(0))
		return NewNull()
	}

	// Calculate new lastIndex (UTF-16)
	matchEndRuneIndex := m.Index + m.Length
	newLastIndex := int64(0)
	currentRuneIdx := 0
	for _, r := range input {
		if currentRuneIdx >= matchEndRuneIndex {
			break
		}
		if r >= 0x10000 {
			newLastIndex += 2
		} else {
			newLastIndex += 1
		}
		currentRuneIdx++
	}

	if isGlobal || isSticky {
		vm.jsMemberSet(reVal, "lastIndex", NewInteger(newLastIndex))
	}

	// Create results array (JScript style)
	resID := vm.allocJSID()
	res := make(map[string]Value)
	res["__js_type"] = NewString("Array")
	// JScript index is the start of the match in UTF-16.
	matchStartUTF16 := int64(0)
	currentRuneIdx = 0
	for _, r := range input {
		if currentRuneIdx >= m.Index {
			break
		}
		if r >= 0x10000 {
			matchStartUTF16 += 2
		} else {
			matchStartUTF16 += 1
		}
		currentRuneIdx++
	}
	res["index"] = NewInteger(matchStartUTF16)
	res["input"] = NewString(input)

	groups := m.Groups()
	vals := make([]Value, len(groups))
	for i, g := range groups {
		if g.Capture.Length < 0 {
			vals[i] = Value{Type: VTJSUndefined}
		} else {
			vals[i] = NewString(g.String())
		}
	}
	vm.jsObjectItems[resID] = res
	vm.jsObjectSlots[resID] = vals
	res["length"] = NewInteger(int64(len(vals)))
	for i, v := range vals {
		res[strconv.Itoa(i)] = v
	}

	// Named capture groups (ES2018)
	var groupsObj map[string]Value
	var groupsObjID int64
	for i, g := range groups {
		if g.Name != "" && g.Name != strconv.Itoa(i) {
			if groupsObj == nil {
				groupsObjID = vm.allocJSID()
				groupsObj = make(map[string]Value)
				res["groups"] = Value{Type: VTJSObject, Num: groupsObjID}
			}
			if g.Capture.Length < 0 {
				groupsObj[g.Name] = Value{Type: VTJSUndefined}
			} else {
				groupsObj[g.Name] = NewString(g.String())
			}
		}
	}
	if groupsObj != nil {
		vm.jsObjectItems[groupsObjID] = groupsObj
	} else {
		res["groups"] = Value{Type: VTJSUndefined}
	}

	return Value{Type: VTJSObject, Num: resID}
}

func (vm *VM) jsRuneToByteOffset(s string, runeIdx int) int {
	if runeIdx <= 0 {
		return 0
	}
	byteIdx := 0
	for range runeIdx {
		if byteIdx >= len(s) {
			break
		}
		_, size := utf8.DecodeRuneInString(s[byteIdx:])
		byteIdx += size
	}
	return byteIdx
}

func (vm *VM) jsRegExpGetFlags(re Value) string {
	flags := vm.jsObjectStringProperty(re, "flags")
	res := ""
	if strings.Contains(flags, "g") {
		res += "g"
	}
	if strings.Contains(flags, "i") {
		res += "i"
	}
	if strings.Contains(flags, "m") {
		res += "m"
	}
	if strings.Contains(flags, "s") {
		res += "s"
	}
	if strings.Contains(flags, "u") {
		res += "u"
	}
	if strings.Contains(flags, "y") {
		res += "y"
	}
	return res
}

func jsStringLength(s string) int64 {
	length := int64(0)
	for _, r := range s {
		if r >= 0x10000 {
			length += 2
		} else {
			length += 1
		}
	}
	return length
}

func (vm *VM) jsSuperGet(member string) Value {
	if len(vm.jsCallStack) == 0 {
		return Value{Type: VTJSUndefined}
	}
	currentFn := vm.jsCallStack[len(vm.jsCallStack)-1].fn
	if currentFn.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}
	}
	closure := vm.jsFunctionItems[currentFn.Num]
	if closure == nil || closure.homeObjID == 0 {
		// Not inside a class method
		vm.jsThrowTypeError("super member access outside of class method")
		return Value{Type: VTJSUndefined}
	}

	homeObj := Value{Type: VTJSObject, Num: closure.homeObjID}
	superProto := vm.jsGetPrototypeValue(homeObj)
	if superProto.Type == VTJSUndefined {
		return Value{Type: VTJSUndefined}
	}

	val, deferred := vm.jsMemberGet(superProto, member)
	if deferred {
		// Getter call was initiated
		return Value{Type: VTJSUndefined}
	}
	return val
}

func (vm *VM) jsSuperSet(member string, val Value) {
	if len(vm.jsCallStack) == 0 {
		return
	}
	currentFn := vm.jsCallStack[len(vm.jsCallStack)-1].fn
	if currentFn.Type != VTJSFunction {
		return
	}
	closure := vm.jsFunctionItems[currentFn.Num]
	if closure == nil || closure.homeObjID == 0 {
		vm.jsThrowTypeError("super member assignment outside of class method")
		return
	}

	homeObj := Value{Type: VTJSObject, Num: closure.homeObjID}
	superProto := vm.jsGetPrototypeValue(homeObj)
	if superProto.Type == VTJSUndefined {
		return
	}

	// Super assignment receiver is 'this'
	vm.jsMemberSet(vm.jsThisValue, member, val)
}

func (vm *VM) jsSuperCallMember(member string, args []Value) Value {
	if len(vm.jsCallStack) == 0 {
		return Value{Type: VTJSUndefined}
	}
	currentFn := vm.jsCallStack[len(vm.jsCallStack)-1].fn
	if currentFn.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}
	}
	closure := vm.jsFunctionItems[currentFn.Num]
	if closure == nil || closure.homeObjID == 0 {
		vm.jsThrowTypeError("super member call outside of class method")
		return Value{Type: VTJSUndefined}
	}

	homeObj := Value{Type: VTJSObject, Num: closure.homeObjID}
	superProto := vm.jsGetPrototypeValue(homeObj)
	if superProto.Type == VTJSUndefined {
		vm.jsThrowTypeError("Super prototype is undefined")
		return Value{Type: VTJSUndefined}
	}

	// Resolve the method from superProto
	method, deferred := vm.jsMemberGet(superProto, member)
	if deferred {
		return Value{Type: VTJSUndefined}
	}

	if method.Type != VTJSFunction && method.Type != VTNativeObject && method.Type != VTBuiltin {
		vm.jsThrowTypeError(fmt.Sprintf("super.%s is not a function", member))
		return Value{Type: VTJSUndefined}
	}

	// Call the method with 'this' set to current instance
	return vm.jsCall(method, vm.jsThisValue, args)
}

func (vm *VM) jsSuperIndexGet(index Value) Value {
	return vm.jsSuperGet(vm.jsPropertyKeyFromValue(index))
}

func (vm *VM) jsSuperIndexSet(index Value, val Value) {
	vm.jsSuperSet(vm.jsPropertyKeyFromValue(index), val)
}

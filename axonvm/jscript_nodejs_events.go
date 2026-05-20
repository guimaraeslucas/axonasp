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
	_ "embed"
)

//go:embed node_js/events.js
var jsEventsPolyfillSource string

// jsEventsModuleKey is the synthetic key used to cache the events module in jsModuleInstances.
const jsEventsModuleKey = "__builtin__:events"

// jsGetOrCreateEventsModule returns the cached events module object (containing the
// EventEmitter constructor), running the polyfill on first call.
// Returns undefined if Node.js compatibility is disabled.
func (vm *VM) jsGetOrCreateEventsModule() Value {
	// Check if an already-cached module env exists
	if env, ok := vm.jsModuleInstances[jsEventsModuleKey]; ok && env != nil {
		// Return the cached constructor directly
		if ctorVal, ok := env.bindings["EventEmitter"]; ok {
			return ctorVal
		}
	}

	// Run the polyfill to obtain the EventEmitter constructor
	ctor := vm.jsRunNodeEventsPolyfill()
	if ctor.Type == VTJSUndefined {
		return Value{Type: VTJSUndefined}
	}

	// In Node.js, require('events') returns the EventEmitter constructor itself,
	// and it also has a property named 'EventEmitter' pointing back to itself.
	if ctor.Type == VTJSFunction {
		vm.jsMemberSet(ctor, "EventEmitter", ctor)
	}

	// Cache the constructor in a dedicated env frame so repeated require("events")
	// returns the same constructor reference without re-running the polyfill.
	moduleEnv := &jsEnvFrame{
		parentID: vm.jsRootEnvID,
		bindings: map[string]Value{
			"EventEmitter": ctor,
		},
	}
	vm.jsModuleInstances[jsEventsModuleKey] = moduleEnv

	// Also expose EventEmitter in the root environment so it is accessible as a
	// global when Node.js compatibility is active (matches common usage patterns).
	vm.ensureJSRootEnv()
	if root := vm.jsEnvItems[vm.jsRootEnvID]; root != nil {
		root.bindings["EventEmitter"] = ctor
	}

	return ctor
}

// jsBuildEventsModuleObject wraps the EventEmitter constructor in a plain object that
// mirrors the Node.js module export shape: { EventEmitter: <ctor> }.
func (vm *VM) jsBuildEventsModuleObject(ctor Value) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 2)
	obj["EventEmitter"] = ctor
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 2)
	return Value{Type: VTJSObject, Num: objID}
}

// jsRunNodeEventsPolyfill compiles and executes the embedded events.js polyfill
// as a normal JScript block, then reads the EventEmitter binding from the root
// environment.
func (vm *VM) jsRunNodeEventsPolyfill() Value {
	source := jsEventsPolyfillSource
	if source == "" {
		return Value{Type: VTJSUndefined}
	}

	compiler := NewASPCompiler("")
	compiler.sourceName = "__builtin__:events"
	compiler.compileJScriptBlock(source)
	compiler.emit(OpHalt)

	if len(compiler.bytecode) == 0 {
		return Value{Type: VTJSUndefined}
	}

	startIP := vm.appendExecuteProgram(compiler.GlobalsCount(), compiler.constants, compiler.bytecode)
	if startIP < 0 || startIP >= len(vm.bytecode) {
		return Value{Type: VTJSUndefined}
	}

	child := vm.cloneForExecuteLocal(startIP)
	child.sourceName = "__builtin__:events"
	if err := child.Run(); err != nil {
		vm.syncExecuteGlobalState(child)
		return Value{Type: VTJSUndefined}
	}

	vm.syncExecuteGlobalState(child)
	return vm.jsNodeGetRootBinding("EventEmitter")
}

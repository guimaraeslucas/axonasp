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
	"fmt"
)

type jsGeneratorState int

const (
	jsGeneratorSuspendedStart jsGeneratorState = iota
	jsGeneratorSuspendedYield
	jsGeneratorExecuting
	jsGeneratorCompleted
)

type jsYieldError struct {
	value    Value
	delegate bool
}

func (e *jsYieldError) Error() string { return "js yield" }

type jsGeneratorObject struct {
	state   jsGeneratorState
	fn      Value
	thisVal Value
	args    []Value
	childVM *VM
	stack   []Value
	sp      int
	fp      int
}

func (vm *VM) jsCreateGeneratorObject(fn Value, thisVal Value, args []Value) Value {
	genID := vm.allocJSID()
	gen := &jsGeneratorObject{
		state:   jsGeneratorSuspendedStart,
		fn:      fn,
		thisVal: thisVal,
		args:    args,
	}
	vm.jsGeneratorItems[genID] = gen

	objID := vm.allocJSID()
	vm.jsObjectItems[objID] = make(map[string]Value)
	vm.jsObjectItems[objID]["__js_generator_id"] = NewInteger(genID)
	vm.jsObjectItems[objID]["__js_type"] = NewString("Generator")

	vm.jsSetDescriptor(objID, "next", jsPropertyDescriptor{
		Value:    vm.jsCreateIntrinsicFunction("Generator.prototype.next", "GeneratorPrototypeNext"),
		HasValue: true,
	})
	vm.jsSetDescriptor(objID, "throw", jsPropertyDescriptor{
		Value:    vm.jsCreateIntrinsicFunction("Generator.prototype.throw", "GeneratorPrototypeThrow"),
		HasValue: true,
	})
	vm.jsSetDescriptor(objID, "return", jsPropertyDescriptor{
		Value:    vm.jsCreateIntrinsicFunction("Generator.prototype.return", "GeneratorPrototypeReturn"),
		HasValue: true,
	})

	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsHandleGeneratorNext(thisVal Value, args []Value) Value {
	genIDVal := vm.jsObjectItems[thisVal.Num]["__js_generator_id"]
	gen := vm.jsGeneratorItems[genIDVal.Num]

	if gen.state == jsGeneratorCompleted {
		return vm.jsCreateIteratorResult(Value{Type: VTJSUndefined}, true)
	}

	if gen.state == jsGeneratorExecuting {
		vm.jsThrowTypeError("Generator is already executing")
		return Value{Type: VTJSUndefined}
	}

	resumeVal := jsArgOrUndefined(args, 0)

	if gen.state == jsGeneratorSuspendedStart {
		gen.childVM = vm.cloneForExecuteLocal(len(vm.bytecode))
		gen.childVM.jsBeginFunctionCall(gen.fn, gen.thisVal, gen.args, Value{Type: VTJSUndefined}, false, Value{Type: VTJSUndefined}, false)
	} else {
		// Restore stack and state before resuming
		copy(gen.childVM.stack, gen.stack)
		gen.childVM.sp = gen.sp
		gen.childVM.fp = gen.fp
		// Sync ID counter from parent in case it moved
		gen.childVM.nextDynamicNativeID = vm.nextDynamicNativeID
		// Push the resume value
		gen.childVM.push(resumeVal)
	}

	gen.state = jsGeneratorExecuting
	err := func() (err error) {
		defer func() {
			if r := recover(); r != nil {
				if vme, ok := r.(*VMError); ok {
					err = vme
				} else if ye, ok := r.(*jsYieldError); ok {
					err = ye
				} else {
					err = fmt.Errorf("%v", r)
				}
			}
		}()
		return gen.childVM.Run()
	}()

	// Sync back ID counter and other state to parent
	vm.syncExecuteGlobalState(gen.childVM)

	if err != nil {
		if y, ok := err.(*jsYieldError); ok {
			gen.state = jsGeneratorSuspendedYield
			// Save stack and state
			if len(gen.stack) < len(gen.childVM.stack) {
				gen.stack = make([]Value, len(gen.childVM.stack))
			}
			copy(gen.stack, gen.childVM.stack)
			gen.sp = gen.childVM.sp
			gen.fp = gen.childVM.fp
			return vm.jsCreateIteratorResult(y.value, false)
		}
		gen.state = jsGeneratorCompleted
		panic(err)
	}

	gen.state = jsGeneratorCompleted
	retVal := Value{Type: VTJSUndefined}
	if gen.childVM.sp >= 0 {
		retVal = gen.childVM.stack[gen.childVM.sp]
	}
	return vm.jsCreateIteratorResult(retVal, true)
}

func (vm *VM) jsCreateIteratorResult(value Value, done bool) Value {
	id := vm.allocJSID()
	obj := make(map[string]Value, 2)
	obj["value"] = value
	obj["done"] = NewBool(done)
	vm.jsObjectItems[id] = obj
	return Value{Type: VTJSObject, Num: id}
}

func (vm *VM) jsHandleGeneratorThrow(thisVal Value, args []Value) Value {
	genIDVal := vm.jsObjectItems[thisVal.Num]["__js_generator_id"]
	gen := vm.jsGeneratorItems[genIDVal.Num]
	gen.state = jsGeneratorCompleted
	reason := jsArgOrUndefined(args, 0)
	panic(&VMError{Msg: vm.valueToString(reason)})
}

func (vm *VM) jsHandleGeneratorReturn(thisVal Value, args []Value) Value {
	genIDVal := vm.jsObjectItems[thisVal.Num]["__js_generator_id"]
	gen := vm.jsGeneratorItems[genIDVal.Num]
	gen.state = jsGeneratorCompleted
	retVal := jsArgOrUndefined(args, 0)
	return vm.jsCreateIteratorResult(retVal, true)
}

func (vm *VM) jsYield(val Value, delegate bool) {
	panic(&jsYieldError{value: val, delegate: delegate})
}

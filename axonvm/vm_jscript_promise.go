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
 */
package axonvm

import (
	"fmt"
)

func (vm *VM) jsNewPromise(args []Value) Value {
	if len(args) == 0 || args[0].Type != VTJSFunction {
		vm.jsThrowTypeError("Promise resolver undefined is not a function")
		return Value{Type: VTJSUndefined}
	}

	executor := args[0]
	promiseID := vm.allocJSID()
	promise := Value{Type: VTJSPromise, Num: promiseID}

	pObj := &jsPromiseObject{
		state: jsPromisePending,
	}
	vm.jsPromiseItems[promiseID] = pObj

	// 25.4.3.1.1 If NewPromiseCapability is used, we'd have resolve/reject.
	// Here we implement the Promise constructor directly.

	resolve := vm.jsCreatePromiseResolveFunction(promise)
	reject := vm.jsCreatePromiseRejectFunction(promise)

	// Call executor(resolve, reject)
	// We wrap in try/catch to reject on throw
	vm.jsTryCall(executor, Value{Type: VTJSUndefined}, []Value{resolve, reject}, func(reason Value) {
		vm.jsCall(reject, Value{Type: VTJSUndefined}, []Value{reason})
	})

	return promise
}

func (vm *VM) jsTryCall(callee Value, thisVal Value, args []Value, onCatch func(Value)) Value {
	// Simple try/call wrapper
	// In a real VM this would use the jsTryStack mechanism
	// but here we can just use a helper if we are not in the main loop.

	// If we are calling from Go, we need to handle potential panics or errors
	// that would normally be caught by the VM loop.

	defer func() {
		if r := recover(); r != nil {
			if vmErr, ok := r.(*VMError); ok {
				onCatch(NewString(vmErr.Error()))
			} else {
				onCatch(NewString(fmt.Sprintf("%v", r)))
			}
		}
	}()

	return vm.jsCall(callee, thisVal, args)
}

func (vm *VM) jsCreatePromiseResolveFunction(promise Value) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("PromiseResolve")
	obj["__js_promise"] = promise
	vm.jsObjectItems[objID] = obj
	return Value{Type: VTJSFunction, Num: objID}
}

func (vm *VM) jsCreatePromiseRejectFunction(promise Value) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("PromiseReject")
	obj["__js_promise"] = promise
	vm.jsObjectItems[objID] = obj
	return Value{Type: VTJSFunction, Num: objID}
}

func (vm *VM) jsResolvePromise(promise Value, resolution Value) {
	pObj := vm.jsPromiseItems[promise.Num]
	if pObj == nil || pObj.state != jsPromisePending {
		return
	}

	if resolution.Type == VTJSPromise && resolution.Num == promise.Num {
		vm.jsRejectPromise(promise, NewString("TypeError: Chaining cycle detected for promise"))
		return
	}

	if resolution.Type == VTJSObject || resolution.Type == VTJSFunction || resolution.Type == VTJSPromise {
		// Check for thenable
		then, deferred := vm.jsMemberGet(resolution, "then")
		if !deferred && then.Type == VTJSFunction {
			vm.jsEnqueueMicrotask(func() {
				resolve := vm.jsCreatePromiseResolveFunction(promise)
				reject := vm.jsCreatePromiseRejectFunction(promise)
				vm.jsTryCall(then, resolution, []Value{resolve, reject}, func(reason Value) {
					vm.jsCall(reject, Value{Type: VTJSUndefined}, []Value{reason})
				})
			})
			return
		}
	}

	pObj.state = jsPromiseFulfilled
	pObj.result = resolution
	vm.jsTriggerPromiseReactions(pObj)
}

func (vm *VM) jsRejectPromise(promise Value, reason Value) {
	pObj := vm.jsPromiseItems[promise.Num]
	if pObj == nil || pObj.state != jsPromisePending {
		return
	}

	pObj.state = jsPromiseRejected
	pObj.result = reason
	vm.jsTriggerPromiseReactions(pObj)
}

func (vm *VM) jsTriggerPromiseReactions(pObj *jsPromiseObject) {
	reactions := pObj.reactions
	pObj.reactions = nil
	state := pObj.state
	result := pObj.result

	for _, reaction := range reactions {
		reac := reaction // copy
		vm.jsEnqueueMicrotask(func() {
			handler := reac.onFulfilled
			if state == jsPromiseRejected {
				handler = reac.onRejected
			}

			if handler.Type != VTJSFunction {
				if state == jsPromiseFulfilled {
					vm.jsResolvePromise(reac.capability.promise, result)
				} else {
					vm.jsRejectPromise(reac.capability.promise, result)
				}
				return
			}

			// Call handler(result)
			var handlerResult Value
			caught := false

			func() {
				defer func() {
					if r := recover(); r != nil {
						caught = true
						if vmErr, ok := r.(*VMError); ok {
							vm.jsRejectPromise(reac.capability.promise, NewString(vmErr.Error()))
						} else {
							vm.jsRejectPromise(reac.capability.promise, NewString(fmt.Sprintf("%v", r)))
						}
					}
				}()
				handlerResult = vm.jsCall(handler, Value{Type: VTJSUndefined}, []Value{result})
			}()

			if !caught {
				vm.jsResolvePromise(reac.capability.promise, handlerResult)
			}
		})
	}
}

func (vm *VM) jsPromiseThen(promise Value, args []Value) Value {
	if promise.Type != VTJSPromise {
		vm.jsThrowTypeError("Method Promise.prototype.then called on incompatible receiver")
		return Value{Type: VTJSUndefined}
	}

	onFulfilled := jsArgOrUndefined(args, 0)
	onRejected := jsArgOrUndefined(args, 1)

	pObj := vm.jsPromiseItems[promise.Num]

	// Create the new promise for chaining
	resultPromiseID := vm.allocJSID()
	resultPromise := Value{Type: VTJSPromise, Num: resultPromiseID}
	resultPObj := &jsPromiseObject{state: jsPromisePending}
	vm.jsPromiseItems[resultPromiseID] = resultPObj

	capability := &jsPromiseCapability{promise: resultPromise}

	reaction := jsPromiseReaction{
		onFulfilled: onFulfilled,
		onRejected:  onRejected,
		capability:  capability,
	}

	if pObj.state == jsPromisePending {
		pObj.reactions = append(pObj.reactions, reaction)
	} else {
		vm.jsEnqueueMicrotask(func() {
			state := pObj.state
			result := pObj.result
			handler := onFulfilled
			if state == jsPromiseRejected {
				handler = onRejected
			}

			if handler.Type != VTJSFunction {
				if state == jsPromiseFulfilled {
					vm.jsResolvePromise(resultPromise, result)
				} else {
					vm.jsRejectPromise(resultPromise, result)
				}
				return
			}

			var handlerResult Value
			caught := false
			func() {
				defer func() {
					if r := recover(); r != nil {
						caught = true
						if vmErr, ok := r.(*VMError); ok {
							vm.jsRejectPromise(resultPromise, NewString(vmErr.Error()))
						} else {
							vm.jsRejectPromise(resultPromise, NewString(fmt.Sprintf("%v", r)))
						}
					}
				}()
				handlerResult = vm.jsCall(handler, Value{Type: VTJSUndefined}, []Value{result})
			}()

			if !caught {
				vm.jsResolvePromise(resultPromise, handlerResult)
			}
		})
	}

	return resultPromise
}

func (vm *VM) jsPromiseCatch(promise Value, args []Value) Value {
	return vm.jsPromiseThen(promise, []Value{Value{Type: VTJSUndefined}, jsArgOrUndefined(args, 0)})
}

func (vm *VM) jsPromiseFinally(promise Value, args []Value) Value {
	onFinally := jsArgOrUndefined(args, 0)
	if onFinally.Type != VTJSFunction {
		return vm.jsPromiseThen(promise, []Value{onFinally, onFinally})
	}

	fulfilledHandler := vm.jsCreateFinallyHandler(onFinally, true)
	rejectedHandler := vm.jsCreateFinallyHandler(onFinally, false)

	return vm.jsPromiseThen(promise, []Value{fulfilledHandler, rejectedHandler})
}

func (vm *VM) jsCreateFinallyHandler(onFinally Value, fulfilled bool) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("PromiseFinallyHandler")
	obj["__js_finally_callback"] = onFinally
	obj["__js_finally_fulfilled"] = NewBool(fulfilled)
	vm.jsObjectItems[objID] = obj
	return Value{Type: VTJSFunction, Num: objID}
}

func (vm *VM) jsHandlePromiseFinallyHandler(callee Value, args []Value) Value {
	callback := vm.jsObjectItems[callee.Num]["__js_finally_callback"]
	fulfilled := vm.jsObjectItems[callee.Num]["__js_finally_fulfilled"].Num != 0
	val := jsArgOrUndefined(args, 0)

	// result = callback()
	result := vm.jsCall(callback, Value{Type: VTJSUndefined}, nil)

	// return Promise.resolve(result).then(() => (fulfilled ? val : throw val))
	p := vm.jsPromiseStaticResolve([]Value{result})

	var thenHandler Value
	if fulfilled {
		thenHandler = vm.jsCreateConstantHandler(val, true)
	} else {
		thenHandler = vm.jsCreateConstantHandler(val, false)
	}

	return vm.jsPromiseThen(p, []Value{thenHandler, thenHandler})
}

func (vm *VM) jsCreateConstantHandler(val Value, resolve bool) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("PromiseConstantHandler")
	obj["__js_const_val"] = val
	obj["__js_const_resolve"] = NewBool(resolve)
	vm.jsObjectItems[objID] = obj
	return Value{Type: VTJSFunction, Num: objID}
}

func (vm *VM) jsHandlePromiseConstantHandler(callee Value, args []Value) Value {
	val := vm.jsObjectItems[callee.Num]["__js_const_val"]
	resolve := vm.jsObjectItems[callee.Num]["__js_const_resolve"].Num != 0

	if resolve {
		return val
	}
	panic(&VMError{Msg: vm.valueToString(val)}) // Simple throw
}

func (vm *VM) jsPromiseStaticResolve(args []Value) Value {
	resolution := jsArgOrUndefined(args, 0)
	if resolution.Type == VTJSPromise {
		return resolution
	}

	promiseID := vm.allocJSID()
	promise := Value{Type: VTJSPromise, Num: promiseID}
	pObj := &jsPromiseObject{
		state:  jsPromiseFulfilled,
		result: resolution,
	}
	vm.jsPromiseItems[promiseID] = pObj
	return promise
}

func (vm *VM) jsPromiseStaticReject(args []Value) Value {
	reason := jsArgOrUndefined(args, 0)
	promiseID := vm.allocJSID()
	promise := Value{Type: VTJSPromise, Num: promiseID}
	pObj := &jsPromiseObject{
		state:  jsPromiseRejected,
		result: reason,
	}
	vm.jsPromiseItems[promiseID] = pObj
	return promise
}

func (vm *VM) jsPromiseStaticAll(args []Value) Value {
	if len(args) == 0 || args[0].Type != VTArray {
		return vm.jsPromiseStaticReject([]Value{NewString("TypeError: Promise.all requires an array")})
	}

	values := args[0].Arr.Values
	if len(values) == 0 {
		return vm.jsPromiseStaticResolve([]Value{ValueFromVBArray(NewVBArrayFromValues(0, nil))})
	}

	resultPromiseID := vm.allocJSID()
	resultPromise := Value{Type: VTJSPromise, Num: resultPromiseID}
	resultPObj := &jsPromiseObject{state: jsPromisePending}
	vm.jsPromiseItems[resultPromiseID] = resultPObj

	results := make([]Value, len(values))
	remaining := len(values)

	// Create a shared state for the resolvers
	allStateID := vm.allocJSID()
	allState := make(map[string]Value, 3)
	allState["results"] = ValueFromVBArray(NewVBArrayFromValues(0, results))
	allState["remaining"] = NewInteger(int64(remaining))
	allState["promise"] = resultPromise
	vm.jsObjectItems[allStateID] = allState

	for i, v := range values {
		idx := i
		// We need to wrap each element in a promise
		p := vm.jsPromiseStaticResolve([]Value{v})

		resolver := vm.jsCreatePromiseAllResolver(allStateID, idx)
		rejecter := vm.jsCreatePromiseRejectFunction(resultPromise)

		vm.jsPromiseThen(p, []Value{resolver, rejecter})
	}

	return resultPromise
}

func (vm *VM) jsCreatePromiseAllResolver(stateID int64, index int) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("PromiseAllResolver")
	obj["__js_all_state"] = Value{Type: VTJSObject, Num: stateID}
	obj["__js_all_index"] = NewInteger(int64(index))
	vm.jsObjectItems[objID] = obj
	return Value{Type: VTJSFunction, Num: objID}
}

func (vm *VM) jsHandlePromiseAllResolver(callee Value, args []Value) {
	stateVal := vm.jsObjectItems[callee.Num]["__js_all_state"]
	indexVal := vm.jsObjectItems[callee.Num]["__js_all_index"]

	state := vm.jsObjectItems[stateVal.Num]
	resultsArr := state["results"].Arr
	remaining := state["remaining"].Num
	promise := state["promise"]

	index := int(indexVal.Num)
	val := jsArgOrUndefined(args, 0)

	resultsArr.Values[index] = val
	remaining--
	state["remaining"] = NewInteger(remaining)

	if remaining == 0 {
		vm.jsResolvePromise(promise, state["results"])
	}
}

func (vm *VM) jsPromiseStaticRace(args []Value) Value {
	if len(args) == 0 || args[0].Type != VTArray {
		return vm.jsPromiseStaticReject([]Value{NewString("TypeError: Promise.race requires an array")})
	}

	values := args[0].Arr.Values
	resultPromiseID := vm.allocJSID()
	resultPromise := Value{Type: VTJSPromise, Num: resultPromiseID}
	resultPObj := &jsPromiseObject{state: jsPromisePending}
	vm.jsPromiseItems[resultPromiseID] = resultPObj

	resolve := vm.jsCreatePromiseResolveFunction(resultPromise)
	reject := vm.jsCreatePromiseRejectFunction(resultPromise)

	for _, v := range values {
		p := vm.jsPromiseStaticResolve([]Value{v})
		vm.jsPromiseThen(p, []Value{resolve, reject})
	}

	return resultPromise
}

// jsEnqueueMicrotask adds a task to the microtask queue.
func (vm *VM) jsEnqueueMicrotask(task func()) {
	vm.jsMicrotaskQueue = append(vm.jsMicrotaskQueue, task)
}

// jsProcessMicrotasks executes all pending microtasks until the queue is empty.
func (vm *VM) jsProcessMicrotasks() {
	for len(vm.jsMicrotaskQueue) > 0 {
		task := vm.jsMicrotaskQueue[0]
		vm.jsMicrotaskQueue = vm.jsMicrotaskQueue[1:]
		func() {
			defer func() {
				if r := recover(); r != nil {
					// In a real browser/Node environment, this would trigger
					// an 'unhandledrejection' event or print to console.
					if vmErr, ok := r.(*VMError); ok {
						fmt.Printf("Unhandled Microtask Error: %v\n", vmErr)
					} else {
						fmt.Printf("Unhandled Microtask Panic: %v\n", r)
					}
				}
			}()
			task()
		}()
	}
}

func (vm *VM) jsGetPromiseState(promise Value) jsPromiseState {
	if promise.Type != VTJSPromise {
		return jsPromiseFulfilled
	}
	pObj, ok := vm.jsPromiseItems[promise.Num]
	if !ok {
		return jsPromiseFulfilled
	}
	return pObj.state
}

func (vm *VM) jsGetPromiseResult(promise Value) Value {
	if promise.Type != VTJSPromise {
		return promise
	}
	pObj, ok := vm.jsPromiseItems[promise.Num]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	return pObj.result
}

type jsAsyncRejectionError struct {
	reason Value
}

func (e *jsAsyncRejectionError) Error() string { return "async rejection" }

func (vm *VM) jsAsyncCall(callee Value, thisVal Value, args []Value) Value {
	// 1. Create the promise that will be returned
	promiseID := vm.allocJSID()
	promise := Value{Type: VTJSPromise, Num: promiseID}
	pObj := &jsPromiseObject{state: jsPromisePending}
	vm.jsPromiseItems[promiseID] = pObj

	// 2. Clone the VM and run the function synchronously (blocking at awaits)
	child := vm.cloneForExecuteLocal(len(vm.bytecode))
	if child.jsBeginFunctionCall(callee, thisVal, args, Value{Type: VTJSUndefined}, false, Value{Type: VTJSUndefined}, false) {
		err := func() (err error) {
			defer func() {
				if r := recover(); r != nil {
					if are, ok := r.(*jsAsyncRejectionError); ok {
						err = are
					} else if vme, ok := r.(*VMError); ok {
						err = vme
					} else {
						err = fmt.Errorf("%v", r)
					}
				}
			}()
			return child.Run()
		}()

		if are, ok := err.(*jsAsyncRejectionError); ok {
			vm.jsRejectPromise(promise, are.reason)
		} else if err != nil {
			vm.jsRejectPromise(promise, NewString(err.Error()))
		} else {
			result := Value{Type: VTJSUndefined}
			if child.sp >= 0 {
				result = child.stack[child.sp]
			}
			vm.jsResolvePromise(promise, result)
		}
		vm.syncExecuteGlobalState(child)
	}

	return promise
}

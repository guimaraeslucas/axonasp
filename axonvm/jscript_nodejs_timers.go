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
	"strings"
	"time"
)

// jsTimerResultQueueSize is the capacity of the timer-fired channel.
const jsTimerResultQueueSize = 256

// jsTimerItem holds the Go-side state for a single setTimeout or setInterval handle.
type jsTimerItem struct {
	timer      *time.Timer   // non-nil for setTimeout
	ticker     *time.Ticker  // non-nil for setInterval
	stopCh     chan struct{} // closed to stop the interval goroutine
	callback   Value
	args       []Value
	isInterval bool
	unrefed    bool // true after .unref() — does not block VM teardown
	cancelled  bool
	intervalMs int64
}

// jsTimerFiredResult is sent from a goroutine when a timer/interval fires.
type jsTimerFiredResult struct {
	timerID    int64
	isInterval bool
}

// jsImmediateItem holds the state for a single setImmediate handle.
type jsImmediateItem struct {
	id       int64
	callback Value
	args     []Value
}

// jsNextTickItem holds the state for a single process.nextTick call.
type jsNextTickItem struct {
	callback Value
	args     []Value
}

// jsNodeTimerMinDelay is the minimum timer resolution (4 ms, matching browser/Node spec).
const jsNodeTimerMinDelay = 4 * time.Millisecond

// jsNodeTimerMaxDelay is Node.js maximum setTimeout delay (about 24.8 days).
const jsNodeTimerMaxDelay = (1<<31 - 1) * time.Millisecond

// jsAllocTimerID returns the next unique timer ID, using the shared dynamic ID counter.
func (vm *VM) jsAllocTimerID() int64 {
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	return id
}

// jsCreateTimeoutObject builds the Timeout native object returned by setTimeout/setInterval.
// The object exposes ref(), unref(), hasRef(), and refresh() per Node.js spec.
func (vm *VM) jsCreateTimeoutObject(timerID int64) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("Timeout")
	obj["__js_timer_id"] = NewInteger(timerID)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	return Value{Type: VTJSObject, Num: objID}
}

// jsGetTimerIDFromObject extracts the timer ID from a Timeout object.
// Returns (id, true) on success; (0, false) if the value is not a Timeout object.
func (vm *VM) jsGetTimerIDFromObject(v Value) (int64, bool) {
	if v.Type != VTJSObject {
		return 0, false
	}
	obj, ok := vm.jsObjectItems[v.Num]
	if !ok {
		return 0, false
	}
	t, ok := obj["__js_type"]
	if !ok || t.Type != VTString || t.Str != "Timeout" {
		return 0, false
	}
	idVal, ok := obj["__js_timer_id"]
	if !ok {
		return 0, false
	}
	return idVal.Num, true
}

// jsCallTimeoutMethod handles ref/unref/hasRef/refresh on a Timeout object.
func (vm *VM) jsCallTimeoutMethod(target Value, member string, args []Value) (Value, bool) {
	timerID, ok := vm.jsGetTimerIDFromObject(target)
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	item := vm.jsTimerItems[timerID]

	switch strings.ToLower(member) {
	case "ref":
		if item != nil {
			item.unrefed = false
		}
		return target, true

	case "unref":
		if item != nil {
			item.unrefed = true
		}
		return target, true

	case "hasref":
		if item == nil || item.cancelled {
			return NewBool(false), true
		}
		return NewBool(!item.unrefed), true

	case "refresh":
		if item != nil && !item.cancelled && item.timer != nil {
			dur := jsDurationFromMS(item.intervalMs)
			item.timer.Reset(dur)
		}
		return target, true

	case "close":
		// Undocumented but widely used alias for clearTimeout
		vm.jsCancelTimer(timerID)
		return Value{Type: VTJSUndefined}, true

	case "[symbol.toprimitive]", "tostring":
		return NewString("Timeout"), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsDurationFromMS converts a millisecond delay to a Go duration,
// clamping to [jsNodeTimerMinDelay, jsNodeTimerMaxDelay].
func jsDurationFromMS(ms int64) time.Duration {
	if ms < 4 {
		ms = 4
	}
	dur := min(time.Duration(ms)*time.Millisecond, jsNodeTimerMaxDelay)
	return dur
}

// jsSetTimeout schedules a one-shot callback after `delay` ms.
// Returns a Timeout object matching Node.js semantics.
func (vm *VM) jsSetTimeout(args []Value) Value {
	if len(args) < 1 || !vm.jsIsCallable(args[0]) {
		vm.jsThrowTypeError("setTimeout requires a callback function")
		return Value{Type: VTJSUndefined}
	}
	callback := args[0]
	delayMS := int64(0)
	if len(args) > 1 {
		delayMS = int64(vm.jsToNumber(args[1]).Flt)
	}
	extraArgs := []Value(nil)
	if len(args) > 2 {
		extraArgs = args[2:]
	}

	timerID := vm.jsAllocTimerID()
	item := &jsTimerItem{
		callback:   callback,
		args:       extraArgs,
		isInterval: false,
		intervalMs: delayMS,
	}
	vm.jsTimerItems[timerID] = item

	dur := jsDurationFromMS(delayMS)
	ch := vm.jsTimerResultQueue
	id := timerID
	item.timer = time.AfterFunc(dur, func() {
		select {
		case ch <- jsTimerFiredResult{timerID: id, isInterval: false}:
		default:
			// Queue full — drop the callback to prevent goroutine leak.
		}
	})

	return vm.jsCreateTimeoutObject(timerID)
}

// jsClearTimeout cancels a pending setTimeout or setInterval identified by a Timeout object.
// Passing null/undefined is a no-op per Node.js spec.
func (vm *VM) jsClearTimeout(args []Value) Value {
	if len(args) == 0 {
		return Value{Type: VTJSUndefined}
	}
	v := args[0]
	if v.Type == VTNull || v.Type == VTJSUndefined || v.Type == VTEmpty {
		return Value{Type: VTJSUndefined}
	}
	timerID, ok := vm.jsGetTimerIDFromObject(v)
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	vm.jsCancelTimer(timerID)
	return Value{Type: VTJSUndefined}
}

// jsCancelTimer stops the underlying Go timer/ticker for the given ID.
func (vm *VM) jsCancelTimer(timerID int64) {
	item, ok := vm.jsTimerItems[timerID]
	if !ok {
		return
	}
	item.cancelled = true
	if item.timer != nil {
		item.timer.Stop()
	}
	if item.ticker != nil {
		item.ticker.Stop()
	}
	if item.stopCh != nil {
		close(item.stopCh)
		item.stopCh = nil
	}
	delete(vm.jsTimerItems, timerID)
}

// jsSetInterval schedules a recurring callback every `delay` ms.
// Returns a Timeout object matching Node.js semantics.
func (vm *VM) jsSetInterval(args []Value) Value {
	if len(args) < 1 || !vm.jsIsCallable(args[0]) {
		vm.jsThrowTypeError("setInterval requires a callback function")
		return Value{Type: VTJSUndefined}
	}
	callback := args[0]
	delayMS := int64(0)
	if len(args) > 1 {
		delayMS = int64(vm.jsToNumber(args[1]).Flt)
	}
	extraArgs := []Value(nil)
	if len(args) > 2 {
		extraArgs = args[2:]
	}

	timerID := vm.jsAllocTimerID()
	stopCh := make(chan struct{})
	item := &jsTimerItem{
		callback:   callback,
		args:       extraArgs,
		isInterval: true,
		intervalMs: delayMS,
		stopCh:     stopCh,
	}
	item.ticker = time.NewTicker(jsDurationFromMS(delayMS))
	vm.jsTimerItems[timerID] = item

	ch := vm.jsTimerResultQueue
	id := timerID
	ticker := item.ticker
	go func() {
		for {
			select {
			case <-ticker.C:
				select {
				case ch <- jsTimerFiredResult{timerID: id, isInterval: true}:
				default:
					// Queue full — skip this tick.
				}
			case <-stopCh:
				return
			}
		}
	}()

	return vm.jsCreateTimeoutObject(timerID)
}

// jsClearInterval stops a recurring interval. Alias behavior: clearTimeout/clearInterval
// are interchangeable in Node.js.
func (vm *VM) jsClearInterval(args []Value) Value {
	return vm.jsClearTimeout(args)
}

// jsSetImmediate enqueues callback to run after the current synchronous code
// and before the next timer tick, at the end of the current event-loop pass.
func (vm *VM) jsSetImmediate(args []Value) Value {
	if len(args) < 1 || !vm.jsIsCallable(args[0]) {
		vm.jsThrowTypeError("setImmediate requires a callback function")
		return Value{Type: VTJSUndefined}
	}
	id := vm.jsAllocTimerID()
	extraArgs := []Value(nil)
	if len(args) > 1 {
		extraArgs = args[1:]
	}
	vm.jsImmediateQueue = append(vm.jsImmediateQueue, jsImmediateItem{
		id:       id,
		callback: args[0],
		args:     extraArgs,
	})
	// Return a lightweight Timeout-compatible object so clearImmediate works.
	return vm.jsCreateTimeoutObject(id)
}

// jsClearImmediate cancels a pending setImmediate callback.
func (vm *VM) jsClearImmediate(args []Value) Value {
	if len(args) == 0 {
		return Value{Type: VTJSUndefined}
	}
	v := args[0]
	timerID, ok := vm.jsGetTimerIDFromObject(v)
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	for i, item := range vm.jsImmediateQueue {
		if item.id == timerID {
			// Remove by swapping with last element.
			last := len(vm.jsImmediateQueue) - 1
			vm.jsImmediateQueue[i] = vm.jsImmediateQueue[last]
			vm.jsImmediateQueue = vm.jsImmediateQueue[:last]
			return Value{Type: VTJSUndefined}
		}
	}
	return Value{Type: VTJSUndefined}
}

// jsPumpTimerResults drains up to max pending timer-fired results and
// schedules their callbacks as microtasks.
func (vm *VM) jsPumpTimerResults(max int) {
	for range max {
		select {
		case result := <-vm.jsTimerResultQueue:
			vm.jsHandleTimerFired(result)
		default:
			return
		}
	}
}

// jsHandleTimerFired resolves a fired timer result into a microtask callback.
func (vm *VM) jsHandleTimerFired(result jsTimerFiredResult) {
	item, ok := vm.jsTimerItems[result.timerID]
	if !ok || item.cancelled {
		return
	}
	// For one-shot timers remove them immediately.
	if !result.isInterval {
		delete(vm.jsTimerItems, result.timerID)
	}
	cb := item.callback
	cbArgs := item.args
	vm.jsEnqueueMicrotask(func() {
		vm.jsCall(cb, Value{Type: VTJSUndefined}, cbArgs)
	})
}

// jsProcessNextTickQueue moves all pending process.nextTick callbacks to the front of the
// microtask queue so they execute before Promise callbacks per Node.js spec.
// Using the microtask queue avoids direct jsCall recursion that would overflow the stack.
func (vm *VM) jsProcessNextTickQueue() {
	if len(vm.jsNextTickQueue) == 0 {
		return
	}
	// Build microtask wrappers for each nextTick item, preserving order.
	nextTickMicrotasks := make([]func(), 0, len(vm.jsNextTickQueue))
	for _, item := range vm.jsNextTickQueue {
		cb := item.callback
		cbArgs := item.args
		nextTickMicrotasks = append(nextTickMicrotasks, func() {
			if vm.jsIsCallable(cb) {
				vm.jsCall(cb, Value{Type: VTJSUndefined}, cbArgs)
			}
		})
	}
	vm.jsNextTickQueue = vm.jsNextTickQueue[:0]
	// Prepend to microtask queue so they run before existing microtasks.
	vm.jsMicrotaskQueue = append(nextTickMicrotasks, vm.jsMicrotaskQueue...)
}

// jsProcessImmediateQueue appends all pending setImmediate callbacks to the microtask queue.
// They run after the current microtask queue is drained, called via jsPumpNodeAsyncTasks.
func (vm *VM) jsProcessImmediateQueue() {
	if len(vm.jsImmediateQueue) == 0 {
		return
	}
	for _, item := range vm.jsImmediateQueue {
		cb := item.callback
		cbArgs := item.args
		vm.jsEnqueueMicrotask(func() {
			if vm.jsIsCallable(cb) {
				vm.jsCall(cb, Value{Type: VTJSUndefined}, cbArgs)
			}
		})
	}
	vm.jsImmediateQueue = vm.jsImmediateQueue[:0]
}

// jsStopAllTimers cancels every active timer/interval and drains the result channel.
// Called during VM reset to prevent goroutine leaks across ASP requests.
func (vm *VM) jsStopAllTimers() {
	for id, item := range vm.jsTimerItems {
		item.cancelled = true
		if item.timer != nil {
			item.timer.Stop()
		}
		if item.ticker != nil {
			item.ticker.Stop()
		}
		if item.stopCh != nil {
			close(item.stopCh)
			item.stopCh = nil
		}
		delete(vm.jsTimerItems, id)
	}
	// Drain any pending fired results that arrived before cancellation.
	for {
		select {
		case <-vm.jsTimerResultQueue:
		default:
			return
		}
	}
}

// jsHasPendingRefedOneShotTimer reports whether at least one active, referenced
// one-shot timer is still pending. Interval timers are excluded to avoid
// blocking script termination indefinitely.
func (vm *VM) jsHasPendingRefedOneShotTimer() bool {
	if len(vm.jsTimerItems) == 0 {
		return false
	}
	for _, item := range vm.jsTimerItems {
		if item == nil || item.cancelled || item.unrefed {
			continue
		}
		if !item.isInterval {
			return true
		}
	}
	return false
}

// jsDrainNodeAsyncOnExit gives one final event-loop drain pass before script
// termination so short-lived setTimeout callbacks can run in CLI/server mode.
func (vm *VM) jsDrainNodeAsyncOnExit() {
	if !vm.enableNodeCompatibility() {
		return
	}
	if vm.jsHasPendingRefedOneShotTimer() == false && len(vm.jsNextTickQueue) == 0 && len(vm.jsMicrotaskQueue) == 0 && len(vm.jsImmediateQueue) == 0 {
		return
	}

	maxWait := 5 * time.Second
	deadline := time.Now().Add(maxWait)
	for {
		vm.jsPumpNodeAsyncTasks(256)

		if !vm.jsHasPendingRefedOneShotTimer() && len(vm.jsNextTickQueue) == 0 && len(vm.jsMicrotaskQueue) == 0 && len(vm.jsImmediateQueue) == 0 {
			return
		}
		if time.Now().After(deadline) {
			return
		}
		time.Sleep(time.Millisecond)
	}
}

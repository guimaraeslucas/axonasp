/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 */

package axonvm

import (
	"math/big"
	"strings"
	"sync/atomic"
	"unsafe"
)

// jsCreateAtomicsObject creates the global Atomics object.
func (vm *VM) jsCreateAtomicsObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 16)
	obj["__js_type"] = NewString("Atomics")
	obj["__js_class"] = NewString("Atomics")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 16)

	createAtomicsMethod := func(name string, ctor string, length int) {
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

	createAtomicsMethod("add", "AtomicsAdd", 3)
	createAtomicsMethod("sub", "AtomicsSub", 3)
	createAtomicsMethod("and", "AtomicsAnd", 3)
	createAtomicsMethod("or", "AtomicsOr", 3)
	createAtomicsMethod("xor", "AtomicsXor", 3)
	createAtomicsMethod("load", "AtomicsLoad", 2)
	createAtomicsMethod("store", "AtomicsStore", 3)
	createAtomicsMethod("exchange", "AtomicsExchange", 3)
	createAtomicsMethod("compareExchange", "AtomicsCompareExchange", 4)
	createAtomicsMethod("isLockFree", "AtomicsIsLockFree", 1)

	return Value{Type: VTJSObject, Num: objID}
}

// jsAtomicsValidateTypedArray validates that the argument is an integer TypedArray 
// backed by a SharedArrayBuffer.
func (vm *VM) jsAtomicsValidateTypedArray(obj Value) (buf []byte, offset int, elemSize int, typeName string, ok bool) {
	if obj.Type != VTJSObject {
		vm.jsThrowTypeError("Atomics: argument is not an object")
		return nil, 0, 0, "", false
	}
	items, exists := vm.jsObjectItems[obj.Num]
	if !exists {
		vm.jsThrowTypeError("Atomics: invalid object")
		return nil, 0, 0, "", false
	}

	t, _ := items["__js_type"]
	typeName = t.Str
	if !jsIsIntegerTypedArray(typeName) {
		vm.jsThrowTypeError("Atomics: argument must be an integer TypedArray")
		return nil, 0, 0, "", false
	}

	bufIDVal, hasBufID := items["__js_buffer_id"]
	if !hasBufID {
		vm.jsThrowTypeError("Atomics: argument is not a TypedArray")
		return nil, 0, 0, "", false
	}

	// Check if it's a SharedArrayBuffer
	bufID := bufIDVal.Num
	buf, isShared := vm.jsSharedArrayBuffers[bufID]
	if !isShared {
		// Strictly enforce SharedArrayBuffer as per prompt instruction
		vm.jsThrowTypeError("Atomics: argument must be backed by a SharedArrayBuffer")
		return nil, 0, 0, "", false
	}

	offsetVal := items["__js_byte_offset"]
	offset = int(offsetVal.Num)
	elemSize = jsTypedArrayElementSize(typeName)

	return buf, offset, elemSize, typeName, true
}

func jsIsIntegerTypedArray(typeName string) bool {
	switch typeName {
	case "Int8Array", "Uint8Array", "Int16Array", "Uint16Array", "Int32Array", "Uint32Array", "BigInt64Array", "BigUint64Array":
		return true
	}
	return false
}

// jsAtomicsValidateAccess validates the index and returns the byte offset.
func (vm *VM) jsAtomicsValidateAccess(obj Value, indexVal Value, buf []byte, offset int, elemSize int) (int, bool) {
	index := int(vm.jsToNumber(indexVal).Flt)
	items := vm.jsObjectItems[obj.Num]
	lenVal := items["__js_byte_length"]
	byteLength := int(lenVal.Num)
	
	if index < 0 || index * elemSize >= byteLength {
		vm.jsThrowRangeError("Atomics: index out of bounds")
		return 0, false
	}

	return offset + (index * elemSize), true
}

// jsAtomicsAdd implements Atomics.add(typedArray, index, value)
func (vm *VM) jsAtomicsAdd(args []Value) Value {
	if len(args) < 3 {
		return Value{Type: VTJSUndefined}
	}
	buf, offset, elemSize, typeName, ok := vm.jsAtomicsValidateTypedArray(args[0])
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	bytePos, ok := vm.jsAtomicsValidateAccess(args[0], args[1], buf, offset, elemSize)
	if !ok {
		return Value{Type: VTJSUndefined}
	}

	val := args[2]

	switch typeName {
	case "Int8Array", "Uint8Array":
		fval := vm.jsToNumber(val).Flt
		old := buf[bytePos]
		buf[bytePos] = old + byte(fval)
		return NewInteger(int64(int8(old)))
	case "Int16Array", "Uint16Array":
		fval := vm.jsToNumber(val).Flt
		ptr := (*uint16)(unsafe.Pointer(&buf[bytePos]))
		old := *ptr
		*ptr = old + uint16(fval)
		return NewInteger(int64(int16(old)))
	case "Int32Array":
		fval := vm.jsToNumber(val).Flt
		ptr := (*int32)(unsafe.Pointer(&buf[bytePos]))
		old := atomic.AddInt32(ptr, int32(fval)) - int32(fval)
		return NewInteger(int64(old))
	case "Uint32Array":
		fval := vm.jsToNumber(val).Flt
		ptr := (*uint32)(unsafe.Pointer(&buf[bytePos]))
		old := atomic.AddUint32(ptr, uint32(fval)) - uint32(fval)
		return NewInteger(int64(old))
	case "BigInt64Array":
		if val.Type != VTJSBigInt {
			vm.jsThrowTypeError("Atomics: argument must be a BigInt")
			return Value{Type: VTJSUndefined}
		}
		ptr := (*int64)(unsafe.Pointer(&buf[bytePos]))
		old := atomic.AddInt64(ptr, val.Big.Int64()) - val.Big.Int64()
		return NewBigInt(big.NewInt(old))
	case "BigUint64Array":
		if val.Type != VTJSBigInt {
			vm.jsThrowTypeError("Atomics: argument must be a BigInt")
			return Value{Type: VTJSUndefined}
		}
		ptr := (*uint64)(unsafe.Pointer(&buf[bytePos]))
		old := atomic.AddUint64(ptr, val.Big.Uint64()) - val.Big.Uint64()
		bi := new(big.Int).SetUint64(old)
		return NewBigInt(bi)
	}

	return Value{Type: VTJSUndefined}
}

// jsAtomicsLoad implements Atomics.load(typedArray, index)
func (vm *VM) jsAtomicsLoad(args []Value) Value {
	if len(args) < 2 {
		return Value{Type: VTJSUndefined}
	}
	buf, offset, elemSize, typeName, ok := vm.jsAtomicsValidateTypedArray(args[0])
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	bytePos, ok := vm.jsAtomicsValidateAccess(args[0], args[1], buf, offset, elemSize)
	if !ok {
		return Value{Type: VTJSUndefined}
	}

	switch typeName {
	case "Int8Array":
		return NewInteger(int64(int8(buf[bytePos])))
	case "Uint8Array":
		return NewInteger(int64(buf[bytePos]))
	case "Int16Array":
		ptr := (*int16)(unsafe.Pointer(&buf[bytePos]))
		return NewInteger(int64(*ptr))
	case "Uint16Array":
		ptr := (*uint16)(unsafe.Pointer(&buf[bytePos]))
		return NewInteger(int64(*ptr))
	case "Int32Array":
		ptr := (*int32)(unsafe.Pointer(&buf[bytePos]))
		return NewInteger(int64(atomic.LoadInt32(ptr)))
	case "Uint32Array":
		ptr := (*uint32)(unsafe.Pointer(&buf[bytePos]))
		return NewInteger(int64(atomic.LoadUint32(ptr)))
	case "BigInt64Array":
		ptr := (*int64)(unsafe.Pointer(&buf[bytePos]))
		return NewInteger(atomic.LoadInt64(ptr))
	case "BigUint64Array":
		ptr := (*uint64)(unsafe.Pointer(&buf[bytePos]))
		return NewInteger(int64(atomic.LoadUint64(ptr)))
	}

	return Value{Type: VTJSUndefined}
}

// jsAtomicsStore implements Atomics.store(typedArray, index, value)
func (vm *VM) jsAtomicsStore(args []Value) Value {
	if len(args) < 3 {
		return Value{Type: VTJSUndefined}
	}
	buf, offset, elemSize, typeName, ok := vm.jsAtomicsValidateTypedArray(args[0])
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	bytePos, ok := vm.jsAtomicsValidateAccess(args[0], args[1], buf, offset, elemSize)
	if !ok {
		return Value{Type: VTJSUndefined}
	}

	val := args[2]

	switch typeName {
	case "Int8Array", "Uint8Array":
		buf[bytePos] = byte(vm.jsToNumber(val).Flt)
	case "Int16Array", "Uint16Array":
		ptr := (*uint16)(unsafe.Pointer(&buf[bytePos]))
		*ptr = uint16(vm.jsToNumber(val).Flt)
	case "Int32Array":
		ptr := (*int32)(unsafe.Pointer(&buf[bytePos]))
		atomic.StoreInt32(ptr, int32(vm.jsToNumber(val).Flt))
	case "Uint32Array":
		ptr := (*uint32)(unsafe.Pointer(&buf[bytePos]))
		atomic.StoreUint32(ptr, uint32(vm.jsToNumber(val).Flt))
	case "BigInt64Array":
		if val.Type != VTJSBigInt {
			vm.jsThrowTypeError("Atomics: argument must be a BigInt")
			return Value{Type: VTJSUndefined}
		}
		ptr := (*int64)(unsafe.Pointer(&buf[bytePos]))
		atomic.StoreInt64(ptr, val.Big.Int64())
	case "BigUint64Array":
		if val.Type != VTJSBigInt {
			vm.jsThrowTypeError("Atomics: argument must be a BigInt")
			return Value{Type: VTJSUndefined}
		}
		ptr := (*uint64)(unsafe.Pointer(&buf[bytePos]))
		atomic.StoreUint64(ptr, val.Big.Uint64())
	}

	return val // Returns the value stored
}

// jsAtomicsExchange implements Atomics.exchange(typedArray, index, value)
func (vm *VM) jsAtomicsExchange(args []Value) Value {
	if len(args) < 3 {
		return Value{Type: VTJSUndefined}
	}
	buf, offset, elemSize, typeName, ok := vm.jsAtomicsValidateTypedArray(args[0])
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	bytePos, ok := vm.jsAtomicsValidateAccess(args[0], args[1], buf, offset, elemSize)
	if !ok {
		return Value{Type: VTJSUndefined}
	}

	val := args[2]

	switch typeName {
	case "Int8Array", "Uint8Array":
		fval := vm.jsToNumber(val).Flt
		old := buf[bytePos]
		buf[bytePos] = byte(fval)
		return NewInteger(int64(int8(old)))
	case "Int16Array", "Uint16Array":
		fval := vm.jsToNumber(val).Flt
		ptr := (*uint16)(unsafe.Pointer(&buf[bytePos]))
		old := *ptr
		*ptr = uint16(fval)
		return NewInteger(int64(int16(old)))
	case "Int32Array":
		fval := vm.jsToNumber(val).Flt
		ptr := (*int32)(unsafe.Pointer(&buf[bytePos]))
		old := atomic.SwapInt32(ptr, int32(fval))
		return NewInteger(int64(old))
	case "Uint32Array":
		fval := vm.jsToNumber(val).Flt
		ptr := (*uint32)(unsafe.Pointer(&buf[bytePos]))
		old := atomic.SwapUint32(ptr, uint32(fval))
		return NewInteger(int64(old))
	case "BigInt64Array":
		if val.Type != VTJSBigInt {
			vm.jsThrowTypeError("Atomics: argument must be a BigInt")
			return Value{Type: VTJSUndefined}
		}
		ptr := (*int64)(unsafe.Pointer(&buf[bytePos]))
		old := atomic.SwapInt64(ptr, val.Big.Int64())
		return NewBigInt(big.NewInt(old))
	case "BigUint64Array":
		if val.Type != VTJSBigInt {
			vm.jsThrowTypeError("Atomics: argument must be a BigInt")
			return Value{Type: VTJSUndefined}
		}
		ptr := (*uint64)(unsafe.Pointer(&buf[bytePos]))
		old := atomic.SwapUint64(ptr, val.Big.Uint64())
		return NewBigInt(new(big.Int).SetUint64(old))
	}

	return Value{Type: VTJSUndefined}
}

// jsAtomicsCompareExchange implements Atomics.compareExchange(typedArray, index, expectedValue, replacementValue)
func (vm *VM) jsAtomicsCompareExchange(args []Value) Value {
	if len(args) < 4 {
		return Value{Type: VTJSUndefined}
	}
	buf, offset, elemSize, typeName, ok := vm.jsAtomicsValidateTypedArray(args[0])
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	bytePos, ok := vm.jsAtomicsValidateAccess(args[0], args[1], buf, offset, elemSize)
	if !ok {
		return Value{Type: VTJSUndefined}
	}

	expected := args[2]
	replacement := args[3]

	switch typeName {
	case "Int8Array", "Uint8Array":
		fexp := vm.jsToNumber(expected).Flt
		frepl := vm.jsToNumber(replacement).Flt
		old := buf[bytePos]
		if old == byte(fexp) {
			buf[bytePos] = byte(frepl)
		}
		return NewInteger(int64(int8(old)))
	case "Int16Array", "Uint16Array":
		fexp := vm.jsToNumber(expected).Flt
		frepl := vm.jsToNumber(replacement).Flt
		ptr := (*uint16)(unsafe.Pointer(&buf[bytePos]))
		old := *ptr
		if old == uint16(fexp) {
			*ptr = uint16(frepl)
		}
		return NewInteger(int64(int16(old)))
	case "Int32Array":
		fexp := vm.jsToNumber(expected).Flt
		frepl := vm.jsToNumber(replacement).Flt
		ptr := (*int32)(unsafe.Pointer(&buf[bytePos]))
		if atomic.CompareAndSwapInt32(ptr, int32(fexp), int32(frepl)) {
			return NewInteger(int64(int32(fexp)))
		}
		return NewInteger(int64(atomic.LoadInt32(ptr)))
	case "Uint32Array":
		fexp := vm.jsToNumber(expected).Flt
		frepl := vm.jsToNumber(replacement).Flt
		ptr := (*uint32)(unsafe.Pointer(&buf[bytePos]))
		if atomic.CompareAndSwapUint32(ptr, uint32(fexp), uint32(frepl)) {
			return NewInteger(int64(uint32(fexp)))
		}
		return NewInteger(int64(atomic.LoadUint32(ptr)))
	case "BigInt64Array":
		if expected.Type != VTJSBigInt || replacement.Type != VTJSBigInt {
			vm.jsThrowTypeError("Atomics: arguments must be BigInt")
			return Value{Type: VTJSUndefined}
		}
		ptr := (*int64)(unsafe.Pointer(&buf[bytePos]))
		expVal := expected.Big.Int64()
		replVal := replacement.Big.Int64()
		if atomic.CompareAndSwapInt64(ptr, expVal, replVal) {
			return NewBigInt(big.NewInt(expVal))
		}
		return NewBigInt(big.NewInt(atomic.LoadInt64(ptr)))
	case "BigUint64Array":
		if expected.Type != VTJSBigInt || replacement.Type != VTJSBigInt {
			vm.jsThrowTypeError("Atomics: arguments must be BigInt")
			return Value{Type: VTJSUndefined}
		}
		ptr := (*uint64)(unsafe.Pointer(&buf[bytePos]))
		expVal := expected.Big.Uint64()
		replVal := replacement.Big.Uint64()
		if atomic.CompareAndSwapUint64(ptr, expVal, replVal) {
			return NewBigInt(new(big.Int).SetUint64(expVal))
		}
		return NewBigInt(new(big.Int).SetUint64(atomic.LoadUint64(ptr)))
	}

	return Value{Type: VTJSUndefined}
}

// jsAtomicsCall implements static methods of the Atomics object.
func (vm *VM) jsAtomicsCall(member string, args []Value) (Value, bool) {
	switch strings.ToLower(member) {
	case "add":
		return vm.jsAtomicsAdd(args), true
	case "sub":
		// Implement sub as add with negative value
		if len(args) >= 3 {
			v := vm.jsToNumber(args[2])
			args[2] = NewDouble(-v.Flt)
		}
		return vm.jsAtomicsAdd(args), true
	case "and":
		// ... logic for and ...
		return vm.jsAtomicsBitwise(member, args), true
	case "or":
		return vm.jsAtomicsBitwise(member, args), true
	case "xor":
		return vm.jsAtomicsBitwise(member, args), true
	case "load":
		return vm.jsAtomicsLoad(args), true
	case "store":
		return vm.jsAtomicsStore(args), true
	case "exchange":
		return vm.jsAtomicsExchange(args), true
	case "compareexchange":
		return vm.jsAtomicsCompareExchange(args), true
	case "islockfree":
		// Simplified: 1, 2, 4, 8 byte operations are typically lock-free on modern CPUs.
		if len(args) > 0 {
			size := int(vm.jsToNumber(args[0]).Flt)
			return NewBool(size == 1 || size == 2 || size == 4 || size == 8), true
		}
		return NewBool(false), true
	}
	return Value{Type: VTJSUndefined}, false
}

func (vm *VM) jsAtomicsBitwise(op string, args []Value) Value {
	if len(args) < 3 {
		return Value{Type: VTJSUndefined}
	}
	buf, offset, elemSize, typeName, ok := vm.jsAtomicsValidateTypedArray(args[0])
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	bytePos, ok := vm.jsAtomicsValidateAccess(args[0], args[1], buf, offset, elemSize)
	if !ok {
		return Value{Type: VTJSUndefined}
	}

	val := uint64(vm.jsToNumber(args[2]).Flt)

	switch typeName {
	case "Int8Array", "Uint8Array":
		old := buf[bytePos]
		switch strings.ToLower(op) {
		case "and": buf[bytePos] = old & byte(val)
		case "or":  buf[bytePos] = old | byte(val)
		case "xor": buf[bytePos] = old ^ byte(val)
		}
		return NewInteger(int64(int8(old)))
	case "Int16Array", "Uint16Array":
		ptr := (*uint16)(unsafe.Pointer(&buf[bytePos]))
		old := *ptr
		switch strings.ToLower(op) {
		case "and": *ptr = old & uint16(val)
		case "or":  *ptr = old | uint16(val)
		case "xor": *ptr = old ^ uint16(val)
		}
		return NewInteger(int64(int16(old)))
	case "Int32Array", "Uint32Array":
		// Go's sync/atomic doesn't have And/Or/Xor until Go 1.23 for some types,
		// or at all for others. For compliance and single-threaded safety, 
		// we can use a loop with CompareAndSwap.
		ptr := (*uint32)(unsafe.Pointer(&buf[bytePos]))
		for {
			old := atomic.LoadUint32(ptr)
			var next uint32
			switch strings.ToLower(op) {
			case "and": next = old & uint32(val)
			case "or":  next = old | uint32(val)
			case "xor": next = old ^ uint32(val)
			}
			if atomic.CompareAndSwapUint32(ptr, old, next) {
				return NewInteger(int64(int32(old)))
			}
		}
	case "BigInt64Array", "BigUint64Array":
		ptr := (*uint64)(unsafe.Pointer(&buf[bytePos]))
		for {
			old := atomic.LoadUint64(ptr)
			var next uint64
			switch strings.ToLower(op) {
			case "and": next = old & uint64(val)
			case "or":  next = old | uint64(val)
			case "xor": next = old ^ uint64(val)
			}
			if atomic.CompareAndSwapUint64(ptr, old, next) {
				return NewInteger(int64(old))
			}
		}
	}

	return Value{Type: VTJSUndefined}
}

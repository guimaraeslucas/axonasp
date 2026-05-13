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
	"strconv"
)

// jsArrayIterator represents the state of an array iterator.
type jsArrayIterator struct {
	target Value
	values []Value // cached values if provided
	index  int
	kind   int // 0: values, 1: keys, 2: entries
}

// jsStringIterator represents the state of a string iterator.
type jsStringIterator struct {
	target string
	runes  []rune
	index  int
}

// jsCreateArrayIterator creates a new Array Iterator object.
func (vm *VM) jsCreateArrayIterator(target Value, kind int) Value {
	id := vm.allocJSID()
	vm.jsObjectItems[id] = map[string]Value{
		"__js_type": NewString("Array Iterator"),
		"__js_ctor": NewString("Array Iterator"),
	}
	vm.jsPropertyItems[id] = make(map[string]jsPropertyDescriptor, 2)

	vm.jsArrayIterators[id] = &jsArrayIterator{
		target: target,
		index:  0,
		kind:   kind,
	}
	return Value{Type: VTJSObject, Num: id}
}

// jsCreateValuesIterator creates an iterator from a slice of values.
func (vm *VM) jsCreateValuesIterator(values []Value) Value {
	id := vm.allocJSID()
	vm.jsObjectItems[id] = map[string]Value{
		"__js_type": NewString("Array Iterator"),
		"__js_ctor": NewString("Array Iterator"),
	}
	vm.jsPropertyItems[id] = make(map[string]jsPropertyDescriptor, 2)

	vm.jsArrayIterators[id] = &jsArrayIterator{
		target: Value{Type: VTJSUndefined},
		values: values,
		index:  0,
		kind:   0,
	}
	return Value{Type: VTJSObject, Num: id}
}

// jsCreateStringIterator creates a new String Iterator object.
func (vm *VM) jsCreateStringIterator(target string) Value {
	id := vm.allocJSID()
	vm.jsObjectItems[id] = map[string]Value{
		"__js_type": NewString("String Iterator"),
		"__js_ctor": NewString("String Iterator"),
	}
	vm.jsPropertyItems[id] = make(map[string]jsPropertyDescriptor, 2)

	vm.jsStringIterators[id] = &jsStringIterator{
		target: target,
		runes:  []rune(target),
		index:  0,
	}
	return Value{Type: VTJSObject, Num: id}
}

// jsIteratorNextResult creates the { value: ..., done: ... } object.
func (vm *VM) jsIteratorNextResult(value Value, done bool) Value {
	id := vm.allocJSID()
	obj := make(map[string]Value, 2)
	obj["value"] = value
	obj["done"] = NewBool(done)
	vm.jsObjectItems[id] = obj

	props := make(map[string]jsPropertyDescriptor, 2)
	props["value"] = jsPropertyDescriptor{Value: value, HasValue: true, Enumerable: true, Configurable: true, Writable: true}
	props["done"] = jsPropertyDescriptor{Value: obj["done"], HasValue: true, Enumerable: true, Configurable: true, Writable: true}
	vm.jsPropertyItems[id] = props

	return Value{Type: VTJSObject, Num: id}
}

// jsPopulatePrototypes adds ES6+ methods and well-known symbols to built-in prototypes.
func (vm *VM) jsPopulatePrototypes(bindings map[string]Value) {
	// Array.prototype[Symbol.iterator] = Array.prototype.values
	if arrayCtor, ok := bindings["Array"]; ok {
		if proto, deferred := vm.jsMemberGet(arrayCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			valuesFn := vm.jsCreateNativeFunction("values", "ArrayValues")
			vm.jsSetDescriptor(proto.Num, "values", jsDefaultPropertyDescriptor(valuesFn))

			itKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolIterator, 10)
			vm.jsSetDescriptor(proto.Num, itKey, jsPropertyDescriptor{
				Value:        valuesFn,
				HasValue:     true,
				Enumerable:   false,
				Configurable: true,
				Writable:     true,
			})
		}
	}

	// String.prototype[Symbol.iterator]
	if stringCtor, ok := bindings["String"]; ok {
		if proto, deferred := vm.jsMemberGet(stringCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			itKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolIterator, 10)
			itFn := vm.jsCreateNativeFunction("[Symbol.iterator]", "StringIteratorFactory")
			vm.jsSetDescriptor(proto.Num, itKey, jsPropertyDescriptor{
				Value:        itFn,
				HasValue:     true,
				Enumerable:   false,
				Configurable: true,
				Writable:     true,
			})
		}
	}

	// Object.prototype methods as callable function objects.
	if objectCtor, ok := bindings["Object"]; ok {
		if proto, deferred := vm.jsMemberGet(objectCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			for _, name := range []string{"hasOwnProperty", "propertyIsEnumerable", "isPrototypeOf", "toString", "toLocaleString", "valueOf"} {
				vm.jsSetDescriptor(proto.Num, name, jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction(name, "ObjectPrototype")))
			}
		}
	}

	// WeakSet.prototype
	if wsCtor, ok := bindings["WeakSet"]; ok {
		if proto, deferred := vm.jsMemberGet(wsCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			for _, name := range []string{"add", "has", "delete"} {
				vm.jsSetDescriptor(proto.Num, name, jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction(name, "WeakSet")))
			}
		}
	}

	// WeakRef.prototype
	if wrCtor, ok := bindings["WeakRef"]; ok {
		if proto, deferred := vm.jsMemberGet(wrCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			vm.jsSetDescriptor(proto.Num, "deref", jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction("deref", "WeakRef")))
		}
	}

	// FinalizationRegistry.prototype
	if frCtor, ok := bindings["FinalizationRegistry"]; ok {
		if proto, deferred := vm.jsMemberGet(frCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			for _, name := range []string{"register", "unregister"} {
				vm.jsSetDescriptor(proto.Num, name, jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction(name, "FinalizationRegistry")))
			}
		}
	}

	// Map.prototype
	if mapCtor, ok := bindings["Map"]; ok {
		if proto, deferred := vm.jsMemberGet(mapCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			for _, name := range []string{"set", "get", "has", "delete", "clear"} {
				vm.jsSetDescriptor(proto.Num, name, jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction(name, "Map")))
			}
		}
	}

	// WeakMap.prototype
	if weakMapCtor, ok := bindings["WeakMap"]; ok {
		if proto, deferred := vm.jsMemberGet(weakMapCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			vm.jsSetDescriptor(proto.Num, "get", jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction("get", "WeakMap")))
			vm.jsSetDescriptor(proto.Num, "set", jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction("set", "WeakMap")))
			vm.jsSetDescriptor(proto.Num, "has", jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction("has", "WeakMap")))
			vm.jsSetDescriptor(proto.Num, "delete", jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction("delete", "WeakMap")))
		}
	}

	// WeakSet.prototype
	if weakSetCtor, ok := bindings["WeakSet"]; ok {
		if proto, deferred := vm.jsMemberGet(weakSetCtor, "prototype"); !deferred && proto.Type == VTJSObject {
			vm.jsSetDescriptor(proto.Num, "add", jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction("add", "WeakSet")))
			vm.jsSetDescriptor(proto.Num, "has", jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction("has", "WeakSet")))
			vm.jsSetDescriptor(proto.Num, "delete", jsDefaultPropertyDescriptor(vm.jsCreateNativeFunction("delete", "WeakSet")))
		}
	}
}

// jsCreateNativeFunction creates a dummy JS function object that jsCall redirects to.
func (vm *VM) jsCreateNativeFunction(name string, ctorName string) Value {
	id := vm.allocJSID()
	vm.jsObjectItems[id] = map[string]Value{
		"__js_type": NewString("function"),
		"__js_ctor": NewString(ctorName),
		"name":      NewString(name),
	}
	vm.jsPropertyItems[id] = make(map[string]jsPropertyDescriptor, 2)
	// Even if it's not a full closure, it's better to use VTJSFunction so typeof is correct.
	return Value{Type: VTJSFunction, Num: id}
}

// jsArrayIteratorNext implements the next() method for Array Iterators.
func (vm *VM) jsArrayIteratorNext(itObj Value) Value {
	it, ok := vm.jsArrayIterators[itObj.Num]
	if !ok {
		return vm.jsIteratorNextResult(Value{Type: VTJSUndefined}, true)
	}

	length := 0
	var values []Value

	if it.values != nil {
		values = it.values
		length = len(values)
	} else if it.target.Type == VTArray && it.target.Arr != nil {
		values = it.target.Arr.Values
		length = len(values)
	} else if it.target.Type == VTJSObject {
		lenVal, _ := vm.jsMemberGet(it.target, "length")
		length = int(vm.jsToNumber(lenVal).Flt)
	}

	if it.index >= length {
		return vm.jsIteratorNextResult(Value{Type: VTJSUndefined}, true)
	}

	var val Value
	switch it.kind {
	case 1: // keys
		val = NewInteger(int64(it.index))
	case 2: // entries
		entryVal := vm.allocJSID()
		entryArr := NewVBArrayFromValues(0, []Value{NewInteger(int64(it.index)), vm.jsArrayIteratorGetVal(it.target, values, it.index)})
		vm.jsObjectItems[entryVal] = map[string]Value{"__js_vbarray_source": ValueFromVBArray(entryArr)}
		val = Value{Type: VTJSObject, Num: entryVal}
	default: // values
		val = vm.jsArrayIteratorGetVal(it.target, values, it.index)
	}

	it.index++
	return vm.jsIteratorNextResult(val, false)
}

func (vm *VM) jsArrayIteratorGetVal(target Value, values []Value, index int) Value {
	if values != nil && index < len(values) {
		return values[index]
	}
	return vm.jsIndexGet(target, NewInteger(int64(index)))
}

// jsStringIteratorNext implements the next() method for String Iterators.
func (vm *VM) jsStringIteratorNext(itObj Value) Value {
	it, ok := vm.jsStringIterators[itObj.Num]
	if !ok {
		return vm.jsIteratorNextResult(Value{Type: VTJSUndefined}, true)
	}

	if it.index >= len(it.runes) {
		return vm.jsIteratorNextResult(Value{Type: VTJSUndefined}, true)
	}

	val := NewString(string(it.runes[it.index]))
	it.index++
	return vm.jsIteratorNextResult(val, false)
}

// jsGetIterator obtains an iterator from an object via Symbol.iterator.
func (vm *VM) jsGetIterator(source Value) Value {
	if source.Type == VTNull || source.Type == VTJSUndefined {
		vm.jsThrowTypeError("Cannot destructure null or undefined")
		return Value{Type: VTJSUndefined}
	}

	// Fast paths for basic types
	if source.Type == VTArray {
		return vm.jsCreateArrayIterator(source, 0)
	}
	if source.Type == VTString {
		return vm.jsCreateStringIterator(source.Str)
	}

	itKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolIterator, 10)
	itFn, _ := vm.jsMemberGet(source, itKey)
	if itFn.Type != VTJSFunction && itFn.Type != VTJSObject {
		// Fallback for native types that might not have prototypes yet or are handled specially
		if source.Type == VTJSObject {
			class := vm.jsObjectStringProperty(source, "__js_type")
			if class == "Map" || class == "Set" || jsIsTypedArrayType(class) {
				vals := vm.jsEnumerateForOfValues(source)
				return vm.jsCreateValuesIterator(vals)
			}
		}
		vm.jsThrowTypeError("Object is not iterable")
		return Value{Type: VTJSUndefined}
	}
	itObj := vm.jsCall(itFn, source, nil)
	if itObj.Type != VTJSObject {
		vm.jsThrowTypeError("Iterator result is not an object")
		return Value{Type: VTJSUndefined}
	}
	return itObj
}

// jsCollectRest consumes all remaining values from an iterator and returns them as a JS Array.
func (vm *VM) jsCollectRest(itObj Value) Value {
	values := make([]Value, 0)
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
		values = append(values, val)
	}
	return ValueFromVBArray(NewVBArrayFromValues(0, values))
}

// jsObjectRest creates a new object with all enumerable properties of source EXCEPT excluded keys.
func (vm *VM) jsObjectRest(source Value, excluded map[string]struct{}) Value {
	targetID := vm.allocJSID()
	targetObj := make(map[string]Value, 8)
	targetProps := make(map[string]jsPropertyDescriptor, 8)

	// In ES spec, source is coerced to object.
	// For now we handle VTJSObject enumerable own properties.
	if source.Type == VTJSObject {
		// Use jsObjectOwnEnumerableKeys to correctly respect property descriptors
		keys := vm.jsObjectOwnEnumerableKeys(source.Num)
		for _, k := range keys {
			if _, isExcluded := excluded[k]; isExcluded {
				continue
			}
			val, _ := vm.jsMemberGet(source, k)
			targetObj[k] = val
			targetProps[k] = jsDefaultPropertyDescriptor(val)
		}
	} else if source.Type == VTArray && source.Arr != nil {
		// Special case: internal arrays are somewhat object-like in destructuring
		for i, v := range source.Arr.Values {
			k := strconv.Itoa(source.Arr.Lower + i)
			if _, isExcluded := excluded[k]; isExcluded {
				continue
			}
			targetObj[k] = v
			targetProps[k] = jsDefaultPropertyDescriptor(v)
		}
	}

	vm.jsObjectItems[targetID] = targetObj
	vm.jsPropertyItems[targetID] = targetProps
	return Value{Type: VTJSObject, Num: targetID}
}

// jsIteratorNextValue calls .next() on an iterator and returns the yielded value.
// If the iterator is done, returns undefined.
func (vm *VM) jsIteratorNextValue(itObj Value) Value {
	result, handled := vm.jsCallMember(itObj, "next", nil)
	if !handled || result.Type != VTJSObject {
		vm.jsThrowTypeError("Iterator result is not an object")
		return Value{Type: VTJSUndefined}
	}
	doneVal, _ := vm.jsMemberGet(result, "done")
	if vm.jsTruthy(doneVal) {
		return Value{Type: VTJSUndefined}
	}
	val, _ := vm.jsMemberGet(result, "value")
	return val
}

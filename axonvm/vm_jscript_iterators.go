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

	if it.target.Type == VTArray && it.target.Arr != nil {
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

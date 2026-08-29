package axonvm

import (
	"strings"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// vbsCollection is the runtime state for one Collection instance.
type vbsCollection struct {
	items []Value
}

// vbsCollectionEnumerator is the runtime state for one Collection enumerator.
type vbsCollectionEnumerator struct {
	items []Value
}

// newVBSCollection allocates a fresh, empty Collection instance.
func newVBSCollection() *vbsCollection {
	return &vbsCollection{
		items: make([]Value, 0, 8),
	}
}

// newCollectionObject stores a fresh Collection instance in the VM and returns its handle.
func (vm *VM) newCollectionObject() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.collectionItems[objID] = newVBSCollection()
	return Value{Type: VTNativeObject, Num: objID}
}

// dispatchCollectionMethod handles all method calls and the default property (member == "") for Collection instances.
func (vm *VM) dispatchCollectionMethod(objID int64, member string, args []Value) (Value, bool) {
	c, exists := vm.collectionItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	if member == "" || strings.EqualFold(member, "Item") {
		if len(args) == 1 {
			idx := vm.asInt(args[0])
			if idx < 1 || idx > len(c.items) {
				vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
				return Value{Type: VTEmpty}, true
			}
			return c.items[idx-1], true
		}
		if len(args) >= 2 {
			vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Collection does not support item assignment")
			return Value{Type: VTEmpty}, true
		}
		vm.raise(vbscript.WrongNumberOfParameters, "Collection.Item requires 1 argument: index")
		return Value{Type: VTEmpty}, true
	}

	switch {
	case strings.EqualFold(member, "Add"):
		if len(args) < 1 {
			vm.raise(vbscript.WrongNumberOfParameters, "Collection.Add requires at least 1 argument")
			return Value{Type: VTEmpty}, true
		}
		c.items = append(c.items, args[0])
		return Value{Type: VTEmpty}, true

	case strings.EqualFold(member, "Remove"):
		if len(args) < 1 {
			vm.raise(vbscript.WrongNumberOfParameters, "Collection.Remove requires 1 argument: index")
			return Value{Type: VTEmpty}, true
		}
		idx := vm.asInt(args[0])
		if idx < 1 || idx > len(c.items) {
			vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
			return Value{Type: VTEmpty}, true
		}
		c.items = append(c.items[:idx-1], c.items[idx:]...)
		return Value{Type: VTEmpty}, true

	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(c.items))), true

	case strings.EqualFold(member, "_NewEnum") || strings.EqualFold(member, "NewEnum"):
		enumID := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		// Snapshot items at the time the enumerator is obtained
		enumItems := make([]Value, len(c.items))
		copy(enumItems, c.items)
		vm.collectionEnumeratorItems[enumID] = &vbsCollectionEnumerator{items: enumItems}
		return Value{Type: VTNativeObject, Num: enumID}, true
	}

	vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Collection does not support '"+member+"'")
	return Value{Type: VTEmpty}, true
}

// dispatchCollectionPropertyGet handles property reads for Collection.
func (vm *VM) dispatchCollectionPropertyGet(objID int64, member string) (Value, bool) {
	c, exists := vm.collectionItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	switch {
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(c.items))), true
	case strings.EqualFold(member, "Item"):
		return vm.newNativeObjectProxy(objID, "Item", nil), true
	case strings.EqualFold(member, "_NewEnum") || strings.EqualFold(member, "NewEnum"):
		enumID := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		enumItems := make([]Value, len(c.items))
		copy(enumItems, c.items)
		vm.collectionEnumeratorItems[enumID] = &vbsCollectionEnumerator{items: enumItems}
		return Value{Type: VTNativeObject, Num: enumID}, true
	}

	return Value{Type: VTEmpty}, false
}

// dispatchCollectionPropertySet handles property Let/Set assignments for Collection.
func (vm *VM) dispatchCollectionPropertySet(objID int64, member string, val Value) bool {
	_, exists := vm.collectionItems[objID]
	if !exists {
		return false
	}
	// Collection properties are read-only
	vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Collection property '"+member+"' is read-only")
	return true
}

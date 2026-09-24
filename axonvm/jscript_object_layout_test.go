package axonvm

import (
	"strconv"
	"testing"
)

func TestJScriptObjectShapeTransitionsAreShared(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	firstID := vm.allocJSID()
	secondID := vm.allocJSID()
	for _, id := range []int64{firstID, secondID} {
		vm.jsObjectItems[id] = make(map[string]Value)
		vm.jsObjectShape[id] = 0
		vm.jsObjectSlots[id] = nil
		vm.jsMemberSet(Value{Type: VTJSObject, Num: id}, "alpha", NewInteger(id))
		vm.jsMemberSet(Value{Type: VTJSObject, Num: id}, "beta", NewString("value"))
	}

	firstShape := vm.jsObjectShape[firstID]
	if firstShape == 0 || vm.jsObjectShape[secondID] != firstShape {
		t.Fatalf("equivalent objects did not share a shape: %d and %d", firstShape, vm.jsObjectShape[secondID])
	}
	if len(vm.jsShapeTransitions) != 2 {
		t.Fatalf("expected two shared shape transitions, got %d", len(vm.jsShapeTransitions))
	}
	if got := vm.jsObjectSlots[firstID][0]; got.Type != VTInteger || got.Num != firstID {
		t.Fatalf("first object did not retain its own slot value: %#v", got)
	}
	if got := vm.jsObjectSlots[secondID][0]; got.Type != VTInteger || got.Num != secondID {
		t.Fatalf("second object did not retain its own slot value: %#v", got)
	}
}

func TestJScriptObjectShapeTransitionsSurvivePooledReset(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	id := vm.allocJSID()
	vm.jsObjectItems[id] = make(map[string]Value)
	vm.jsObjectShape[id] = 0
	vm.jsObjectSlots[id] = nil
	vm.jsMemberSet(Value{Type: VTJSObject, Num: id}, "alpha", NewInteger(1))
	shape := vm.jsObjectShape[id]

	vm.captureBaseProgramState()
	vm.resetForReuse()
	if shape == 0 {
		t.Fatal("object did not receive a shape")
	}
	if len(vm.jsShapeTransitions) == 0 || len(vm.jsShapeSlots) == 0 || len(vm.jsShapeSlotIndex) == 0 {
		t.Fatal("pooled reset discarded reusable shape metadata")
	}
	if len(vm.jsObjectShape) != 0 || len(vm.jsObjectSlots) != 0 || len(vm.jsObjectShapeDisabled) != 0 {
		t.Fatal("pooled reset retained request-specific object shape state")
	}

	secondID := vm.allocJSID()
	vm.jsObjectItems[secondID] = make(map[string]Value)
	vm.jsObjectShape[secondID] = 0
	vm.jsObjectSlots[secondID] = nil
	vm.jsMemberSet(Value{Type: VTJSObject, Num: secondID}, "alpha", NewInteger(2))
	if vm.jsObjectShape[secondID] != shape {
		t.Fatalf("new request did not reuse shape: got %d want %d", vm.jsObjectShape[secondID], shape)
	}
}

func TestJScriptObjectShapeCacheIsBounded(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	for i := range jsObjectShapeCacheLimit + 1 {
		id := vm.allocJSID()
		vm.jsObjectItems[id] = make(map[string]Value)
		vm.jsObjectShape[id] = 0
		vm.jsObjectSlots[id] = nil
		vm.jsMemberSet(Value{Type: VTJSObject, Num: id}, "unique_"+strconv.Itoa(i), NewInteger(int64(i)))
	}
	if len(vm.jsShapeTransitions) != jsObjectShapeCacheLimit {
		t.Fatalf("shape cache exceeded limit: got %d want %d", len(vm.jsShapeTransitions), jsObjectShapeCacheLimit)
	}
	if !vm.jsShapeCacheSaturated {
		t.Fatal("shape cache was not marked saturated")
	}
}

func TestJScriptDefaultPropertiesDoNotAllocateDescriptors(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	id := vm.allocJSID()
	vm.jsObjectItems[id] = make(map[string]Value)
	obj := Value{Type: VTJSObject, Num: id}

	vm.jsMemberSet(obj, "ordinary", NewInteger(1))
	if props := vm.jsPropertyItems[id]; props != nil {
		t.Fatal("ordinary property unexpectedly allocated descriptor storage")
	}
	desc, exists := vm.jsGetDescriptor(id, "ordinary")
	if !exists || !desc.HasValue || !desc.Writable || !desc.Enumerable || !desc.Configurable || desc.Value.Num != 1 {
		t.Fatalf("ordinary property did not synthesize its default descriptor: %#v, %v", desc, exists)
	}

	vm.jsSetDescriptor(id, "fixed", jsPropertyDescriptor{
		Value:      NewInteger(2),
		HasValue:   true,
		Enumerable: true,
	})
	if _, exists := vm.jsPropertyItems[id]["fixed"]; !exists {
		t.Fatal("non-default property did not retain descriptor storage")
	}
	vm.jsMemberSet(obj, "fixed", NewInteger(3))
	if got := vm.jsObjectItems[id]["fixed"]; got.Num != 2 {
		t.Fatalf("non-writable property changed: %#v", got)
	}
}

func TestJScriptShapeMutationAndDeleteSemantics(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var a = {x: 1, y: 2}; var b = {x: 3, y: 4};` +
		`a.x = 5; delete a.y; a.y = 6;` +
		`Object.defineProperty(b, "x", {value: 7, writable: false, enumerable: false, configurable: false});` +
		`b.x = 8;` +
		`Response.Write(a.x + ":" + a.y + ":" + b.x + ":" + b.y + ":" + Object.keys(b).join(",") + ":" + delete b.x);` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "5:6:7:4:y:false" {
		t.Fatalf("unexpected property semantics output: %q", out)
	}
}

func TestJScriptLargeObjectStopsTrackingShapeTransitions(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	id := vm.allocJSID()
	vm.jsObjectItems[id] = make(map[string]Value)
	vm.jsObjectShape[id] = 0
	vm.jsObjectSlots[id] = nil
	obj := Value{Type: VTJSObject, Num: id}

	for i := 0; i <= jsObjectShapePropertyLimit; i++ {
		vm.jsMemberSet(obj, "key_"+strconv.Itoa(i), NewInteger(int64(i)))
	}
	if _, tracked := vm.jsObjectShape[id]; tracked {
		t.Fatal("large object retained incremental shape tracking past the limit")
	}
	if _, disabled := vm.jsObjectShapeDisabled[id]; !disabled {
		t.Fatal("large object was not marked as permanently unshaped")
	}
	lastKey := "key_" + strconv.Itoa(jsObjectShapePropertyLimit)
	if got := vm.jsObjectItems[id][lastKey]; got.Type != VTInteger || got.Num != int64(jsObjectShapePropertyLimit) {
		t.Fatalf("large object lost property value: %#v", got)
	}
	transitionCount := len(vm.jsShapeTransitions)
	if vm.jsEnsureObjectICLayout(id) {
		t.Fatal("oversized object unexpectedly rebuilt an inline-cache layout")
	}
	if len(vm.jsShapeTransitions) != transitionCount {
		t.Fatal("oversized object created more shape transitions during a later lookup")
	}
}

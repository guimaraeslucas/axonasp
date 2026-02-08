package experimental

import "fmt"

const StackSize = 2048
const GlobalsSize = 65536
const MaxFrames = 1024

// Value represents a VBScript value (can be anything: int, string, object, etc.)
// In a real implementation, we might use a specific struct to avoid boxing/unboxing overhead
// or use a tagged pointer approach, but interface{} is fine for phase 1.
type Value interface{}

// Bytecode represents a compiled chunk of code
type Bytecode struct {
	Instructions []byte
	Constants    []Value
}

// CallFrame represents a function call frame
type CallFrame struct {
	Closure     *Bytecode // In real impl, this would be a Function object
	IP          int       // Instruction Pointer
	BasePointer int       // Where on the stack this frame's locals start
}

// VM is the Virtual Machine
type VM struct {
	Constants []Value
	Globals   map[string]Value // Symbol table for globals

	Stack []Value
	SP    int // Stack Pointer (points to the next free slot)

	Frames      []*CallFrame
	FramesIndex int
}

func NewVM(bytecode *Bytecode) *VM {
	mainFrame := &CallFrame{
		Closure:     bytecode,
		IP:          0,
		BasePointer: 0,
	}

	vm := &VM{
		Constants:   bytecode.Constants,
		Globals:     make(map[string]Value),
		Stack:       make([]Value, StackSize),
		SP:          0,
		Frames:      make([]*CallFrame, MaxFrames),
		FramesIndex: 1,
	}
	vm.Frames[0] = mainFrame

	return vm
}

func (vm *VM) StackTop() Value {
	if vm.SP == 0 {
		return nil
	}
	return vm.Stack[vm.SP-1]
}

func (vm *VM) Push(v Value) error {
	if vm.SP >= StackSize {
		return fmt.Errorf("stack overflow")
	}
	vm.Stack[vm.SP] = v
	vm.SP++
	return nil
}

func (vm *VM) Pop() Value {
	if vm.SP == 0 {
		return nil // Underflow?
	}
	val := vm.Stack[vm.SP-1]
	vm.SP--
	return val
}

func (vm *VM) LastPopped() Value {
	return vm.Stack[vm.SP]
}

package experimental

import "fmt"

const StackSize = 2048
const GlobalsSize = 65536
const MaxFrames = 1024

// VBScriptEmpty represents the VBScript Empty type
type VBScriptEmpty struct{}

func (e VBScriptEmpty) String() string { return "Empty" }

var EmptyValue = VBScriptEmpty{}

// HostEnvironment defines the interface for bridging VM with the host (AxonASP server)
type HostEnvironment interface {
	GetVariable(name string) (interface{}, bool)
	SetVariable(name string, value interface{}) error
	CallFunction(name string, args []interface{}) (interface{}, error)
	CreateObject(progID string) (interface{}, error)
}

// BuiltinFunction represents a host-provided function (e.g., Len, Mid)
type BuiltinFunction struct {
	Name string
}

// Function represents a compiled VBScript function or sub
type Function struct {
	Name           string
	Bytecode       *Bytecode
	ParameterCount int
	LocalCount     int // Number of local variables including parameters
}

// Value represents a VBScript value (can be anything: int, string, object, etc.)
type Value interface{}

// Bytecode represents a compiled chunk of code
type Bytecode struct {
	Instructions []byte
	Constants    []Value
	GlobalNames  []string
}

// CallFrame represents a function call frame
type CallFrame struct {
	Func        *Function
	IP          int // Instruction Pointer
	BasePointer int // Where on the stack this frame's locals start
}

// VM is the Virtual Machine
type VM struct {
	Constants   []Value
	Globals     map[string]Value // Symbol table for globals
	GlobalSlots []Value

	Stack []Value
	SP    int // Stack Pointer (points to the next free slot)

	Frames      []*CallFrame
	FramesIndex int

	Host HostEnvironment // Interface to the host system
}

func NewVM(mainFunc *Function, host HostEnvironment) *VM {
	mainFrame := &CallFrame{
		Func:        mainFunc,
		IP:          0,
		BasePointer: 0,
	}

	var globalSlots []Value
	if mainFunc != nil && mainFunc.Bytecode != nil && len(mainFunc.Bytecode.GlobalNames) > 0 {
		globalSlots = make([]Value, len(mainFunc.Bytecode.GlobalNames))
		for i := range globalSlots {
			globalSlots[i] = EmptyValue
		}
	}

	vm := &VM{
		Constants:   mainFunc.Bytecode.Constants,
		Globals:     make(map[string]Value),
		GlobalSlots: globalSlots,
		Stack:       make([]Value, StackSize),
		SP:          0,
		Frames:      make([]*CallFrame, MaxFrames),
		FramesIndex: 1,
		Host:        host,
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

func (vm *VM) lookupGlobalFunction(name string, frame *CallFrame) *Function {
	if name == "" {
		return nil
	}
	if fnVal, ok := vm.Globals[name]; ok {
		if fn, ok := fnVal.(*Function); ok {
			return fn
		}
	}
	if frame != nil && frame.Func != nil && frame.Func.Bytecode != nil {
		for idx, globalName := range frame.Func.Bytecode.GlobalNames {
			if globalName == name {
				if idx < len(vm.GlobalSlots) {
					if fn, ok := vm.GlobalSlots[idx].(*Function); ok {
						return fn
					}
				}
				break
			}
		}
	}
	return nil
}

func (vm *VM) lookupGlobalValue(name string, frame *CallFrame) Value {
	if name == "" {
		return nil
	}
	if val, ok := vm.Globals[name]; ok {
		return val
	}
	if frame != nil && frame.Func != nil && frame.Func.Bytecode != nil {
		for idx, globalName := range frame.Func.Bytecode.GlobalNames {
			if globalName == name {
				if idx < len(vm.GlobalSlots) {
					return vm.GlobalSlots[idx]
				}
				break
			}
		}
	}
	return nil
}

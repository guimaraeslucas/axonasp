package experimental

import (
	"encoding/binary"
	"fmt"
	"math"
)

// Run executes the bytecode starting from the current frame
func (vm *VM) Run() error {
	for vm.FramesIndex > 0 {
		frame := vm.Frames[vm.FramesIndex-1]
		instructions := frame.Func.Bytecode.Instructions

		if frame.IP >= len(instructions) {
			// Implicit return at end of function
			vm.FramesIndex--
			if vm.FramesIndex == 0 {
				return nil
			}
			continue
		}

		// Fetch
		op := Opcode(instructions[frame.IP])
		frame.IP++

		// Decode & Execute
		switch op {
		case OP_CONSTANT:
			constIndex := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			err := vm.Push(vm.Constants[constIndex])
			if err != nil {
				return err
			}

		case OP_TRUE:
			err := vm.Push(true)
			if err != nil {
				return err
			}
		case OP_FALSE:
			err := vm.Push(false)
			if err != nil {
				return err
			}
		case OP_NULL:
			err := vm.Push(nil) // VBScript Null is different from Empty
			if err != nil {
				return err
			}
		case OP_EMPTY:
			err := vm.Push(EmptyValue)
			if err != nil {
				return err
			}
		case OP_NOTHING:
			err := vm.Push(nil)
			if err != nil {
				return err
			}

		case OP_ADD:
			right := vm.Pop()
			left := vm.Pop()
			res, err := add(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_SUB:
			right := vm.Pop()
			left := vm.Pop()
			res, err := sub(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_MUL:
			right := vm.Pop()
			left := vm.Pop()
			res, err := mul(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_DIV:
			right := vm.Pop()
			left := vm.Pop()
			res, err := div(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_IDIV:
			right := vm.Pop()
			left := vm.Pop()
			res, err := idiv(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_MOD:
			right := vm.Pop()
			left := vm.Pop()
			res, err := mod(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_NEG:
			val := vm.Pop()
			res, err := neg(val)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_NOT:
			val := vm.Pop()
			res, err := notOp(val)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_EQUAL:
			right := vm.Pop()
			left := vm.Pop()
			res := isEqual(left, right)
			err := vm.Push(res)
			if err != nil {
				return err
			}

		case OP_GREATER:
			right := vm.Pop()
			left := vm.Pop()
			res, err := isGreater(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_LESS:
			right := vm.Pop()
			left := vm.Pop()
			res, err := isLess(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_CONCAT:
			right := vm.Pop()
			left := vm.Pop()
			res := concat(left, right)
			err := vm.Push(res)
			if err != nil {
				return err
			}

		case OP_GET_GLOBAL:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			nameVal := vm.Constants[nameIdx]
			nameStr := nameVal.(string)

			val, ok := vm.Globals[nameStr]
			if vm.Host != nil {
				if hVal, exists := vm.Host.GetVariable(nameStr); exists {
					val = hVal
					ok = true
				}
			}
			if !ok {
				val = EmptyValue
			}
			err := vm.Push(val)
			if err != nil {
				return err
			}

		case OP_SET_GLOBAL:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			nameVal := vm.Constants[nameIdx]
			nameStr := nameVal.(string)
			val := vm.Pop()

			if vm.Host != nil {
				_ = vm.Host.SetVariable(nameStr, val)
			}
			vm.Globals[nameStr] = val

		case OP_GET_GLOBAL_FAST:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			if int(nameIdx) >= len(frame.Func.Bytecode.GlobalNames) {
				return fmt.Errorf("global index out of range: %d", nameIdx)
			}
			nameStr := frame.Func.Bytecode.GlobalNames[nameIdx]
			var val Value = EmptyValue
			if vm.Host != nil {
				if hVal, ok := vm.Host.GetVariable(nameStr); ok {
					val = hVal
				}
			} else if int(nameIdx) < len(vm.GlobalSlots) {
				val = vm.GlobalSlots[nameIdx]
			} else if gVal, ok := vm.Globals[nameStr]; ok {
				val = gVal
			}
			err := vm.Push(val)
			if err != nil {
				return err
			}

		case OP_SET_GLOBAL_FAST:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			if int(nameIdx) >= len(frame.Func.Bytecode.GlobalNames) {
				return fmt.Errorf("global index out of range: %d", nameIdx)
			}
			nameStr := frame.Func.Bytecode.GlobalNames[nameIdx]
			val := vm.Pop()
			if vm.Host != nil {
				_ = vm.Host.SetVariable(nameStr, val)
			}
			if int(nameIdx) < len(vm.GlobalSlots) {
				vm.GlobalSlots[nameIdx] = val
			}
			vm.Globals[nameStr] = val

		case OP_GET_LOCAL:
			localIdx := int(instructions[frame.IP])
			frame.IP++
			err := vm.Push(vm.Stack[frame.BasePointer+localIdx])
			if err != nil {
				return err
			}

		case OP_SET_LOCAL:
			localIdx := int(instructions[frame.IP])
			frame.IP++
			vm.Stack[frame.BasePointer+localIdx] = vm.Pop()

		case OP_INC_LOCAL:
			localIdx := int(instructions[frame.IP])
			frame.IP++
			current := vm.Stack[frame.BasePointer+localIdx]
			res, err := add(current, int64(1))
			if err != nil {
				return err
			}
			vm.Stack[frame.BasePointer+localIdx] = res

		case OP_JUMP:
			offset := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			frame.IP += int(offset)

		case OP_JUMP_IF_FALSE:
			offset := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			condition := vm.Pop()
			if isFalsey(condition) {
				frame.IP += int(offset)
			}

		case OP_CALL:
			argCount := int(instructions[frame.IP])
			frame.IP++

			// Function object is below arguments on stack
			funcVal := vm.Stack[vm.SP-argCount-1]

			switch f := funcVal.(type) {
			case *Function:
				if argCount != f.ParameterCount {
					return fmt.Errorf("expected %d arguments, got %d", f.ParameterCount, argCount)
				}

				// New frame
				newFrame := &CallFrame{
					Func:        f,
					IP:          0,
					BasePointer: vm.SP - argCount,
				}

				// Ensure enough stack space for locals
				localsToAllocate := f.LocalCount - f.ParameterCount
				for i := 0; i < localsToAllocate; i++ {
					vm.Push(EmptyValue)
				}

				if vm.FramesIndex >= MaxFrames {
					return fmt.Errorf("stack overflow (frames)")
				}
				vm.Frames[vm.FramesIndex] = newFrame
				vm.FramesIndex++

			case *BuiltinFunction:
				// Collect args
				args := make([]interface{}, argCount)
				// Stack: [Func, Arg1, Arg2, ...]
				// SP points after last arg.
				// Args start at SP - argCount.
				for i := 0; i < argCount; i++ {
					args[i] = vm.Stack[vm.SP-argCount+i]
				}

				// Call Host
				if vm.Host == nil {
					return fmt.Errorf("host environment not available for builtin call")
				}
				res, err := vm.Host.CallFunction(f.Name, args)
				if err != nil {
					if fn := vm.lookupGlobalFunction(f.Name, frame); fn != nil {
						// Replace function object and dispatch as a user-defined function.
						vm.Stack[vm.SP-argCount-1] = fn
						if argCount != fn.ParameterCount {
							return fmt.Errorf("expected %d arguments, got %d", fn.ParameterCount, argCount)
						}
						newFrame := &CallFrame{
							Func:        fn,
							IP:          0,
							BasePointer: vm.SP - argCount,
						}
						localsToAllocate := fn.LocalCount - fn.ParameterCount
						for i := 0; i < localsToAllocate; i++ {
							vm.Push(EmptyValue)
						}
						if vm.FramesIndex >= MaxFrames {
							return fmt.Errorf("stack overflow (frames)")
						}
						vm.Frames[vm.FramesIndex] = newFrame
						vm.FramesIndex++
						break
					}
					if obj := vm.lookupGlobalValue(f.Name, frame); obj != nil {
						switch callable := obj.(type) {
						case interface {
							CallMethod(string, ...interface{}) (interface{}, error)
						}:
							result, callErr := callable.CallMethod("", args...)
							if callErr != nil {
								return callErr
							}
							vm.SP -= (argCount + 1)
							if result == nil {
								result = EmptyValue
							}
							vm.Push(result)
							break
						case interface {
							CallMethod(string, ...interface{}) interface{}
						}:
							result := callable.CallMethod("", args...)
							vm.SP -= (argCount + 1)
							if result == nil {
								result = EmptyValue
							}
							vm.Push(result)
							break
						}
					}
					return err
				}

				// Pop Func and Args
				vm.SP -= (argCount + 1)

				// Push Result
				if res == nil {
					res = EmptyValue
				}
				vm.Push(res)

			case interface {
				CallMethod(string, ...interface{}) (interface{}, error)
			}:
				args := make([]interface{}, argCount)
				for i := 0; i < argCount; i++ {
					args[i] = vm.Stack[vm.SP-argCount+i]
				}
				res, err := f.CallMethod("", args...)
				if err != nil {
					return err
				}
				vm.SP -= (argCount + 1)
				if res == nil {
					res = EmptyValue
				}
				vm.Push(res)

			case interface {
				CallMethod(string, ...interface{}) interface{}
			}:
				args := make([]interface{}, argCount)
				for i := 0; i < argCount; i++ {
					args[i] = vm.Stack[vm.SP-argCount+i]
				}
				res := f.CallMethod("", args...)
				vm.SP -= (argCount + 1)
				if res == nil {
					res = EmptyValue
				}
				vm.Push(res)

			default:
				return fmt.Errorf("can only call functions, got %T", funcVal)
			}

		case OP_RETURN_VALUE:
			result := vm.Pop()

			// Pop frame
			vm.FramesIndex--
			frame = vm.Frames[vm.FramesIndex-1] // Get previous frame to restore IP later

			// Pop locals and arguments and function object
			oldFrame := vm.Frames[vm.FramesIndex]
			vm.SP = oldFrame.BasePointer - 1

			// Push result
			vm.Push(result)

		case OP_RETURN:
			// Pop frame
			vm.FramesIndex--
			if vm.FramesIndex > 0 {
				oldFrame := vm.Frames[vm.FramesIndex]
				vm.SP = oldFrame.BasePointer - 1
				vm.Push(EmptyValue)
			}

		case OP_POP:
			vm.Pop()

		case OP_INC_GLOBAL_FAST:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			if int(nameIdx) >= len(frame.Func.Bytecode.GlobalNames) {
				return fmt.Errorf("global index out of range: %d", nameIdx)
			}
			nameStr := frame.Func.Bytecode.GlobalNames[nameIdx]
			var current Value = EmptyValue
			if vm.Host != nil {
				if hVal, ok := vm.Host.GetVariable(nameStr); ok {
					current = hVal
				}
			} else if int(nameIdx) < len(vm.GlobalSlots) {
				current = vm.GlobalSlots[nameIdx]
			} else if gVal, ok := vm.Globals[nameStr]; ok {
				current = gVal
			}
			res, err := add(current, int64(1))
			if err != nil {
				return err
			}
			if vm.Host != nil {
				_ = vm.Host.SetVariable(nameStr, res)
			}
			if int(nameIdx) < len(vm.GlobalSlots) {
				vm.GlobalSlots[nameIdx] = res
			}
			vm.Globals[nameStr] = res

		case OP_NEW:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			nameVal := vm.Constants[nameIdx]
			progID := nameVal.(string)

			if vm.Host == nil {
				return fmt.Errorf("host environment not available for object creation")
			}
			obj, err := vm.Host.CreateObject(progID)
			if err != nil {
				return err
			}
			err = vm.Push(obj)
			if err != nil {
				return err
			}

		default:
			return fmt.Errorf("unknown opcode execution: %d", op)
		}
	}
	return nil
}

// Helper for boolean logic
func isFalsey(v Value) bool {
	switch val := v.(type) {
	case bool:
		return !val
	case int, int64:
		return val == 0
	case VBScriptEmpty:
		return true
	case nil:
		return true // Null is falsey
	default:
		return false
	}
}

// --- Arithmetic Helpers ---

func toInt64(v Value) (int64, error) {
	switch val := v.(type) {
	case int:
		return int64(val), nil
	case int64:
		return val, nil
	case float64:
		// VBScript rounds to nearest even for some operations, but simple round for now
		return int64(math.Round(val)), nil
	case VBScriptEmpty:
		return 0, nil
	default:
		return 0, fmt.Errorf("cannot convert %T to int64", v)
	}
}

func toFloat(v Value) (float64, error) {
	switch val := v.(type) {
	case int:
		return float64(val), nil
	case int64:
		return float64(val), nil
	case float64:
		return val, nil
	case VBScriptEmpty:
		return 0, nil
	default:
		return 0, fmt.Errorf("type mismatch: expected number, got %T", v)
	}
}

func add(a, b Value) (Value, error) {
	if a == EmptyValue {
		a = int64(0)
	}
	if b == EmptyValue {
		b = int64(0)
	}

	i1, ok1 := a.(int64)
	i2, ok2 := b.(int64)
	if ok1 && ok2 {
		return i1 + i2, nil
	}

	f1, err := toFloat(a)
	if err != nil {
		return nil, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return nil, err
	}
	return f1 + f2, nil
}

func sub(a, b Value) (Value, error) {
	if a == EmptyValue {
		a = int64(0)
	}
	if b == EmptyValue {
		b = int64(0)
	}

	i1, ok1 := a.(int64)
	i2, ok2 := b.(int64)
	if ok1 && ok2 {
		return i1 - i2, nil
	}
	f1, err := toFloat(a)
	if err != nil {
		return nil, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return nil, err
	}
	return f1 - f2, nil
}

func mul(a, b Value) (Value, error) {
	if a == EmptyValue {
		a = int64(0)
	}
	if b == EmptyValue {
		b = int64(0)
	}

	i1, ok1 := a.(int64)
	i2, ok2 := b.(int64)
	if ok1 && ok2 {
		return i1 * i2, nil
	}
	f1, err := toFloat(a)
	if err != nil {
		return nil, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return nil, err
	}
	return f1 * f2, nil
}

func div(a, b Value) (Value, error) {
	f1, err := toFloat(a)
	if err != nil {
		return nil, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return nil, err
	}
	if f2 == 0 {
		return nil, fmt.Errorf("division by zero")
	}
	return f1 / f2, nil
}

func mod(a, b Value) (Value, error) {
	// VBScript rounds both operands to the nearest integer before Mod
	v1, err := toInt64(a)
	if err != nil {
		return nil, err
	}
	v2, err := toInt64(b)
	if err != nil {
		return nil, err
	}

	if v2 == 0 {
		return nil, fmt.Errorf("division by zero")
	}
	return v1 % v2, nil
}

func idiv(a, b Value) (Value, error) {
	// VBScript integer division rounds both operands before dividing
	v1, err := toInt64(a)
	if err != nil {
		return nil, err
	}
	v2, err := toInt64(b)
	if err != nil {
		return nil, err
	}
	if v2 == 0 {
		return nil, fmt.Errorf("division by zero")
	}
	return v1 / v2, nil
}

func neg(a Value) (Value, error) {
	switch val := a.(type) {
	case int64:
		return -val, nil
	case float64:
		return -val, nil
	case VBScriptEmpty:
		return int64(0), nil
	default:
		return nil, fmt.Errorf("invalid type for negation: %T", a)
	}
}

func notOp(a Value) (Value, error) {
	if b, ok := a.(bool); ok {
		return !b, nil
	}
	val, err := toInt64(a)
	if err != nil {
		return nil, err
	}
	return int64(^int32(val)), nil
}

func isEqual(a, b Value) bool {
	if a == EmptyValue && b == EmptyValue {
		return true
	}
	return a == b
}

func isGreater(a, b Value) (bool, error) {
	f1, err := toFloat(a)
	if err != nil {
		return false, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return false, err
	}
	return f1 > f2, nil
}

func isLess(a, b Value) (bool, error) {
	f1, err := toFloat(a)
	if err != nil {
		return false, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return false, err
	}
	return f1 < f2, nil
}

func concat(a, b Value) Value {
	return toString(a) + toString(b)
}

func toString(v Value) string {
	if v == nil {
		return ""
	}
	switch val := v.(type) {
	case string:
		return val
	case []byte:
		return string(val)
	case int:
		return fmt.Sprintf("%d", val)
	case int64:
		return fmt.Sprintf("%d", val)
	case float64:
		return fmt.Sprintf("%g", val)
	case bool:
		if val {
			return "True"
		}
		return "False"
	case VBScriptEmpty:
		return ""
	default:
		return fmt.Sprintf("%v", val)
	}
}

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

		case OP_POW:
			right := vm.Pop()
			left := vm.Pop()
			res, err := pow(left, right)
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

		case OP_NOT_EQUAL:
			right := vm.Pop()
			left := vm.Pop()
			res := !isEqual(left, right)
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

		case OP_GREATER_EQUAL:
			right := vm.Pop()
			left := vm.Pop()
			res, err := isGreaterEqual(left, right)
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

		case OP_LESS_EQUAL:
			right := vm.Pop()
			left := vm.Pop()
			res, err := isLessEqual(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_IS:
			right := vm.Pop()
			left := vm.Pop()
			// VBScript 'Is' operator checks if two object references refer to the same object
			err := vm.Push(left == right)
			if err != nil {
				return err
			}

		case OP_AND:
			right := vm.Pop()
			left := vm.Pop()
			res, err := andOp(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_OR:
			right := vm.Pop()
			left := vm.Pop()
			res, err := orOp(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_XOR:
			right := vm.Pop()
			left := vm.Pop()
			res, err := xorOp(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_EQV:
			right := vm.Pop()
			left := vm.Pop()
			res, err := eqvOp(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_IMP:
			right := vm.Pop()
			left := vm.Pop()
			res, err := impOp(left, right)
			if err != nil {
				return err
			}
			err = vm.Push(res)
			if err != nil {
				return err
			}

		case OP_GET_MEMBER:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			nameVal := vm.Constants[nameIdx]
			nameStr := nameVal.(string)
			obj := vm.Pop()

			if obj == nil || obj == EmptyValue {
				return fmt.Errorf("object required for member access: %s", nameStr)
			}

			var val Value = EmptyValue
			switch o := obj.(type) {
			case interface {
				GetProperty(string) interface{}
			}:
				val = o.GetProperty(nameStr)
			case map[string]interface{}:
				val = o[nameStr]
			default:
				return fmt.Errorf("object does not support member access: %T", obj)
			}
			err := vm.Push(val)
			if err != nil {
				return err
			}

		case OP_SET_MEMBER:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			nameVal := vm.Constants[nameIdx]
			nameStr := nameVal.(string)
			val := vm.Pop()
			obj := vm.Pop()

			if obj == nil || obj == EmptyValue {
				return fmt.Errorf("object required for member assignment: %s", nameStr)
			}

			switch o := obj.(type) {
			case interface {
				SetProperty(string, interface{}) error
			}:
				err := o.SetProperty(nameStr, val)
				if err != nil {
					return err
				}
			case map[string]interface{}:
				o[nameStr] = val
			default:
				return fmt.Errorf("object does not support member assignment: %T", obj)
			}

		case OP_CALL_MEMBER:
			nameIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			argCount := int(instructions[frame.IP])
			frame.IP++

			nameVal := vm.Constants[nameIdx]
			methodName := nameVal.(string)

			// Stack: [obj, arg1, arg2, ..., argN]
			// obj is at vm.SP - argCount - 1
			obj := vm.Stack[vm.SP-argCount-1]

			if obj == nil || obj == EmptyValue {
				return fmt.Errorf("object required for method call: %s", methodName)
			}

			// Try to get compiled method for direct VM execution
			if getter, ok := obj.(interface {
				GetMethod(string) *Function
			}); ok {
				fn := getter.GetMethod(methodName)
				if fn != nil {
					if argCount != fn.ParameterCount {
						return fmt.Errorf("expected %d arguments, got %d", fn.ParameterCount, argCount)
					}

					newFrame := &CallFrame{
						Func:          fn,
						IP:            0,
						BasePointer:   vm.SP - argCount,
						ContextObject: obj,
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
					continue
				}
			}

			args := make([]interface{}, argCount)
			for i := 0; i < argCount; i++ {
				args[i] = vm.Stack[vm.SP-argCount+i]
			}

			var res interface{}
			var err error

			switch o := obj.(type) {
			case interface {
				CallMethod(string, ...interface{}) (interface{}, error)
			}:
				res, err = o.CallMethod(methodName, args...)
				if err != nil {
					return err
				}
			case interface {
				CallMethod(string, ...interface{}) interface{}
			}:
				res = o.CallMethod(methodName, args...)
			default:
				return fmt.Errorf("object does not support method calls: %T", obj)
			}

			// Pop obj and args
			vm.SP -= (argCount + 1)

			if res == nil {
				res = EmptyValue
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

		case OP_SET_INDEXED:
			indexCount := int(instructions[frame.IP])
			frame.IP++

			// Stack: [obj, index1, ..., indexN, value]
			value := vm.Pop()
			indexes := make([]interface{}, indexCount)
			for i := indexCount - 1; i >= 0; i-- {
				indexes[i] = vm.Pop()
			}
			obj := vm.Pop()

			if vm.Host == nil {
				return fmt.Errorf("host environment not available for indexed assignment")
			}
			err := vm.Host.SetIndexed(obj, indexes, value)
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

			// 1. Check Context Object (Class Members)
			found := false
			if frame.ContextObject != nil {
				switch o := frame.ContextObject.(type) {
				case interface {
					GetMember(string) (interface{}, bool, error)
				}:
					if mVal, ok, _ := o.GetMember(nameStr); ok {
						val = mVal
						found = true
					}
				case interface {
					GetProperty(string) interface{}
				}:
					if mVal := o.GetProperty(nameStr); mVal != nil {
						val = mVal
						found = true
					}
				}
			}

			if !found {
				if vm.Host != nil {
					if hVal, ok := vm.Host.GetVariable(nameStr); ok {
						val = hVal
						found = true
					}
				}
			}

			if !found {
				if int(nameIdx) < len(vm.GlobalSlots) {
					val = vm.GlobalSlots[nameIdx]
				} else if gVal, ok := vm.Globals[nameStr]; ok {
					val = gVal
				}
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

			// 1. Check Context Object (Class Members)
			handled := false
			if frame.ContextObject != nil {
				if o, ok := frame.ContextObject.(interface {
					SetMember(string, interface{}) (bool, error)
				}); ok {
					if h, _ := o.SetMember(nameStr, val); h {
						handled = true
					}
				} else if o, ok := frame.ContextObject.(interface {
					SetProperty(string, interface{}) error
				}); ok {
					if err := o.SetProperty(nameStr, val); err == nil {
						handled = true
					}
				}
			}

			if !handled {
				if vm.Host != nil {
					_ = vm.Host.SetVariable(nameStr, val)
				}
				if int(nameIdx) < len(vm.GlobalSlots) {
					vm.GlobalSlots[nameIdx] = val
				}
				vm.Globals[nameStr] = val
			}

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
				continue

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
						continue
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
							continue
						case interface {
							CallMethod(string, ...interface{}) interface{}
						}:
							result := callable.CallMethod("", args...)
							vm.SP -= (argCount + 1)
							if result == nil {
								result = EmptyValue
							}
							vm.Push(result)
							continue
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
			if vm.FramesIndex == 0 {
				return nil
			}
			// Get previous frame to restore local frame variable for next loop iteration
			// although it will be re-assigned at top of loop.
			frame = vm.Frames[vm.FramesIndex-1]

			// Pop locals and arguments and function object
			oldFrame := vm.Frames[vm.FramesIndex]
			vm.SP = oldFrame.BasePointer - 1

			// Push result
			vm.Push(result)

		case OP_RETURN:
			// Pop frame
			vm.FramesIndex--
			if vm.FramesIndex == 0 {
				return nil
			}
			oldFrame := vm.Frames[vm.FramesIndex]
			vm.SP = oldFrame.BasePointer - 1
			vm.Push(EmptyValue)

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

		case OP_FALLBACK:
			nodeIdx := binary.BigEndian.Uint16(instructions[frame.IP:])
			frame.IP += 2
			node := vm.Constants[nodeIdx]

			if vm.Host == nil {
				return fmt.Errorf("host environment not available for fallback execution")
			}
			res, err := vm.Host.ExecuteAST(node)
			if err != nil {
				return err
			}
			if res != nil {
				err = vm.Push(res)
				if err != nil {
					return err
				}
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

func pow(a, b Value) (Value, error) {
	f1, err := toFloat(a)
	if err != nil {
		return nil, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return nil, err
	}
	return math.Pow(f1, f2), nil
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

func isGreaterEqual(a, b Value) (bool, error) {
	f1, err := toFloat(a)
	if err != nil {
		return false, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return false, err
	}
	return f1 >= f2, nil
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

func isLessEqual(a, b Value) (bool, error) {
	f1, err := toFloat(a)
	if err != nil {
		return false, err
	}
	f2, err := toFloat(b)
	if err != nil {
		return false, err
	}
	return f1 <= f2, nil
}

func concat(a, b Value) Value {
	return toString(a) + toString(b)
}

func andOp(a, b Value) (Value, error) {
	if b1, ok1 := a.(bool); ok1 {
		if b2, ok2 := b.(bool); ok2 {
			return b1 && b2, nil
		}
	}
	v1, err := toInt64(a)
	if err != nil {
		return nil, err
	}
	v2, err := toInt64(b)
	if err != nil {
		return nil, err
	}
	return v1 & v2, nil
}

func orOp(a, b Value) (Value, error) {
	if b1, ok1 := a.(bool); ok1 {
		if b2, ok2 := b.(bool); ok2 {
			return b1 || b2, nil
		}
	}
	v1, err := toInt64(a)
	if err != nil {
		return nil, err
	}
	v2, err := toInt64(b)
	if err != nil {
		return nil, err
	}
	return v1 | v2, nil
}

func xorOp(a, b Value) (Value, error) {
	if b1, ok1 := a.(bool); ok1 {
		if b2, ok2 := b.(bool); ok2 {
			return b1 != b2, nil
		}
	}
	v1, err := toInt64(a)
	if err != nil {
		return nil, err
	}
	v2, err := toInt64(b)
	if err != nil {
		return nil, err
	}
	return v1 ^ v2, nil
}

func eqvOp(a, b Value) (Value, error) {
	if b1, ok1 := a.(bool); ok1 {
		if b2, ok2 := b.(bool); ok2 {
			return b1 == b2, nil
		}
	}
	v1, err := toInt64(a)
	if err != nil {
		return nil, err
	}
	v2, err := toInt64(b)
	if err != nil {
		return nil, err
	}
	return ^(v1 ^ v2), nil
}

func impOp(a, b Value) (Value, error) {
	if b1, ok1 := a.(bool); ok1 {
		if b2, ok2 := b.(bool); ok2 {
			return !b1 || b2, nil
		}
	}
	v1, err := toInt64(a)
	if err != nil {
		return nil, err
	}
	v2, err := toInt64(b)
	if err != nil {
		return nil, err
	}
	return (^v1) | v2, nil
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

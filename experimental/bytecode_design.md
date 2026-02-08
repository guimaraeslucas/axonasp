# Bytecode Design for AxonASP VM

## Architecture
- **Type:** Stack-based Virtual Machine.
- **Values:** `interface{}` (wrapping VBScript types: int64, float64, string, bool, etc.) or a custom `Value` struct.
- **Instructions:** Fixed-width opcode (1 byte) + variable-width operands (usually 2 bytes for indices).

## Data Structures

### VM
- `Stack`: Array of values.
- `SP`: Stack Pointer (index of the top).
- `Globals`: Map or Array of global variables.
- `Frames`: Stack of `CallFrame` for function calls.

### CallFrame
- `Func`: Pointer to the compiled Function/Script.
- `IP`: Instruction Pointer (current bytecode index).
- `BasePointer`: Index in the global stack where this frame's locals begin.

### Bytecode / Chunk
- `Instructions`: `[]byte` (the code).
- `Constants`: `[]Value` (literals: numbers, strings).

## Instruction Set (Draft)

| Opcode | Operands | Description |
| :--- | :--- | :--- |
| `OP_CONSTANT` | `idx (uint16)` | Push `Constants[idx]` to stack. |
| `OP_POP` | - | Pop top value (statement end). |
| `OP_ADD` | - | Pop b, Pop a, Push a + b. |
| `OP_SUB` | - | Pop b, Pop a, Push a - b. |
| `OP_MUL` | - | Pop b, Pop a, Push a * b. |
| `OP_DIV` | - | Pop b, Pop a, Push a / b. |
| `OP_NEG` | - | Pop a, Push -a. |
| `OP_TRUE` | - | Push `True`. |
| `OP_FALSE` | - | Push `False`. |
| `OP_NULL` | - | Push `Null`. |
| `OP_EMPTY` | - | Push `Empty`. |
| `OP_NOT` | - | Pop a, Push Not a. |
| `OP_EQUAL` | - | Pop b, Pop a, Push a = b. |
| `OP_GREATER` | - | Pop b, Pop a, Push a > b. |
| `OP_LESS` | - | Pop b, Pop a, Push a < b. |
| `OP_JUMP` | `offset (uint16)` | IP += offset. |
| `OP_JUMP_IF_FALSE`| `offset (uint16)` | Pop a, if a is False, IP += offset. |
| `OP_GET_GLOBAL` | `idx (uint16)` | Push `Globals[name_from_consts[idx]]`. |
| `OP_SET_GLOBAL` | `idx (uint16)` | Store top to `Globals[name_from_consts[idx]]`. |
| `OP_GET_LOCAL` | `idx (uint8)` | Push value from stack at `BasePointer + idx`. |
| `OP_SET_LOCAL` | `idx (uint8)` | Store top to stack at `BasePointer + idx`. |
| `OP_CALL` | `arg_count (uint8)` | Call function. |
| `OP_RETURN_VALUE` | - | Return with value on top of stack. |
| `OP_RETURN` | - | Return `Empty`. |
| `OP_NEW` | `idx (uint16)` | Create object `Constants[idx]` and push. |

## Handling VBScript Specifics
- **Case Insensitivity:** Compiler normalizes identifiers to lowercase before looking up constant indices.
- **Dynamic Types:** Operations like `ADD` must check types at runtime (int vs float vs string).
- **ByRef:** Locals might need to be boxed if passed ByRef, or the VM needs a mechanism to pass pointers to stack slots.

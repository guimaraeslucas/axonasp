package experimental

type Opcode byte

const (
	// Constants and Literals
	OP_CONSTANT Opcode = iota
	OP_NULL
	OP_EMPTY
	OP_NOTHING
	OP_TRUE
	OP_FALSE

	// Arithmetic
	OP_ADD
	OP_SUB
	OP_MUL
	OP_DIV
	OP_IDIV // Integer division
	OP_MOD
	OP_POW // ^
	OP_NEG // -X

	// String
	OP_CONCAT // &

	// Comparison
	OP_EQUAL
	OP_NOT_EQUAL
	OP_GREATER
	OP_GREATER_EQUAL
	OP_LESS
	OP_LESS_EQUAL
	OP_IS // Object identity

	// Logical
	OP_NOT
	OP_AND
	OP_OR
	OP_XOR
	OP_EQV
	OP_IMP

	// Variables
	OP_GET_GLOBAL
	OP_SET_GLOBAL
	OP_GET_LOCAL
	OP_SET_LOCAL

	// Stack Manipulation
	OP_POP

	// Control Flow
	OP_JUMP
	OP_JUMP_IF_FALSE
	OP_LOOP // Jump back

	// Functions/Procedures
	OP_CALL
	OP_RETURN
	OP_RETURN_VALUE

	// Objects
	OP_NEW
	OP_GET_MEMBER
	OP_SET_MEMBER

	// Optimized Globals/Locals
	OP_GET_GLOBAL_FAST
	OP_SET_GLOBAL_FAST
	OP_INC_LOCAL
	OP_INC_GLOBAL_FAST
)

// Definition helps in debugging and disassembly
type Definition struct {
	Name          string
	OperandWidths []int
}

var definitions = map[Opcode]*Definition{
	OP_CONSTANT:        {"OP_CONSTANT", []int{2}}, // 2-byte index into constants pool
	OP_NULL:            {"OP_NULL", []int{}},
	OP_EMPTY:           {"OP_EMPTY", []int{}},
	OP_NOTHING:         {"OP_NOTHING", []int{}},
	OP_TRUE:            {"OP_TRUE", []int{}},
	OP_FALSE:           {"OP_FALSE", []int{}},
	OP_ADD:             {"OP_ADD", []int{}},
	OP_SUB:             {"OP_SUB", []int{}},
	OP_MUL:             {"OP_MUL", []int{}},
	OP_DIV:             {"OP_DIV", []int{}},
	OP_IDIV:            {"OP_IDIV", []int{}},
	OP_MOD:             {"OP_MOD", []int{}},
	OP_POW:             {"OP_POW", []int{}},
	OP_NEG:             {"OP_NEG", []int{}},
	OP_CONCAT:          {"OP_CONCAT", []int{}},
	OP_EQUAL:           {"OP_EQUAL", []int{}},
	OP_NOT_EQUAL:       {"OP_NOT_EQUAL", []int{}},
	OP_GREATER:         {"OP_GREATER", []int{}},
	OP_GREATER_EQUAL:   {"OP_GREATER_EQUAL", []int{}},
	OP_LESS:            {"OP_LESS", []int{}},
	OP_LESS_EQUAL:      {"OP_LESS_EQUAL", []int{}},
	OP_IS:              {"OP_IS", []int{}},
	OP_NOT:             {"OP_NOT", []int{}},
	OP_AND:             {"OP_AND", []int{}},
	OP_OR:              {"OP_OR", []int{}},
	OP_XOR:             {"OP_XOR", []int{}},
	OP_EQV:             {"OP_EQV", []int{}},
	OP_IMP:             {"OP_IMP", []int{}},
	OP_GET_GLOBAL:      {"OP_GET_GLOBAL", []int{2}}, // 2-byte index constant (name)
	OP_SET_GLOBAL:      {"OP_SET_GLOBAL", []int{2}},
	OP_GET_LOCAL:       {"OP_GET_LOCAL", []int{1}}, // 1-byte stack slot index
	OP_SET_LOCAL:       {"OP_SET_LOCAL", []int{1}},
	OP_POP:             {"OP_POP", []int{}},
	OP_JUMP:            {"OP_JUMP", []int{2}}, // 2-byte offset
	OP_JUMP_IF_FALSE:   {"OP_JUMP_IF_FALSE", []int{2}},
	OP_LOOP:            {"OP_LOOP", []int{2}},
	OP_CALL:            {"OP_CALL", []int{1}}, // 1-byte arg count
	OP_RETURN:          {"OP_RETURN", []int{}},
	OP_RETURN_VALUE:    {"OP_RETURN_VALUE", []int{}},
	OP_NEW:             {"OP_NEW", []int{2}},        // 2-byte index constant (class name)
	OP_GET_MEMBER:      {"OP_GET_MEMBER", []int{2}}, // 2-byte index constant (prop name)
	OP_SET_MEMBER:      {"OP_SET_MEMBER", []int{2}},
	OP_GET_GLOBAL_FAST: {"OP_GET_GLOBAL_FAST", []int{2}},
	OP_SET_GLOBAL_FAST: {"OP_SET_GLOBAL_FAST", []int{2}},
	OP_INC_LOCAL:       {"OP_INC_LOCAL", []int{1}},
	OP_INC_GLOBAL_FAST: {"OP_INC_GLOBAL_FAST", []int{2}},
}

func Lookup(op byte) (*Definition, bool) {
	def, ok := definitions[Opcode(op)]
	return def, ok
}

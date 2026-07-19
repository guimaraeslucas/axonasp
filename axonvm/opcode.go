/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package axonvm

type OpCode byte

const (
	OpHalt OpCode = iota

	// Data Movement
	OpConstant
	OpPop
	OpGetGlobal
	OpSetGlobal
	OpGetLocal // [OpCode, OffsetHigh, OffsetLow]
	OpSetLocal // [OpCode, OffsetHigh, OffsetLow]
	OpGetClassMember
	OpSetClassMember
	OpEraseGlobal
	OpEraseLocal
	OpEraseClassMember
	OpSet

	// Arithmetic
	OpAdd
	OpSub
	OpMul
	OpDiv
	OpMod
	OpPow
	OpIAdd
	OpISub
	OpIMul
	OpIDiv
	// OpIRightShift performs arithmetic right shift on two integer operands.
	// Stack before: [..., left, right]
	// Stack after:  [..., left >> right]
	OpIRightShift

	// Dedicated zero-allocation unary math opcodes for hot VBScript builtins.
	OpMathSin
	OpMathCos
	OpMathTan
	OpMathAtn
	OpMathSqr
	OpMathAbs
	OpMathExp
	OpMathLog
	OpMathRound
	OpMathInt

	// String & Logical
	OpConcat
	OpEq
	OpIsRef
	OpIsNotRef
	OpNeq
	OpLt
	OpGt
	OpLte
	OpGte
	OpAnd
	OpOr
	OpXor
	OpNot
	OpNeg
	OpEqv
	OpImp

	// Control Flow
	OpJump        // [OpCode, Target3, Target2, Target1, Target0] - absolute bytecode target
	OpJumpIfFalse // [OpCode, Target3, Target2, Target1, Target0] - absolute bytecode target
	OpJumpIfTrue  // [OpCode, Target3, Target2, Target1, Target0] - absolute bytecode target

	// Output
	OpWrite
	OpWriteStatic
	// OpWriteN is the zero-allocation Response.Write optimisation emitted by the
	// compiler when the argument to Response.Write is a top-level & concatenation
	// chain.  Instead of building intermediate string Values with OpConcat and then
	// passing a single concatenated Value to OpCallMember(Write,1), the compiler
	// pushes every individual operand onto the stack and emits OpWriteN(N).  The VM
	// pops N values, converts each to its string representation using the per-VM
	// stringWorkBuffer, and writes all parts in a single Response.Write call — one
	// mutex acquisition, no intermediate string allocations.
	// [OpCode, CountHigh, CountLow]
	OpWriteN

	// Configuration
	OpSetOption           // [OpCode, OptionID, ValueID]
	OpSetDirective        // [OpCode, NameConstIdxHigh, NameConstIdxLow, ValueConstIdxHigh, ValueConstIdxLow]
	OpRegisterClass       // [OpCode, ClassNameConstIdxHigh, ClassNameConstIdxLow]
	OpRegisterClassMethod // [OpCode, ClassNameConstIdxHigh, ClassNameConstIdxLow, MethodNameConstIdxHigh, MethodNameConstIdxLow, UserSubConstIdxHigh, UserSubConstIdxLow, IsPublicHigh, IsPublicLow]
	OpRegisterClassField
	OpRegisterClassPropertyGet
	// OpInitClassArrayField registers fixed-size array dimensions for a class member field.
	// Dims are popped from the stack (dim0..dimN-1 pushed in order), then stored in the
	// class metadata so every new instance gets a pre-allocated array.
	// [OpCode, ClassNameConstIdxHigh, ClassNameConstIdxLow, FieldNameConstIdxHigh, FieldNameConstIdxLow, DimCountHigh, DimCountLow]
	// Stack before: [..., dim0, dim1, ..., dimN-1]  (dimN-1 at TOS)
	OpInitClassArrayField
	OpRegisterClassPropertyLet
	OpRegisterClassPropertySet

	// Error Handling & Debug
	OpOnErrorResumeNext // Enables error suppression
	OpOnErrorGoto0      // Disables error suppression
	OpLine              // [OpCode, LineHigh, LineLow, ColHigh, ColLow] - Sets current line/column for debugging
	OpLabel             // [OpCode, LabelIDHigh, LabelIDLow] - Marker (no-op)
	OpGotoLabel         // [OpCode, Target3, Target2, Target1, Target0] - Jump to absolute bytecode target

	// Member Access & Calls
	OpMemberGet
	OpMemberSet
	OpMemberSetSet
	OpMe // Pushes the current class instance (activeClassObjectID) as VTObject
	// OpCallMember now carries one inline 32-bit cache slot populated by the VM
	// on first execution for monomorphic call-site fast paths.
	// [OpCode, ConstMemberIdxHigh, ConstMemberIdxLow, ArgCountHigh, ArgCountLow, CacheB3, CacheB2, CacheB1, CacheB0]
	OpCallMember
	OpCallBuiltin // [OpCode, RegistryIdxHigh, RegistryIdxLow, ArgCountHigh, ArgCountLow]
	OpCall
	OpNewClass // [OpCode, ClassNameConstIdxHigh, ClassNameConstIdxLow]
	OpArraySet // [OpCode, ArgCountHigh, ArgCountLow] ; stack: [..., targetArray, idx1..idxN, value]
	OpRet
	OpSwap // Swaps the top two elements on the stack.
)

const (
	// OpArgGlobalRef pushes a VTArgRef for a global slot. Used at call sites so that
	// ByRef parameters can write back to the caller's global variable on return.
	// [OpCode, IdxHigh, IdxLow]
	OpArgGlobalRef OpCode = iota + OpSwap + 1
	// OpArgLocalRef pushes a VTArgRef for a local slot. Used at call sites so that
	// ByRef parameters can write back to the caller's local variable on return.
	// [OpCode, IdxHigh, IdxLow]
	OpArgLocalRef

	// OpArgClassMemberRef pushes a VTArgRef for the current class member slot.
	// Used at call sites so that ByRef parameters can write back to the caller's
	// class field on return.
	// [OpCode, MemberConstIdxHigh, MemberConstIdxLow]
	OpArgClassMemberRef

	// OpWithEnter pops the TOS object and stores it on the VM with-object stack.
	// Emitted once at the start of a With block (after evaluating the subject).
	// [OpCode]
	OpWithEnter

	// OpWithLeave removes the innermost entry from the VM with-object stack.
	// Emitted once at the end of a With block (End With).
	// [OpCode]
	OpWithLeave

	// OpWithLoad pushes a copy of the innermost with-object onto the data stack.
	// Emitted before every '.Member' access or '.Prop = value' inside a With block.
	// [OpCode]
	OpWithLoad

	// OpLetGlobal writes TOS to a global slot for plain VBScript "name = value"
	// (not Set). Global variables are mutable Variant slots and are overwritten.
	// [OpCode, IdxHigh, IdxLow]
	OpLetGlobal

	// OpLetLocal writes TOS to a local frame slot for plain VBScript "name = value"
	// (not Set). Local variables are mutable Variant slots and are overwritten.
	// [OpCode, OffsetHigh, OffsetLow]
	OpLetLocal

	// OpLetClassMember writes TOS to a class member field; default Property Let dispatch
	// applies when the current member value is a VTObject.
	// [OpCode, MemberConstIdxHigh, MemberConstIdxLow]
	OpLetClassMember

	// OpCoerceToValue pops TOS and, if it is a VTObject with a default Property Get
	// (zero explicit arguments), starts the getter call and pushes its result.
	// For VTObject with Num == 0 (Nothing) or any non-object, re-pushes value unchanged.
	// Used for implicit object-to-value coercion in arithmetic, concatenation, and output.
	// [OpCode]
	OpCoerceToValue

	// JScript opcodes (isolated execution path)
	OpJSDeclareName            // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSGetName                // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSSetName                // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSImport                 // [OpCode, ModuleConstIdxHigh, ModuleConstIdxLow, SpecCountHigh, SpecCountLow, ImportedIdxHigh, ImportedIdxLow, LocalIdxHigh, LocalIdxLow, ...]
	OpJSExport                 // [OpCode, LocalNameConstIdxHigh, LocalNameConstIdxLow, ExportNameConstIdxHigh, ExportNameConstIdxLow]
	OpJSGetLocal               // [OpCode, OffsetHigh, OffsetLow]
	OpJSSetLocal               // [OpCode, OffsetHigh, OffsetLow]
	OpJSIncLocal               // [OpCode, OffsetHigh, OffsetLow]
	OpJSDecLocal               // [OpCode, OffsetHigh, OffsetLow]
	OpJSMemberGet              // [OpCode, NameConstIdxHigh, NameConstIdxLow, ICNodeIDHigh, ICNodeIDLow]
	OpJSMemberSet              // [OpCode, NameConstIdxHigh, NameConstIdxLow, ICNodeIDHigh, ICNodeIDLow]
	OpJSCall                   // [OpCode, ArgCountHigh, ArgCountLow]
	OpJSCallMember             // [OpCode, NameConstIdxHigh, NameConstIdxLow, ArgCountHigh, ArgCountLow]
	OpJSTailCall               // [OpCode, ArgCountHigh, ArgCountLow]
	OpJSTailCallMember         // [OpCode, NameConstIdxHigh, NameConstIdxLow, ArgCountHigh, ArgCountLow]
	OpJSCallComputedMember     // [OpCode, ArgCountHigh, ArgCountLow]
	OpJSTailCallComputedMember // [OpCode, ArgCountHigh, ArgCountLow]
	OpJSCreateClosure          // [OpCode, FunctionTemplateConstIdxHigh, FunctionTemplateConstIdxLow]
	OpJSAdd                    // [OpCode]
	OpJSStrictEq               // [OpCode]
	OpJSStrictNeq              // [OpCode]
	OpJSTryEnter               // [OpCode, CatchTarget3, CatchTarget2, CatchTarget1, CatchTarget0]
	OpJSTryLeave               // [OpCode]
	OpJSThrow                  // [OpCode]
	OpJSNewObject              // [OpCode]
	OpJSNewArray               // [OpCode, CountHigh, CountLow]
	OpJSTypeof                 // [OpCode]
	OpJSInstanceOf             // [OpCode]
	OpJSIn                     // [OpCode]
	OpJSDelete                 // [OpCode] - computed delete; stack: [..., obj, key]
	OpJSReturn                 // [OpCode]
	OpJSLoadUndefined          // [OpCode]
	OpJSLoadThis               // [OpCode]
	OpJSSetThis                // [OpCode] ; pops RHS, assigns to 'this'
	OpJSDup                    // [OpCode]
	OpJSRequireObject          // [OpCode] ; throws TypeError if TOS is null or undefined
	OpJSGetIterator            // [OpCode] ; pops RHS, pushes iterator (RHS[Symbol.iterator]())
	OpJSIteratorNext           // [OpCode] ; peek iterator, call .next(), push value (or undefined if done)
	OpJSCollectRest            // [OpCode] ; pops iterator, pushes array of remaining values
	OpJSObjectRest             // [OpCode, StaticCountH, StaticCountL, ..., DynamicCountH, DynamicCountL] ; pops obj and DynamicCount keys, pushes rest object
	OpJSPop                    // [OpCode]
	OpJSRot                    // [OpCode, Count] ; rotates top N elements of the stack
	OpJSJump                   // [OpCode, Target3, Target2, Target1, Target0]
	OpJSJumpIfFalse            // [OpCode, Target3, Target2, Target1, Target0]
	OpJSJumpIfTrue             // [OpCode, Target3, Target2, Target1, Target0]
	OpJSJumpIfNotUndefined     // [OpCode, Target3, Target2, Target1, Target0]
	OpJSLoadCatchError         // [OpCode]

	OpJSStoreCatchError         // [OpCode]
	OpJSBreak                   // [OpCode] - Break from loop (jump target managed by compiler)
	OpJSContinue                // [OpCode] - Continue to next iteration (jump target managed by compiler)
	OpJSForIn                   // [OpCode, EnumVarIdxHigh, EnumVarIdxLow, LoopStartTarget3, LoopStartTarget2, LoopStartTarget1, LoopStartTarget0] - for...in setup
	OpJSForInCleanup            // [OpCode, ForInPos3, ForInPos2, ForInPos1, ForInPos0] - clears for...in enumerator state and source value
	OpJSSwitch                  // [OpCode] - switch statement (value already on stack)
	OpJSCase                    // [OpCode, Target3, Target2, Target1, Target0] - case label with jump target
	OpJSDefault                 // [OpCode, Target3, Target2, Target1, Target0] - default label with jump target
	OpJSSubtract                // [OpCode] - JScript binary subtraction (for type coercion compatibility)
	OpJSMultiply                // [OpCode] - JScript binary multiplication
	OpJSDivide                  // [OpCode] - JScript binary division
	OpJSModulo                  // [OpCode] - JScript binary modulo
	OpJSNegate                  // [OpCode] - JScript unary negation
	OpJSUnaryPlus               // [OpCode] - JScript unary + (ToNumber coercion)
	OpJSBitwiseAnd              // [OpCode] - JScript bitwise AND
	OpJSBitwiseOr               // [OpCode] - JScript bitwise OR
	OpJSBitwiseXor              // [OpCode] - JScript bitwise XOR
	OpJSBitwiseNot              // [OpCode] - JScript bitwise NOT
	OpJSLeftShift               // [OpCode] - JScript left shift
	OpJSRightShift              // [OpCode] - JScript right shift
	OpJSUnsignedRightShift      // [OpCode] - JScript unsigned right shift
	OpJSLess                    // [OpCode] - JScript less than
	OpJSGreater                 // [OpCode] - JScript greater than
	OpJSLessEqual               // [OpCode] - JScript less than or equal
	OpJSGreaterEqual            // [OpCode] - JScript greater than or equal
	OpJSLooseEqual              // [OpCode] - JScript == (loose equality)
	OpJSLooseNotEqual           // [OpCode] - JScript != (loose inequality)
	OpJSLogicalAnd              // [OpCode] - JScript && (logical AND)
	OpJSLogicalOr               // [OpCode] - JScript || (logical OR)
	OpJSLogicalNot              // [OpCode] - JScript ! (logical NOT)
	OpJSNew                     // [OpCode, ArgCountHigh, ArgCountLow] - new constructor call
	OpJSMemberDelete            // [OpCode, NameConstIdxHigh, NameConstIdxLow] - delete property
	OpJSIndexGet                // [OpCode] - array/object index access (bracket notation)
	OpJSIndexSet                // [OpCode] - array/object index assignment
	OpJSComma                   // [OpCode] - comma operator
	OpJSPostIncrement           // [OpCode, NameConstIdxHigh, NameConstIdxLow] - post-increment (var++)
	OpJSPostDecrement           // [OpCode, NameConstIdxHigh, NameConstIdxLow] - post-decrement (var--)
	OpJSPreIncrement            // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pre-increment (++var)
	OpJSPreDecrement            // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pre-decrement (--var)
	OpJSAddAssign               // [OpCode, NameConstIdxHigh, NameConstIdxLow] - +=
	OpJSSubtractAssign          // [OpCode, NameConstIdxHigh, NameConstIdxLow] - -=
	OpJSMultiplyAssign          // [OpCode, NameConstIdxHigh, NameConstIdxLow] - *=
	OpJSDivideAssign            // [OpCode, NameConstIdxHigh, NameConstIdxLow] - /=
	OpJSModuloAssign            // [OpCode, NameConstIdxHigh, NameConstIdxLow] - %=
	OpJSMemberIndexGet          // [OpCode, NameConstIdxHigh, NameConstIdxLow] - member[index] get
	OpJSMemberIndexSet          // [OpCode, NameConstIdxHigh, NameConstIdxLow] - member[index] set
	OpJSPostMemberIncrement     // [OpCode, NameConstIdxHigh, NameConstIdxLow] - post-increment (obj.prop++)
	OpJSPostMemberDecrement     // [OpCode, NameConstIdxHigh, NameConstIdxLow] - post-decrement (obj.prop--)
	OpJSPreMemberIncrement      // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pre-increment (++obj.prop)
	OpJSPreMemberDecrement      // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pre-decrement (--obj.prop)
	OpJSPostIndexIncrement      // [OpCode] - post-increment (obj[idx]++)
	OpJSPostIndexDecrement      // [OpCode] - post-decrement (obj[idx]--)
	OpJSPreIndexIncrement       // [OpCode] - pre-increment (++obj[idx])
	OpJSPreIndexDecrement       // [OpCode] - pre-decrement (--obj[idx])
	OpJSExponent                // [OpCode] - JScript exponentiation (**)
	OpJSCoalesce                // [OpCode] - JScript ??
	OpJSJumpIfNullish           // [OpCode, Target3, Target2, Target1, Target0] - jump if null or undefined
	OpJSJumpIfNotNullish        // [OpCode, Target3, Target2, Target1, Target0] - jump if not null and not undefined
	OpJSExponentAssign          // [OpCode, NameConstIdxHigh, NameConstIdxLow] - **=
	OpJSLogicalAndAssign        // [OpCode, NameConstIdxHigh, NameConstIdxLow] - &&=
	OpJSLogicalOrAssign         // [OpCode, NameConstIdxHigh, NameConstIdxLow] - ||=
	OpJSCoalesceAssign          // [OpCode, NameConstIdxHigh, NameConstIdxLow] - ??=
	OpJSDefineProperty          // [OpCode, NameConstIdxHigh, NameConstIdxLow, KindHigh, KindLow] ; stack: [..., target, value]
	OpJSSetProto                // [OpCode] ; stack: [..., target, proto]
	OpJSClassInherit            // [OpCode] ; stack: [..., superclass, subclass]
	OpJSSuperCall               // [OpCode, ArgCountHigh, ArgCountLow] ; stack: [..., arg1, arg2, ...]
	OpJSSuperMemberGet          // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSSuperMemberSet          // [OpCode, NameConstIdxHigh, NameConstIdxLow] ; stack: [..., value]
	OpJSSuperCallMember         // [OpCode, NameConstIdxHigh, NameConstIdxLow, ArgCountHigh, ArgCountLow] ; stack: [..., arg1, arg2, ...]
	OpJSSuperIndexGet           // [OpCode] ; stack: [..., index]
	OpJSSuperIndexSet           // [OpCode] ; stack: [..., value, index]
	OpJSSuperCallComputedMember // [OpCode, ArgCountHigh, ArgCountLow] ; stack: [..., index, arg1, arg2, ...]

	OpJSYield         // [OpCode]
	OpJSYieldDelegate // [OpCode]
	OpJSAwait         // [OpCode]

	// Strict Mode & Block Scoping
	OpJSStrictModeEnter  // [OpCode] - enter strict mode scope
	OpJSStrictModeExit   // [OpCode] - exit strict mode scope
	OpJSBlockScopeEnter  // [OpCode] - enter new block scope for let/const
	OpJSBlockScopeExit   // [OpCode] - exit block scope
	OpJSLetDeclare       // [OpCode, NameConstIdxHigh, NameConstIdxLow] - clear TDZ for let variable
	OpJSTDZRegisterLet   // [OpCode, NameConstIdxHigh, NameConstIdxLow] - register let variable in TDZ
	OpJSTDZRegisterConst // [OpCode, NameConstIdxHigh, NameConstIdxLow] - register const variable in TDZ
	OpJSConstInitialize  // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pop value, set const, clear TDZ
	OpJSForIterEnter     // [OpCode, NumVarsHigh, NumVarsLow, NameIdx1Hi, NameIdx1Lo, ...] - per-iteration env enter
	OpJSForIterExit      // [OpCode, NumVarsHigh, NumVarsLow, NameIdx1Hi, NameIdx1Lo, ...] - per-iteration env exit+writeback

	// OpJSLoadNewTarget pushes the newTarget value for the current function call onto the stack.
	// If the current call is not a constructor call, it pushes undefined.
	// [OpCode]
	OpJSLoadNewTarget
)

const (
	// OpIncLocalInt increments one local numeric slot in place.
	// [OpCode, OffsetHigh, OffsetLow]
	OpIncLocalInt OpCode = iota + OpJSLoadNewTarget + 1
	// OpDecLocalInt decrements one local numeric slot in place.
	// [OpCode, OffsetHigh, OffsetLow]
	OpDecLocalInt
	// OpIncGlobalInt increments one global numeric slot in place.
	// [OpCode, IdxHigh, IdxLow]
	OpIncGlobalInt
	// OpDecGlobalInt decrements one global numeric slot in place.
	// [OpCode, IdxHigh, IdxLow]
	OpDecGlobalInt
	// OpJSIncLocalInt increments one JScript identifier in place without stack push.
	// [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSIncLocalInt
	// OpJSDecLocalInt decrements one JScript identifier in place without stack push.
	// [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSDecLocalInt
	// OpJSForFastIntEnter validates once (pre-loop) that both local slots used by
	// fast loop paths contain VTInteger values. It also performs the initial
	// loop condition test (counter < limit) and jumps to the exit target if false.
	// [OpCode, CounterSlotHigh, CounterSlotLow, LimitSlotHigh, LimitSlotLow, Target3, Target2, Target1, Target0]
	OpJSForFastIntEnter

	// OpJSForFastInt is a fused super-instruction for JScript `for (let i = 0; i < N; i++)`
	// loops backed by local slots. It atomically increments the counter slot, compares the
	// updated value against the local limit slot, and performs a contiguous in-dispatch
	// relative back-jump when the loop is still in range.
	//
	// Stack: unchanged — no values pushed or popped.
	// Format: [OpCode(1), counterSlotH(1), counterSlotL(1), limitSlotH(1), limitSlotL(1),
	//          jumpOffsetB3(1), jumpOffsetB2(1), jumpOffsetB1(1), jumpOffsetB0(1)]
	// Total: 9 bytes (1 opcode + 8 operand bytes).
	OpJSForFastInt
	// OpJSForIterEnterFast enters a lexical for-loop iteration without allocating a
	// child environment frame. It is only emitted when closure capture analysis proves
	// loop bindings do not escape.
	// [OpCode]
	OpJSForIterEnterFast
	// OpJSForIterExitFast exits a lexical for-loop iteration for the non-capturing fast path.
	// [OpCode]
	OpJSForIterExitFast

	// OpJSRootFrameEnter reserves one contiguous root-frame local area for top-level
	// JScript local-slot lowered identifiers. It initializes each slot with undefined.
	// [OpCode, LocalCountHigh, LocalCountLow]
	OpJSRootFrameEnter

	// OpJSRootFrameLeave releases one top-level JScript root-frame local area and
	// restores the stack pointer to the pre-reservation depth.
	// [OpCode, LocalCountHigh, LocalCountLow]
	OpJSRootFrameLeave

	// OpNop is a no-operation placeholder emitted by the peephole optimizer to
	// fill bytes that were made redundant by constant folding.  The VM advances
	// the instruction pointer past it at near-zero cost; jump offsets are
	// preserved because the bytecode array is never shrunk.
	// [OpCode]  (0 operand bytes)
	OpNop

	// OpForNextFastInt is a fused super-instruction for integer For...Next loops with ±1 step.
	// It atomically increments or decrements a local counter slot, compares it to a local
	// limit slot, and jumps directly back to the loop body when still in range — eliminating
	// the per-iteration direction check and the multi-opcode condition sequence used by the
	// generic path.  The pre-loop bounds check is emitted separately (OpLte/OpGte + OpJumpIfFalse)
	// to guard against zero-iteration ranges; this opcode only runs once the loop is entered.
	//
	// Stack: unchanged — no values pushed or popped.
	// Format: [OpCode(1), varLocalIdxH(1), varLocalIdxL(1), endLocalIdxH(1), endLocalIdxL(1),
	//          stepSign(1), bodyTargetB3(1), bodyTargetB2(1), bodyTargetB1(1), bodyTargetB0(1)]
	// stepSign: 0x01 = increment (+1), 0xFF = decrement (-1).
	// Total: 10 bytes (1 opcode + 9 operand bytes).
	OpForNextFastInt

	// OpForNextFastGlobalInt is a fused super-instruction for global For...Next loops with ±1 step.
	// It atomically increments or decrements a global counter slot, compares it to a global limit
	// slot, and jumps directly back to the loop body when still in range.
	//
	// Stack: unchanged — no values pushed or popped.
	// Format: [OpCode(1), varGlobalIdxH(1), varGlobalIdxL(1), endGlobalIdxH(1), endGlobalIdxL(1),
	//          stepSign(1), bodyTargetB3(1), bodyTargetB2(1), bodyTargetB1(1), bodyTargetB0(1)]
	// stepSign: 0x01 = increment (+1), 0xFF = decrement (-1).
	// Total: 10 bytes (1 opcode + 9 operand bytes).
	OpForNextFastGlobalInt

	// OpJSJumpIfLessFast is a fused test-and-branch super-instruction for JScript numeric
	// for-loops that use the pattern `identifier < numericLiteral` as their test condition.
	// It reads the named variable from the JS environment, compares it numerically to a
	// constant limit, and jumps to the exit target when the variable is NOT less than the
	// limit (loop condition false).  When the variable IS less the instruction falls through
	// to the loop body — zero stack impact on the hot iteration path.
	//
	// Stack: unchanged — no values pushed or popped.
	// Format: [OpCode(1), nameConstIdxH(1), nameConstIdxL(1), limitConstIdxH(1), limitConstIdxL(1),
	//          exitTargetB3(1), exitTargetB2(1), exitTargetB1(1), exitTargetB0(1)]
	// Total: 9 bytes (1 opcode + 8 operand bytes).
	OpJSJumpIfLessFast

	// OpJSComputedPropertySet sets a computed property key on an object literal being built.
	// It pops the key (top), then the value (next), then the target object (next),
	// and calls jsIndexSet(obj, key, value). The outer object reference below on the stack
	// is left intact so subsequent properties can continue to reference the object.
	//
	// Stack before: ..., obj (outer), obj (dup), value, key
	// Stack after:  ..., obj (outer)
	// Format: [OpCode(1)]  (0 operand bytes).
	OpJSComputedPropertySet

	// OpJSForOf iterates over the values of an iterable (Array, String, Map, Set, etc.).
	// On first encounter the source value is popped from the stack and its values are
	// collected into an enumerator keyed by the opcode position. On each pass the next
	// value is assigned to the named variable; when exhausted the opcode jumps to exitTarget
	// and removes the enumerator.
	//
	// Stack: source value consumed on first pass; no further stack impact per iteration.
	// Format: [OpCode(1), nameConstIdxH(1), nameConstIdxL(1),
	//          exitTargetB3(1), exitTargetB2(1), exitTargetB1(1), exitTargetB0(1)]
	// Total: 7 bytes (1 opcode + 6 operand bytes).
	OpJSForOf

	// OpJSForOfCleanup removes the for-of enumerator created at the given bytecode position.
	// Emitted after a break out of the loop so the enumerator does not leak.
	//
	// Format: [OpCode(1), forOfPosB3(1), forOfPosB2(1), forOfPosB1(1), forOfPosB0(1)]
	// Total: 5 bytes (1 opcode + 4 operand bytes).
	OpJSForOfCleanup
	// OpJSExportAll exports all members from a module source into the current module.
	// [OpCode, ModuleConstIdxHigh, ModuleConstIdxLow]
	OpJSExportAll

	// Fused Branching Super-Instructions
	// These instructions combine a comparison and a conditional jump to reduce dispatch overhead.
	// Format: [OpCode(1), Target3(1), Target2(1), Target1(1), Target0(1)] - absolute bytecode target
	// Total: 5 bytes.

	// OpJumpIfNotEq fuses OpEq + OpJumpIfFalse.
	OpJumpIfNotEq
	// OpJumpIfEq fuses OpNeq + OpJumpIfFalse (Jump if Equal).
	OpJumpIfEq
	// OpJumpIfNotLt fuses OpLt + OpJumpIfFalse (Jump if Greater or Equal).
	OpJumpIfNotLt
	// OpJumpIfLte fuses OpGt + OpJumpIfFalse (Jump if Less or Equal).
	OpJumpIfLte
	// OpJumpIfNotIs fuses OpIsRef + OpJumpIfFalse.
	OpJumpIfNotIs

	// JScript Fused Branching Super-Instructions
	// Format: [OpCode(1), Target3(1), Target2(1), Target1(1), Target0(1)]
	OpJSJumpIfLooseNotEq
	OpJSJumpIfLooseEq
	OpJSJumpIfStrictNotEq
	OpJSJumpIfStrictEq
	OpJSJumpIfNotLess
	OpJSJumpIfLessEqual

	// OpExtPrefix is the escape opcode for extended instruction space.
	//
	// Why this exists:
	// - OpCode is byte-based, so the primary opcode space is capped at 256 values.
	// - The VM hit that hard cap and could not add more primary opcodes safely.
	//
	// How it works:
	// - The compiler emits [OpExtPrefix, ExtOpCode, operands...].
	// - The VM decodes OpExtPrefix and dispatches the second byte through one
	//   dedicated extended-opcode switch.
	// - Existing one-byte primary opcodes stay unchanged and keep hot-path speed.
	//
	// Current encoding contract:
	// - Every extended opcode currently carries one uint16 operand.
	// - Encoded size is 4 bytes total: prefix(1) + ext(1) + operand(2).
	// - opcodeOperandSize(OpExtPrefix) is therefore 3 bytes.
	OpExtPrefix
)

// ExtOpCode is the second-byte opcode namespace reached via OpExtPrefix.
type ExtOpCode byte

const (
	// ExtOpInitRecord creates one zero-initialized UDT record instance from a compiled
	// record declaration index and pushes it onto the stack.
	// [OpExtPrefix, ExtOpInitRecord, DefIdxHigh, DefIdxLow]
	ExtOpInitRecord ExtOpCode = iota
	// ExtOpGetRecordMember loads one UDT member value by fixed member index.
	// [OpExtPrefix, ExtOpGetRecordMember, MemberIdxHigh, MemberIdxLow]
	ExtOpGetRecordMember
	// ExtOpSetRecordMember writes one UDT member value by fixed member index.
	// [OpExtPrefix, ExtOpSetRecordMember, MemberIdxHigh, MemberIdxLow]
	ExtOpSetRecordMember

	// ExtOpAxonASP pushes the string "G3pix AxonASP VBScript Engine" onto the stack.
	// [OpExtPrefix, ExtOpAxonASP] (0 operand bytes)
	ExtOpAxonASP

	// JScript Math Optimizations
	// [OpExtPrefix, ExtOpJSMath*, (optional operands)]
	ExtOpJSMathSin
	ExtOpJSMathCos
	ExtOpJSMathTan
	ExtOpJSMathAbs
	ExtOpJSMathFloor
	ExtOpJSMathCeil
	ExtOpJSMathRound
	ExtOpJSMathSqrt
	ExtOpJSMathMin
	ExtOpJSMathMax

	// Phase 4: Events
	// ExtOpRegisterClassEvent registers one event name for a class.
	// [OpExtPrefix, ExtOpRegisterClassEvent, ClassNameIdxHigh, ClassNameIdxLow, EventNameIdxHigh, EventNameIdxLow]
	ExtOpRegisterClassEvent
	// ExtOpRaiseEvent raises an event in the current class instance.
	// [OpExtPrefix, ExtOpRaiseEvent, EventNameIdxHigh, EventNameIdxLow, ArgCountHigh, ArgCountLow]
	ExtOpRaiseEvent
	// ExtOpWithEventsRegister registers a WithEvents binding for a variable in a class.
	// [OpExtPrefix, ExtOpWithEventsRegister, ClassNameIdxHigh, ClassNameIdxLow, VarNameIdxHigh, VarNameIdxLow]
	ExtOpWithEventsRegister
	// ExtOpRegisterClassInterface registers an interface implemented by a class.
	// [OpExtPrefix, ExtOpRegisterClassInterface, ClassNameIdxHigh, ClassNameIdxLow, InterfaceNameIdxHigh, InterfaceNameIdxLow]
	ExtOpRegisterClassInterface

	// Fused Variable Load + Branch
	ExtOpJumpLocalIfFalse  // [OpExtPrefix, ExtOpJumpLocalIfFalse, OffsetHigh, OffsetLow, Target3, Target2, Target1, Target0]
	ExtOpJumpGlobalIfFalse // [OpExtPrefix, ExtOpJumpGlobalIfFalse, IdxHigh, IdxLow, Target3, Target2, Target1, Target0]
	ExtOpJSJumpNameIfFalse // [OpExtPrefix, ExtOpJSJumpNameIfFalse, NameConstIdxHigh, NameConstIdxLow, Target3, Target2, Target1, Target0]

	// Fused Variable Load + Const + Binary Op + Store
	ExtOpAddLocalConst    // [OpExtPrefix, ExtOpAddLocalConst, OffsetHigh, OffsetLow, ConstIdxHigh, ConstIdxLow]
	ExtOpSubGlobalConst   // [OpExtPrefix, ExtOpSubGlobalConst, IdxHigh, IdxLow, ConstIdxHigh, ConstIdxLow]
	ExtOpConcatLocalConst // [OpExtPrefix, ExtOpConcatLocalConst, OffsetHigh, OffsetLow, ConstIdxHigh, ConstIdxLow]

	// Constant Pooling
	ExtOpConstant2 // [OpExtPrefix, ExtOpConstant2, Const1High, Const1Low, Const2High, Const2Low]
	ExtOpConstant3 // [OpExtPrefix, ExtOpConstant3, Const1High, Const1Low, Const2High, Const2Low, Const3High, Const3Low]
	ExtOpConstant4 // [OpExtPrefix, ExtOpConstant4, Const1High, Const1Low, Const2High, Const2Low, Const3High, Const3Low, Const4High, Const4Low]

	// Phase 3: File I/O
	ExtOpFileOpen      // [OpExtPrefix, ExtOpFileOpen] (path, mode, num on stack)
	ExtOpFileClose     // [OpExtPrefix, ExtOpFileClose] (num on stack)
	ExtOpFilePrint     // [OpExtPrefix, ExtOpFilePrint, ArgCountH, ArgCountL] (num, arg1..argN on stack)
	ExtOpFileWrite     // [OpExtPrefix, ExtOpFileWrite, ArgCountH, ArgCountL] (num, arg1..argN on stack)
	ExtOpFileLineInput // [OpExtPrefix, ExtOpFileLineInput] (num on stack)
	ExtOpFilePut       // [OpExtPrefix, ExtOpFilePut] (num, pos, value on stack)
	ExtOpFileGet       // [OpExtPrefix, ExtOpFileGet] (num, pos on stack)
	ExtOpFileFreeFile  // [OpExtPrefix, ExtOpFileFreeFile] (pushes result to stack)

	// ExtOpJSReThrow re-throws a pending exception from jsErrStack after a finally
	// block has been executed.  If no exception is pending (normal path), it is a
	// no-op and execution continues at the next instruction.
	// [OpExtPrefix, ExtOpJSReThrow] (0 operand bytes beyond ext opcode)
	ExtOpJSReThrow

	// ExtOpCloneRecord clones the UDT record on top of the stack.
	// [OpExtPrefix, ExtOpCloneRecord] (0 operand bytes)
	ExtOpCloneRecord

	// ExtOpShiftLeft performs a logical left shift (<<). Both operands on stack.
	// Stack before: [..., left, right]
	// Stack after:  [..., left << right]
	// [OpExtPrefix, ExtOpShiftLeft] (0 operand bytes beyond ext opcode)
	ExtOpShiftLeft

	// ExtOpShiftRight performs a logical right shift (>>). Both operands on stack.
	// Stack before: [..., left, right]
	// Stack after:  [..., left >> right]
	// [OpExtPrefix, ExtOpShiftRight] (0 operand bytes beyond ext opcode)
	ExtOpShiftRight
)

func (op OpCode) String() string {
	switch op {
	case OpHalt:
		return "OpHalt"
	case OpConstant:
		return "OpConstant"
	case OpPop:
		return "OpPop"
	case OpGetGlobal:
		return "OpGetGlobal"
	case OpSetGlobal:
		return "OpSetGlobal"
	case OpGetClassMember:
		return "OpGetClassMember"
	case OpSetClassMember:
		return "OpSetClassMember"
	case OpEraseGlobal:
		return "OpEraseGlobal"
	case OpEraseLocal:
		return "OpEraseLocal"
	case OpEraseClassMember:
		return "OpEraseClassMember"
	case OpAdd:
		return "OpAdd"
	case OpSub:
		return "OpSub"
	case OpMul:
		return "OpMul"
	case OpDiv:
		return "OpDiv"
	case OpMod:
		return "OpMod"
	case OpPow:
		return "OpPow"
	case OpIAdd:
		return "OpIAdd"
	case OpISub:
		return "OpISub"
	case OpIMul:
		return "OpIMul"
	case OpIDiv:
		return "OpIDiv"
	case OpIRightShift:
		return "OpIRightShift"
	case OpMathSin:
		return "OpMathSin"
	case OpMathCos:
		return "OpMathCos"
	case OpMathTan:
		return "OpMathTan"
	case OpMathAtn:
		return "OpMathAtn"
	case OpMathSqr:
		return "OpMathSqr"
	case OpMathAbs:
		return "OpMathAbs"
	case OpMathExp:
		return "OpMathExp"
	case OpMathLog:
		return "OpMathLog"
	case OpMathRound:
		return "OpMathRound"
	case OpMathInt:
		return "OpMathInt"
	case OpConcat:
		return "OpConcat"
	case OpEq:
		return "OpEq"
	case OpIsRef:
		return "OpIsRef"
	case OpIsNotRef:
		return "OpIsNotRef"
	case OpLt:
		return "OpLt"
	case OpJump:
		return "OpJump"
	case OpJumpIfFalse:
		return "OpJumpIfFalse"
	case OpOnErrorResumeNext:
		return "OpOnErrorResumeNext"
	case OpOnErrorGoto0:
		return "OpOnErrorGoto0"
	case OpLine:
		return "OpLine"
	case OpWrite:
		return "OpWrite"
	case OpWriteStatic:
		return "OpWriteStatic"
	case OpWriteN:
		return "OpWriteN"
	case OpSetDirective:
		return "OpSetDirective"
	case OpRegisterClass:
		return "OpRegisterClass"
	case OpRegisterClassMethod:
		return "OpRegisterClassMethod"
	case OpRegisterClassField:
		return "OpRegisterClassField"
	case OpInitClassArrayField:
		return "OpInitClassArrayField"
	case OpRegisterClassPropertyGet:
		return "OpRegisterClassPropertyGet"
	case OpRegisterClassPropertyLet:
		return "OpRegisterClassPropertyLet"
	case OpRegisterClassPropertySet:
		return "OpRegisterClassPropertySet"
	case OpCallMember:
		return "OpCallMember"
	case OpCallBuiltin:
		return "OpCallBuiltin"
	case OpCall:
		return "OpCall"
	case OpNewClass:
		return "OpNewClass"
	case OpArraySet:
		return "OpArraySet"
	case OpRet:
		return "OpRet"
	case OpSwap:
		return "OpSwap"
	case OpArgGlobalRef:
		return "OpArgGlobalRef"
	case OpArgLocalRef:
		return "OpArgLocalRef"
	case OpArgClassMemberRef:
		return "OpArgClassMemberRef"
	case OpWithEnter:
		return "OpWithEnter"
	case OpWithLeave:
		return "OpWithLeave"
	case OpWithLoad:
		return "OpWithLoad"
	case OpLetGlobal:
		return "OpLetGlobal"
	case OpLetLocal:
		return "OpLetLocal"
	case OpLetClassMember:
		return "OpLetClassMember"
	case OpCoerceToValue:
		return "OpCoerceToValue"
	case OpJSDeclareName:
		return "OpJSDeclareName"
	case OpJSGetName:
		return "OpJSGetName"
	case OpJSSetName:
		return "OpJSSetName"
	case OpJSImport:
		return "OpJSImport"
	case OpJSExport:
		return "OpJSExport"
	case OpJSGetLocal:
		return "OpJSGetLocal"
	case OpJSSetLocal:
		return "OpJSSetLocal"
	case OpJSIncLocal:
		return "OpJSIncLocal"
	case OpJSDecLocal:
		return "OpJSDecLocal"
	case OpJSMemberGet:
		return "OpJSMemberGet"
	case OpJSMemberSet:
		return "OpJSMemberSet"
	case OpJSCall:
		return "OpJSCall"
	case OpJSCallMember:
		return "OpJSCallMember"
	case OpJSTailCall:
		return "OpJSTailCall"
	case OpJSTailCallMember:
		return "OpJSTailCallMember"
	case OpJSCallComputedMember:
		return "OpJSCallComputedMember"
	case OpJSTailCallComputedMember:
		return "OpJSTailCallComputedMember"
	case OpJSCreateClosure:
		return "OpJSCreateClosure"
	case OpJSAdd:
		return "OpJSAdd"
	case OpJSStrictEq:
		return "OpJSStrictEq"
	case OpJSStrictNeq:
		return "OpJSStrictNeq"
	case OpJSTryEnter:
		return "OpJSTryEnter"
	case OpJSTryLeave:
		return "OpJSTryLeave"
	case OpJSThrow:
		return "OpJSThrow"
	case OpJSNewObject:
		return "OpJSNewObject"
	case OpJSNewArray:
		return "OpJSNewArray"
	case OpJSTypeof:
		return "OpJSTypeof"
	case OpJSInstanceOf:
		return "OpJSInstanceOf"
	case OpJSIn:
		return "OpJSIn"
	case OpJSDelete:
		return "OpJSDelete"
	case OpJSReturn:
		return "OpJSReturn"
	case OpJSLoadUndefined:
		return "OpJSLoadUndefined"
	case OpJSLoadThis:
		return "OpJSLoadThis"
	case OpJSSetThis:
		return "OpJSSetThis"
	case OpJSDup:
		return "OpJSDup"
	case OpJSRequireObject:
		return "OpJSRequireObject"
	case OpJSGetIterator:
		return "OpJSGetIterator"
	case OpJSIteratorNext:
		return "OpJSIteratorNext"
	case OpJSCollectRest:
		return "OpJSCollectRest"
	case OpJSObjectRest:
		return "OpJSObjectRest"
	case OpJSPop:
		return "OpJSPop"
	case OpJSRot:
		return "OpJSRot"
	case OpJSJump:
		return "OpJSJump"
	case OpJSJumpIfFalse:
		return "OpJSJumpIfFalse"
	case OpJSJumpIfTrue:
		return "OpJSJumpIfTrue"
	case OpJSJumpIfNotUndefined:
		return "OpJSJumpIfNotUndefined"
	case OpJSLoadCatchError:
		return "OpJSLoadCatchError"
	case OpJSStoreCatchError:
		return "OpJSStoreCatchError"
	case OpJSBreak:
		return "OpJSBreak"
	case OpJSContinue:
		return "OpJSContinue"
	case OpJSForIn:
		return "OpJSForIn"
	case OpJSForInCleanup:
		return "OpJSForInCleanup"
	case OpJSSwitch:
		return "OpJSSwitch"
	case OpJSCase:
		return "OpJSCase"
	case OpJSDefault:
		return "OpJSDefault"
	case OpJSSubtract:
		return "OpJSSubtract"
	case OpJSMultiply:
		return "OpJSMultiply"
	case OpJSDivide:
		return "OpJSDivide"
	case OpJSModulo:
		return "OpJSModulo"
	case OpJSNegate:
		return "OpJSNegate"
	case OpJSUnaryPlus:
		return "OpJSUnaryPlus"
	case OpJSBitwiseAnd:
		return "OpJSBitwiseAnd"
	case OpJSBitwiseOr:
		return "OpJSBitwiseOr"
	case OpJSBitwiseXor:
		return "OpJSBitwiseXor"
	case OpJSBitwiseNot:
		return "OpJSBitwiseNot"
	case OpJSLeftShift:
		return "OpJSLeftShift"
	case OpJSRightShift:
		return "OpJSRightShift"
	case OpJSUnsignedRightShift:
		return "OpJSUnsignedRightShift"
	case OpJSLess:
		return "OpJSLess"
	case OpJSGreater:
		return "OpJSGreater"
	case OpJSLessEqual:
		return "OpJSLessEqual"
	case OpJSGreaterEqual:
		return "OpJSGreaterEqual"
	case OpJSLooseEqual:
		return "OpJSLooseEqual"
	case OpJSLooseNotEqual:
		return "OpJSLooseNotEqual"
	case OpJSLogicalAnd:
		return "OpJSLogicalAnd"
	case OpJSLogicalOr:
		return "OpJSLogicalOr"
	case OpJSLogicalNot:
		return "OpJSLogicalNot"
	case OpJSNew:
		return "OpJSNew"
	case OpJSMemberDelete:
		return "OpJSMemberDelete"
	case OpJSIndexGet:
		return "OpJSIndexGet"
	case OpJSIndexSet:
		return "OpJSIndexSet"
	case OpJSComma:
		return "OpJSComma"
	case OpJSPostIncrement:
		return "OpJSPostIncrement"
	case OpJSPostDecrement:
		return "OpJSPostDecrement"
	case OpJSPreIncrement:
		return "OpJSPreIncrement"
	case OpJSPreDecrement:
		return "OpJSPreDecrement"
	case OpJSAddAssign:
		return "OpJSAddAssign"
	case OpJSSubtractAssign:
		return "OpJSSubtractAssign"
	case OpJSMultiplyAssign:
		return "OpJSMultiplyAssign"
	case OpJSDivideAssign:
		return "OpJSDivideAssign"
	case OpJSModuloAssign:
		return "OpJSModuloAssign"
	case OpJSMemberIndexGet:
		return "OpJSMemberIndexGet"
	case OpJSMemberIndexSet:
		return "OpJSMemberIndexSet"
	case OpJSPostMemberIncrement:
		return "OpJSPostMemberIncrement"
	case OpJSPostMemberDecrement:
		return "OpJSPostMemberDecrement"
	case OpJSPreMemberIncrement:
		return "OpJSPreMemberIncrement"
	case OpJSPreMemberDecrement:
		return "OpJSPreMemberDecrement"
	case OpJSPostIndexIncrement:
		return "OpJSPostIndexIncrement"
	case OpJSPostIndexDecrement:
		return "OpJSPostIndexDecrement"
	case OpJSPreIndexIncrement:
		return "OpJSPreIndexIncrement"
	case OpJSPreIndexDecrement:
		return "OpJSPreIndexDecrement"
	case OpJSExponent:
		return "OpJSExponent"
	case OpJSCoalesce:
		return "OpJSCoalesce"
	case OpJSJumpIfNullish:
		return "OpJSJumpIfNullish"
	case OpJSJumpIfNotNullish:
		return "OpJSJumpIfNotNullish"
	case OpJSExponentAssign:
		return "OpJSExponentAssign"
	case OpJSLogicalAndAssign:
		return "OpJSLogicalAndAssign"
	case OpJSLogicalOrAssign:
		return "OpJSLogicalOrAssign"
	case OpJSCoalesceAssign:
		return "OpJSCoalesceAssign"
	case OpJSDefineProperty:
		return "OpJSDefineProperty"
	case OpJSSetProto:
		return "OpJSSetProto"
	case OpJSClassInherit:
		return "OpJSClassInherit"
	case OpJSSuperCall:
		return "OpJSSuperCall"
	case OpJSSuperMemberGet:
		return "OpJSSuperMemberGet"
	case OpJSSuperMemberSet:
		return "OpJSSuperMemberSet"
	case OpJSSuperCallMember:
		return "OpJSSuperCallMember"
	case OpJSSuperIndexGet:
		return "OpJSSuperIndexGet"
	case OpJSSuperIndexSet:
		return "OpJSSuperIndexSet"
	case OpJSSuperCallComputedMember:
		return "OpJSSuperCallComputedMember"
	case OpJSYield:
		return "OpJSYield"
	case OpJSYieldDelegate:
		return "OpJSYieldDelegate"
	case OpJSAwait:
		return "OpJSAwait"
	case OpJSStrictModeEnter:
		return "OpJSStrictModeEnter"
	case OpJSStrictModeExit:
		return "OpJSStrictModeExit"
	case OpJSBlockScopeEnter:
		return "OpJSBlockScopeEnter"
	case OpJSBlockScopeExit:
		return "OpJSBlockScopeExit"
	case OpJSLetDeclare:
		return "OpJSLetDeclare"
	case OpJSTDZRegisterLet:
		return "OpJSTDZRegisterLet"
	case OpJSTDZRegisterConst:
		return "OpJSTDZRegisterConst"
	case OpJSConstInitialize:
		return "OpJSConstInitialize"
	case OpJSForIterEnter:
		return "OpJSForIterEnter"
	case OpJSForIterExit:
		return "OpJSForIterExit"
	case OpJSLoadNewTarget:
		return "OpJSLoadNewTarget"
	case OpIncLocalInt:
		return "OpIncLocalInt"
	case OpDecLocalInt:
		return "OpDecLocalInt"
	case OpIncGlobalInt:
		return "OpIncGlobalInt"
	case OpDecGlobalInt:
		return "OpDecGlobalInt"
	case OpJSIncLocalInt:
		return "OpJSIncLocalInt"
	case OpJSDecLocalInt:
		return "OpJSDecLocalInt"
	case OpJSForFastIntEnter:
		return "OpJSForFastIntEnter"
	case OpJSForFastInt:
		return "OpJSForFastInt"
	case OpJSForIterEnterFast:
		return "OpJSForIterEnterFast"
	case OpJSForIterExitFast:
		return "OpJSForIterExitFast"
	case OpJSRootFrameEnter:
		return "OpJSRootFrameEnter"
	case OpJSRootFrameLeave:
		return "OpJSRootFrameLeave"
	case OpNop:
		return "OpNop"
	case OpForNextFastInt:
		return "OpForNextFastInt"
	case OpForNextFastGlobalInt:
		return "OpForNextFastGlobalInt"
	case OpJSJumpIfLessFast:
		return "OpJSJumpIfLessFast"
	case OpJSComputedPropertySet:
		return "OpJSComputedPropertySet"
	case OpJSForOf:
		return "OpJSForOf"
	case OpJSForOfCleanup:
		return "OpJSForOfCleanup"
	case OpJSExportAll:
		return "OpJSExportAll"
	case OpJumpIfNotEq:
		return "OpJumpIfNotEq"
	case OpJumpIfEq:
		return "OpJumpIfEq"
	case OpJumpIfNotLt:
		return "OpJumpIfNotLt"
	case OpJumpIfLte:
		return "OpJumpIfLte"
	case OpJumpIfNotIs:
		return "OpJumpIfNotIs"
	case OpJSJumpIfLooseNotEq:
		return "OpJSJumpIfLooseNotEq"
	case OpJSJumpIfLooseEq:
		return "OpJSJumpIfLooseEq"
	case OpJSJumpIfStrictNotEq:
		return "OpJSJumpIfStrictNotEq"
	case OpJSJumpIfStrictEq:
		return "OpJSJumpIfStrictEq"
	case OpJSJumpIfNotLess:
		return "OpJSJumpIfNotLess"
	case OpJSJumpIfLessEqual:
		return "OpJSJumpIfLessEqual"
	case OpExtPrefix:
		return "OpExtPrefix"
	default:
		return "OpUnknown"
	}
}

func (op ExtOpCode) String() string {
	switch op {
	case ExtOpInitRecord:
		return "ExtOpInitRecord"
	case ExtOpGetRecordMember:
		return "ExtOpGetRecordMember"
	case ExtOpSetRecordMember:
		return "ExtOpSetRecordMember"
	case ExtOpAxonASP:
		return "ExtOpAxonASP"
	case ExtOpJSMathSin:
		return "ExtOpJSMathSin"
	case ExtOpJSMathCos:
		return "ExtOpJSMathCos"
	case ExtOpJSMathTan:
		return "ExtOpJSMathTan"
	case ExtOpJSMathAbs:
		return "ExtOpJSMathAbs"
	case ExtOpJSMathFloor:
		return "ExtOpJSMathFloor"
	case ExtOpJSMathCeil:
		return "ExtOpJSMathCeil"
	case ExtOpJSMathRound:
		return "ExtOpJSMathRound"
	case ExtOpJSMathSqrt:
		return "ExtOpJSMathSqrt"
	case ExtOpJSMathMin:
		return "ExtOpJSMathMin"
	case ExtOpJSMathMax:
		return "ExtOpJSMathMax"
	case ExtOpRegisterClassEvent:
		return "ExtOpRegisterClassEvent"
	case ExtOpRaiseEvent:
		return "ExtOpRaiseEvent"
	case ExtOpWithEventsRegister:
		return "ExtOpWithEventsRegister"
	case ExtOpRegisterClassInterface:
		return "ExtOpRegisterClassInterface"
	case ExtOpJumpLocalIfFalse:
		return "ExtOpJumpLocalIfFalse"
	case ExtOpJumpGlobalIfFalse:
		return "ExtOpJumpGlobalIfFalse"
	case ExtOpJSJumpNameIfFalse:
		return "ExtOpJSJumpNameIfFalse"
	case ExtOpAddLocalConst:
		return "ExtOpAddLocalConst"
	case ExtOpSubGlobalConst:
		return "ExtOpSubGlobalConst"
	case ExtOpConcatLocalConst:
		return "ExtOpConcatLocalConst"
	case ExtOpConstant2:
		return "ExtOpConstant2"
	case ExtOpConstant3:
		return "ExtOpConstant3"
	case ExtOpConstant4:
		return "ExtOpConstant4"
	case ExtOpFileOpen:
		return "ExtOpFileOpen"
	case ExtOpFileClose:
		return "ExtOpFileClose"
	case ExtOpFilePrint:
		return "ExtOpFilePrint"
	case ExtOpFileWrite:
		return "ExtOpFileWrite"
	case ExtOpFileLineInput:
		return "ExtOpFileLineInput"
	case ExtOpFilePut:
		return "ExtOpFilePut"
	case ExtOpFileGet:
		return "ExtOpFileGet"
	case ExtOpFileFreeFile:
		return "ExtOpFileFreeFile"
	case ExtOpJSReThrow:
		return "ExtOpJSReThrow"
	case ExtOpCloneRecord:
		return "ExtOpCloneRecord"
	case ExtOpShiftLeft:
		return "ExtOpShiftLeft"
	case ExtOpShiftRight:
		return "ExtOpShiftRight"
	default:
		return "ExtOpUnknown"
	}
}

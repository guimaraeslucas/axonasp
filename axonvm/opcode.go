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
	OpMe          // Pushes the current class instance (activeClassObjectID) as VTObject
	OpCallMember  // [OpCode, ConstMemberIdxHigh, ConstMemberIdxLow, ArgCountHigh, ArgCountLow]
	OpCallBuiltin // [OpCode, RegistryIdxHigh, RegistryIdxLow, ArgCountHigh, ArgCountLow]
	OpCall
	OpNewClass // [OpCode, ClassNameConstIdxHigh, ClassNameConstIdxLow]
	OpArraySet // [OpCode, ArgCountHigh, ArgCountLow] ; stack: [..., targetArray, idx1..idxN, value]
	OpRet
)

const (
	// OpArgGlobalRef pushes a VTArgRef for a global slot. Used at call sites so that
	// ByRef parameters can write back to the caller's global variable on return.
	// [OpCode, IdxHigh, IdxLow]
	OpArgGlobalRef OpCode = iota + OpRet + 1
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
	OpJSDeclareName         // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSGetName             // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSSetName             // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSMemberGet           // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSMemberSet           // [OpCode, NameConstIdxHigh, NameConstIdxLow]
	OpJSCall                // [OpCode, ArgCountHigh, ArgCountLow]
	OpJSCallMember          // [OpCode, NameConstIdxHigh, NameConstIdxLow, ArgCountHigh, ArgCountLow]
	OpJSCreateClosure       // [OpCode, FunctionTemplateConstIdxHigh, FunctionTemplateConstIdxLow]
	OpJSAdd                 // [OpCode]
	OpJSStrictEq            // [OpCode]
	OpJSStrictNeq           // [OpCode]
	OpJSTryEnter            // [OpCode, CatchTarget3, CatchTarget2, CatchTarget1, CatchTarget0]
	OpJSTryLeave            // [OpCode]
	OpJSThrow               // [OpCode]
	OpJSNewObject           // [OpCode]
	OpJSNewArray            // [OpCode, CountHigh, CountLow]
	OpJSTypeof              // [OpCode]
	OpJSInstanceOf          // [OpCode]
	OpJSIn                  // [OpCode]
	OpJSDelete              // [OpCode]
	OpJSReturn              // [OpCode]
	OpJSLoadUndefined       // [OpCode]
	OpJSLoadThis            // [OpCode]
	OpJSDup                 // [OpCode]
	OpJSPop                 // [OpCode]
	OpJSJump                // [OpCode, Target3, Target2, Target1, Target0]
	OpJSJumpIfFalse         // [OpCode, Target3, Target2, Target1, Target0]
	OpJSJumpIfTrue          // [OpCode, Target3, Target2, Target1, Target0]
	OpJSLoadCatchError      // [OpCode]
	OpJSStoreCatchError     // [OpCode]
	OpJSBreak               // [OpCode] - Break from loop (jump target managed by compiler)
	OpJSContinue            // [OpCode] - Continue to next iteration (jump target managed by compiler)
	OpJSForIn               // [OpCode, EnumVarIdxHigh, EnumVarIdxLow, LoopStartTarget3, LoopStartTarget2, LoopStartTarget1, LoopStartTarget0] - for...in setup
	OpJSForInCleanup        // [OpCode, ForInPos3, ForInPos2, ForInPos1, ForInPos0] - clears for...in enumerator state and source value
	OpJSSwitch              // [OpCode] - switch statement (value already on stack)
	OpJSCase                // [OpCode, Target3, Target2, Target1, Target0] - case label with jump target
	OpJSDefault             // [OpCode, Target3, Target2, Target1, Target0] - default label with jump target
	OpJSSubtract            // [OpCode] - JScript binary subtraction (for type coercion compatibility)
	OpJSMultiply            // [OpCode] - JScript binary multiplication
	OpJSDivide              // [OpCode] - JScript binary division
	OpJSModulo              // [OpCode] - JScript binary modulo
	OpJSNegate              // [OpCode] - JScript unary negation
	OpJSBitwiseAnd          // [OpCode] - JScript bitwise AND
	OpJSBitwiseOr           // [OpCode] - JScript bitwise OR
	OpJSBitwiseXor          // [OpCode] - JScript bitwise XOR
	OpJSBitwiseNot          // [OpCode] - JScript bitwise NOT
	OpJSLeftShift           // [OpCode] - JScript left shift
	OpJSRightShift          // [OpCode] - JScript right shift
	OpJSUnsignedRightShift  // [OpCode] - JScript unsigned right shift
	OpJSLess                // [OpCode] - JScript less than
	OpJSGreater             // [OpCode] - JScript greater than
	OpJSLessEqual           // [OpCode] - JScript less than or equal
	OpJSGreaterEqual        // [OpCode] - JScript greater than or equal
	OpJSLooseEqual          // [OpCode] - JScript == (loose equality)
	OpJSLooseNotEqual       // [OpCode] - JScript != (loose inequality)
	OpJSLogicalAnd          // [OpCode] - JScript && (logical AND)
	OpJSLogicalOr           // [OpCode] - JScript || (logical OR)
	OpJSLogicalNot          // [OpCode] - JScript ! (logical NOT)
	OpJSNew                 // [OpCode, ArgCountHigh, ArgCountLow] - new constructor call
	OpJSMemberDelete        // [OpCode, NameConstIdxHigh, NameConstIdxLow] - delete property
	OpJSIndexGet            // [OpCode] - array/object index access (bracket notation)
	OpJSIndexSet            // [OpCode] - array/object index assignment
	OpJSComma               // [OpCode] - comma operator
	OpJSPostIncrement       // [OpCode, NameConstIdxHigh, NameConstIdxLow] - post-increment (var++)
	OpJSPostDecrement       // [OpCode, NameConstIdxHigh, NameConstIdxLow] - post-decrement (var--)
	OpJSPreIncrement        // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pre-increment (++var)
	OpJSPreDecrement        // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pre-decrement (--var)
	OpJSAddAssign           // [OpCode, NameConstIdxHigh, NameConstIdxLow] - +=
	OpJSSubtractAssign      // [OpCode, NameConstIdxHigh, NameConstIdxLow] - -=
	OpJSMultiplyAssign      // [OpCode, NameConstIdxHigh, NameConstIdxLow] - *=
	OpJSDivideAssign        // [OpCode, NameConstIdxHigh, NameConstIdxLow] - /=
	OpJSModuloAssign        // [OpCode, NameConstIdxHigh, NameConstIdxLow] - %=
	OpJSMemberIndexGet      // [OpCode, NameConstIdxHigh, NameConstIdxLow] - member[index] get
	OpJSMemberIndexSet      // [OpCode, NameConstIdxHigh, NameConstIdxLow] - member[index] set
	OpJSPostMemberIncrement // [OpCode, NameConstIdxHigh, NameConstIdxLow] - post-increment (obj.prop++)
	OpJSPostMemberDecrement // [OpCode, NameConstIdxHigh, NameConstIdxLow] - post-decrement (obj.prop--)
	OpJSPreMemberIncrement  // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pre-increment (++obj.prop)
	OpJSPreMemberDecrement  // [OpCode, NameConstIdxHigh, NameConstIdxLow] - pre-decrement (--obj.prop)
	OpJSPostIndexIncrement  // [OpCode] - post-increment (obj[idx]++)
	OpJSPostIndexDecrement  // [OpCode] - post-decrement (obj[idx]--)
	OpJSPreIndexIncrement   // [OpCode] - pre-increment (++obj[idx])
	OpJSPreIndexDecrement   // [OpCode] - pre-decrement (--obj[idx])
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
	case OpDiv:
		return "OpDiv"
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
	case OpCall:
		return "OpCall"
	case OpNewClass:
		return "OpNewClass"
	case OpArraySet:
		return "OpArraySet"
	case OpMemberGet:
		return "OpMemberGet"
	case OpMe:
		return "OpMe"
	case OpMemberSet:
		return "OpMemberSet"
	case OpMemberSetSet:
		return "OpMemberSetSet"
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
	case OpJSMemberGet:
		return "OpJSMemberGet"
	case OpJSMemberSet:
		return "OpJSMemberSet"
	case OpJSCall:
		return "OpJSCall"
	case OpJSCallMember:
		return "OpJSCallMember"
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
	case OpJSDup:
		return "OpJSDup"
	case OpJSPop:
		return "OpJSPop"
	case OpJSJump:
		return "OpJSJump"
	case OpJSJumpIfFalse:
		return "OpJSJumpIfFalse"
	case OpJSJumpIfTrue:
		return "OpJSJumpIfTrue"
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
	case OpJSMemberIndexSet:
		return "OpJSMemberIndexSet"
	default:
		return "OpUnknown"
	}
}

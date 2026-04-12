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
	default:
		return "OpUnknown"
	}
}

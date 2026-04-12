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
 * ----------------------------------------------------------------------------
 * THIRD PARTY ATTRIBUTION / ORIGINAL SOURCE
 * ----------------------------------------------------------------------------
 * This code was adapted from https://github.com/kmvi/vbscript-parser/
 * ensuring compatibility with VBScript language specifications.
 *
 * Original Copyright (c) [ANO] kmvi (and/or original authors).
 * Licensed under the BSD 3-Clause License.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * 1. Redistributions of source code must retain the above copyright notice,
 * this list of conditions and the following disclaimer.
 *
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 * this list of conditions and the following disclaimer in the documentation
 * and/or other materials provided with the distribution.
 *
 * 3. Neither the name of the copyright holder nor the names of its
 * contributors may be used to endorse or promote products derived from
 * this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 */
package vbscript

// VBSyntaxErrorCode represents VBScript syntax error codes
type VBSyntaxErrorCode int

const (
	// Error codes that match VBScript's standard error codes
	InvalidProcedureCallOrArgument                   VBSyntaxErrorCode = 5
	Overflow                                         VBSyntaxErrorCode = 6
	OutOfMemory                                      VBSyntaxErrorCode = 7
	SubscriptOutOfRange                              VBSyntaxErrorCode = 9
	TheArrayIsOfFixedLengthOrTemporarilyLocked       VBSyntaxErrorCode = 10
	DivisionByZero                                   VBSyntaxErrorCode = 11
	TypeMismatch                                     VBSyntaxErrorCode = 13
	OutOfStringSpace                                 VBSyntaxErrorCode = 14
	CannotPerformTheRequestedOperation               VBSyntaxErrorCode = 17
	StackOverflow                                    VBSyntaxErrorCode = 28
	UndefinedSubProcedureOrFunction                  VBSyntaxErrorCode = 35
	ErrorLoadingDLL                                  VBSyntaxErrorCode = 48
	InternalError                                    VBSyntaxErrorCode = 51
	BadFileNameOrNumber                              VBSyntaxErrorCode = 52
	FileNotFound                                     VBSyntaxErrorCode = 53
	BadFileMode                                      VBSyntaxErrorCode = 54
	FileIsAlreadyOpen                                VBSyntaxErrorCode = 55
	DeviceIOError                                    VBSyntaxErrorCode = 57
	FileAlreadyExists                                VBSyntaxErrorCode = 58
	DiskSpaceIsFull                                  VBSyntaxErrorCode = 61
	InputBeyondTheEndOfTheFile                       VBSyntaxErrorCode = 62
	TooManyFiles                                     VBSyntaxErrorCode = 67
	DeviceNotAvailable                               VBSyntaxErrorCode = 68
	PermissionDenied                                 VBSyntaxErrorCode = 70
	DiskNotReady                                     VBSyntaxErrorCode = 71
	CannotRenameWithDifferentDrive                   VBSyntaxErrorCode = 74
	PathFileAccessError                              VBSyntaxErrorCode = 75
	PathNotFound                                     VBSyntaxErrorCode = 76
	ObjectVariableNotSet                             VBSyntaxErrorCode = 91
	ForLoopIsNotInitialized                          VBSyntaxErrorCode = 92
	InvalidUseOfNull                                 VBSyntaxErrorCode = 94
	CouldNotCreateTheRequiredTemporaryFile           VBSyntaxErrorCode = 322
	CouldNotFindTargetObject                         VBSyntaxErrorCode = 424
	ActiveXCannotCreateObject                        VBSyntaxErrorCode = 429
	ClassDoesNotSupportAutomation                    VBSyntaxErrorCode = 430
	FileNameOrClassNameNotFound                      VBSyntaxErrorCode = 432
	ObjectDoesntSupportThisPropertyOrMethod          VBSyntaxErrorCode = 438
	AutomationError                                  VBSyntaxErrorCode = 440
	ObjectDoesNotSupportThisAction                   VBSyntaxErrorCode = 445
	ObjectDoesNotSupportTheNamedArguments            VBSyntaxErrorCode = 446
	ObjectDoesNotSupportTheCurrentLocale             VBSyntaxErrorCode = 447
	NamedArgumentNotFound                            VBSyntaxErrorCode = 448
	ParametersAreNotOptional                         VBSyntaxErrorCode = 449
	WrongNumberOfParameters                          VBSyntaxErrorCode = 450
	IsNotACollectionOfObjects                        VBSyntaxErrorCode = 451
	TheSpecifiedDLLFunctionWasNotFound               VBSyntaxErrorCode = 453
	CodeResourceLockError                            VBSyntaxErrorCode = 455
	ThisKeyAlreadyAssociatedWithAnElement            VBSyntaxErrorCode = 457
	VariableUsesAnAutomationTypeNotSupported         VBSyntaxErrorCode = 458
	TheRemoteServerDoesNotExist                      VBSyntaxErrorCode = 462
	ImageIsInvalid                                   VBSyntaxErrorCode = 481
	VariableNotDefined                               VBSyntaxErrorCode = 500
	IllegalAssignment                                VBSyntaxErrorCode = 501
	TheObjectIsNotSafeForScripting                   VBSyntaxErrorCode = 502
	ObjectNotSafeForInitializing                     VBSyntaxErrorCode = 503
	ObjectCanNotCreateASecureEnvironment             VBSyntaxErrorCode = 504
	InvalidOrUnqualifiedReference                    VBSyntaxErrorCode = 505
	ClassTypeIsNotDefined                            VBSyntaxErrorCode = 506
	UnexpectedError                                  VBSyntaxErrorCode = 507
	InsufficientMemory                               VBSyntaxErrorCode = 1001
	SyntaxError                                      VBSyntaxErrorCode = 1002
	ExpectedColon                                    VBSyntaxErrorCode = 1003
	ExpectedLParen                                   VBSyntaxErrorCode = 1005
	ExpectedRParen                                   VBSyntaxErrorCode = 1006
	ExpectedRBracket                                 VBSyntaxErrorCode = 1007
	ExpectedIdentifier                               VBSyntaxErrorCode = 1010
	ExpectedEqual                                    VBSyntaxErrorCode = 1011
	ExpectedIf                                       VBSyntaxErrorCode = 1012
	ExpectedTo                                       VBSyntaxErrorCode = 1013
	ExpectedEnd                                      VBSyntaxErrorCode = 1014
	ExpectedFunction                                 VBSyntaxErrorCode = 1015
	ExpectedSub                                      VBSyntaxErrorCode = 1016
	ExpectedThen                                     VBSyntaxErrorCode = 1017
	ExpectedWend                                     VBSyntaxErrorCode = 1018
	ExpectedLoop                                     VBSyntaxErrorCode = 1019
	ExpectedNext                                     VBSyntaxErrorCode = 1020
	ExpectedCase                                     VBSyntaxErrorCode = 1021
	ExpectedSelect                                   VBSyntaxErrorCode = 1022
	ExpectedExpression                               VBSyntaxErrorCode = 1023
	ExpectedStatement                                VBSyntaxErrorCode = 1024
	ExpectedEndOfStatement                           VBSyntaxErrorCode = 1025
	ExpectedInteger                                  VBSyntaxErrorCode = 1026
	ExpectedWhileOrUntil                             VBSyntaxErrorCode = 1027
	ExpectedWhileUntilOrEndOfStatement               VBSyntaxErrorCode = 1028
	ExpectedWith                                     VBSyntaxErrorCode = 1029
	IdentifierTooLong                                VBSyntaxErrorCode = 1030
	InvalidNumber                                    VBSyntaxErrorCode = 1031
	InvalidCharacter                                 VBSyntaxErrorCode = 1032
	UnterminatedStringConstant                       VBSyntaxErrorCode = 1033
	UnterminatedComment                              VBSyntaxErrorCode = 1034
	InvalidUseOfMeKeyword                            VBSyntaxErrorCode = 1037
	LoopWithoutDo                                    VBSyntaxErrorCode = 1038
	InvalidExitStatement                             VBSyntaxErrorCode = 1039
	InvalidForLoopControlVariable                    VBSyntaxErrorCode = 1040
	NameRedefined                                    VBSyntaxErrorCode = 1041
	MustBeFirstStatementOnTheLine                    VBSyntaxErrorCode = 1042
	CannotAssignToNonByValVariable                   VBSyntaxErrorCode = 1043
	CannotUseParenthesesWhenCallingSub               VBSyntaxErrorCode = 1044
	ExpectedLiteral                                  VBSyntaxErrorCode = 1045
	ExpectedIn                                       VBSyntaxErrorCode = 1046
	ExpectedClass                                    VBSyntaxErrorCode = 1047
	MustBeDefinedInsideClass                         VBSyntaxErrorCode = 1048
	ExpectedLetGetSet                                VBSyntaxErrorCode = 1049
	ExpectedProperty                                 VBSyntaxErrorCode = 1050
	InconsistentNumberOfArguments                    VBSyntaxErrorCode = 1051
	CannotHaveMultipleDefault                        VBSyntaxErrorCode = 1052
	ClassInitializeOrTerminateDoNotHaveArguments     VBSyntaxErrorCode = 1053
	PropertySetOrLetMustHaveArguments                VBSyntaxErrorCode = 1054
	UnexpectedNext                                   VBSyntaxErrorCode = 1055
	DefaultCanBeSpecifiedOnlyOnPropertyFunctionOrSub VBSyntaxErrorCode = 1056
	DefaultMustAlsoSpecifyPublic                     VBSyntaxErrorCode = 1057
	DefaultCanOnlyBeOnPropertyGet                    VBSyntaxErrorCode = 1058
	RequiresARegularExpressionObject                 VBSyntaxErrorCode = 5016
	RegularExpressionSyntaxError                     VBSyntaxErrorCode = 5017
	TheNumberOfWordsError                            VBSyntaxErrorCode = 5018
	RegularExpressionIsMissingClosingBracket         VBSyntaxErrorCode = 5019
	RegularExpressionIsMissingClosingParen           VBSyntaxErrorCode = 5020
	CharacterSetCrossBorder                          VBSyntaxErrorCode = 5021
	True                                             VBSyntaxErrorCode = 32766
	False                                            VBSyntaxErrorCode = 32767
	ElementWasNotFound                               VBSyntaxErrorCode = 32811
	SpecifiedDateNotAvailableInLocale                VBSyntaxErrorCode = 32812
)

var VBScriptErrorMessages = map[VBSyntaxErrorCode]string{
	InvalidProcedureCallOrArgument: "Invalid procedure call or argument",
	Overflow:                       "Overflow",
	OutOfMemory:                    "Out of memory",
	SubscriptOutOfRange:            "Subscript out of range",
	TheArrayIsOfFixedLengthOrTemporarilyLocked: "The Array is of fixed length or temporarily locked",
	DivisionByZero:                               "Division by zero",
	TypeMismatch:                                 "Type mismatch",
	OutOfStringSpace:                             "Out of string space (overflow)",
	CannotPerformTheRequestedOperation:           "cannot perform the requested operation",
	StackOverflow:                                "Stack overflow",
	UndefinedSubProcedureOrFunction:              "Undefined SUB procedure or Function",
	ErrorLoadingDLL:                              "Error loading DLL",
	InternalError:                                "Internal error",
	BadFileNameOrNumber:                          "bad file name or number",
	FileNotFound:                                 "File not found",
	BadFileMode:                                  "Bad file mode",
	FileIsAlreadyOpen:                            "File is already open",
	DeviceIOError:                                "Device I/O error",
	FileAlreadyExists:                            "File already exists",
	DiskSpaceIsFull:                              "Disk space is full",
	InputBeyondTheEndOfTheFile:                   "input beyond the end of the file",
	TooManyFiles:                                 "Too many files",
	DeviceNotAvailable:                           "Device not available",
	PermissionDenied:                             "Permission denied",
	DiskNotReady:                                 "Disk not ready",
	CannotRenameWithDifferentDrive:               "Cannot rename with different drive",
	PathFileAccessError:                          "Path/ file access error",
	PathNotFound:                                 "Path not found",
	ObjectVariableNotSet:                         "Object variable not set",
	ForLoopIsNotInitialized:                      "For loop is not initialized",
	InvalidUseOfNull:                             "Invalid use of Null",
	CouldNotCreateTheRequiredTemporaryFile:       "Could not create the required temporary file",
	CouldNotFindTargetObject:                     "Could not find target object",
	ActiveXCannotCreateObject:                    "ActiveX cannot create object",
	ClassDoesNotSupportAutomation:                "Class does not support Automation",
	FileNameOrClassNameNotFound:                  "File name or class name not found during Automation operation",
	ObjectDoesntSupportThisPropertyOrMethod:      "Object doesn’t support this property or method",
	AutomationError:                              "Automation error",
	ObjectDoesNotSupportThisAction:               "Object does not support this action",
	ObjectDoesNotSupportTheNamedArguments:        "Object does not support the named arguments",
	ObjectDoesNotSupportTheCurrentLocale:         "Object does not support the current locale",
	NamedArgumentNotFound:                        "Named argument not found",
	ParametersAreNotOptional:                     "parameters are not optional",
	WrongNumberOfParameters:                      "Wrong number of parameters or invalid property assignment",
	IsNotACollectionOfObjects:                    "is not a collection of objects",
	TheSpecifiedDLLFunctionWasNotFound:           "The specified DLL function was not found",
	CodeResourceLockError:                        "Code resource lock error",
	ThisKeyAlreadyAssociatedWithAnElement:        "This key already associated with an element of this collection",
	VariableUsesAnAutomationTypeNotSupported:     "Variable uses an Automation type not supported in VBScript",
	TheRemoteServerDoesNotExist:                  "The remote server does not exist or is not available",
	ImageIsInvalid:                               "Image is invalid.",
	VariableNotDefined:                           "variable not defined",
	IllegalAssignment:                            "illegal assignment",
	TheObjectIsNotSafeForScripting:               "The object is not safe for scripting",
	ObjectNotSafeForInitializing:                 "Object not safe for initializing",
	ObjectCanNotCreateASecureEnvironment:         "Object can not create a secure environment",
	InvalidOrUnqualifiedReference:                "invalid or unqualified reference",
	ClassTypeIsNotDefined:                        "Class/Type is not defined",
	UnexpectedError:                              "Unexpected error",
	InsufficientMemory:                           "Insufficient memory",
	SyntaxError:                                  "syntax error",
	ExpectedColon:                                "Missing ':'",
	ExpectedLParen:                               "Missing '('",
	ExpectedRParen:                               "Missing ')'",
	ExpectedRBracket:                             "Missing ']'",
	ExpectedIdentifier:                           "Missing Identifier",
	ExpectedEqual:                                "Missing =",
	ExpectedIf:                                   "Missing 'If'",
	ExpectedTo:                                   "Missing 'To'",
	ExpectedEnd:                                  "Missing 'End'",
	ExpectedFunction:                             "Missing 'Function'",
	ExpectedSub:                                  "Missing 'Sub'",
	ExpectedThen:                                 "Missing 'Then'",
	ExpectedWend:                                 "Missing 'Wend'",
	ExpectedLoop:                                 "Missing 'Loop'",
	ExpectedNext:                                 "Missing 'Next'",
	ExpectedCase:                                 "Missing 'Case'",
	ExpectedSelect:                               "Missing 'Select'",
	ExpectedExpression:                           "Missing expression",
	ExpectedStatement:                            "Missing statement",
	ExpectedEndOfStatement:                       "Missing end of statement",
	ExpectedInteger:                              "Requires an integer constant",
	ExpectedWhileOrUntil:                         "Missing 'While' or 'Until'",
	ExpectedWhileUntilOrEndOfStatement:           "Missing 'While, 'Until,' or End of statement",
	ExpectedWith:                                 "Too many locals or arguments",
	IdentifierTooLong:                            "Identifier Too long",
	InvalidNumber:                                "The number of is invalid",
	InvalidCharacter:                             "Invalid character",
	UnterminatedStringConstant:                   "Unterminated string constant",
	UnterminatedComment:                          "Unterminated comment",
	InvalidUseOfMeKeyword:                        "Invalid use of 'Me' Keyword",
	LoopWithoutDo:                                "'loop' statement is missing 'do'",
	InvalidExitStatement:                         "invalid 'exit' statement",
	InvalidForLoopControlVariable:                "invalid 'for' loop control variable",
	NameRedefined:                                "Name redefined",
	MustBeFirstStatementOnTheLine:                "Must be the first statement line",
	CannotAssignToNonByValVariable:               "Cannot be assigned to non - Byval argument",
	CannotUseParenthesesWhenCallingSub:           "Cannot use parentheses when calling Sub",
	ExpectedLiteral:                              "Requires a literal constant",
	ExpectedIn:                                   "Missing 'In'",
	ExpectedClass:                                "Missing 'Class'",
	MustBeDefinedInsideClass:                     "Must be inside a class definition",
	ExpectedLetGetSet:                            "Missing Let, Set or Get in the property declaration",
	ExpectedProperty:                             "Missing 'Property'",
	InconsistentNumberOfArguments:                "The Number of parameters must be consistent with the attribute description",
	CannotHaveMultipleDefault:                    "cannot have more than one default attribute / method in a class",
	ClassInitializeOrTerminateDoNotHaveArguments: "Class did not initialize or terminate the process parameters",
	PropertySetOrLetMustHaveArguments:            "Property Let or Set should have at least one parameter",
	UnexpectedNext:                               "Error at 'Next'",
	DefaultCanBeSpecifiedOnlyOnPropertyFunctionOrSub: "'Default' can only be in the 'Property', 'Function' or 'Sub'",
	DefaultMustAlsoSpecifyPublic:                     "'Default' must be 'Public'",
	DefaultCanOnlyBeOnPropertyGet:                    "Can only specify the Property Get in the 'Default'",
	RequiresARegularExpressionObject:                 "Requires a regular expression object",
	RegularExpressionSyntaxError:                     "Regular expression syntax error",
	TheNumberOfWordsError:                            "The number of words error",
	RegularExpressionIsMissingClosingBracket:         "Regular expressions is missing ']'",
	RegularExpressionIsMissingClosingParen:           "Regular expressions is missing ')'",
	CharacterSetCrossBorder:                          "Character set Cross-border",
	True:                                             "True",
	False:                                            "False",
	ElementWasNotFound:                               "Element was not found",
	SpecifiedDateNotAvailableInLocale:                "The specified date is not available in the current locale’s calendar",
}

func (e VBSyntaxErrorCode) String() string {
	if msg, ok := VBScriptErrorMessages[e]; ok {
		return msg
	}
	return "Unknown VBScript Error"
}

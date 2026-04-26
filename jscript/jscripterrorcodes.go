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
 */
package jscript

// JSSyntaxErrorCode represents JScript runtime and syntax error codes
type JSSyntaxErrorCode int

const (
	// Standard JScript runtime and syntax errors
	InvalidProcedureCallOrArgument        JSSyntaxErrorCode = 5
	Overflow                              JSSyntaxErrorCode = 6
	OutOfMemory                           JSSyntaxErrorCode = 7
	SubscriptOutOfRange                   JSSyntaxErrorCode = 9
	ArrayIsFixedOrLocked                  JSSyntaxErrorCode = 10
	DivisionByZero                        JSSyntaxErrorCode = 11
	TypeMismatch                          JSSyntaxErrorCode = 13
	OutOfStringSpace                      JSSyntaxErrorCode = 14
	CannotPerformRequestedOperation       JSSyntaxErrorCode = 17
	OutOfStackSpace                       JSSyntaxErrorCode = 28
	SubOrFunctionNotDefined               JSSyntaxErrorCode = 35
	ErrorInLoadingDLL                     JSSyntaxErrorCode = 48
	InternalError                         JSSyntaxErrorCode = 51
	BadFileNameOrNumber                   JSSyntaxErrorCode = 52
	FileNotFound                          JSSyntaxErrorCode = 53
	BadFileMode                           JSSyntaxErrorCode = 54
	FileAlreadyOpen                       JSSyntaxErrorCode = 55
	DeviceIOError                         JSSyntaxErrorCode = 57
	FileAlreadyExists                     JSSyntaxErrorCode = 58
	DiskFull                              JSSyntaxErrorCode = 61
	InputPastEndOfFile                    JSSyntaxErrorCode = 62
	TooManyFiles                          JSSyntaxErrorCode = 67
	DeviceUnavailable                     JSSyntaxErrorCode = 68
	PermissionDenied                      JSSyntaxErrorCode = 70
	DiskNotReady                          JSSyntaxErrorCode = 71
	CannotRenameWithDifferentDrive        JSSyntaxErrorCode = 74
	PathFileAccessError                   JSSyntaxErrorCode = 75
	PathNotFound                          JSSyntaxErrorCode = 76
	ObjectVariableNotSet                  JSSyntaxErrorCode = 91
	ForLoopNotInitialized                 JSSyntaxErrorCode = 92
	InvalidUseOfNull                      JSSyntaxErrorCode = 94
	CannotCreateTemporaryFile             JSSyntaxErrorCode = 322
	ObjectRequired                        JSSyntaxErrorCode = 424
	AutomationServerCannotCreateObject    JSSyntaxErrorCode = 429
	ClassDoesNotSupportAutomation         JSSyntaxErrorCode = 430
	FileNameOrClassNameNotFound           JSSyntaxErrorCode = 432
	ObjectDoesntSupportPropertyOrMethod   JSSyntaxErrorCode = 438
	AutomationError                       JSSyntaxErrorCode = 440
	ObjectDoesntSupportAction             JSSyntaxErrorCode = 445
	ObjectDoesntSupportNamedArguments     JSSyntaxErrorCode = 446
	ObjectDoesntSupportCurrentLocale      JSSyntaxErrorCode = 447
	NamedArgumentNotFound                 JSSyntaxErrorCode = 448
	ArgumentNotOptional                   JSSyntaxErrorCode = 449
	WrongNumberOfArguments                JSSyntaxErrorCode = 450
	ObjectNotACollection                  JSSyntaxErrorCode = 451
	SpecifiedDLLFunctionNotFound          JSSyntaxErrorCode = 453
	VariableUsesUnsupportedAutomationType JSSyntaxErrorCode = 458
	RemoteServerMachineDoesNotExist       JSSyntaxErrorCode = 462
	CannotAssignToVariable                JSSyntaxErrorCode = 501
	ObjectNotSafeForScripting             JSSyntaxErrorCode = 502
	ObjectNotSafeForInitializing          JSSyntaxErrorCode = 503
	ObjectNotSafeForCreating              JSSyntaxErrorCode = 504
	ExceptionOccurred                     JSSyntaxErrorCode = 507

	// JScript-specific 5000+ series errors
	CannotAssignToThis                      JSSyntaxErrorCode = 5000
	NumberExpected                          JSSyntaxErrorCode = 5001
	FunctionExpected                        JSSyntaxErrorCode = 5002
	CannotAssignToFunctionResult            JSSyntaxErrorCode = 5003
	CannotIndexObject                       JSSyntaxErrorCode = 5004
	StringExpected                          JSSyntaxErrorCode = 5005
	DateObjectExpected                      JSSyntaxErrorCode = 5006
	ObjectExpectedJScript                   JSSyntaxErrorCode = 5007
	IllegalAssignment                       JSSyntaxErrorCode = 5008
	UndefinedIdentifier                     JSSyntaxErrorCode = 5009
	BooleanExpected                         JSSyntaxErrorCode = 5010
	CannotExecuteCodeFromFreedScript        JSSyntaxErrorCode = 5011
	ObjectMemberExpected                    JSSyntaxErrorCode = 5012
	VBArrayExpected                         JSSyntaxErrorCode = 5013
	JScriptObjectExpected                   JSSyntaxErrorCode = 5014
	EnumeratorObjectExpected                JSSyntaxErrorCode = 5015
	RegularExpressionObjectExpected         JSSyntaxErrorCode = 5016
	SyntaxErrorInRegularExpression          JSSyntaxErrorCode = 5017
	UnexpectedQuantifier                    JSSyntaxErrorCode = 5018
	ExpectedBracketInRegularExpression      JSSyntaxErrorCode = 5019
	ExpectedParenInRegularExpression        JSSyntaxErrorCode = 5020
	InvalidRangeInCharacterSet              JSSyntaxErrorCode = 5021
	ExceptionThrownAndNotCaught             JSSyntaxErrorCode = 5022
	FunctionDoesNotHaveValidPrototypeObject JSSyntaxErrorCode = 5023
)

var JSErrorMessages = map[JSSyntaxErrorCode]string{
	InvalidProcedureCallOrArgument:          "Invalid procedure call or argument",
	Overflow:                                "Overflow",
	OutOfMemory:                             "Out of memory",
	SubscriptOutOfRange:                     "Subscript out of range",
	ArrayIsFixedOrLocked:                    "This array is fixed or temporarily locked",
	DivisionByZero:                          "Division by zero",
	TypeMismatch:                            "Type mismatch",
	OutOfStringSpace:                        "Out of string space",
	CannotPerformRequestedOperation:         "Can't perform requested operation",
	OutOfStackSpace:                         "Out of stack space",
	SubOrFunctionNotDefined:                 "Sub or Function not defined",
	ErrorInLoadingDLL:                       "Error in loading DLL",
	InternalError:                           "Internal error",
	BadFileNameOrNumber:                     "Bad file name or number",
	FileNotFound:                            "File not found",
	BadFileMode:                             "Bad file mode",
	FileAlreadyOpen:                         "File already open",
	DeviceIOError:                           "Device I/O error",
	FileAlreadyExists:                       "File already exists",
	DiskFull:                                "Disk full",
	InputPastEndOfFile:                      "Input past end of file",
	TooManyFiles:                            "Too many files",
	DeviceUnavailable:                       "Device unavailable",
	PermissionDenied:                        "Permission denied",
	DiskNotReady:                            "Disk not ready",
	CannotRenameWithDifferentDrive:          "Can't rename with different drive",
	PathFileAccessError:                     "Path/File access error",
	PathNotFound:                            "Path not found",
	ObjectVariableNotSet:                    "Object variable or With block variable not set",
	ForLoopNotInitialized:                   "For loop not initialized",
	InvalidUseOfNull:                        "Invalid use of Null",
	CannotCreateTemporaryFile:               "Can't create necessary temporary file",
	ObjectRequired:                          "Object required",
	AutomationServerCannotCreateObject:      "Automation server can't create object",
	ClassDoesNotSupportAutomation:           "Class doesn't support Automation",
	FileNameOrClassNameNotFound:             "File name or class name not found during Automation operation",
	ObjectDoesntSupportPropertyOrMethod:     "Object doesn't support this property or method",
	AutomationError:                         "Automation error",
	ObjectDoesntSupportAction:               "Object doesn't support this action",
	ObjectDoesntSupportNamedArguments:       "Object doesn't support named arguments",
	ObjectDoesntSupportCurrentLocale:        "Object doesn't support current locale setting",
	NamedArgumentNotFound:                   "Named argument not found",
	ArgumentNotOptional:                     "Argument not optional",
	WrongNumberOfArguments:                  "Wrong number of arguments or invalid property assignment",
	ObjectNotACollection:                    "Object not a collection",
	SpecifiedDLLFunctionNotFound:            "Specified DLL function not found",
	VariableUsesUnsupportedAutomationType:   "Variable uses an Automation type not supported in JScript",
	RemoteServerMachineDoesNotExist:         "The remote server machine does not exist or is unavailable",
	CannotAssignToVariable:                  "Cannot assign to variable",
	ObjectNotSafeForScripting:               "Object not safe for scripting",
	ObjectNotSafeForInitializing:            "Object not safe for initializing",
	ObjectNotSafeForCreating:                "Object not safe for creating",
	ExceptionOccurred:                       "An exception occurred",
	CannotAssignToThis:                      "Cannot assign to 'this'",
	NumberExpected:                          "Number expected",
	FunctionExpected:                        "Function expected",
	CannotAssignToFunctionResult:            "Cannot assign to a function result",
	CannotIndexObject:                       "Cannot index object",
	StringExpected:                          "String expected",
	DateObjectExpected:                      "Date object expected",
	ObjectExpectedJScript:                   "Object expected",
	IllegalAssignment:                       "Illegal assignment",
	UndefinedIdentifier:                     "Undefined identifier",
	BooleanExpected:                         "Boolean expected",
	CannotExecuteCodeFromFreedScript:        "Can't execute code from a freed script",
	ObjectMemberExpected:                    "Object member expected",
	VBArrayExpected:                         "VBArray expected",
	JScriptObjectExpected:                   "JScript object expected",
	EnumeratorObjectExpected:                "Enumerator object expected",
	RegularExpressionObjectExpected:         "Regular Expression object expected",
	SyntaxErrorInRegularExpression:          "Syntax error in regular expression",
	UnexpectedQuantifier:                    "Unexpected quantifier",
	ExpectedBracketInRegularExpression:      "Expected ']' in regular expression",
	ExpectedParenInRegularExpression:        "Expected ')' in regular expression",
	InvalidRangeInCharacterSet:              "Invalid range in character set",
	ExceptionThrownAndNotCaught:             "Exception thrown and not caught",
	FunctionDoesNotHaveValidPrototypeObject: "Function does not have a valid prototype object",
}

func (e JSSyntaxErrorCode) String() string {
	if msg, ok := JSErrorMessages[e]; ok {
		return msg
	}
	return "Unknown JScript Error"
}

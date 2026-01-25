/*
 * AxonASP Server - Version 1.0
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
	SyntaxError VBSyntaxErrorCode = 1002
	ExpectedColon VBSyntaxErrorCode = 1003
	ExpectedLParen VBSyntaxErrorCode = 1005
	ExpectedRParen VBSyntaxErrorCode = 1006
	ExpectedRBracket VBSyntaxErrorCode = 1007
	ExpectedIdentifier VBSyntaxErrorCode = 1010
	ExpectedEqual VBSyntaxErrorCode = 1011
	ExpectedIf VBSyntaxErrorCode = 1012
	ExpectedTo VBSyntaxErrorCode = 1013
	ExpectedEnd VBSyntaxErrorCode = 1014
	ExpectedFunction VBSyntaxErrorCode = 1015
	ExpectedSub VBSyntaxErrorCode = 1016
	ExpectedThen VBSyntaxErrorCode = 1017
	ExpectedWend VBSyntaxErrorCode = 1018
	ExpectedLoop VBSyntaxErrorCode = 1019
	ExpectedNext VBSyntaxErrorCode = 1020
	ExpectedCase VBSyntaxErrorCode = 1021
	ExpectedSelect VBSyntaxErrorCode = 1022
	ExpectedExpression VBSyntaxErrorCode = 1023
	ExpectedStatement VBSyntaxErrorCode = 1024
	ExpectedEndOfStatement VBSyntaxErrorCode = 1025
	ExpectedInteger VBSyntaxErrorCode = 1026
	ExpectedWhileOrUntil VBSyntaxErrorCode = 1027
	ExpectedWhileUntilOrEndOfStatement VBSyntaxErrorCode = 1028
	ExpectedWith VBSyntaxErrorCode = 1029
	IdentifierTooLong VBSyntaxErrorCode = 1030
	InvalidNumber VBSyntaxErrorCode = 1031
	InvalidCharacter VBSyntaxErrorCode = 1032
	UnterminatedStringConstant VBSyntaxErrorCode = 1033
	UnterminatedComment VBSyntaxErrorCode = 1034
	InvalidUseOfMeKeyword VBSyntaxErrorCode = 1037
	LoopWithoutDo VBSyntaxErrorCode = 1038
	InvalidExitStatement VBSyntaxErrorCode = 1039
	InvalidForLoopControlVariable VBSyntaxErrorCode = 1040
	NameRedefined VBSyntaxErrorCode = 1041
	MustBeFirstStatementOnTheLine VBSyntaxErrorCode = 1042
	CannotAssignToNonByValVariable VBSyntaxErrorCode = 1043
	CannotUseParenthesesWhenCallingSub VBSyntaxErrorCode = 1044
	ExpectedLiteral VBSyntaxErrorCode = 1045
	ExpectedIn VBSyntaxErrorCode = 1046
	ExpectedClass VBSyntaxErrorCode = 1047
	MustBeDefinedInsideClass VBSyntaxErrorCode = 1048
	ExpectedLetGetSet VBSyntaxErrorCode = 1049
	ExpectedProperty VBSyntaxErrorCode = 1050
	InconsistentNumberOfArguments VBSyntaxErrorCode = 1051
	CannotHaveMultipleDefault VBSyntaxErrorCode = 1052
	ClassInitializeOrTerminateDoNotHaveArguments VBSyntaxErrorCode = 1053
	PropertySetOrLetMustHaveArguments VBSyntaxErrorCode = 1054
	UnexpectedNext VBSyntaxErrorCode = 1055
	DefaultCanBeSpecifiedOnlyOnPropertyFunctionOrSub VBSyntaxErrorCode = 1056
	DefaultMustAlsoSpecifyPublic VBSyntaxErrorCode = 1057
	DefaultCanOnlyBeOnPropertyGet VBSyntaxErrorCode = 1058
)

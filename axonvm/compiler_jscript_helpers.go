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

import (
	"math"
	"math/big"
	"slices"
	"strconv"

	jsast "g3pix.com.br/axonasp/jscript/ast"
	jstoken "g3pix.com.br/axonasp/jscript/token"
	jsunistring "g3pix.com.br/axonasp/jscript/unistring"
)

func jsFunctionPreventsLocalSlots(fn *jsast.FunctionLiteral) bool {
	if fn == nil || fn.Body == nil {
		return false
	}
	return slices.ContainsFunc(fn.Body.List, jsStatementPreventsLocalSlots)
}

func jsProgramPreventsLocalSlots(stmts []jsast.Statement) bool {
	return slices.ContainsFunc(stmts, jsStatementPreventsLocalSlots)
}

func jsStatementPreventsLocalSlots(stmt jsast.Statement) bool {
	if stmt == nil {
		return false
	}
	switch node := stmt.(type) {
	case *jsast.FunctionDeclaration:
		return true
	case *jsast.WithStatement:
		return true
	case *jsast.BlockStatement:
		if slices.ContainsFunc(node.List, jsStatementPreventsLocalSlots) {
			return true
		}
	case *jsast.ExpressionStatement:
		return jsExpressionPreventsLocalSlots(node.Expression)
	case *jsast.IfStatement:
		if jsExpressionPreventsLocalSlots(node.Test) {
			return true
		}
		if jsStatementPreventsLocalSlots(node.Consequent) {
			return true
		}
		if jsStatementPreventsLocalSlots(node.Alternate) {
			return true
		}
	case *jsast.ForStatement:
		if node.Initializer != nil {
			switch init := node.Initializer.(type) {
			case *jsast.ForLoopInitializerExpression:
				if jsExpressionPreventsLocalSlots(init.Expression) {
					return true
				}
			case *jsast.ForLoopInitializerVarDeclList:
				for _, b := range init.List {
					if b != nil && jsExpressionPreventsLocalSlots(b.Initializer) {
						return true
					}
				}
			case *jsast.ForLoopInitializerLexicalDecl:
				for _, b := range init.LexicalDeclaration.List {
					if b != nil && jsExpressionPreventsLocalSlots(b.Initializer) {
						return true
					}
				}
			}
		}
		if jsExpressionPreventsLocalSlots(node.Test) || jsExpressionPreventsLocalSlots(node.Update) {
			return true
		}
		return jsStatementPreventsLocalSlots(node.Body)
	case *jsast.ReturnStatement:
		return jsExpressionPreventsLocalSlots(node.Argument)
	case *jsast.ThrowStatement:
		return jsExpressionPreventsLocalSlots(node.Argument)
	case *jsast.WhileStatement:
		return jsExpressionPreventsLocalSlots(node.Test) || jsStatementPreventsLocalSlots(node.Body)
	case *jsast.DoWhileStatement:
		return jsExpressionPreventsLocalSlots(node.Test) || jsStatementPreventsLocalSlots(node.Body)
	case *jsast.SwitchStatement:
		if jsExpressionPreventsLocalSlots(node.Discriminant) {
			return true
		}
		for i := range node.Body {
			clause := node.Body[i]
			if clause == nil {
				continue
			}
			if jsExpressionPreventsLocalSlots(clause.Test) {
				return true
			}
			if slices.ContainsFunc(clause.Consequent, jsStatementPreventsLocalSlots) {
				return true
			}
		}
	case *jsast.TryStatement:
		if jsStatementPreventsLocalSlots(node.Body) {
			return true
		}
		if node.Catch != nil && jsStatementPreventsLocalSlots(node.Catch.Body) {
			return true
		}
		if node.Finally != nil && jsStatementPreventsLocalSlots(node.Finally) {
			return true
		}
	case *jsast.UsingDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionPreventsLocalSlots(b.Initializer) {
				return true
			}
		}
	case *jsast.ForInStatement:
		return jsExpressionPreventsLocalSlots(node.Source) || jsStatementPreventsLocalSlots(node.Body)
	case *jsast.ForOfStatement:
		return jsExpressionPreventsLocalSlots(node.Source) || jsStatementPreventsLocalSlots(node.Body)
	}
	return false
}

func jsExpressionPreventsLocalSlots(expr jsast.Expression) bool {
	if expr == nil {
		return false
	}
	switch node := expr.(type) {
	case *jsast.FunctionLiteral, *jsast.ArrowFunctionLiteral:
		return true
	case *jsast.AssignExpression:
		return jsExpressionPreventsLocalSlots(node.Left) || jsExpressionPreventsLocalSlots(node.Right)
	case *jsast.BinaryExpression:
		return jsExpressionPreventsLocalSlots(node.Left) || jsExpressionPreventsLocalSlots(node.Right)
	case *jsast.UnaryExpression:
		return jsExpressionPreventsLocalSlots(node.Operand)
	case *jsast.DotExpression:
		return jsExpressionPreventsLocalSlots(node.Left)
	case *jsast.BracketExpression:
		return jsExpressionPreventsLocalSlots(node.Left) || jsExpressionPreventsLocalSlots(node.Member)
	case *jsast.CallExpression:
		if jsExpressionPreventsLocalSlots(node.Callee) {
			return true
		}
		if slices.ContainsFunc(node.ArgumentList, jsExpressionPreventsLocalSlots) {
			return true
		}
	case *jsast.NewExpression:
		if jsExpressionPreventsLocalSlots(node.Callee) {
			return true
		}
		if slices.ContainsFunc(node.ArgumentList, jsExpressionPreventsLocalSlots) {
			return true
		}
	case *jsast.ObjectLiteral:
		for i := range node.Value {
			switch prop := node.Value[i].(type) {
			case *jsast.PropertyShort:
				if jsExpressionPreventsLocalSlots(prop.Initializer) {
					return true
				}
			case *jsast.PropertyKeyed:
				if jsExpressionPreventsLocalSlots(prop.Key) || jsExpressionPreventsLocalSlots(prop.Value) {
					return true
				}
			}
		}
	case *jsast.ArrayLiteral:
		if slices.ContainsFunc(node.Value, jsExpressionPreventsLocalSlots) {
			return true
		}
	case *jsast.ConditionalExpression:
		return jsExpressionPreventsLocalSlots(node.Test) || jsExpressionPreventsLocalSlots(node.Consequent) || jsExpressionPreventsLocalSlots(node.Alternate)
	case *jsast.SequenceExpression:
		if slices.ContainsFunc(node.Sequence, jsExpressionPreventsLocalSlots) {
			return true
		}
	case *jsast.TemplateLiteral:
		if slices.ContainsFunc(node.Expressions, jsExpressionPreventsLocalSlots) {
			return true
		}
	}
	return false
}

func jsStatementCapturesLoopNames(stmt jsast.Statement, names map[string]struct{}) bool {
	if stmt == nil || len(names) == 0 {
		return false
	}
	switch node := stmt.(type) {
	case *jsast.FunctionDeclaration:
		return jsFunctionCapturesLoopNames(node.Function, names)
	case *jsast.BlockStatement:
		for i := range node.List {
			if jsStatementCapturesLoopNames(node.List[i], names) {
				return true
			}
		}
	case *jsast.ExpressionStatement:
		return jsExpressionCapturesLoopNames(node.Expression, names)
	case *jsast.IfStatement:
		return jsExpressionCapturesLoopNames(node.Test, names) || jsStatementCapturesLoopNames(node.Consequent, names) || jsStatementCapturesLoopNames(node.Alternate, names)
	case *jsast.ForStatement:
		if node.Initializer != nil {
			switch init := node.Initializer.(type) {
			case *jsast.ForLoopInitializerExpression:
				if jsExpressionCapturesLoopNames(init.Expression, names) {
					return true
				}
			case *jsast.ForLoopInitializerVarDeclList:
				for _, b := range init.List {
					if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
						return true
					}
				}
			case *jsast.ForLoopInitializerLexicalDecl:
				for _, b := range init.LexicalDeclaration.List {
					if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
						return true
					}
				}
			}
		}
		return jsExpressionCapturesLoopNames(node.Test, names) || jsExpressionCapturesLoopNames(node.Update, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.ForInStatement:
		return jsExpressionCapturesLoopNames(node.Source, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.ForOfStatement:
		return jsExpressionCapturesLoopNames(node.Source, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.ReturnStatement:
		return jsExpressionCapturesLoopNames(node.Argument, names)
	case *jsast.ThrowStatement:
		return jsExpressionCapturesLoopNames(node.Argument, names)
	case *jsast.WhileStatement:
		return jsExpressionCapturesLoopNames(node.Test, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.DoWhileStatement:
		return jsExpressionCapturesLoopNames(node.Test, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.SwitchStatement:
		if jsExpressionCapturesLoopNames(node.Discriminant, names) {
			return true
		}
		for i := range node.Body {
			clause := node.Body[i]
			if clause == nil {
				continue
			}
			if jsExpressionCapturesLoopNames(clause.Test, names) {
				return true
			}
			for j := range clause.Consequent {
				if jsStatementCapturesLoopNames(clause.Consequent[j], names) {
					return true
				}
			}
		}
	case *jsast.TryStatement:
		if jsStatementCapturesLoopNames(node.Body, names) {
			return true
		}
		if node.Catch != nil && jsStatementCapturesLoopNames(node.Catch.Body, names) {
			return true
		}
		if node.Finally != nil && jsStatementCapturesLoopNames(node.Finally, names) {
			return true
		}
	case *jsast.VariableStatement:
		for _, b := range node.List {
			if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
				return true
			}
		}
	case *jsast.LexicalDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
				return true
			}
		}
	case *jsast.UsingDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
				return true
			}
		}
	}
	return false
}

func jsExpressionCapturesLoopNames(expr jsast.Expression, names map[string]struct{}) bool {
	if expr == nil || len(names) == 0 {
		return false
	}
	switch node := expr.(type) {
	case *jsast.FunctionLiteral:
		return jsFunctionCapturesLoopNames(node, names)
	case *jsast.ArrowFunctionLiteral:
		return jsArrowFunctionCapturesLoopNames(node, names)
	case *jsast.ClassExpression:
		if node.Class == nil {
			return false
		}
		return jsClassCapturesLoopNames(node.Class, names)
	case *jsast.AssignExpression:
		return jsExpressionCapturesLoopNames(node.Left, names) || jsExpressionCapturesLoopNames(node.Right, names)
	case *jsast.BinaryExpression:
		return jsExpressionCapturesLoopNames(node.Left, names) || jsExpressionCapturesLoopNames(node.Right, names)
	case *jsast.UnaryExpression:
		return jsExpressionCapturesLoopNames(node.Operand, names)
	case *jsast.DotExpression:
		return jsExpressionCapturesLoopNames(node.Left, names)
	case *jsast.BracketExpression:
		return jsExpressionCapturesLoopNames(node.Left, names) || jsExpressionCapturesLoopNames(node.Member, names)
	case *jsast.CallExpression:
		if jsExpressionCapturesLoopNames(node.Callee, names) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionCapturesLoopNames(node.ArgumentList[i], names) {
				return true
			}
		}
	case *jsast.NewExpression:
		if jsExpressionCapturesLoopNames(node.Callee, names) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionCapturesLoopNames(node.ArgumentList[i], names) {
				return true
			}
		}
	case *jsast.ObjectLiteral:
		for i := range node.Value {
			switch p := node.Value[i].(type) {
			case *jsast.PropertyShort:
				if p.Initializer != nil {
					if jsExpressionCapturesLoopNames(p.Initializer, names) {
						return true
					}
				}
			case *jsast.PropertyKeyed:
				if jsExpressionCapturesLoopNames(p.Key, names) || jsExpressionCapturesLoopNames(p.Value, names) {
					return true
				}
			}
		}
	case *jsast.ArrayLiteral:
		for i := range node.Value {
			if jsExpressionCapturesLoopNames(node.Value[i], names) {
				return true
			}
		}
	case *jsast.ConditionalExpression:
		return jsExpressionCapturesLoopNames(node.Test, names) || jsExpressionCapturesLoopNames(node.Consequent, names) || jsExpressionCapturesLoopNames(node.Alternate, names)
	case *jsast.TemplateLiteral:
		for i := range node.Expressions {
			if jsExpressionCapturesLoopNames(node.Expressions[i], names) {
				return true
			}
		}
	case *jsast.SequenceExpression:
		for i := range node.Sequence {
			if jsExpressionCapturesLoopNames(node.Sequence[i], names) {
				return true
			}
		}
	}
	return false
}

func jsFunctionCapturesLoopNames(fn *jsast.FunctionLiteral, names map[string]struct{}) bool {
	if fn == nil || fn.Body == nil || len(names) == 0 {
		return false
	}
	visible := jsVisibleCaptureNamesForFunction(names, fn.Name, fn.ParameterList)
	if len(visible) == 0 {
		return false
	}
	for i := range fn.Body.List {
		if jsStatementReferencesLoopNames(fn.Body.List[i], visible) {
			return true
		}
	}
	return false
}

func jsArrowFunctionCapturesLoopNames(fn *jsast.ArrowFunctionLiteral, names map[string]struct{}) bool {
	if fn == nil || len(names) == 0 {
		return false
	}
	visible := jsVisibleCaptureNamesForFunction(names, nil, fn.ParameterList)
	if len(visible) == 0 {
		return false
	}
	switch body := fn.Body.(type) {
	case *jsast.ExpressionBody:
		return jsExpressionReferencesLoopNames(body.Expression, visible)
	case *jsast.BlockStatement:
		for i := range body.List {
			if jsStatementReferencesLoopNames(body.List[i], visible) {
				return true
			}
		}
	}
	return false
}

func jsClassCapturesLoopNames(class *jsast.ClassLiteral, names map[string]struct{}) bool {
	if class == nil || len(names) == 0 {
		return false
	}
	if jsExpressionCapturesLoopNames(class.SuperClass, names) {
		return true
	}
	for i := range class.Body {
		switch el := class.Body[i].(type) {
		case *jsast.FieldDefinition:
			if jsExpressionCapturesLoopNames(el.Key, names) || jsExpressionCapturesLoopNames(el.Initializer, names) {
				return true
			}
		case *jsast.MethodDefinition:
			if jsExpressionCapturesLoopNames(el.Key, names) || jsFunctionCapturesLoopNames(el.Body, names) {
				return true
			}
		case *jsast.ClassStaticBlock:
			if el.Block != nil {
				for j := range el.Block.List {
					if jsStatementCapturesLoopNames(el.Block.List[j], names) {
						return true
					}
				}
			}
		}
	}
	return false
}

func jsVisibleCaptureNamesForFunction(names map[string]struct{}, fnName *jsast.Identifier, params *jsast.ParameterList) map[string]struct{} {
	if len(names) == 0 {
		return nil
	}
	visible := make(map[string]struct{}, len(names))
	for name := range names {
		visible[name] = struct{}{}
	}
	if fnName != nil {
		delete(visible, fnName.Name.String())
	}
	if params != nil {
		for _, b := range params.List {
			if b == nil || b.Target == nil {
				continue
			}
			paramNames := make([]string, 0, 2)
			jsExtractBindingNames(b.Target, &paramNames)
			for _, n := range paramNames {
				delete(visible, n)
			}
		}
		if params.Rest != nil {
			restNames := make([]string, 0, 2)
			if target, ok := params.Rest.(jsast.BindingTarget); ok {
				jsExtractBindingNames(target, &restNames)
			}
			for _, n := range restNames {
				delete(visible, n)
			}
		}
	}
	if len(visible) == 0 {
		return nil
	}
	return visible
}

func jsStatementReferencesLoopNames(stmt jsast.Statement, names map[string]struct{}) bool {
	if stmt == nil || len(names) == 0 {
		return false
	}
	switch node := stmt.(type) {
	case *jsast.FunctionDeclaration:
		return jsFunctionCapturesLoopNames(node.Function, names)
	case *jsast.BlockStatement:
		for i := range node.List {
			if jsStatementReferencesLoopNames(node.List[i], names) {
				return true
			}
		}
	case *jsast.ExpressionStatement:
		return jsExpressionReferencesLoopNames(node.Expression, names)
	case *jsast.IfStatement:
		return jsExpressionReferencesLoopNames(node.Test, names) || jsStatementReferencesLoopNames(node.Consequent, names) || jsStatementReferencesLoopNames(node.Alternate, names)
	case *jsast.ForStatement:
		if node.Initializer != nil {
			switch init := node.Initializer.(type) {
			case *jsast.ForLoopInitializerExpression:
				if jsExpressionReferencesLoopNames(init.Expression, names) {
					return true
				}
			case *jsast.ForLoopInitializerVarDeclList:
				for _, b := range init.List {
					if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
						return true
					}
				}
			case *jsast.ForLoopInitializerLexicalDecl:
				for _, b := range init.LexicalDeclaration.List {
					if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
						return true
					}
				}
			}
		}
		return jsExpressionReferencesLoopNames(node.Test, names) || jsExpressionReferencesLoopNames(node.Update, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.ForInStatement:
		return jsExpressionReferencesLoopNames(node.Source, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.ForOfStatement:
		return jsExpressionReferencesLoopNames(node.Source, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.ReturnStatement:
		return jsExpressionReferencesLoopNames(node.Argument, names)
	case *jsast.ThrowStatement:
		return jsExpressionReferencesLoopNames(node.Argument, names)
	case *jsast.WhileStatement:
		return jsExpressionReferencesLoopNames(node.Test, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.DoWhileStatement:
		return jsExpressionReferencesLoopNames(node.Test, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.SwitchStatement:
		if jsExpressionReferencesLoopNames(node.Discriminant, names) {
			return true
		}
		for i := range node.Body {
			clause := node.Body[i]
			if clause == nil {
				continue
			}
			if jsExpressionReferencesLoopNames(clause.Test, names) {
				return true
			}
			for j := range clause.Consequent {
				if jsStatementReferencesLoopNames(clause.Consequent[j], names) {
					return true
				}
			}
		}
	case *jsast.TryStatement:
		if jsStatementReferencesLoopNames(node.Body, names) {
			return true
		}
		if node.Catch != nil && jsStatementReferencesLoopNames(node.Catch.Body, names) {
			return true
		}
		if node.Finally != nil && jsStatementReferencesLoopNames(node.Finally, names) {
			return true
		}
	case *jsast.VariableStatement:
		for _, b := range node.List {
			if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
				return true
			}
		}
	case *jsast.LexicalDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
				return true
			}
		}
	case *jsast.UsingDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
				return true
			}
		}
	}
	return false
}

func jsExpressionReferencesLoopNames(expr jsast.Expression, names map[string]struct{}) bool {
	if expr == nil || len(names) == 0 {
		return false
	}
	switch node := expr.(type) {
	case *jsast.Identifier:
		_, ok := names[node.Name.String()]
		return ok
	case *jsast.FunctionLiteral:
		return jsFunctionCapturesLoopNames(node, names)
	case *jsast.ArrowFunctionLiteral:
		return jsArrowFunctionCapturesLoopNames(node, names)
	case *jsast.ClassExpression:
		if node.Class == nil {
			return false
		}
		return jsClassCapturesLoopNames(node.Class, names)
	case *jsast.AssignExpression:
		return jsExpressionReferencesLoopNames(node.Left, names) || jsExpressionReferencesLoopNames(node.Right, names)
	case *jsast.BinaryExpression:
		return jsExpressionReferencesLoopNames(node.Left, names) || jsExpressionReferencesLoopNames(node.Right, names)
	case *jsast.UnaryExpression:
		return jsExpressionReferencesLoopNames(node.Operand, names)
	case *jsast.DotExpression:
		return jsExpressionReferencesLoopNames(node.Left, names)
	case *jsast.BracketExpression:
		return jsExpressionReferencesLoopNames(node.Left, names) || jsExpressionReferencesLoopNames(node.Member, names)
	case *jsast.CallExpression:
		if jsExpressionReferencesLoopNames(node.Callee, names) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionReferencesLoopNames(node.ArgumentList[i], names) {
				return true
			}
		}
	case *jsast.NewExpression:
		if jsExpressionReferencesLoopNames(node.Callee, names) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionReferencesLoopNames(node.ArgumentList[i], names) {
				return true
			}
		}
	case *jsast.ObjectLiteral:
		for i := range node.Value {
			switch p := node.Value[i].(type) {
			case *jsast.PropertyShort:
				if p.Initializer != nil {
					if jsExpressionReferencesLoopNames(p.Initializer, names) {
						return true
					}
				} else {
					if _, ok := names[p.Name.Name.String()]; ok {
						return true
					}
				}
			case *jsast.PropertyKeyed:
				if jsExpressionReferencesLoopNames(p.Key, names) || jsExpressionReferencesLoopNames(p.Value, names) {
					return true
				}
			}
		}
	case *jsast.ArrayLiteral:
		for i := range node.Value {
			if jsExpressionReferencesLoopNames(node.Value[i], names) {
				return true
			}
		}
	case *jsast.ConditionalExpression:
		return jsExpressionReferencesLoopNames(node.Test, names) || jsExpressionReferencesLoopNames(node.Consequent, names) || jsExpressionReferencesLoopNames(node.Alternate, names)
	case *jsast.TemplateLiteral:
		for i := range node.Expressions {
			if jsExpressionReferencesLoopNames(node.Expressions[i], names) {
				return true
			}
		}
	case *jsast.SequenceExpression:
		for i := range node.Sequence {
			if jsExpressionReferencesLoopNames(node.Sequence[i], names) {
				return true
			}
		}
	}
	return false
}

// foldJSExpr recursively attempts to fold a JScript AST expression to a
// simpler literal at compile time. It mutates BinaryExpression children
// in-place and returns either the original node or a new literal node.
func foldJSExpr(expr jsast.Expression) jsast.Expression {
	bin, ok := expr.(*jsast.BinaryExpression)
	if !ok {
		return expr
	}
	// Fold children first (depth-first), enabling chained folding.
	bin.Left = foldJSExpr(bin.Left)
	bin.Right = foldJSExpr(bin.Right)
	// Try to evaluate this binary expression at compile time.
	if folded := foldJSBinaryLiterals(bin.Operator, bin.Left, bin.Right); folded != nil {
		return folded
	}
	return bin
}

// foldJSBinaryLiterals evaluates a binary operation over two literal AST nodes
// at compile time. Returns a new literal expression, or nil if not foldable.
func foldJSBinaryLiterals(op jstoken.Token, left, right jsast.Expression) jsast.Expression {
	lf, lok := jsExprAsFloat(left)
	rf, rok := jsExprAsFloat(right)
	ls, lsok := jsExprAsString(left)
	rs, rsok := jsExprAsString(right)

	switch op {
	case jstoken.PLUS:
		// + with any string operand coerces both sides to string.
		if lsok || rsok {
			if !lsok {
				if !lok {
					return nil
				}
				ls = jsFloatToString(lf)
			}
			if !rsok {
				if !rok {
					return nil
				}
				rs = jsFloatToString(rf)
			}
			return &jsast.StringLiteral{Value: jsunistring.String(ls + rs)}
		}
		if lok && rok {
			return jsNewNumLiteral(lf + rf)
		}
	case jstoken.MINUS:
		if lok && rok {
			return jsNewNumLiteral(lf - rf)
		}
	case jstoken.MULTIPLY:
		if lok && rok {
			return jsNewNumLiteral(lf * rf)
		}
	case jstoken.SLASH:
		if lok && rok && rf != 0 {
			return jsNewNumLiteral(lf / rf)
		}
	case jstoken.REMAINDER:
		if lok && rok && rf != 0 {
			return jsNewNumLiteral(math.Mod(lf, rf))
		}
	}
	return nil
}

// jsExprAsFloat extracts a float64 from a JS NumberLiteral node.
func jsExprAsFloat(expr jsast.Expression) (float64, bool) {
	num, ok := expr.(*jsast.NumberLiteral)
	if !ok {
		return 0, false
	}
	switch v := num.Value.(type) {
	case int64:
		return float64(v), true
	case int:
		return float64(v), true
	case float64:
		return v, true
	}
	return 0, false
}

// jsExprAsString extracts a plain Go string from a JS StringLiteral node.
func jsExprAsString(expr jsast.Expression) (string, bool) {
	sl, ok := expr.(*jsast.StringLiteral)
	if !ok {
		return "", false
	}
	return sl.Value.String(), true
}

// jsNewNumLiteral constructs a JS NumberLiteral carrying a float64 result.
// Idx and Literal are left at zero/empty since this node is never used for
// error reporting.
func jsNewNumLiteral(f float64) *jsast.NumberLiteral {
	return &jsast.NumberLiteral{Value: f}
}

func jsNumericLiteralInt64(expr jsast.Expression) (int64, bool) {
	num, ok := expr.(*jsast.NumberLiteral)
	if !ok {
		return 0, false
	}
	switch v := num.Value.(type) {
	case int:
		return int64(v), true
	case int32:
		return int64(v), true
	case int64:
		return v, true
	case float32:
		if float64(int64(v)) == float64(v) {
			return int64(v), true
		}
	case float64:
		if math.Trunc(v) == v {
			return int64(v), true
		}
	case *big.Int:
		if v.IsInt64() {
			return v.Int64(), true
		}
	}
	return 0, false
}

func jsFloatToString(f float64) string {
	if f == math.Trunc(f) && !math.IsInf(f, 0) && !math.IsNaN(f) {
		return strconv.FormatInt(int64(f), 10)
	}
	return strconv.FormatFloat(f, 'f', -1, 64)
}

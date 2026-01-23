package vbscript

import (
	"github.com/guimaraeslucas/vbscript-go/ast"
)

//Adapted from https://github.com/kmvi/vbscript-parser/tree/master

// Parser represents a VBScript parser
type Parser struct {
	options     *ParsingOptions
	lexer       *Lexer
	next        Token
	startMarker Marker
	lastMarker  Marker
	comments    []*CommentToken
	inWithBlock bool
}

// NewParser creates a new Parser instance
func NewParser(code string) *Parser {
	return NewParserWithOptions(code, NewParsingOptions())
}

// NewParserWithOptions creates a new Parser with custom options
func NewParserWithOptions(code string, options *ParsingOptions) *Parser {
	if code == "" {
		panic("code cannot be empty")
	}

	if options == nil {
		options = NewParsingOptions()
	}

	lexer := NewLexer(code)

	p := &Parser{
		options:     options,
		lexer:       lexer,
		startMarker: NewMarker(0, lexer.CurrentLine, 0),
		lastMarker:  NewMarker(0, lexer.CurrentLine, 0),
		comments:    []*CommentToken{},
		inWithBlock: false,
		next:        &InvalidToken{},
	}

	return p
}

// Parse parses VBScript code and returns a Program AST node
func (p *Parser) Parse() *ast.Program {
	p.reset()

	p.createMarker() // marker
	p.skipCommentsAndNewlines()

	optionExplicit, optionCompare, optionBase := p.parseOptions()
	program := ast.NewProgram(optionExplicit, optionCompare, optionBase)

	for !p.matchEof() {
		p.skipCommentsAndNewlines()
		if p.matchEof() {
			break
		}
		program.Body = append(program.Body, p.parseGlobalStatement())
	}

	p.skipCommentsAndNewlines()

	return program
}

// Private parsing methods

func (p *Parser) reset() {
	p.lexer.Reset()
	p.startMarker = NewMarker(0, p.lexer.CurrentLine, 0)
	p.lastMarker = NewMarker(0, p.lexer.CurrentLine, 0)
	p.next = p.lexer.NextToken()
	p.lastMarker = NewMarker(p.lexer.Index, p.lexer.CurrentLine, p.lexer.LineIndex())
}

func (p *Parser) parseOptions() (bool, ast.OptionCompareMode, int) {
	optionExplicit := false
	optionCompare := ast.OptionCompareBinary
	optionBase := 0

	for {
		p.skipCommentsAndNewlines()
		if !p.optKeyword(KeywordOption) {
			break
		}
		switch {
		case p.matchKeyword(KeywordExplicit):
			p.move()
			optionExplicit = true
			p.skipComments()
			p.expectEofOrLineTermination()
		case p.matchKeyword(KeywordCompare):
			p.move()
			if p.matchKeyword(KeywordText) {
				p.move()
				optionCompare = ast.OptionCompareText
			} else if p.matchKeyword(KeywordBinary) {
				p.move()
				optionCompare = ast.OptionCompareBinary
			} else {
				panic(p.vbSyntaxError(SyntaxError))
			}
			p.skipComments()
			p.expectEofOrLineTermination()
		case p.matchKeyword(KeywordBase):
			p.move()
			baseValue := 0
			switch lit := p.next.(type) {
			case *DecIntegerLiteralToken:
				baseValue = int(lit.Value)
			case *OctIntegerLiteralToken:
				baseValue = int(lit.Value)
			case *HexIntegerLiteralToken:
				baseValue = int(lit.Value)
			default:
				panic(p.vbSyntaxError(SyntaxError))
			}

			if baseValue != 0 && baseValue != 1 {
				panic(p.vbSyntaxError(SyntaxError))
			}

			optionBase = baseValue
			p.move()
			p.skipComments()
			p.expectEofOrLineTermination()
		default:
			p.skipComments()
			p.expectEofOrLineTermination()
		}
	}

	return optionExplicit, optionCompare, optionBase
}

// consumeOptionStatement eats an inline Option Explicit (or any Option statement) as a no-op stub.
// VBScript only honors Option at the top of the file; we accept it elsewhere to avoid parse errors.
func (p *Parser) consumeOptionStatement() {
	p.move() // consume Option
	if p.matchKeyword(KeywordExplicit) {
		p.move()
	} else if p.matchKeyword(KeywordCompare) {
		p.move()
		if p.matchKeyword(KeywordText) || p.matchKeyword(KeywordBinary) {
			p.move()
		}
	} else if p.matchKeyword(KeywordBase) {
		p.move()
		if _, ok := p.next.(*DecIntegerLiteralToken); ok {
			p.move()
		} else if _, ok := p.next.(*OctIntegerLiteralToken); ok {
			p.move()
		} else if _, ok := p.next.(*HexIntegerLiteralToken); ok {
			p.move()
		}
	}
	p.skipComments()
	p.optLineTermination()
}

func (p *Parser) parseGlobalStatement() ast.Statement {
	p.createMarker() // marker

	var stmt ast.Statement
	if p.matchKeyword(KeywordClass) {
		stmt = p.parseClassDeclaration()
	} else if p.matchKeyword(KeywordSub) {
		stmt = p.parseSubDeclaration(ast.MethodAccessModifierNone, false, false)
	} else if p.matchKeyword(KeywordFunction) {
		stmt = p.parseFunctionDeclaration(ast.MethodAccessModifierNone, false, false)
	} else if p.matchKeyword(KeywordPrivate) || p.matchKeyword(KeywordPublic) {
		stmt = p.parsePublicOrPrivate(true, false)
	} else {
		stmt = p.parseBlockStatement(true)
	}

	return stmt
}

// Marker and token movement methods

func (p *Parser) createMarker() Marker {
	return NewMarker(p.startMarker.Index, p.startMarker.Line, p.startMarker.Column)
}

func (p *Parser) move() Token {
	token := p.next

	p.lastMarker.Index = p.lexer.Index
	p.lastMarker.Line = p.lexer.CurrentLine
	p.lastMarker.Column = p.lexer.LineIndex()

	p.lexer.skipWhitespaces()

	if p.lexer.Index != p.startMarker.Index {
		p.startMarker.Index = p.lexer.Index
		p.startMarker.Line = p.lexer.CurrentLine
		p.startMarker.Column = p.lexer.LineIndex()
	}

	p.next = p.lexer.NextToken()

	return token
}

// Token matching and expecting methods

func (p *Parser) matchEof() bool {
	_, ok := p.next.(*EOFToken)
	return ok
}

func (p *Parser) matchLineTermination() bool {
	switch p.next.(type) {
	case *LineTerminationToken, *ColonLineTerminationToken:
		return true
	default:
		return false
	}
}

func (p *Parser) matchColonLineTermination() bool {
	_, ok := p.next.(*ColonLineTerminationToken)
	return ok
}

func (p *Parser) matchKeyword(kw Keyword) bool {
	switch t := p.next.(type) {
	case *KeywordToken:
		return t.Keyword == kw
	case *KeywordOrIdentifierToken:
		return t.Keyword == kw
	default:
		return false
	}
}

func (p *Parser) optKeyword(kw Keyword) bool {
	if p.matchKeyword(kw) {
		p.move()
		return true
	}
	return false
}

func (p *Parser) expectKeyword(kw Keyword) {
	token := p.move()
	switch t := token.(type) {
	case *KeywordToken:
		if t.Keyword != kw {
			panic(p.vbSyntaxError(SyntaxError))
		}
	case *KeywordOrIdentifierToken:
		if t.Keyword != kw {
			panic(p.vbSyntaxError(SyntaxError))
		}
	default:
		panic(p.vbSyntaxError(SyntaxError))
	}
}

func (p *Parser) matchPunctuation(punc Punctuation) bool {
	t, ok := p.next.(*PunctuationToken)
	return ok && t.Type == punc
}

func (p *Parser) optPunctuation(punc Punctuation) bool {
	if p.matchPunctuation(punc) {
		p.move()
		return true
	}
	return false
}

func (p *Parser) expectPunctuation(punc Punctuation) {
	token := p.move()
	t, ok := token.(*PunctuationToken)
	if !ok || t.Type != punc {
		panic(p.vbSyntaxError(SyntaxError))
	}
}

func (p *Parser) matchIdentifier() bool {
	switch t := p.next.(type) {
	case *IdentifierToken, *KeywordOrIdentifierToken, *ExtendedIdentifierToken:
		return true
	case *KeywordToken:
		// Allow keywords to be identifiers unless they are strict block terminators
		// that rely on this check in parseMultiInlineStatement
		switch t.Keyword {
		case KeywordEnd, KeywordElse, KeywordElseIf, KeywordCase, KeywordNext, KeywordLoop, KeywordWEnd:
			return false
		}
		return true
	default:
		return false
	}
}

func (p *Parser) expectIdentifier() string {
	token := p.move()
	// fmt.Printf("DEBUG: expectIdentifier got %T: %v\n", token, token)
	switch t := token.(type) {
	case *IdentifierToken:
		return t.Name
	case *KeywordOrIdentifierToken:
		return t.Name
	case *ExtendedIdentifierToken:
		return t.Name
	case *KeywordToken:
		// Allow keywords to be identifiers unless they are strict block terminators
		switch t.Keyword {
		case KeywordEnd, KeywordElse, KeywordElseIf, KeywordCase, KeywordNext, KeywordLoop, KeywordWEnd:
			panic(p.vbSyntaxError(ExpectedIdentifier))
		}
		return t.Keyword.String()
	default:
		panic(p.vbSyntaxError(ExpectedIdentifier))
	}
}

func (p *Parser) optLineTermination() bool {
	if p.matchLineTermination() {
		p.move()
		return true
	}
	return false
}

func (p *Parser) optColonLineTermination() bool {
	if p.matchColonLineTermination() {
		p.move()
		return true
	}
	return false
}

func (p *Parser) expectLineTermination() {
	token := p.move()
	switch token.(type) {
	case *LineTerminationToken, *ColonLineTerminationToken:
		return
	default:
		panic(p.vbSyntaxError(SyntaxError))
	}
}

func (p *Parser) expectEofOrLineTermination() {
	token := p.move()
	switch token.(type) {
	case *LineTerminationToken, *ColonLineTerminationToken, *EOFToken:
		return
	default:
		panic(p.vbSyntaxError(SyntaxError))
	}
}

// Skip and expect methods

func (p *Parser) skipComments() {
	for _, ok := p.next.(*CommentToken); ok; _, ok = p.next.(*CommentToken) {
		p.move()
	}
}

func (p *Parser) skipCommentsAndNewlines() {
	for {
		switch p.next.(type) {
		case *CommentToken:
			p.move()
		case *LineTerminationToken:
			p.move()
		case *ColonLineTerminationToken:
			p.move()
		default:
			return
		}
	}
}

// Stub methods for statements (to be implemented)

func (p *Parser) parseBlockStatement(inGlobal bool) ast.Statement {
	p.createMarker() // marker

	// Handle stray comments/colons that should be treated as empty statements
	if _, ok := p.next.(*CommentToken); ok {
		p.skipComments()
		if inGlobal {
			p.expectEofOrLineTermination()
		} else {
			// If we hit EOF inside a block after a comment, we accept it as an empty statement
			// The surrounding loop (e.g. in parseSubDeclaration) will then check for matchEof()
			// and trigger the appropriate "Expected End ..." error instead of a generic SyntaxError.
			if !p.matchEof() {
				p.expectLineTermination()
			}
		}
		return nil
	}
	if _, ok := p.next.(*ColonLineTerminationToken); ok {
		p.move()
		return nil
	}

	var stmt ast.Statement
	if k, ok := p.next.(*KeywordToken); ok {
		switch k.Keyword {
		case KeywordIf:
			stmt = p.parseIfStatement()
		case KeywordFor:
			stmt = p.parseForOrForEachStatement()
		case KeywordDo:
			stmt = p.parseDoStatement()
		case KeywordSelect:
			stmt = p.parseSelectStatement()
		case KeywordWhile:
			stmt = p.parseWhileStatement()
		case KeywordWith:
			stmt = p.parseWithStatement()
		default:
			stmt = p.parseInlineStatement()
		}
	} else {
		stmt = p.parseInlineStatement()
	}

	p.skipComments()
	if inGlobal {
		p.expectEofOrLineTermination()
	} else {
		p.expectLineTermination()
	}

	return stmt
}

func (p *Parser) parsePublicOrPrivate(inGlobal, inlineOnly bool) ast.Statement {
	token1 := p.next
	p.move()
	token2 := p.next

	var isDefault bool
	if p.optKeyword(KeywordDefault) {
		isDefault = true
	}

	// Check if token1 is Private and Default is used
	if isDefault && token1.(*KeywordToken).Keyword == KeywordPrivate {
		panic(p.vbSyntaxError(SyntaxError))
	}

	var stmt ast.Statement
	if p.matchKeyword(KeywordSub) {
		modifier := p.getMethodAccessModifier(token1, token2)
		stmt = p.parseSubDeclaration(modifier, !inGlobal, inlineOnly)
	} else if p.matchKeyword(KeywordFunction) {
		modifier := p.getMethodAccessModifier(token1, token2)
		stmt = p.parseFunctionDeclaration(modifier, !inGlobal, inlineOnly)
	} else if (!inGlobal || !isDefault) && p.matchKeyword(KeywordProperty) {
		modifier := p.getMethodAccessModifier(token1, token2)
		stmt = p.parsePropertyDeclaration(modifier)
	} else if !isDefault && p.matchKeyword(KeywordConst) {
		stmt = p.parseConstDeclaration(ast.MemberAccessModifierNone)
	} else if !isDefault && p.matchIdentifier() {
		stmt = p.parseFieldsDeclaration(p.getFieldAccessModifier(token1))
	} else if isDefault {
		panic(p.vbSyntaxError(SyntaxError))
	} else {
		panic(p.vbSyntaxError(SyntaxError))
	}

	return stmt
}

func (p *Parser) parseClassDeclaration() ast.Statement {
	p.createMarker() // marker

	p.expectKeyword(KeywordClass)
	id := p.parseIdentifier()

	p.skipComments()
	p.expectLineTermination()

	stmt := ast.NewClassDeclaration(id)

	for {
		p.skipCommentsAndNewlines()
		if p.matchEof() || p.matchKeyword(KeywordEnd) {
			break
		}

		var memberStmt ast.Statement
		if p.matchKeyword(KeywordPublic) || p.matchKeyword(KeywordPrivate) {
			memberStmt = p.parsePublicOrPrivate(false, false)
		} else if p.matchKeyword(KeywordDim) {
			memberStmt = p.parseVariablesDeclaration()
		} else if p.matchKeyword(KeywordConst) {
			memberStmt = p.parseConstDeclaration(ast.MemberAccessModifierNone)
		} else if p.matchKeyword(KeywordFunction) {
			memberStmt = p.parseFunctionDeclaration(ast.MethodAccessModifierNone, true, false)
		} else if p.matchKeyword(KeywordSub) {
			memberStmt = p.parseSubDeclaration(ast.MethodAccessModifierNone, true, false)
		} else if p.matchKeyword(KeywordProperty) {
			memberStmt = p.parsePropertyDeclaration(ast.MethodAccessModifierNone)
		} else {
			panic(p.vbSyntaxError(SyntaxError))
		}

		p.skipComments()
		p.expectLineTermination()

		stmt.AddMember(memberStmt)
	}

	p.expectKeyword(KeywordEnd)
	p.expectKeyword(KeywordClass)

	return stmt
}

func (p *Parser) parseSubDeclaration(modifier ast.MethodAccessModifier, isMethod, inlineOnly bool) ast.Statement {
	return p.parseProcedure(KeywordSub, modifier, isMethod, inlineOnly, func(id *ast.Identifier, body ast.Statement) ast.Statement {
		if id.Name == "Class_Initialize" {
			return ast.NewInitializeSubDeclaration(modifier, body)
		} else if id.Name == "Class_Terminate" {
			return ast.NewTerminateSubDeclaration(modifier, body)
		}
		return ast.NewSubDeclaration(modifier, id, body)
	})
}

func (p *Parser) parseFunctionDeclaration(modifier ast.MethodAccessModifier, isMethod, inlineOnly bool) ast.Statement {
	return p.parseProcedure(KeywordFunction, modifier, isMethod, inlineOnly, func(id *ast.Identifier, body ast.Statement) ast.Statement {
		return ast.NewFunctionDeclaration(modifier, id, body)
	})
}

func (p *Parser) parseProcedure(kw Keyword, modifier ast.MethodAccessModifier, isMethod, inlineOnly bool, ctor func(*ast.Identifier, ast.Statement) ast.Statement) ast.Statement {
	p.createMarker() // marker
	line := p.startMarker.Line

	p.expectKeyword(kw)

	id := p.parseIdentifier()

	hasLParen := false
	var args []*ast.Parameter
	if p.optPunctuation(PunctLParen) {
		hasLParen = true
		args = p.parseParameterList()
		p.expectPunctuation(PunctRParen)
	}

	p.skipComments()
	inline := p.matchColonLineTermination() || !p.matchLineTermination()
	hasLParen = hasLParen || p.matchColonLineTermination()
	p.optLineTermination()

	if !inline && inlineOnly {
		panic(p.vbSyntaxError(SyntaxError))
	}

	if inline && !hasLParen {
		panic(p.vbSyntaxError(SyntaxError))
	}

	var body ast.Statement
	if inline {
		body = p.parseMultiInlineStatement(false, line)
	} else {
		list := ast.NewStatementList()
		for {
			p.skipCommentsAndNewlines()
			if p.matchEof() || p.matchKeyword(KeywordEnd) {
				break
			}

			list.Add(p.parseBlockStatement(false))
		}
		body = list
	}

	stmt := ctor(id, body)

	// Add Parameters to the procedure
	switch s := stmt.(type) {
	case *ast.SubDeclaration:
		s.Parameters = args
	case *ast.InitializeSubDeclaration:
		s.Parameters = args
	case *ast.TerminateSubDeclaration:
		s.Parameters = args
	case *ast.FunctionDeclaration:
		s.Parameters = args
	}

	p.expectKeyword(KeywordEnd)

	expectedCode := SyntaxError
	if kw == KeywordSub {
		expectedCode = SyntaxError
	} else {
		expectedCode = SyntaxError
	}
	p.expectKeyword(kw)
	if expectedCode != SyntaxError {
		// Just to use the variable
	}

	return stmt
}

func (p *Parser) parsePropertyDeclaration(modifier ast.MethodAccessModifier) ast.Statement {
	p.createMarker() // marker

	p.expectKeyword(KeywordProperty)

	var ctor func(ast.MethodAccessModifier, *ast.Identifier) ast.Statement
	if p.optKeyword(KeywordGet) {
		ctor = func(m ast.MethodAccessModifier, id *ast.Identifier) ast.Statement {
			return ast.NewPropertyGetDeclaration(m, id)
		}
	} else if p.optKeyword(KeywordSet) {
		ctor = func(m ast.MethodAccessModifier, id *ast.Identifier) ast.Statement {
			return ast.NewPropertySetDeclaration(m, id)
		}
	} else if p.optKeyword(KeywordLet) {
		ctor = func(m ast.MethodAccessModifier, id *ast.Identifier) ast.Statement {
			return ast.NewPropertyLetDeclaration(m, id)
		}
	} else {
		panic(p.vbSyntaxError(SyntaxError))
	}

	id := p.parseIdentifier()
	stmt := ctor(modifier, id)

	if p.optPunctuation(PunctLParen) {
		if params := p.parseParameterList(); params != nil {
			if decl, ok := stmt.(*ast.BasePropertyDeclaration); ok {
				decl.Parameters = params
			}
		}
		p.expectPunctuation(PunctRParen)
	}

	p.skipComments()
	p.expectLineTermination()

	for {
		p.skipCommentsAndNewlines()
		if p.matchEof() || p.matchKeyword(KeywordEnd) {
			break
		}

		if decl, ok := stmt.(*ast.BasePropertyDeclaration); ok {
			if p.matchKeyword(KeywordConst) {
				decl.Body = append(decl.Body, p.parseConstDeclaration(ast.MemberAccessModifierNone))
			} else {
				decl.Body = append(decl.Body, p.parseBlockStatement(false))
			}
		}
	}

	p.expectKeyword(KeywordEnd)
	p.expectKeyword(KeywordProperty)

	return stmt
}

func (p *Parser) parseInlineStatement() ast.Statement {
	p.createMarker() // marker

	var stmt ast.Statement
	if k, ok := p.next.(*KeywordToken); ok {
		switch k.Keyword {
		case KeywordDim:
			stmt = p.parseVariablesDeclaration()
		case KeywordReDim:
			stmt = p.parseReDimStatement()
		case KeywordConst:
			stmt = p.parseConstDeclaration(ast.MemberAccessModifierNone)
		case KeywordOn:
			stmt = p.parseOnErrorStatement()
		case KeywordExit:
			stmt = p.parseExitStatement()
		case KeywordErase:
			stmt = p.parseEraseStatement()
		case KeywordSet:
			stmt = p.parseSetAssignmentStatement()
		case KeywordCall:
			stmt = p.parseCallStatement()
		case KeywordIf:
			stmt = p.parseIfStatement()
		case KeywordOption:
			p.consumeOptionStatement()
			stmt = nil
		case KeywordPublic, KeywordPrivate:
			stmt = p.parsePublicOrPrivate(false, true)
		case KeywordSub, KeywordFunction:
			panic(p.vbSyntaxError(SyntaxError))
		default:
			// Unhandled keyword in inline statement context
			panic(p.vbSyntaxError(SyntaxError))
		}
	} else if p.inWithBlock && p.matchPunctuation(PunctDot) {
		stmt = p.parseAssignmentOrCallStatement()
	} else if p.matchIdentifier() {
		stmt = p.parseAssignmentOrCallStatement()
	} else {
		panic(p.vbSyntaxError(SyntaxError))
	}

	return stmt
}

func (p *Parser) parseIfStatement() ast.Statement {
	p.createMarker() // marker
	line := p.startMarker.Line

	p.expectKeyword(KeywordIf)
	test := p.parseExpression()
	p.expectKeyword(KeywordThen)

	p.skipComments()
	inline := !p.optLineTermination() || line == p.startMarker.Line

	var consequent, alternate ast.Statement

	if !inline {
		block := ast.NewStatementList()
		for {
			p.skipCommentsAndNewlines()
			if p.matchEof() || p.matchKeyword(KeywordEnd) || p.matchKeyword(KeywordElse) || p.matchKeyword(KeywordElseIf) {
				break
			}
			block.Add(p.parseBlockStatement(false))
		}

		if p.matchKeyword(KeywordElse) {
			alternate = p.parseElseStatement()
		} else if p.matchKeyword(KeywordElseIf) {
			block.Add(p.parseElseIfStatement())
		}

		p.expectKeyword(KeywordEnd)
		p.expectKeyword(KeywordIf)

		consequent = block
	} else {
		consequent = p.parseMultiInlineStatement(true, line)

		p.skipComments()

		// Check for explicit End If
		if p.optKeyword(KeywordEnd) {
			p.expectKeyword(KeywordIf)
		} else if p.optKeyword(KeywordElse) {
			alternate = p.parseMultiInlineStatement(false, line)
			p.skipComments()
		}
	}

	return ast.NewIfStatement(test, consequent, alternate)
}

func (p *Parser) parseElseIfStatement() ast.Statement {
	p.createMarker() // marker
	line := p.startMarker.Line

	p.expectKeyword(KeywordElseIf)

	test := p.parseExpression()
	p.expectKeyword(KeywordThen)

	p.skipComments()
	inline := !p.optLineTermination()

	var consequent, alternate ast.Statement
	if !inline {
		p.skipCommentsAndNewlines()

		block := ast.NewStatementList()
		for {
			p.skipCommentsAndNewlines()
			if p.matchEof() || p.matchKeyword(KeywordEnd) || p.matchKeyword(KeywordElse) || p.matchKeyword(KeywordElseIf) {
				break
			}

			block.Add(p.parseBlockStatement(false))
		}

		if p.matchKeyword(KeywordElse) {
			alternate = p.parseElseStatement()
		} else if p.matchKeyword(KeywordElseIf) {
			block.Add(p.parseElseIfStatement())
		}

		consequent = block
	} else {
		consequent = p.parseMultiInlineStatement(true, line)

		p.skipCommentsAndNewlines()
		if p.matchKeyword(KeywordElse) {
			alternate = p.parseElseStatement()
		} else if p.matchKeyword(KeywordElseIf) {
			block := ast.NewStatementList()
			block.Add(consequent)
			block.Add(p.parseElseIfStatement())
			consequent = block
		}
	}

	return ast.NewElseIfStatement(test, consequent, alternate)
}

func (p *Parser) parseElseStatement() ast.Statement {
	p.createMarker() // marker
	line := p.startMarker.Line

	p.expectKeyword(KeywordElse)

	p.skipComments()
	inline := !p.optLineTermination()

	var stmt ast.Statement
	if !inline {
		elseBlock := ast.NewStatementList()
		for {
			p.skipCommentsAndNewlines()
			if p.matchEof() || p.matchKeyword(KeywordEnd) {
				break
			}
			elseBlock.Add(p.parseBlockStatement(false))
		}
		stmt = elseBlock
	} else {
		stmt = p.parseMultiInlineStatement(false, line)
	}

	return stmt
}

func (p *Parser) parseMultiInlineStatement(matchElse bool, line int) ast.Statement {
	stmts := ast.NewStatementList()

	for {
		p.skipComments()

		// Stop if we reached EOF
		if p.matchEof() {
			break
		}

		// Only break for block terminators if we're clearly not in the middle of parsing an expression
		// Check if the next token is a terminator keyword AND we're not about to parse a member access
		if p.matchKeyword(KeywordEnd) || p.matchKeyword(KeywordElse) || p.matchKeyword(KeywordElseIf) {
			// Don't break if we're at an identifier, as it might be starting a statement like "response.end"
			// Only break if it's really a block terminator
			if !p.matchIdentifier() {
				break
			}
		}

		stmts.Add(p.parseInlineStatement())
		p.skipComments()

		// Separator between inline statements: consume colon or any line termination
		if p.optColonLineTermination() {
			continue
		}
		if p.matchLineTermination() {
			break
		}

		break
	}

	if stmts.Count() == 1 {
		return stmts.Get(0)
	}
	return stmts
}

func (p *Parser) parseForOrForEachStatement() ast.Statement {
	p.createMarker() // marker

	p.expectKeyword(KeywordFor)

	if p.optKeyword(KeywordEach) {
		return p.parseForEachStatement()
	}
	return p.parseForStatement()
}

func (p *Parser) parseForStatement() ast.Statement {
	p.createMarker() // marker

	id := p.parseIdentifier()
	p.expectPunctuation(PunctEqual)
	from := p.parseExpression()
	p.expectKeyword(KeywordTo)
	to := p.parseExpression()

	var step ast.Expression
	if p.optKeyword(KeywordStep) {
		step = p.parseExpression()
	}

	p.skipComments()
	p.expectLineTermination()

	stmt := ast.NewForStatement(id, from, to, step)
	for {
		p.skipCommentsAndNewlines()
		if p.matchEof() || p.matchKeyword(KeywordNext) {
			break
		}
		stmt.Body = append(stmt.Body, p.parseBlockStatement(false))
	}

	p.expectKeyword(KeywordNext)

	return stmt
}

func (p *Parser) parseForEachStatement() ast.Statement {
	p.createMarker() // marker

	id := p.parseIdentifier()
	p.expectKeyword(KeywordIn)
	in := p.parseExpression()

	p.skipComments()
	p.expectLineTermination()

	stmt := ast.NewForEachStatement(id, in)
	for {
		p.skipCommentsAndNewlines()
		if p.matchEof() || p.matchKeyword(KeywordNext) {
			break
		}
		stmt.Body = append(stmt.Body, p.parseBlockStatement(false))
	}

	p.expectKeyword(KeywordNext)

	return stmt
}

func (p *Parser) parseDoStatement() ast.Statement {
	p.createMarker() // marker

	p.expectKeyword(KeywordDo)

	loopType := ast.LoopTypeNone
	testType := ast.ConditionTestTypeNone
	var condition ast.Expression

	// Optional pre-test condition
	if p.optKeyword(KeywordWhile) {
		loopType = ast.LoopTypeWhile
		testType = ast.ConditionTestTypePreTest
		condition = p.parseExpression()
	} else if p.optKeyword(KeywordUntil) {
		loopType = ast.LoopTypeUntil
		testType = ast.ConditionTestTypePreTest
		condition = p.parseExpression()
	}

	p.skipComments()
	if !p.optLineTermination() {
		p.expectLineTermination()
	}

	body := []ast.Statement{}
	for {
		p.skipCommentsAndNewlines()
		if p.matchEof() || p.matchKeyword(KeywordLoop) {
			break
		}
		body = append(body, p.parseBlockStatement(false))
	}

	p.expectKeyword(KeywordLoop)

	// Optional post-test condition if not provided before
	if testType == ast.ConditionTestTypeNone {
		if p.optKeyword(KeywordWhile) {
			loopType = ast.LoopTypeWhile
			condition = p.parseExpression()
			testType = ast.ConditionTestTypePostTest
		} else if p.optKeyword(KeywordUntil) {
			loopType = ast.LoopTypeUntil
			condition = p.parseExpression()
			testType = ast.ConditionTestTypePostTest
		}
	}

	stmt := ast.NewDoStatement(loopType, testType, condition)
	stmt.Body = body

	return stmt
}

func (p *Parser) parseSelectStatement() ast.Statement {
	p.createMarker() // marker

	p.expectKeyword(KeywordSelect)
	p.expectKeyword(KeywordCase)

	condition := p.parseExpression()

	p.skipComments()
	p.expectLineTermination()

	stmt := ast.NewSelectStatement(condition)

	last := false
	for {
		p.skipCommentsAndNewlines()
		if !p.optKeyword(KeywordCase) {
			break
		}
		if last {
			panic(p.vbSyntaxError(SyntaxError))
		}

		caseStmt := ast.NewCaseStatement()
		caseLine := p.startMarker.Line
		if !p.optKeyword(KeywordElse) {
			// Helper to parse case value
			parseCaseValue := func() ast.Expression {
				p.optKeyword(KeywordIs) // Optional Is

				var op ast.BinaryOperation
				hasOp := false

				if p.optPunctuation(PunctEqual) {
					op = ast.BinaryOperationEqual
					hasOp = true
				} else if p.optPunctuation(PunctNotEqual) {
					op = ast.BinaryOperationNotEqual
					hasOp = true
				} else if p.optPunctuation(PunctLess) {
					op = ast.BinaryOperationLess
					hasOp = true
				} else if p.optPunctuation(PunctGreater) {
					op = ast.BinaryOperationGreater
					hasOp = true
				} else if p.optPunctuation(PunctLessOrEqual) {
					op = ast.BinaryOperationLessOrEqual
					hasOp = true
				} else if p.optPunctuation(PunctGreaterOrEqual) {
					op = ast.BinaryOperationGreaterOrEqual
					hasOp = true
				}

				if hasOp {
					rhs := p.parseExpression()
					return ast.NewBinaryExpression(op, ast.NewMissingValueExpression(), rhs)
				}

				expr := p.parseExpression()
				if p.optKeyword(KeywordTo) {
					end := p.parseExpression()
					call := ast.NewIndexOrCallExpression(ast.NewIdentifier("__RANGE__"))
					call.Indexes = []ast.Expression{expr, end}
					return call
				}
				return expr
			}

			caseStmt.Values = append(caseStmt.Values, parseCaseValue())
			for p.optPunctuation(PunctComma) {
				if !p.matchKeyword(KeywordElse) {
					caseStmt.Values = append(caseStmt.Values, parseCaseValue())
				} else {
					panic(p.vbSyntaxError(SyntaxError))
				}
			}
		} else {
			last = true
		}

		p.skipComments()

		// Check if there are inline statements after case value (on same line)
		inline := !p.optLineTermination() || caseLine == p.startMarker.Line

		if inline && !p.matchEof() && !p.matchKeyword(KeywordCase) && !p.matchKeyword(KeywordEnd) {
			// Parse inline statements separated by colons
			caseStmt.Body = append(caseStmt.Body, p.parseMultiInlineStatement(false, caseLine))
		} else {
			// Parse block statements on following lines
			for {
				p.skipCommentsAndNewlines()
				if p.matchEof() || p.matchKeyword(KeywordCase) || p.matchKeyword(KeywordEnd) {
					break
				}

				caseStmt.Body = append(caseStmt.Body, p.parseBlockStatement(false))
			}
		}

		stmt.Cases = append(stmt.Cases, caseStmt)
	}

	p.skipCommentsAndNewlines()

	p.expectKeyword(KeywordEnd)
	p.expectKeyword(KeywordSelect)

	return stmt
}

func (p *Parser) parseWhileStatement() ast.Statement {
	p.createMarker() // marker

	p.expectKeyword(KeywordWhile)

	condition := p.parseExpression()

	p.skipComments()
	p.expectLineTermination()

	stmt := ast.NewWhileStatement(condition)

	for {
		p.skipCommentsAndNewlines()
		if p.matchEof() || p.matchKeyword(KeywordWEnd) {
			break
		}

		stmt.Body = append(stmt.Body, p.parseBlockStatement(false))
	}

	p.expectKeyword(KeywordWEnd)

	return stmt
}

func (p *Parser) parseWithStatement() ast.Statement {
	p.createMarker() // marker

	p.expectKeyword(KeywordWith)
	expr := p.parseExpression()

	p.skipComments()
	p.expectLineTermination()

	p.inWithBlock = true
	stmt := ast.NewWithStatement(expr)

	for {
		p.skipCommentsAndNewlines()
		if p.matchEof() || p.matchKeyword(KeywordEnd) {
			break
		}

		stmt.Body = append(stmt.Body, p.parseBlockStatement(false))
	}

	p.inWithBlock = false

	p.expectKeyword(KeywordEnd)
	p.expectKeyword(KeywordWith)

	return stmt
}

// Expression parsing

func (p *Parser) parseExpression() ast.Expression {
	return p.parseBinaryExpression()
}

func (p *Parser) parseBinaryExpression() ast.Expression {
	return p.parseBinaryExpressionWithPrecedence(0)
}

func (p *Parser) parseBinaryExpressionWithPrecedence(minPrec int) ast.Expression {
	var expr ast.Expression

	// Handle Not operator as a prefix unary operator
	if p.optKeyword(KeywordNot) {
		notPrec := 15 // Not has precedence 15
		// Parse the operand with higher precedence than Not
		// If the operand is a binary expression, it must have higher precedence to be part of the Not
		operand := p.parseBinaryExpressionWithPrecedence(notPrec + 1)
		expr = ast.NewUnaryExpression(ast.UnaryOperationNot, operand)
	} else {
		expr = p.parseExpExpression()
	}

	for {
		op := p.next
		prec := p.binaryPrecedence(op)
		if prec <= minPrec {
			break
		}

		p.move()

		// Parse right side with higher precedence (left-associative)
		right := p.parseBinaryExpressionWithPrecedence(prec)

		expr = ast.NewBinaryExpression(p.getBinaryOperation(op), expr, right)
	}

	return expr
}

func (p *Parser) parseExpExpression() ast.Expression {
	expr := p.parseUnaryExpression()
	if p.optPunctuation(PunctExp) {
		right := p.parseExpExpression()
		expr = ast.NewBinaryExpression(ast.BinaryOperationExponentiation, expr, right)
	}

	return expr
}

func (p *Parser) parseUnaryExpression() ast.Expression {
	if p.optPunctuation(PunctMinus) {
		expr := p.parseUnaryExpression()
		return ast.NewUnaryExpression(ast.UnaryOperationMinus, expr)
	} else if p.optPunctuation(PunctPlus) {
		expr := p.parseUnaryExpression()
		return ast.NewUnaryExpression(ast.UnaryOperationPlus, expr)
	}

	return p.parseValueExpression()
}

func (p *Parser) parseValueExpression() ast.Expression {
	var expr ast.Expression

	// Check if it's a literal token using switch for proper type checking
	switch p.next.(type) {
	case *DecIntegerLiteralToken, *StringLiteralToken, *FloatLiteralToken,
		*DateLiteralToken, *EmptyLiteralToken, *NothingLiteralToken,
		*NullLiteralToken, *TrueLiteralToken, *FalseLiteralToken,
		*HexIntegerLiteralToken, *OctIntegerLiteralToken:
		expr = p.parseConstExpression()
	default:
		if p.optPunctuation(PunctLParen) {
			expr = p.parseExpression()
			p.expectPunctuation(PunctRParen)
		} else if p.optKeyword(KeywordNew) {
			expr = ast.NewNewExpression(p.parseLeftExpression())
		} else {
			expr = p.parseLeftExpression()
		}
	}

	return expr
}

func (p *Parser) parseLeftExpression() ast.Expression {
	var expr ast.Expression
	if p.inWithBlock && p.optPunctuation(PunctDot) {
		prop := p.parsePropertyId()
		expr = ast.NewWithMemberAccessExpression(prop)
	} else {
		expr = p.parseIdentifier()
	}

	for {
		if p.optPunctuation(PunctDot) {
			prop := p.parsePropertyId()
			expr = ast.NewMemberExpression(expr, prop)
		} else if p.optPunctuation(PunctLParen) {
			ix := ast.NewIndexOrCallExpression(expr)

			if p.optPunctuation(PunctComma) {
				ix.Indexes = append(ix.Indexes, ast.NewMissingValueExpression())
			} else if !p.matchPunctuation(PunctRParen) {
				ix.Indexes = append(ix.Indexes, p.parseExpression())
			}

			for p.optPunctuation(PunctComma) {
				if p.matchPunctuation(PunctComma) || p.matchPunctuation(PunctRParen) {
					ix.Indexes = append(ix.Indexes, ast.NewMissingValueExpression())
				} else {
					ix.Indexes = append(ix.Indexes, p.parseExpression())
				}
			}

			p.expectPunctuation(PunctRParen)
			expr = ix
		} else {
			break
		}
	}

	return expr
}

func (p *Parser) parsePropertyId() *ast.Identifier {
	var name string
	if p.matchIdentifier() {
		name = p.expectIdentifier()
	} else {
		name = p.expectAnyKeywordAsIdentifier()
	}
	return ast.NewIdentifier(name)
}

func (p *Parser) parseConstExpression() ast.Expression {
	var expr ast.Expression
	switch t := p.next.(type) {
	case *DecIntegerLiteralToken:
		expr = ast.NewIntegerLiteral(int64(t.Value))
		p.move()
	case *HexIntegerLiteralToken:
		expr = ast.NewIntegerLiteral(int64(t.Value))
		p.move()
	case *OctIntegerLiteralToken:
		expr = ast.NewIntegerLiteral(int64(t.Value))
		p.move()
	case *StringLiteralToken:
		expr = ast.NewStringLiteral(t.Value)
		p.move()
	case *FloatLiteralToken:
		expr = ast.NewFloatLiteral(t.Value)
		p.move()
	case *DateLiteralToken:
		expr = ast.NewDateLiteral(t.Value)
		p.move()
	case *EmptyLiteralToken:
		expr = ast.NewEmptyLiteral()
		p.move()
	case *NothingLiteralToken:
		expr = ast.NewNothingLiteral()
		p.move()
	case *NullLiteralToken:
		expr = ast.NewNullLiteral()
		p.move()
	case *TrueLiteralToken:
		expr = ast.NewBooleanLiteral(true)
		p.move()
	case *FalseLiteralToken:
		expr = ast.NewBooleanLiteral(false)
		p.move()
	default:
		panic(p.vbSyntaxError(SyntaxError))
	}

	return expr
}

func (p *Parser) parseConstInitExpression() ast.Expression {
	var expr ast.Expression
	if p.optPunctuation(PunctPlus) {
		expr = ast.NewUnaryExpression(ast.UnaryOperationPlus, p.parseConstInitExpression())
	} else if p.optPunctuation(PunctMinus) {
		expr = ast.NewUnaryExpression(ast.UnaryOperationMinus, p.parseConstInitExpression())
	} else if p.optPunctuation(PunctLParen) {
		expr = p.parseConstExpression()
		p.expectPunctuation(PunctRParen)
	} else {
		expr = p.parseConstExpression()
	}

	return expr
}

func (p *Parser) parseIdentifier() *ast.Identifier {
	name := p.expectIdentifier()
	if len(name) > ast.IdentifierMaxLength {
		panic(p.vbSyntaxError(SyntaxError))
	}
	return ast.NewIdentifier(name)
}

func (p *Parser) binaryPrecedence(token Token) int {
	if p, ok := token.(*PunctuationToken); ok {
		switch p.Type {
		case PunctStar, PunctSlash:
			return 50
		case PunctBackslash:
			return 49
		case PunctPlus, PunctMinus:
			return 47
		case PunctAmp:
			return 46
		case PunctEqual:
			return 30
		case PunctNotEqual:
			return 29
		case PunctLess:
			return 28
		case PunctGreater:
			return 29
		case PunctLessOrEqual:
			return 27
		case PunctGreaterOrEqual:
			return 26
		}
	}

	if k, ok := token.(*KeywordToken); ok {
		switch k.Keyword {
		case KeywordMod:
			return 48
		case KeywordIs:
			return 25
		case KeywordNot:
			return 15
		case KeywordAnd:
			return 10
		case KeywordOr:
			return 9
		case KeywordXor:
			return 8
		case KeywordEqv:
			return 7
		case KeywordImp:
			return 6
		}
	}

	return 0
}

func (p *Parser) getBinaryOperation(token Token) ast.BinaryOperation {
	if p, ok := token.(*PunctuationToken); ok {
		switch p.Type {
		case PunctExp:
			return ast.BinaryOperationExponentiation
		case PunctStar:
			return ast.BinaryOperationMultiplication
		case PunctSlash:
			return ast.BinaryOperationDivision
		case PunctBackslash:
			return ast.BinaryOperationIntDivision
		case PunctPlus:
			return ast.BinaryOperationAddition
		case PunctMinus:
			return ast.BinaryOperationSubtraction
		case PunctAmp:
			return ast.BinaryOperationConcatenation
		case PunctEqual:
			return ast.BinaryOperationEqual
		case PunctNotEqual:
			return ast.BinaryOperationNotEqual
		case PunctLess:
			return ast.BinaryOperationLess
		case PunctGreater:
			return ast.BinaryOperationGreater
		case PunctLessOrEqual:
			return ast.BinaryOperationLessOrEqual
		case PunctGreaterOrEqual:
			return ast.BinaryOperationGreaterOrEqual
		}
	}

	if k, ok := token.(*KeywordToken); ok {
		switch k.Keyword {
		case KeywordMod:
			return ast.BinaryOperationMod
		case KeywordIs:
			return ast.BinaryOperationIs
		case KeywordAnd:
			return ast.BinaryOperationAnd
		case KeywordOr:
			return ast.BinaryOperationOr
		case KeywordXor:
			return ast.BinaryOperationXor
		case KeywordEqv:
			return ast.BinaryOperationEqv
		case KeywordImp:
			return ast.BinaryOperationImp
		}
	}

	panic(p.vbSyntaxError(SyntaxError))
}

// Statement helper methods

func (p *Parser) parseVariablesDeclaration() ast.Statement {
	p.expectKeyword(KeywordDim)

	stmt := ast.NewVariablesDeclaration()
	stmt.Variables = append(stmt.Variables, p.parseVariableDeclaration())

	for {
		p.skipComments()
		if p.matchEof() || p.matchLineTermination() {
			break
		}

		p.expectPunctuation(PunctComma)
		stmt.Variables = append(stmt.Variables, p.parseVariableDeclaration())
	}

	return stmt
}

func (p *Parser) parseVariableDeclaration() *ast.VariableDeclaration {
	id := p.parseIdentifier()
	decl := ast.NewVariableDeclaration(id, false)

	if p.optPunctuation(PunctLParen) {
		if !p.matchPunctuation(PunctRParen) {
			dim := p.expectInteger()
			decl.ArrayDims = append(decl.ArrayDims, ast.NewIntegerLiteral(int64(dim)))

			for !p.matchEof() && !p.matchPunctuation(PunctRParen) {
				p.expectPunctuation(PunctComma)
				dim := p.expectInteger()
				decl.ArrayDims = append(decl.ArrayDims, ast.NewIntegerLiteral(int64(dim)))
			}
		} else {
			decl.IsDynamicArray = true
		}

		p.expectPunctuation(PunctRParen)
	}

	return decl
}

func (p *Parser) parseFieldDeclaration() *ast.FieldDeclaration {
	id := p.parseIdentifier()
	decl := ast.NewFieldDeclaration(id, false)

	if p.optPunctuation(PunctLParen) {
		if !p.matchPunctuation(PunctRParen) {
			dim := p.expectInteger()
			decl.ArrayDims = append(decl.ArrayDims, ast.NewIntegerLiteral(int64(dim)))

			for !p.matchEof() && !p.matchPunctuation(PunctRParen) {
				p.expectPunctuation(PunctComma)
				dim := p.expectInteger()
				decl.ArrayDims = append(decl.ArrayDims, ast.NewIntegerLiteral(int64(dim)))
			}
		} else {
			decl.IsDynamicArray = true
		}

		p.expectPunctuation(PunctRParen)
	}

	return decl
}

func (p *Parser) parseFieldsDeclaration(modifier ast.FieldAccessModifier) ast.Statement {
	stmt := ast.NewFieldsDeclaration(modifier)
	stmt.Fields = append(stmt.Fields, p.parseFieldDeclaration())

	for {
		p.skipComments()
		if p.matchEof() || p.matchLineTermination() {
			break
		}

		p.expectPunctuation(PunctComma)
		stmt.Fields = append(stmt.Fields, p.parseFieldDeclaration())
	}

	return stmt
}

func (p *Parser) parseConstDeclaration(modifier ast.MemberAccessModifier) ast.Statement {
	p.expectKeyword(KeywordConst)

	stmt := ast.NewConstsDeclaration(modifier)

	id := p.parseIdentifier()
	p.expectPunctuation(PunctEqual)
	expr := p.parseConstInitExpression()
	stmt.Declarations = append(stmt.Declarations, ast.NewConstDeclaration(id, expr))

	for {
		p.skipComments()
		if p.matchEof() || p.matchLineTermination() {
			break
		}

		p.expectPunctuation(PunctComma)
		id := p.parseIdentifier()
		p.expectPunctuation(PunctEqual)
		expr := p.parseConstInitExpression()
		stmt.Declarations = append(stmt.Declarations, ast.NewConstDeclaration(id, expr))
	}

	return stmt
}

func (p *Parser) parseReDimStatement() ast.Statement {
	p.expectKeyword(KeywordReDim)

	preserve := p.optKeyword(KeywordPreserve)
	result := ast.NewReDimStatement(preserve)

	redim := ast.NewReDimDeclaration(p.parseIdentifier())
	p.expectPunctuation(PunctLParen)
	redim.ArrayDims = append(redim.ArrayDims, p.parseExpression())
	for p.optPunctuation(PunctComma) {
		redim.ArrayDims = append(redim.ArrayDims, p.parseExpression())
	}
	p.expectPunctuation(PunctRParen)
	result.ReDims = append(result.ReDims, redim)

	for p.optPunctuation(PunctComma) {
		redim := ast.NewReDimDeclaration(p.parseIdentifier())
		p.expectPunctuation(PunctLParen)
		redim.ArrayDims = append(redim.ArrayDims, p.parseExpression())
		for p.optPunctuation(PunctComma) {
			redim.ArrayDims = append(redim.ArrayDims, p.parseExpression())
		}
		p.expectPunctuation(PunctRParen)
		result.ReDims = append(result.ReDims, redim)
	}

	return result
}

func (p *Parser) parseOnErrorStatement() ast.Statement {
	p.expectKeyword(KeywordOn)
	p.expectKeyword(KeywordError)

	var stmt ast.Statement
	if p.optKeyword(KeywordResume) {
		p.expectKeyword(KeywordNext)
		stmt = ast.NewOnErrorResumeNextStatement()
	} else {
		p.expectKeyword(KeywordGoto)
		i := p.expectInteger()
		if i != 0 {
			panic(p.vbSyntaxError(SyntaxError))
		}
		stmt = ast.NewOnErrorGoTo0Statement()
	}

	return stmt
}

func (p *Parser) parseExitStatement() ast.Statement {
	p.expectKeyword(KeywordExit)

	var stmt ast.Statement
	if p.optKeyword(KeywordDo) {
		stmt = ast.NewExitDoStatement()
	} else if p.optKeyword(KeywordFor) {
		stmt = ast.NewExitForStatement()
	} else if p.optKeyword(KeywordSub) {
		stmt = ast.NewExitSubStatement()
	} else if p.optKeyword(KeywordFunction) {
		stmt = ast.NewExitFunctionStatement()
	} else if p.optKeyword(KeywordProperty) {
		stmt = ast.NewExitPropertyStatement()
	} else {
		panic(p.vbSyntaxError(SyntaxError))
	}

	return stmt
}

func (p *Parser) parseEraseStatement() ast.Statement {
	p.expectKeyword(KeywordErase)
	id := p.parseIdentifier()
	return ast.NewEraseStatement(id)
}

func (p *Parser) parseSetAssignmentStatement() ast.Statement {
	p.optKeyword(KeywordSet)
	left := p.parseLeftExpression()
	p.expectPunctuation(PunctEqual)
	right := p.parseExpression()

	return ast.NewAssignmentStatement(left, right, true)
}

func (p *Parser) parseCallStatement() ast.Statement {
	p.expectKeyword(KeywordCall)

	return ast.NewCallStatement(p.parseLeftExpression())
}

func (p *Parser) parseAssignmentOrCallStatement() ast.Statement {
	left := p.parseLeftExpression()

	var stmt ast.Statement
	if p.optPunctuation(PunctEqual) {
		right := p.parseExpression()
		stmt = ast.NewAssignmentStatement(left, right, false)
	} else if p.matchLineTermination() {
		if indexExpr, ok := left.(*ast.IndexOrCallExpression); ok && len(indexExpr.Indexes) <= 1 {
			callstmt := ast.NewCallSubStatement(indexExpr.Object)
			if len(indexExpr.Indexes) != 0 {
				callstmt.Arguments = append(callstmt.Arguments, indexExpr.Indexes[0])
			}
			stmt = callstmt
		} else {
			stmt = ast.NewCallSubStatement(left)
		}
	} else if p.matchPunctuation(PunctComma) {
		if indexExpr, ok := left.(*ast.IndexOrCallExpression); ok && len(indexExpr.Indexes) <= 1 {
			if len(indexExpr.Indexes) == 0 {
				panic(p.vbSyntaxError(SyntaxError))
			}

			callstmt := ast.NewCallSubStatement(indexExpr.Object)
			callstmt.Arguments = append(callstmt.Arguments, indexExpr.Indexes[0])
			for p.optPunctuation(PunctComma) {
				isEmptyValue := p.matchPunctuation(PunctComma)
				var arg ast.Expression
				if isEmptyValue {
					arg = ast.NewMissingValueExpression()
				} else {
					arg = p.parseExpression()
				}
				callstmt.Arguments = append(callstmt.Arguments, arg)
			}
			stmt = callstmt
		} else if _, ok := left.(*ast.Identifier); ok {
			callstmt := ast.NewCallSubStatement(left)
			callstmt.Arguments = append(callstmt.Arguments, ast.NewMissingValueExpression())
			for p.optPunctuation(PunctComma) {
				isEmptyValue := p.matchPunctuation(PunctComma)
				var arg ast.Expression
				if isEmptyValue {
					arg = ast.NewMissingValueExpression()
				} else {
					arg = p.parseExpression()
				}
				callstmt.Arguments = append(callstmt.Arguments, arg)
			}
			stmt = callstmt
		} else {
			panic(p.vbSyntaxError(SyntaxError))
		}
	} else {
		callstmt := ast.NewCallSubStatement(left)

		p.skipComments()
		// Check for block terminators or else/elseif which end the statement list
		if !p.matchLineTermination() && !p.matchEof() &&
			!p.matchKeyword(KeywordEnd) && !p.matchKeyword(KeywordElse) && !p.matchKeyword(KeywordElseIf) &&
			!p.matchKeyword(KeywordNext) && !p.matchKeyword(KeywordLoop) && !p.matchKeyword(KeywordWEnd) {
			callstmt.Arguments = append(callstmt.Arguments, p.parseExpression())

			for p.optPunctuation(PunctComma) {
				isEmptyValue := p.matchPunctuation(PunctComma)
				var expr ast.Expression
				if isEmptyValue {
					expr = ast.NewMissingValueExpression()
				} else {
					expr = p.parseExpression()
				}
				callstmt.Arguments = append(callstmt.Arguments, expr)
			}
		}

		stmt = callstmt
	}

	return stmt
}

func (p *Parser) parseParameterList() []*ast.Parameter {
	var result []*ast.Parameter

	if p.matchKeyword(KeywordByRef) || p.matchKeyword(KeywordByVal) || p.matchIdentifier() {
		result = append(result, p.parseParameter())
	}

	if len(result) != 0 {
		for p.optPunctuation(PunctComma) {
			result = append(result, p.parseParameter())
		}
	}

	return result
}

func (p *Parser) parseParameter() *ast.Parameter {
	modifier := ast.ParameterModifierNone
	if p.optKeyword(KeywordByRef) {
		modifier = ast.ParameterModifierByRef
	} else if p.optKeyword(KeywordByVal) {
		modifier = ast.ParameterModifierByVal
	}

	id := p.parseIdentifier()

	parens := false
	if p.optPunctuation(PunctLParen) {
		p.expectPunctuation(PunctRParen)
		parens = true
	}

	return ast.NewParameter(id, modifier, parens)
}

func (p *Parser) getMethodAccessModifier(token1, token2 Token) ast.MethodAccessModifier {
	k, ok := token1.(*KeywordToken)
	if !ok {
		return ast.MethodAccessModifierNone
	}

	switch k.Keyword {
	case KeywordPrivate:
		return ast.MethodAccessModifierPrivate
	case KeywordPublic:
		k2, ok := token2.(*KeywordToken)
		if ok && k2.Keyword == KeywordDefault {
			return ast.MethodAccessModifierPublicDefault
		}
		return ast.MethodAccessModifierPublic
	}

	return ast.MethodAccessModifierNone
}

func (p *Parser) getMemberAccessModifier(token Token) ast.MemberAccessModifier {
	k, ok := token.(*KeywordToken)
	if !ok {
		return ast.MemberAccessModifierNone
	}

	switch k.Keyword {
	case KeywordPrivate:
		return ast.MemberAccessModifierPrivate
	case KeywordPublic:
		return ast.MemberAccessModifierPublic
	}

	return ast.MemberAccessModifierNone
}

func (p *Parser) getFieldAccessModifier(token Token) ast.FieldAccessModifier {
	k, ok := token.(*KeywordToken)
	if !ok {
		return ast.FieldAccessModifierNone
	}

	switch k.Keyword {
	case KeywordPrivate:
		return ast.FieldAccessModifierPrivate
	case KeywordPublic:
		return ast.FieldAccessModifierPublic
	}

	return ast.FieldAccessModifierNone
}

func (p *Parser) expectInteger() int {
	token := p.move()
	t, ok := token.(*DecIntegerLiteralToken)
	if !ok {
		panic(p.vbSyntaxError(SyntaxError))
	}
	return int(t.Value)
}

func (p *Parser) expectAnyKeywordAsIdentifier() string {
	token := p.move()
	if t, ok := token.(*KeywordToken); ok {
		return t.Keyword.String()
	}
	if t, ok := token.(*KeywordOrIdentifierToken); ok {
		return t.Name
	}

	panic(p.vbSyntaxError(SyntaxError))
}

// Error handling

func (p *Parser) vbSyntaxError(code VBSyntaxErrorCode) error {
	// Capture current token text if available
	tokenText := ""
	if p.next != nil {
		start := p.next.GetStart()
		end := p.next.GetEnd()
		if start >= 0 && end >= start {
			runes := []rune(p.lexer.Code)
			if end <= len(runes) {
				tokenText = string(runes[start:end])
			}
		}
	}

	// Capture the full line text from the lexer's state
	lineText := ""
	if p.lexer != nil {
		// Determine start and end of current line based on lexer's CurrentLineStart
		start := p.lexer.CurrentLineStart
		end := start
		for end < len([]rune(p.lexer.Code)) {
			ch := p.lexer.getChar(end)
			if ch == '\n' || ch == '\r' || ch == 0 {
				break
			}
			end++
		}
		runes := []rune(p.lexer.Code)
		if start >= 0 && start < len(runes) && end <= len(runes) && end >= start {
			lineText = string(runes[start:end])
		}
	}

	return NewVBSyntaxError(code, p.lastMarker.Line, p.lastMarker.Column, tokenText, lineText)
}

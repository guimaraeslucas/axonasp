package vbscript

import (
	"strconv"
	"strings"
)

// Lexer represents a VBScript lexical analyzer
type Lexer struct {
	Code             string
	Index            int
	CurrentLine      int
	CurrentLineStart int
	Length           int
	sb               strings.Builder
}

// NewLexer creates a new Lexer instance
func NewLexer(code string) *Lexer {
	if code == "" {
		return &Lexer{
			Code:             code,
			Index:            0,
			CurrentLine:      0,
			CurrentLineStart: 0,
			Length:           0,
		}
	}
	return &Lexer{
		Code:             code,
		Index:            0,
		CurrentLine:      1,
		CurrentLineStart: 0,
		Length:           len([]rune(code)),
	}
}

// LineIndex returns the column position in the current line
func (l *Lexer) LineIndex() int {
	return l.Index - l.CurrentLineStart
}

// NextToken returns the next token from the source code
func (l *Lexer) NextToken() Token {
	l.skipWhitespaces()

	if l.isEof() {
		return &EOFToken{
			BaseToken: BaseToken{
				Start:      l.Index,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
		}
	}

	c := l.getChar(l.Index)
	next := l.getChar(l.Index + 1)

	if IsLineTerminator(c) {
		return l.nextLineTermination()
	}

	comment := l.nextComment()
	if comment != nil {
		return comment
	}

	if IsIdentifierStart(c) {
		return l.nextIdentifier()
	}

	if c == '"' {
		return l.nextStringLiteral()
	}

	if c == '.' {
		if IsDecDigit(next) {
			return l.nextNumericLiteral()
		}
		return l.nextPunctuation()
	}

	if IsDecDigit(c) {
		return l.nextNumericLiteral()
	}

	if c == '&' {
		if CharEquals(next, 'h') || CharEquals(next, 'o') || IsDecDigit(next) {
			return l.nextNumericLiteral()
		}
		return l.nextPunctuation()
	}

	if c == '#' {
		return l.nextDateLiteral()
	}

	if c == '[' {
		return l.nextExtendedIdentifier()
	}

	return l.nextPunctuation()
}

// AsSequence returns all tokens as a slice
func (l *Lexer) AsSequence() []Token {
	var tokens []Token
	for !l.isEof() {
		tokens = append(tokens, l.NextToken())
	}
	return tokens
}

// Reset resets the lexer to the beginning
func (l *Lexer) Reset() {
	l.Index = 0
	if l.Length == 0 {
		l.CurrentLine = 0
	} else {
		l.CurrentLine = 1
	}
	l.CurrentLineStart = 0
	l.sb.Reset()
}

// Private helper methods

func (l *Lexer) getChar(pos int) rune {
	if pos < 0 || pos >= len(l.Code) {
		return rune(0)
	}
	runes := []rune(l.Code)
	if pos >= len(runes) {
		return rune(0)
	}
	return runes[pos]
}

func (l *Lexer) isEof() bool {
	return l.Index >= l.Length
}

func (l *Lexer) skipWhitespaces() {
	l.skipWSOnly()

	for !l.isEof() {
		c := l.getChar(l.Index)
		if c == '_' {
			l.Index++
			l.skipWSOnly()
			c = l.getChar(l.Index)
			if IsNewLine(c) {
				l.skipNewline()
			} else {
				panic(l.vbSyntaxError(InvalidCharacter))
			}
		} else {
			break
		}
	}
}

func (l *Lexer) skipWSOnly() {
	c := l.getChar(l.Index)
	for IsWhiteSpace(c) {
		l.Index++
		c = l.getChar(l.Index)
	}
}

func (l *Lexer) skipNewline() {
	c := l.getChar(l.Index)
	l.Index++
	if IsNewLine(c) {
		if c == '\r' && l.getChar(l.Index) == '\n' {
			l.Index++
		}
		l.CurrentLine++
		l.CurrentLineStart = l.Index
	}
}

func (l *Lexer) nextExtendedIdentifier() Token {
	start := l.Index
	l.Index++ // skip '['

	for !l.isEof() {
		c := l.getChar(l.Index)
		if IsExtendedIdentifier(c) {
			l.Index++
		} else {
			break
		}
	}

	if l.getChar(l.Index) != ']' {
		panic(l.vbSyntaxError(ExpectedRBracket))
	}

	runes := []rune(l.Code)
	name := string(runes[start : l.Index+1])

	l.Index++

	return &ExtendedIdentifierToken{
		IdentifierToken: IdentifierToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
			Name: name,
		},
	}
}

func (l *Lexer) nextDateLiteral() Token {
	start := l.Index
	l.Index++ // skip '#'
	l.sb.Reset()

	for !l.isEof() {
		c := l.getChar(l.Index)
		if c == '#' || IsNewLine(c) {
			break
		}
		l.sb.WriteRune(c)
		l.Index++
	}

	if l.getChar(l.Index) != '#' || l.sb.Len() == 0 {
		panic(l.vbSyntaxError(SyntaxError))
	}

	dateStr := l.sb.String()
	date, err := GetDate(dateStr)
	if err != nil {
		panic(l.vbSyntaxError(SyntaxError))
	}

	l.Index++

	return &DateLiteralToken{
		LiteralToken: LiteralToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
		},
		Value: date,
	}
}

func (l *Lexer) nextIdentifier() Token {
	start := l.Index
	id := l.getIdentifierName()

	var result Token

	switch {
	case CIEquals(id, "true"):
		result = &TrueLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case CIEquals(id, "null"):
		result = &NullLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case CIEquals(id, "false"):
		result = &FalseLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case CIEquals(id, "empty"):
		result = &EmptyLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case CIEquals(id, "nothing"):
		result = &NothingLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case IsKeyword(id):
		kw, _ := GetKeyword(id)
		result = &KeywordToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
			Keyword: kw,
			Name:    id,
		}
	case IsKeywordAsIdentifier(id):
		kw, _ := GetKeywordAsIdentifier(id)
		result = &KeywordOrIdentifierToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
			Keyword: kw,
			Name:    id,
		}
	default:
		result = &IdentifierToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
			Name: id,
		}
	}

	return result
}

func (l *Lexer) getIdentifierName() string {
	start := l.Index
	for !l.isEof() {
		c := l.getChar(l.Index)
		if IsIdentifier(c) {
			l.Index++
		} else {
			break
		}
	}
	runes := []rune(l.Code)
	return string(runes[start:l.Index])
}

func (l *Lexer) nextStringLiteral() Token {
	start := l.Index
	l.Index++ // skip opening quote
	l.sb.Reset()
	err := true

	for !l.isEof() {
		c := l.getChar(l.Index)
		if c == '"' {
			c = l.getChar(l.Index + 1)
			l.Index++
			if c == '"' {
				l.Index++
				l.sb.WriteRune(c)
			} else {
				err = false
				break
			}
		} else if IsNewLine(c) {
			break
		} else {
			l.sb.WriteRune(c)
			l.Index++
		}
	}

	if err {
		panic(l.vbSyntaxError(UnterminatedStringConstant))
	}

	return &StringLiteralToken{
		LiteralToken: LiteralToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index - 1,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
		},
		Value: l.sb.String(),
	}
}

func (l *Lexer) nextNumericLiteral() Token {
	start := l.Index
	c := l.getChar(l.Index)
	next := l.getChar(l.Index + 1)

	var decStr string
	var fstr strings.Builder

	if c != '.' {
		if c == '&' {
			if CharEquals(next, 'h') {
				return l.nextHexIntLiteral()
			} else if CharEquals(next, 'o') {
				return l.nextOctIntLiteralPrefix()
			} else if IsOctDigit(next) {
				return l.nextOctIntLiteral()
			} else {
				panic(l.vbSyntaxError(SyntaxError))
			}
		} else {
			decStr = l.getDecStr()
			if IsIdentifierStart(l.getChar(l.Index)) {
				panic(l.vbSyntaxError(ExpectedEndOfStatement))
			}
		}
	}

	c = l.getChar(l.Index)
	if c == '.' {
		l.Index++
		fstr.WriteRune('.')
		fstr.WriteString(l.getDecStr())
		c = l.getChar(l.Index)
	}

	if CharEquals(c, 'e') || CharEquals(c, 'E') {
		fstr.WriteRune('e')
		c = l.getChar(l.Index + 1)
		l.Index++
		if c == '+' || c == '-' {
			l.Index++
			fstr.WriteRune(c)
		}

		c = l.getChar(l.Index)
		if IsDecDigit(c) {
			fstr.WriteString(l.getDecStr())
		} else {
			panic(l.vbSyntaxError(InvalidNumber))
		}
	}

	c = l.getChar(l.Index)
	if IsIdentifierStart(c) {
		panic(l.vbSyntaxError(ExpectedEndOfStatement))
	}

	floatStr := fstr.String()
	if floatStr != "" && decStr != "" {
		floatStr = decStr + floatStr
	}

	if floatStr != "" {
		val, err := strconv.ParseFloat(floatStr, 64)
		if err != nil {
			panic(l.vbSyntaxError(InvalidNumber))
		}

		return &FloatLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
			Value: val,
		}
	}

	result := l.parseInteger(decStr, 10)
	result.SetStart(start)
	return result
}

func (l *Lexer) getDecStr() string {
	return l.getStrByPredicate(IsDecDigit)
}

func (l *Lexer) getOctStr() string {
	return l.getStrByPredicate(IsOctDigit)
}

func (l *Lexer) getHexStr() string {
	return l.getStrByPredicate(IsHexDigit)
}

func (l *Lexer) getStrByPredicate(predicate func(rune) bool) string {
	start := l.Index
	c := l.getChar(l.Index)
	for predicate(c) {
		l.Index++
		c = l.getChar(l.Index)
	}
	runes := []rune(l.Code)
	return string(runes[start:l.Index])
}

func (l *Lexer) parseInteger(str string, base int) Token {
	val, err := strconv.ParseInt(str, base, 64)
	if err != nil {
		if base == 8 || base == 16 {
			panic(l.vbSyntaxError(SyntaxError))
		}

		floatVal, err := strconv.ParseFloat(str, 64)
		if err != nil {
			panic(l.vbSyntaxError(InvalidNumber))
		}

		return &FloatLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
			Value: floatVal,
		}
	}

	var result Token
	switch base {
	case 8:
		result = &OctIntegerLiteralToken{
			DecIntegerLiteralToken: DecIntegerLiteralToken{
				LiteralToken: LiteralToken{
					BaseToken: BaseToken{
						End:        l.Index,
						LineNumber: l.CurrentLine,
						LineStart:  l.CurrentLineStart,
					},
				},
				Value: val,
			},
		}
	case 10:
		result = &DecIntegerLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
			Value: val,
		}
	case 16:
		result = &HexIntegerLiteralToken{
			DecIntegerLiteralToken: DecIntegerLiteralToken{
				LiteralToken: LiteralToken{
					BaseToken: BaseToken{
						End:        l.Index,
						LineNumber: l.CurrentLine,
						LineStart:  l.CurrentLineStart,
					},
				},
				Value: val,
			},
		}
	}

	return result
}

func (l *Lexer) nextOctIntLiteralPrefix() Token {
	start := l.Index
	l.Index += 2 // skip '&o'

	str := l.getOctStr()
	c := l.getChar(l.Index)

	if IsDecDigit(c) && !IsOctDigit(c) {
		panic(l.vbSyntaxError(SyntaxError))
	}

	if IsIdentifierStart(c) {
		panic(l.vbSyntaxError(ExpectedEndOfStatement))
	}

	result := l.parseInteger(str, 8)
	result.SetStart(start)

	return result
}

func (l *Lexer) nextOctIntLiteral() Token {
	start := l.Index
	l.Index++ // skip '&'

	str := l.getOctStr()
	c := l.getChar(l.Index)

	if IsDecDigit(c) && !IsOctDigit(c) {
		panic(l.vbSyntaxError(SyntaxError))
	}

	if IsIdentifierStart(c) {
		panic(l.vbSyntaxError(ExpectedEndOfStatement))
	}

	result := l.parseInteger(str, 8)
	result.SetStart(start)

	return result
}

func (l *Lexer) nextHexIntLiteral() Token {
	start := l.Index
	l.Index += 2 // skip '&h'

	str := l.getHexStr()
	c := l.getChar(l.Index)

	if IsIdentifierStart(c) {
		panic(l.vbSyntaxError(ExpectedEndOfStatement))
	}

	result := l.parseInteger(str, 16)
	result.SetStart(start)

	return result
}

func (l *Lexer) nextComment() Token {
	for !l.isEof() {
		c := l.getChar(l.Index)
		if c == '\'' {
			l.Index++
			return l.nextCommentBody(1, false)
		} else if CharEquals(c, 'r') {
			c2 := l.getChar(l.Index + 1)
			c3 := l.getChar(l.Index + 2)
			c4 := l.getChar(l.Index + 3)
			if CharEquals(c2, 'e') && CharEquals(c3, 'm') && IsWhiteSpace(c4) {
				l.Index += 3
				return l.nextCommentBody(3, true)
			}
			break
		} else {
			break
		}
	}
	return nil
}

func (l *Lexer) nextCommentBody(offset int, isRem bool) Token {
	start := l.Index - offset
	l.sb.Reset()

	for !l.isEof() {
		c := l.getChar(l.Index)
		if IsNewLine(c) {
			break
		}
		l.sb.WriteRune(c)
		l.Index++
	}

	return &CommentToken{
		BaseToken: BaseToken{
			Start:      start,
			End:        l.Index,
			LineNumber: l.CurrentLine,
			LineStart:  l.CurrentLineStart,
		},
		Comment: l.sb.String(),
		IsRem:   isRem,
	}
}

func (l *Lexer) nextLineTermination() Token {
	start := l.Index
	line := l.CurrentLine
	isColon := false

	for !l.isEof() {
		c := l.getChar(l.Index)
		if IsLineTerminator(c) {
			if c == '\r' && l.getChar(l.Index+1) == '\n' {
				l.Index++
			}

			l.Index++
			isColon = isColon || (c == ':')

			if c != ':' {
				l.CurrentLine++
				l.CurrentLineStart = l.Index
			}
		} else {
			break
		}

		l.skipWhitespaces()
	}

	var token Token
	if isColon && line == l.CurrentLine {
		token = &ColonLineTerminationToken{
			LineTerminationToken: LineTerminationToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine - 1,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	} else {
		token = &LineTerminationToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine - 1,
				LineStart:  l.CurrentLineStart,
			},
		}
	}

	return token
}

func (l *Lexer) nextPunctuation() Token {
	start := l.Index
	c := l.getChar(l.Index)
	next := l.getChar(l.Index + 1)

	var punctType *Punctuation

	switch c {
	case '(':
		p := PunctLParen
		punctType = &p
	case ')':
		p := PunctRParen
		punctType = &p
	case '.':
		p := PunctDot
		punctType = &p
	case ',':
		p := PunctComma
		punctType = &p
	case '+':
		p := PunctPlus
		punctType = &p
	case '-':
		p := PunctMinus
		punctType = &p
	case '/':
		p := PunctSlash
		punctType = &p
	case '\\':
		p := PunctBackslash
		punctType = &p
	case '*':
		p := PunctStar
		punctType = &p
	case '&':
		p := PunctAmp
		punctType = &p
	case '^':
		p := PunctExp
		punctType = &p
	case '=':
		if next == '<' {
			l.Index++
			p := PunctLessOrEqual
			punctType = &p
		} else if next == '>' {
			l.Index++
			p := PunctGreaterOrEqual
			punctType = &p
		} else {
			p := PunctEqual
			punctType = &p
		}
	case '<':
		if next == '=' {
			l.Index++
			p := PunctLessOrEqual
			punctType = &p
		} else if next == '>' {
			l.Index++
			p := PunctNotEqual
			punctType = &p
		} else {
			p := PunctLess
			punctType = &p
		}
	case '>':
		if next == '=' {
			l.Index++
			p := PunctGreaterOrEqual
			punctType = &p
		} else if next == '<' {
			l.Index++
			p := PunctNotEqual
			punctType = &p
		} else {
			p := PunctGreater
			punctType = &p
		}
	}

	if punctType == nil {
		panic(l.vbSyntaxError(InvalidCharacter))
	}

	l.Index++

	return &PunctuationToken{
		BaseToken: BaseToken{
			Start:      start,
			End:        l.Index,
			LineNumber: l.CurrentLine,
			LineStart:  l.CurrentLineStart,
		},
		Type: *punctType,
	}
}

func (l *Lexer) vbSyntaxError(code VBSyntaxErrorCode) error {
	// Capture the current offending character/token and full line text
	tokenText := ""
	if l.Index < len([]rune(l.Code)) {
		r := l.getChar(l.Index)
		if r != 0 {
			tokenText = string(r)
		}
	}

	lineText := l.currentLineText()

	return NewVBSyntaxError(code, l.CurrentLine, l.LineIndex(), tokenText, lineText)
}

// VBSyntaxError represents a VBScript syntax error
type VBSyntaxError struct {
	Code      VBSyntaxErrorCode
	Line      int
	Column    int
	TokenText string
	LineText  string
}

// NewVBSyntaxError creates a new syntax error
func NewVBSyntaxError(code VBSyntaxErrorCode, line, column int, tokenText, lineText string) *VBSyntaxError {
	return &VBSyntaxError{
		Code:      code,
		Line:      line,
		Column:    column,
		TokenText: tokenText,
		LineText:  lineText,
	}
}

// Error implements the error interface
func (e *VBSyntaxError) Error() string {
	msg := "VBScript syntax error"
	// Include numeric error code, position, token and line context when available
	if e.Line > 0 && e.Column >= 0 {
		msg = msg + " " + strconv.Itoa(int(e.Code)) + " at line " + strconv.Itoa(e.Line) + ", column " + strconv.Itoa(e.Column)
	}
	if e.TokenText != "" {
		msg = msg + ": '" + e.TokenText + "'"
	}
	if e.LineText != "" {
		msg = msg + "\n" + e.LineText
	}
	return msg
}

// currentLineText returns the full text of the current line
func (l *Lexer) currentLineText() string {
	if len(l.Code) == 0 {
		return ""
	}
	// Find start and end of current line using rune indices
	start := l.CurrentLineStart
	if start < 0 {
		start = 0
	}
	// Scan forward until newline or EOF
	end := start
	for end < len([]rune(l.Code)) {
		ch := l.getChar(end)
		if ch == '\n' || ch == '\r' || ch == 0 {
			break
		}
		end++
	}
	runes := []rune(l.Code)
	if start >= 0 && start < len(runes) && end <= len(runes) && end >= start {
		return string(runes[start:end])
	}
	return ""
}

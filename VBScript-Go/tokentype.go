package vbscript

// TokenType represents the type of a token
type TokenType int

const (
	TokenEOF TokenType = iota
	TokenLineTermination
	TokenComment
	TokenStringLiteral
	TokenDecIntegerLiteral
	TokenHexIntegerLiteral
	TokenOctIntegerLiteral
	TokenDateLiteral
	TokenFloatLiteral
	TokenTrueLiteral
	TokenFalseLiteral
	TokenNullLiteral
	TokenEmptyLiteral
	TokenNothingLiteral
	TokenIdentifier
)

// Punctuation represents punctuation token types
type Punctuation int

const (
	PunctLParen Punctuation = iota
	PunctRParen
	PunctLBracket
	PunctRBracket
	PunctDot
	PunctComma
	PunctPlus
	PunctMinus
	PunctSlash
	PunctBackslash
	PunctStar
	PunctAmp
	PunctExp
	PunctEqual
	PunctNotEqual
	PunctLess
	PunctLessOrEqual
	PunctGreater
	PunctGreaterOrEqual
)

// Keyword represents VBScript keywords
type Keyword int

const (
	KeywordStep Keyword = iota
	KeywordProperty
	KeywordExplicit
	KeywordError
	KeywordErase
	KeywordDefault
	KeywordAnd
	KeywordByRef
	KeywordByVal
	KeywordCall
	KeywordCase
	KeywordClass
	KeywordConst
	KeywordDim
	KeywordDo
	KeywordEach
	KeywordElse
	KeywordElseIf
	KeywordEnd
	KeywordEqv
	KeywordExit
	KeywordFor
	KeywordFunction
	KeywordGet
	KeywordGoto
	KeywordIf
	KeywordImp
	KeywordIn
	KeywordXor
	KeywordWith
	KeywordWhile
	KeywordWEnd
	KeywordTo
	KeywordUntil
	KeywordThen
	KeywordSub
	KeywordSet
	KeywordSelect
	KeywordResume
	KeywordReDim
	KeywordPublic
	KeywordPrivate
	KeywordPreserve
	KeywordOr
	KeywordOption
	KeywordOn
	KeywordNot
	KeywordNext
	KeywordNew
	KeywordMod
	KeywordLoop
	KeywordLet
	KeywordIs
	KeywordBinary
	KeywordCompare
	KeywordText
)

// String returns the string representation of a Keyword
func (k Keyword) String() string {
	switch k {
	case KeywordStep:
		return "Step"
	case KeywordProperty:
		return "Property"
	case KeywordExplicit:
		return "Explicit"
	case KeywordError:
		return "Error"
	case KeywordErase:
		return "Erase"
	case KeywordDefault:
		return "Default"
	case KeywordAnd:
		return "And"
	case KeywordByRef:
		return "ByRef"
	case KeywordByVal:
		return "ByVal"
	case KeywordCall:
		return "Call"
	case KeywordCase:
		return "Case"
	case KeywordClass:
		return "Class"
	case KeywordConst:
		return "Const"
	case KeywordDim:
		return "Dim"
	case KeywordDo:
		return "Do"
	case KeywordEach:
		return "Each"
	case KeywordElse:
		return "Else"
	case KeywordElseIf:
		return "ElseIf"
	case KeywordEnd:
		return "End"
	case KeywordEqv:
		return "Eqv"
	case KeywordExit:
		return "Exit"
	case KeywordFor:
		return "For"
	case KeywordFunction:
		return "Function"
	case KeywordGet:
		return "Get"
	case KeywordGoto:
		return "Goto"
	case KeywordIf:
		return "If"
	case KeywordImp:
		return "Imp"
	case KeywordIn:
		return "In"
	case KeywordXor:
		return "Xor"
	case KeywordWith:
		return "With"
	case KeywordWhile:
		return "While"
	case KeywordWEnd:
		return "WEnd"
	case KeywordTo:
		return "To"
	case KeywordUntil:
		return "Until"
	case KeywordThen:
		return "Then"
	case KeywordSub:
		return "Sub"
	case KeywordSet:
		return "Set"
	case KeywordSelect:
		return "Select"
	case KeywordResume:
		return "Resume"
	case KeywordReDim:
		return "ReDim"
	case KeywordPublic:
		return "Public"
	case KeywordPrivate:
		return "Private"
	case KeywordPreserve:
		return "Preserve"
	case KeywordOr:
		return "Or"
	case KeywordOption:
		return "Option"
	case KeywordOn:
		return "On"
	case KeywordNot:
		return "Not"
	case KeywordNext:
		return "Next"
	case KeywordNew:
		return "New"
	case KeywordMod:
		return "Mod"
	case KeywordLoop:
		return "Loop"
	case KeywordLet:
		return "Let"
	case KeywordIs:
		return "Is"
	case KeywordBinary:
		return "Binary"
	case KeywordCompare:
		return "Compare"
	case KeywordText:
		return "Text"
	default:
		return "Unknown"
	}
}

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

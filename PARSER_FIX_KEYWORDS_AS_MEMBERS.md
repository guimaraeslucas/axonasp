# Parser Fix: Keywords as Member Names

## Problem
The VBScript parser was rejecting code like `response.end` when it appeared in colon-separated inline statements (e.g., `x=1 : response.end`). While keywords after a dot worked in simple statements, they failed when combined with colon separators because the parser didn't properly recognize `ColonLineTerminationToken` as a valid line terminator in all contexts.

## Root Cause
The parser had three functions that checked for line termination:
- `matchLineTermination()` - Used in expression/statement parsing
- `expectLineTermination()` - Used after parsing block statements  
- `expectEofOrLineTermination()` - Used after parsing global statements

All three only recognized `LineTerminationToken` (newlines) but not `ColonLineTerminationToken` (colons). This caused the parser to fail when trying to parse statements after a colon, as it would try to continue parsing the current statement instead of recognizing the colon as a statement separator.

## Solution
Updated all three line termination functions to also accept `ColonLineTerminationToken`:

1. **`matchLineTermination()`** (line ~221):
   ```go
   func (p *Parser) matchLineTermination() bool {
       switch p.next.(type) {
       case *LineTerminationToken, *ColonLineTerminationToken:
           return true
       default:
           return false
       }
   }
   ```

2. **`expectLineTermination()`** (line ~327):
   ```go
   func (p *Parser) expectLineTermination() {
       token := p.move()
       switch token.(type) {
       case *LineTerminationToken, *ColonLineTerminationToken:
           return
       default:
           panic(p.vbSyntaxError(SyntaxError))
       }
   }
   ```

3. **`expectEofOrLineTermination()`** (line ~335):
   ```go
   func (p *Parser) expectEofOrLineTermination() {
       token := p.move()
       switch token.(type) {
       case *LineTerminationToken, *ColonLineTerminationToken, *EOFToken:
           return
       default:
           panic(p.vbSyntaxError(SyntaxError))
       }
   }
   ```

## Files Modified
- `VBScript-Go/parser.go`: Updated three line termination functions

## Testing
All existing tests pass, plus new tests verify:
- `response.end` works as a simple statement
- `x=1 : response.end` works with colon separator
- `select case` blocks with colon-separated statements containing keyword members

## Impact
This fix enables Classic ASP code patterns like those used in aspLite to parse correctly, where `Response.end` and similar keywords-as-members appear after colons in inline statements.

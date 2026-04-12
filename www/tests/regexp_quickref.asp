<html>
<head>
    <title>G3REGEXP - Quick Reference</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; }
        h1 { color: #0066cc; border-bottom: 3px solid #0066cc; padding: 10px 0; }
        h2 { color: #333; margin-top: 30px; border-left: 4px solid #0066cc; padding-left: 10px; }
        .code-block { background: #f4f4f4; padding: 15px; border-radius: 4px; border-left: 4px solid #0066cc; margin: 10px 0; font-family: 'Courier New'; overflow-x: auto; }
        .example { background: #e8f4f8; padding: 15px; border-radius: 4px; margin: 10px 0; }
        .property { color: #c41e3a; font-weight: bold; }
        .method { color: #0066cc; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        table td, table th { padding: 10px; border: 1px solid #ddd; text-align: left; }
        table th { background: #0066cc; color: white; }
        .link { color: #0066cc; text-decoration: none; }
        .link:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3REGEXP - Quick Reference Guide</h1>
        <p>Fast reference for using RegExp (Regular Expressions) in G3 AxonASP</p>

        <h2>Creating a RegExp Object</h2>
        <div class="code-block">
Dim regex
Set regex = Server.CreateObject("G3REGEXP")
' Or use alias:
Set regex = Server.CreateObject("REGEXP")
        </div>

        <h2>Properties (Using SetProperty/GetProperty)</h2>
        <table>
            <tr>
                <th>Property</th>
                <th>Type</th>
                <th>Description</th>
                <th>Example</th>
            </tr>
            <tr>
                <td><span class="property">Pattern</span></td>
                <td>String</td>
                <td>The regular expression pattern</td>
                <td><code>regex.SetProperty "pattern", "\d+"</code></td>
            </tr>
            <tr>
                <td><span class="property">IgnoreCase</span></td>
                <td>Boolean</td>
                <td>Case-insensitive matching (default: False)</td>
                <td><code>regex.SetProperty "ignorecase", True</code></td>
            </tr>
            <tr>
                <td><span class="property">Global</span></td>
                <td>Boolean</td>
                <td>Find all matches (default: False for first match only)</td>
                <td><code>regex.SetProperty "global", True</code></td>
            </tr>
            <tr>
                <td><span class="property">MultiLine</span></td>
                <td>Boolean</td>
                <td>^ and $ match line boundaries (default: False)</td>
                <td><code>regex.SetProperty "multiline", True</code></td>
            </tr>
        </table>

        <h2>Methods</h2>

        <h3><span class="method">Test(inputString)</span> → Boolean</h3>
        <p>Check if the pattern matches the string. Returns True or False.</p>
        <div class="example">
            <strong>Example:</strong>
            <div class="code-block">
Dim regex, result
Set regex = Server.CreateObject("G3REGEXP")
regex.SetProperty "pattern", "^\d{3}-\d{4}$"
result = regex.Test("555-1234")  ' Returns: True
            </div>
        </div>

        <h3><span class="method">Execute(inputString)</span> → RegExpMatches Collection</h3>
        <p>Find all (or first) matches and return a collection of matches.</p>
        <div class="example">
            <strong>Example:</strong>
            <div class="code-block">
Dim regex, matches, match, i
Set regex = Server.CreateObject("G3REGEXP")
regex.SetProperty "pattern", "\d+"
regex.SetProperty "global", True

Set matches = regex.Execute("I have 10 apples and 25 oranges")

For i = 0 To matches.Count - 1
    Set match = matches.Item(i)
    Response.Write match.Value & " at index " & match.FirstIndex & "<br>"
Next

' Output:
' 10 at index 8
' 25 at index 25
            </div>
        </div>

        <p><strong>Match Object Properties:</strong></p>
        <ul>
            <li><code>Value</code> - The matched text</li>
            <li><code>FirstIndex</code> or <code>Index</code> - Position in string (1-based)</li>
            <li><code>Length</code> - Length of the match</li>
            <li><code>Count</code> - Total number of matches (on collection)</li>
        </ul>

        <h3><span class="method">Replace(inputString, replacement)</span> → String</h3>
        <p>Replace matched patterns with replacement text.</p>
        <div class="example">
            <strong>Example:</strong>
            <div class="code-block">
Dim regex, result
Set regex = Server.CreateObject("G3REGEXP")
regex.SetProperty "pattern", "\d+"
regex.SetProperty "global", True

result = regex.Replace("Order 123 with 5 items", "[NUM]")
' Result: "Order [NUM] with [NUM] items"
            </div>
        </div>

        <h2>Common Patterns</h2>
        <table>
            <tr>
                <th>Pattern</th>
                <th>Description</th>
                <th>Example</th>
            </tr>
            <tr>
                <td><code>\d+</code></td>
                <td>One or more digits</td>
                <td>123, 2024</td>
            </tr>
            <tr>
                <td><code>[a-z]+</code></td>
                <td>One or more lowercase letters</td>
                <td>hello, world</td>
            </tr>
            <tr>
                <td><code>^text$</code></td>
                <td>Exact match (start and end)</td>
                <td>Matches "text" exactly</td>
            </tr>
            <tr>
                <td><code>\w+</code></td>
                <td>One or more word characters</td>
                <td>word, test_var</td>
            </tr>
            <tr>
                <td><code>[aeiou]</code></td>
                <td>Any vowel</td>
                <td>a, e, i, o, u</td>
            </tr>
            <tr>
                <td><code>\b\w+\b</code></td>
                <td>Complete words</td>
                <td>Uses word boundaries</td>
            </tr>
        </table>

        <h2>Practical Examples</h2>

        <h3>1. Email Validation</h3>
        <div class="example">
            <div class="code-block">
Function IsValidEmail(email)
    Dim regex
    Set regex = Server.CreateObject("G3REGEXP")
    regex.SetProperty "pattern", "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    IsValidEmail = regex.Test(email)
End Function

If IsValidEmail("user@example.com") Then
    Response.Write "Valid email"
End If
            </div>
        </div>

        <h3>2. Extract All Numbers from Text</h3>
        <div class="example">
            <div class="code-block">
Dim regex, matches, i
Set regex = Server.CreateObject("G3REGEXP")
regex.SetProperty "pattern", "\d+"
regex.SetProperty "global", True

Set matches = regex.Execute("Invoice #2024001 for $1500.99")

Response.Write "Numbers found: "
For i = 0 To matches.Count - 1
    If i > 0 Then Response.Write ", "
    Response.Write matches.Item(i).Value
Next
' Output: "Numbers found: 2024001, 1500, 99"
            </div>
        </div>

        <h3>3. Clean Up Text</h3>
        <div class="example">
            <div class="code-block">
Dim regex, result
Set regex = Server.CreateObject("G3REGEXP")

' Remove extra whitespace
regex.SetProperty "pattern", "\s+"
regex.SetProperty "global", True
result = regex.Replace("hello    world", " ")
' Result: "hello world"

' Remove HTML tags
regex.SetProperty "pattern", "<[^>]+>"
result = regex.Replace("<p>Hello</p>", "")
' Result: "Hello"
            </div>
        </div>

        <h3>4. Phone Number Validation</h3>
        <div class="example">
            <div class="code-block">
Dim regex
Set regex = Server.CreateObject("G3REGEXP")
regex.SetProperty "pattern", "^(\+1)?[-.\s]?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}$"

' All of these would match:
' 555-123-4567
' (555) 123-4567
' 5551234567
' +1-555-123-4567
            </div>
        </div>

        <h3>5. Replace with Multiple Occurrences</h3>
        <div class="example">
            <div class="code-block">
Dim regex, result
Set regex = Server.CreateObject("G3REGEXP")
regex.SetProperty "pattern", "[aeiou]"
regex.SetProperty "ignorecase", True
regex.SetProperty "global", True

result = regex.Replace("Hello World", "*")
' Result: "H*ll* W*rld"
            </div>
        </div>

        <h2>Common Metacharacters</h2>
        <table>
            <tr>
                <th>Metacharacter</th>
                <th>Meaning</th>
            </tr>
            <tr>
                <td><code>.</code></td>
                <td>Any character (except newline)</td>
            </tr>
            <tr>
                <td><code>*</code></td>
                <td>Zero or more of previous character</td>
            </tr>
            <tr>
                <td><code>+</code></td>
                <td>One or more of previous character</td>
            </tr>
            <tr>
                <td><code>?</code></td>
                <td>Zero or one of previous character</td>
            </tr>
            <tr>
                <td><code>^</code></td>
                <td>Start of string</td>
            </tr>
            <tr>
                <td><code>$</code></td>
                <td>End of string</td>
            </tr>
            <tr>
                <td><code>[...]</code></td>
                <td>Character class (one of ...)</td>
            </tr>
            <tr>
                <td><code>[^...]</code></td>
                <td>Negated character class (not ...)</td>
            </tr>
            <tr>
                <td><code>\d</code></td>
                <td>Digit (0-9)</td>
            </tr>
            <tr>
                <td><code>\w</code></td>
                <td>Word character (a-z, A-Z, 0-9, _)</td>
            </tr>
            <tr>
                <td><code>\s</code></td>
                <td>Whitespace</td>
            </tr>
        </table>

        <h2>Test Your Patterns</h2>
        <p>Visit <a href="/test_regexp.asp" class="link">test_regexp.asp</a> for a comprehensive test suite with examples of all features.</p>

        <footer style="margin-top: 40px; padding-top: 20px; border-top: 2px solid #0066cc; text-align: center; font-size: 12px; color: #666;">
            <p><strong>G3 AxonASP</strong> - RegExp Quick Reference<br>
            Powered by Go's regexp engine with VBScript compatibility</p>
        </footer>
    </div>
</body>
</html>

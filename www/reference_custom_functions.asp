<%
' ============================================================================
' G3 AxonASP - Custom Functions Quick Reference
' ============================================================================

' All 51 functions organized by category with quick syntax examples

%>
<!DOCTYPE html>
<html>
<head>
    <title>G3 AxonASP - Custom Functions Reference</title>
    <style>
        * { margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f5f5; padding: 20px; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0066cc; border-bottom: 3px solid #0066cc; padding-bottom: 10px; margin-bottom: 30px; }
        h2 { color: #333; background: #f0f0f0; padding: 10px 15px; margin-top: 30px; margin-bottom: 15px; border-left: 4px solid #0066cc; }
        .category { margin-bottom: 20px; }
        .function { background: #f9f9f9; border: 1px solid #ddd; border-radius: 4px; padding: 15px; margin-bottom: 10px; }
        .function-name { font-weight: bold; color: #0066cc; font-family: 'Courier New', monospace; font-size: 14px; }
        .function-desc { color: #555; margin-top: 5px; font-size: 14px; }
        .function-syntax { background: #f0f0f0; padding: 10px; margin-top: 8px; font-family: 'Courier New', monospace; font-size: 12px; border-left: 3px solid #0066cc; overflow-x: auto; }
        .php-equiv { color: #666; font-size: 12px; margin-top: 5px; }
        .badge { display: inline-block; background: #0066cc; color: white; padding: 2px 8px; border-radius: 3px; font-size: 11px; margin-right: 5px; }
        .badge-array { background: #28a745; }
        .badge-string { background: #fd7e14; }
        .badge-math { background: #6f42c1; }
        .badge-type { background: #dc3545; }
        .badge-datetime { background: #17a2b8; }
        .badge-crypto { background: #343a40; }
        .badge-util { background: #6c757d; }
        .count { font-size: 24px; font-weight: bold; color: #0066cc; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        table th { background: #f0f0f0; padding: 10px; text-align: left; border: 1px solid #ddd; font-weight: bold; }
        table td { padding: 10px; border: 1px solid #ddd; }
        table tr:nth-child(even) { background: #f9f9f9; }
        .footer { text-align: center; color: #999; margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📚 G3 AxonASP Custom Functions Reference</h1>
        <div class="count">51 Functions Total</div>

        <!-- ARRAYS -->
        <h2><span class="badge badge-array">ARRAYS</span> 9 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">AxArrayMerge(arr1, arr2, ...)</div>
                <div class="function-desc">Merge multiple arrays into one contiguous array</div>
                <div class="function-syntax">merged = AxArrayMerge(Array(1,2), Array(3,4))</div>
                <div class="php-equiv">PHP: array_merge()</div>
            </div>
            <div class="function">
                <div class="function-name">AxArrayContains(needle, haystack)</div>
                <div class="function-desc">Search for exact value in array/collection</div>
                <div class="function-syntax">found = AxArrayContains("item", myArray)</div>
                <div class="php-equiv">PHP: in_array()</div>
            </div>
            <div class="function">
                <div class="function-name">AxArrayMap(callback, array)</div>
                <div class="function-desc">Apply callback function to each element</div>
                <div class="function-syntax">result = AxArrayMap("FunctionName", myArray)</div>
                <div class="php-equiv">PHP: array_map()</div>
            </div>
            <div class="function">
                <div class="function-name">AxArrayFilter(callback, array)</div>
                <div class="function-desc">Filter array elements using callback</div>
                <div class="function-syntax">filtered = AxArrayFilter("FilterFunc", myArray)</div>
                <div class="php-equiv">PHP: array_filter()</div>
            </div>
            <div class="function">
                <div class="function-name">AxCount(array)</div>
                <div class="function-desc">Return array length (0 if empty/null)</div>
                <div class="function-syntax">count = AxCount(myArray)</div>
                <div class="php-equiv">PHP: count()</div>
            </div>
            <div class="function">
                <div class="function-name">AxExplode(delimiter, string [, limit])</div>
                <div class="function-desc">Split string by delimiter</div>
                <div class="function-syntax">parts = AxExplode(",", "one,two,three")</div>
                <div class="php-equiv">PHP: explode()</div>
            </div>
            <div class="function">
                <div class="function-name">AxArrayReverse(array)</div>
                <div class="function-desc">Reverse array element order</div>
                <div class="function-syntax">reversed = AxArrayReverse(myArray)</div>
                <div class="php-equiv">PHP: array_reverse()</div>
            </div>
            <div class="function">
                <div class="function-name">AxRange(start, end [, step])</div>
                <div class="function-desc">Create array of sequential values</div>
                <div class="function-syntax">numbers = AxRange(1, 10, 2)</div>
                <div class="php-equiv">PHP: range()</div>
            </div>
            <div class="function">
                <div class="function-name">AxImplode(glue, array)</div>
                <div class="function-desc">Join array elements with separator</div>
                <div class="function-syntax">text = AxImplode(" | ", myArray)</div>
                <div class="php-equiv">PHP: implode(), join()</div>
            </div>
        </div>

        <!-- STRINGS -->
        <h2><span class="badge badge-string">STRINGS</span> 9 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">AxStringReplace(search, replace, subject)</div>
                <div class="function-desc">Replace all occurrences (array support)</div>
                <div class="function-syntax">text = AxStringReplace("old", "new", myText)</div>
                <div class="php-equiv">PHP: str_replace()</div>
            </div>
            <div class="function">
                <div class="function-name">AxSprintf(format, ...args)</div>
                <div class="function-desc">C-style string formatting</div>
                <div class="function-syntax">text = AxSprintf("Age: %d, Score: %f", 25, 95.5)</div>
                <div class="php-equiv">PHP: sprintf()</div>
            </div>
            <div class="function">
                <div class="function-name">AxPad(string, length, pad_string, pad_type)</div>
                <div class="function-desc">Pad string to length (0=left, 1=right, 2=both)</div>
                <div class="function-syntax">padded = AxPad("5", 5, "0", 0)</div>
                <div class="php-equiv">PHP: str_pad()</div>
            </div>
            <div class="function">
                <div class="function-name">AxRepeat(string, multiplier)</div>
                <div class="function-desc">Repeat string N times</div>
                <div class="function-syntax">stars = AxRepeat("*", 10)</div>
                <div class="php-equiv">PHP: str_repeat()</div>
            </div>
            <div class="function">
                <div class="function-name">AxUcFirst(string)</div>
                <div class="function-desc">Uppercase first character</div>
                <div class="function-syntax">text = AxUcFirst("hello world")</div>
                <div class="php-equiv">PHP: ucfirst()</div>
            </div>
            <div class="function">
                <div class="function-name">AxWordCount(string, format)</div>
                <div class="function-desc">Count words (0=count, 1=array)</div>
                <div class="function-syntax">count = AxWordCount("The quick brown fox", 0)</div>
                <div class="php-equiv">PHP: str_word_count()</div>
            </div>
            <div class="function">
                <div class="function-name">AxNewLineToBr(string)</div>
                <div class="function-desc">Convert newlines to HTML &lt;br&gt;</div>
                <div class="function-syntax">html = AxNewLineToBr(textWithNewlines)</div>
                <div class="php-equiv">PHP: nl2br()</div>
            </div>
            <div class="function">
                <div class="function-name">AxTrim(string [, chars])</div>
                <div class="function-desc">Remove whitespace or custom chars</div>
                <div class="function-syntax">clean = AxTrim("  hello world  ")</div>
                <div class="php-equiv">PHP: trim()</div>
            </div>
            <div class="function">
                <div class="function-name">AxStringGetCsv(string [, delimiter])</div>
                <div class="function-desc">Parse CSV string into array</div>
                <div class="function-syntax">values = AxStringGetCsv("col1,col2,col3")</div>
                <div class="php-equiv">PHP: str_getcsv()</div>
            </div>
        </div>

        <!-- MATH -->
        <h2><span class="badge badge-math">MATH</span> 6 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">AxCeil(number)</div>
                <div class="function-desc">Round up to nearest integer</div>
                <div class="function-syntax">result = AxCeil(4.3)</div>
                <div class="php-equiv">PHP: ceil()</div>
            </div>
            <div class="function">
                <div class="function-name">AxFloor(number)</div>
                <div class="function-desc">Round down to nearest integer</div>
                <div class="function-syntax">result = AxFloor(4.8)</div>
                <div class="php-equiv">PHP: floor()</div>
            </div>
            <div class="function">
                <div class="function-name">AxMax(...values)</div>
                <div class="function-desc">Return maximum value</div>
                <div class="function-syntax">max = AxMax(10, 20, 15, 5)</div>
                <div class="php-equiv">PHP: max()</div>
            </div>
            <div class="function">
                <div class="function-name">AxMin(...values)</div>
                <div class="function-desc">Return minimum value</div>
                <div class="function-syntax">min = AxMin(10, 20, 15, 5)</div>
                <div class="php-equiv">PHP: min()</div>
            </div>
            <div class="function">
                <div class="function-name">AxRand([min, max])</div>
                <div class="function-desc">Random integer</div>
                <div class="function-syntax">num = AxRand(1, 100)</div>
                <div class="php-equiv">PHP: rand()</div>
            </div>
            <div class="function">
                <div class="function-name">AxNumberFormat(number, decimals, dec_point, thousands_sep)</div>
                <div class="function-desc">Format number with separators</div>
                <div class="function-syntax">formatted = AxNumberFormat(1234567.89, 2, ".", ",")</div>
                <div class="php-equiv">PHP: number_format()</div>
            </div>
        </div>

        <!-- TYPE CHECKING -->
        <h2><span class="badge badge-type">TYPE CHECKING</span> 6 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">AxIsInt(value)</div>
                <div class="function-desc">Check if value is integer type</div>
                <div class="function-syntax">If AxIsInt(myValue) Then</div>
                <div class="php-equiv">PHP: is_int()</div>
            </div>
            <div class="function">
                <div class="function-name">AxIsFloat(value)</div>
                <div class="function-desc">Check if value is float type</div>
                <div class="function-syntax">If AxIsFloat(myValue) Then</div>
                <div class="php-equiv">PHP: is_float()</div>
            </div>
            <div class="function">
                <div class="function-name">AxCTypeAlpha(string)</div>
                <div class="function-desc">All characters alphabetic</div>
                <div class="function-syntax">If AxCTypeAlpha("hello") Then</div>
                <div class="php-equiv">PHP: ctype_alpha()</div>
            </div>
            <div class="function">
                <div class="function-name">AxCTypeAlnum(string)</div>
                <div class="function-desc">All characters alphanumeric</div>
                <div class="function-syntax">If AxCTypeAlnum("abc123") Then</div>
                <div class="php-equiv">PHP: ctype_alnum()</div>
            </div>
            <div class="function">
                <div class="function-name">AxEmpty(value)</div>
                <div class="function-desc">Check if empty (null, "", 0, False, empty array)</div>
                <div class="function-syntax">If AxEmpty(myValue) Then</div>
                <div class="php-equiv">PHP: empty()</div>
            </div>
            <div class="function">
                <div class="function-name">AxIsset(value)</div>
                <div class="function-desc">Check if value is set (not null/empty)</div>
                <div class="function-syntax">If AxIsset(myValue) Then</div>
                <div class="php-equiv">PHP: isset()</div>
            </div>
        </div>

        <!-- DATE/TIME -->
        <h2><span class="badge badge-datetime">DATE/TIME</span> 2 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">AxTime()</div>
                <div class="function-desc">Current Unix timestamp</div>
                <div class="function-syntax">timestamp = AxTime()</div>
                <div class="php-equiv">PHP: time()</div>
            </div>
            <div class="function">
                <div class="function-name">AxDate(format [, timestamp])</div>
                <div class="function-desc">Format date (Y-m-d, H:i:s, etc)</div>
                <div class="function-syntax">date = AxDate("Y-m-d H:i:s")</div>
                <div class="php-equiv">PHP: date()</div>
            </div>
        </div>

        <!-- CRYPTO/ENCODING -->
        <h2><span class="badge badge-crypto">CRYPTO & ENCODING</span> 10 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">AxMd5(string)</div>
                <div class="function-desc">MD5 hash</div>
                <div class="function-syntax">hash = AxMd5("password")</div>
                <div class="php-equiv">PHP: md5()</div>
            </div>
            <div class="function">
                <div class="function-name">AxSha1(string)</div>
                <div class="function-desc">SHA1 hash</div>
                <div class="function-syntax">hash = AxSha1("password")</div>
                <div class="php-equiv">PHP: sha1()</div>
            </div>
            <div class="function">
                <div class="function-name">AxHash(algorithm, string)</div>
                <div class="function-desc">Hash with algorithm (sha256, etc)</div>
                <div class="function-syntax">hash = AxHash("sha256", "password")</div>
                <div class="php-equiv">PHP: hash()</div>
            </div>
            <div class="function">
                <div class="function-name">AxBase64Encode(string)</div>
                <div class="function-desc">Base64 encode</div>
                <div class="function-syntax">encoded = AxBase64Encode("text")</div>
                <div class="php-equiv">PHP: base64_encode()</div>
            </div>
            <div class="function">
                <div class="function-name">AxBase64Decode(string)</div>
                <div class="function-desc">Base64 decode</div>
                <div class="function-syntax">decoded = AxBase64Decode(encoded)</div>
                <div class="php-equiv">PHP: base64_decode()</div>
            </div>
            <div class="function">
                <div class="function-name">AxUrlDecode(string)</div>
                <div class="function-desc">URL decode</div>
                <div class="function-syntax">text = AxUrlDecode("Hello%20World")</div>
                <div class="php-equiv">PHP: urldecode()</div>
            </div>
            <div class="function">
                <div class="function-name">AxRawUrlDecode(string)</div>
                <div class="function-desc">Raw URL decode</div>
                <div class="function-syntax">text = AxRawUrlDecode("Hello+World")</div>
                <div class="php-equiv">PHP: rawurldecode()</div>
            </div>
            <div class="function">
                <div class="function-name">AxRgbToHex(red, green, blue)</div>
                <div class="function-desc">Convert RGB to hex color</div>
                <div class="function-syntax">hex = AxRgbToHex(255, 128, 0)</div>
                <div class="php-equiv">Custom function</div>
            </div>
            <div class="function">
                <div class="function-name">AxHtmlSpecialChars(string)</div>
                <div class="function-desc">Escape HTML special characters</div>
                <div class="function-syntax">safe = AxHtmlSpecialChars(userInput)</div>
                <div class="php-equiv">PHP: htmlspecialchars()</div>
            </div>
            <div class="function">
                <div class="function-name">AxStripTags(string)</div>
                <div class="function-desc">Remove HTML tags</div>
                <div class="function-syntax">plain = AxStripTags(htmlContent)</div>
                <div class="php-equiv">PHP: strip_tags()</div>
            </div>
        </div>

        <!-- VALIDATION -->
        <h2><span class="badge badge-util">VALIDATION</span> 2 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">AxFilterValidateIp(ip)</div>
                <div class="function-desc">Validate IP address</div>
                <div class="function-syntax">If AxFilterValidateIp("192.168.1.1") Then</div>
                <div class="php-equiv">PHP: filter_var() with FILTER_VALIDATE_IP</div>
            </div>
            <div class="function">
                <div class="function-name">AxFilterValidateEmail(email)</div>
                <div class="function-desc">Validate email address</div>
                <div class="function-syntax">If AxFilterValidateEmail("user@example.com") Then</div>
                <div class="php-equiv">PHP: filter_var() with FILTER_VALIDATE_EMAIL</div>
            </div>
        </div>

        <!-- REQUEST -->
        <h2><span class="badge badge-util">REQUEST</span> 3 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">AxGetRequest()</div>
                <div class="function-desc">Get all parameters (GET + POST merged)</div>
                <div class="function-syntax">params = AxGetRequest()</div>
                <div class="php-equiv">PHP: $_REQUEST</div>
            </div>
            <div class="function">
                <div class="function-name">AxGetGet()</div>
                <div class="function-desc">Get only GET parameters</div>
                <div class="function-syntax">params = AxGetGet()</div>
                <div class="php-equiv">PHP: $_GET</div>
            </div>
            <div class="function">
                <div class="function-name">AxGetPost()</div>
                <div class="function-desc">Get only POST parameters</div>
                <div class="function-syntax">params = AxGetPost()</div>
                <div class="php-equiv">PHP: $_POST</div>
            </div>
        </div>

        <!-- UTILITIES -->
        <h2><span class="badge badge-util">UTILITIES</span> 4 Functions</h2>
        <div class="category">
            <div class="function">
                <div class="function-name">Document.Write(string)</div>
                <div class="function-desc">Safe HTML-encoded Response.Write</div>
                <div class="function-syntax">Document.Write userInput</div>
                <div class="php-equiv">Custom function - prevents XSS</div>
            </div>
            <div class="function">
                <div class="function-name">AxVarDump(value)</div>
                <div class="function-desc">Debug output any value recursively</div>
                <div class="function-syntax">AxVarDump myArray</div>
                <div class="php-equiv">PHP: var_dump()</div>
            </div>
            <div class="function">
                <div class="function-name">AxGenerateGuid()</div>
                <div class="function-desc">Generate unique GUID</div>
                <div class="function-syntax">guid = AxGenerateGuid()</div>
                <div class="php-equiv">PHP: uniqid(), uuid_v4()</div>
            </div>
            <div class="function">
                <div class="function-name">AxBuildQueryString(dictionary)</div>
                <div class="function-desc">Build URL query string from dict</div>
                <div class="function-syntax">query = AxBuildQueryString(params)</div>
                <div class="php-equiv">PHP: http_build_query()</div>
            </div>
        </div>

        <!-- SUMMARY TABLE -->
        <h2>Quick Reference Table</h2>
        <table>
            <tr>
                <th>Category</th>
                <th>Functions</th>
                <th>Count</th>
            </tr>
            <tr>
                <td>Arrays</td>
                <td>AxArrayMerge, AxArrayContains, AxArrayMap, AxArrayFilter, AxCount, AxExplode, AxArrayReverse, AxRange, AxImplode</td>
                <td>9</td>
            </tr>
            <tr>
                <td>Strings</td>
                <td>AxStringReplace, AxSprintf, AxPad, AxRepeat, AxUcFirst, AxWordCount, AxNewLineToBr, AxTrim, AxStringGetCsv</td>
                <td>9</td>
            </tr>
            <tr>
                <td>Math</td>
                <td>AxCeil, AxFloor, AxMax, AxMin, AxRand, AxNumberFormat</td>
                <td>6</td>
            </tr>
            <tr>
                <td>Type Checking</td>
                <td>AxIsInt, AxIsFloat, AxCTypeAlpha, AxCTypeAlnum, AxEmpty, AxIsset</td>
                <td>6</td>
            </tr>
            <tr>
                <td>Date/Time</td>
                <td>AxTime, AxDate</td>
                <td>2</td>
            </tr>
            <tr>
                <td>Crypto & Encoding</td>
                <td>AxMd5, AxSha1, AxHash, AxBase64Encode, AxBase64Decode, AxUrlDecode, AxRawUrlDecode, AxRgbToHex, AxHtmlSpecialChars, AxStripTags</td>
                <td>10</td>
            </tr>
            <tr>
                <td>Validation</td>
                <td>AxFilterValidateIp, AxFilterValidateEmail</td>
                <td>2</td>
            </tr>
            <tr>
                <td>Request</td>
                <td>AxGetRequest, AxGetGet, AxGetPost</td>
                <td>3</td>
            </tr>
            <tr>
                <td>Utilities</td>
                <td>Document.Write, AxVarDump, AxGenerateGuid, AxBuildQueryString</td>
                <td>4</td>
            </tr>
            <tr style="background: #e8f4f8; font-weight: bold;">
                <td>TOTAL</td>
                <td colspan="1"></td>
                <td>51</td>
            </tr>
        </table>

        <div class="footer">
            <p><strong>G3 AxonASP Custom Functions Reference</strong></p>
            <p>All functions follow VBScript conventions with Ax prefix and PascalCase naming.</p>
            <p>For complete documentation, see CUSTOM_FUNCTIONS.md or CUSTOM_FUNCTIONS_PT-BR.md</p>
            <p>Generated: 17 January 2026</p>
        </div>
    </div>
</body>
</html>

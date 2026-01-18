<%@ Page Language="VBScript" %>
<%
' G3 AxonASP Custom Functions - Quick Reference Examples

' ============================================================================
' ARRAY FUNCTIONS EXAMPLES
' ============================================================================

' Merge arrays
Dim arr1, arr2, result
arr1 = Array(1, 2, 3)
arr2 = Array(4, 5, 6)
result = AxArrayMerge(arr1, arr2)
' result = [1, 2, 3, 4, 5, 6]

' Check if value exists in array
Dim fruits
fruits = Array("apple", "banana", "cherry")
If AxArrayContains("banana", fruits) Then
    ' Banana found!
End If

' Count array items
Dim count
count = AxCount(fruits)  ' 3

' Split string into array
Dim parts
parts = AxExplode(",", "one,two,three")
' parts = ["one", "two", "three"]

' Reverse array
Dim numbers, reversed
numbers = Array(1, 2, 3, 4, 5)
reversed = AxArrayReverse(numbers)
' reversed = [5, 4, 3, 2, 1]

' Create range
Dim range_result
range_result = AxRange(1, 10, 2)
' range_result = [1, 3, 5, 7, 9]

' Join array with separator
Dim joined
joined = AxImplode(" | ", Array("a", "b", "c"))
' joined = "a | b | c"

' ============================================================================
' STRING FUNCTIONS EXAMPLES
' ============================================================================

' Replace text
Dim original, replaced
original = "Hello World"
replaced = AxStringReplace("World", "ASP", original)
' replaced = "Hello ASP"

' Format string (printf-style)
Dim formatted
formatted = AxSprintf("Name: %s, Age: %d, Score: %f", "John", 30, 95.5)
' formatted = "Name: John, Age: 30, Score: 95.500000"

' Pad string with zeros
Dim padded
padded = AxPad("5", 5, "0", 0)  ' Left pad
' padded = "00005"

' Repeat string
Dim repeated
repeated = AxRepeat("*", 10)
' repeated = "**********"

' Uppercase first letter
Dim ucfirst_result
ucfirst_result = AxUcFirst("hello world")
' ucfirst_result = "Hello world"

' Count words
Dim word_count
word_count = AxWordCount("The quick brown fox", 0)
' word_count = 4

' Convert newlines to HTML
Dim text_with_newlines, html_text
text_with_newlines = "Line 1" & vbCrLf & "Line 2" & vbCrLf & "Line 3"
html_text = AxNewLineToBr(text_with_newlines)
' html_text contains <br> tags instead of newlines

' Trim whitespace
Dim trimmed
trimmed = AxTrim("  hello world  ")
' trimmed = "hello world"

' ============================================================================
' MATH FUNCTIONS EXAMPLES
' ============================================================================

' Rounding
Response.Write AxCeil(4.2) & "<br>"    ' 5
Response.Write AxFloor(4.8) & "<br>"   ' 4

' Min/Max
Response.Write AxMax(10, 20, 15, 5) & "<br>"   ' 20
Response.Write AxMin(10, 20, 15, 5) & "<br>"   ' 5

' Random number
Dim random_num
random_num = AxRand(1, 100)  ' 1-100

' Format number
Dim formatted_number
formatted_number = AxNumberFormat(1234567.89, 2, ".", ",")
' formatted_number = "1,234,567.89"

' ============================================================================
' TYPE CHECKING EXAMPLES
' ============================================================================

' Check types
Response.Write AxIsInt(5) & "<br>"              ' True
Response.Write AxIsFloat(5.5) & "<br>"          ' True
Response.Write AxCTypeAlpha("hello") & "<br>"   ' True
Response.Write AxCTypeAlnum("abc123") & "<br>"  ' True

' Check empty values
Response.Write AxEmpty("") & "<br>"     ' True
Response.Write AxEmpty(0) & "<br>"      ' True
Response.Write AxEmpty(Array()) & "<br>"  ' True

' Check if set
Dim test_var
test_var = "something"
Response.Write AxIsset(test_var) & "<br>"  ' True

' ============================================================================
' DATE/TIME EXAMPLES
' ============================================================================

' Get current Unix timestamp
Dim timestamp
timestamp = AxTime()

' Format current date
Response.Write AxDate("Y-m-d") & "<br>"           ' 2024-01-16
Response.Write AxDate("Y-m-d H:i:s") & "<br>"     ' 2024-01-16 14:30:45
Response.Write AxDate("d/m/Y") & "<br>"           ' 16/01/2024
Response.Write AxDate("l, F j, Y") & "<br>"       ' Tuesday, January 16, 2024

' Format specific timestamp
Dim past_timestamp
past_timestamp = 1672531200
Response.Write AxDate("Y-m-d", past_timestamp) & "<br>"

' ============================================================================
' HASHING & ENCODING EXAMPLES
' ============================================================================

Dim password, secret
password = "myPassword123"

' Generate hash
Response.Write AxMd5(password) & "<br>"
Response.Write AxSha1(password) & "<br>"
Response.Write AxHash("sha256", password) & "<br>"

' Base64
Dim encoded_text, decoded_text
encoded_text = AxBase64Encode("Hello, World!")
decoded_text = AxBase64Decode(encoded_text)
' decoded_text = "Hello, World!"

' URL encoding
Dim encoded_url, decoded_url
encoded_url = "Hello%20World%21"
decoded_url = AxUrlDecode(encoded_url)
' decoded_url = "Hello World!"

' RGB to Hex
Dim color_hex
color_hex = AxRgbToHex(255, 128, 0)
' color_hex = "#FF8000"

' HTML escaping
Dim html_input, escaped_html
html_input = "<script>alert('xss')</script>"
escaped_html = AxHtmlSpecialChars(html_input)

' Strip HTML tags
Dim html_content, plain_text
html_content = "<p>Hello <b>World</b></p>"
plain_text = AxStripTags(html_content)
' plain_text = "Hello World"

' ============================================================================
' VALIDATION EXAMPLES
' ============================================================================

' Validate IP
If AxFilterValidateIp("192.168.1.1") Then
    Response.Write "Valid IP address"
End If

' Validate Email
If AxFilterValidateEmail("user@example.com") Then
    Response.Write "Valid email address"
End If

' ============================================================================
' REQUEST ARRAYS EXAMPLES
' ============================================================================

' Get all request parameters
Dim all_params, get_params, post_params
all_params = AxGetRequest()   ' GET + POST
get_params = AxGetGet()        ' Only GET
post_params = AxGetPost()      ' Only POST

' Access specific parameter
Dim user_name
user_name = all_params("username")

' ============================================================================
' UTILITY FUNCTIONS EXAMPLES
' ============================================================================

' Generate unique ID
Dim unique_id
unique_id = AxGenerateGuid()

' Build query string
Dim query_params, query_string
Set query_params = CreateObject("Scripting.Dictionary")
query_params("page") = 1
query_params("limit") = 20
query_params("search") = "test value"
query_string = AxBuildQueryString(query_params)
' query_string = "page=1&limit=20&search=test%20value"

' ============================================================================
' DOCUMENT.WRITE (SAFE HTML) EXAMPLE
' ============================================================================

Dim user_input
user_input = "<img src=x onerror='alert(1)'>"

' Normal Response.Write (UNSAFE)
Response.Write user_input  ' Would execute JavaScript!

' Using Document.Write (SAFE - HTML escaped)
Document.Write user_input  ' Displays: &lt;img src=x onerror=&#39;alert(1)&#39;&gt;

' ============================================================================
' PRACTICAL EXAMPLE: USER FORM PROCESSING
' ============================================================================

' Simulate form submission
Dim form_data, errors, clean_data

Set form_data = CreateObject("Scripting.Dictionary")
Set clean_data = CreateObject("Scripting.Dictionary")
Set errors = CreateObject("Scripting.Dictionary")

' Input from form
form_data("name") = "  John Doe  "
form_data("email") = "john@example.com"
form_data("age") = "25"
form_data("website") = "http://example.com"

' Validation and cleaning
If AxEmpty(form_data("name")) Then
    errors("name") = "Name is required"
Else
    clean_data("name") = AxTrim(form_data("name"))
End If

If Not AxFilterValidateEmail(form_data("email")) Then
    errors("email") = "Invalid email format"
Else
    clean_data("email") = form_data("email")
End If

If Not AxIsInt(form_data("age")) Then
    errors("age") = "Age must be a number"
Else
    clean_data("age") = CInt(form_data("age"))
End If

' Output results
If errors.Count = 0 Then
    Response.Write "Form is valid!<br>"
    Response.Write "Name: " & clean_data("name") & "<br>"
    Response.Write "Email: " & clean_data("email") & "<br>"
    Response.Write "Age: " & clean_data("age") & "<br>"
Else
    Response.Write "Form has errors:<br>"
    Dim key
    For Each key In errors.Keys()
        Response.Write "- " & key & ": " & errors(key) & "<br>"
    Next
End If

' ============================================================================
' PRACTICAL EXAMPLE: CSV EXPORT
' ============================================================================

' Sample data
Dim records, headers, csv_output
ReDim records(2)
records(0) = Array("John", "Doe", 30)
records(1) = Array("Jane", "Smith", 25)
records(2) = Array("Bob", "Johnson", 35)

headers = Array("First Name", "Last Name", "Age")

' Build CSV
csv_output = AxImplode(",", headers) & vbCrLf
Dim i
For i = 0 To UBound(records)
    csv_output = csv_output & AxImplode(",", records(i)) & vbCrLf
Next

' Output or save
' Response.AddHeader "Content-Disposition", "attachment; filename=data.csv"
' Response.Write csv_output

' ============================================================================
' PRACTICAL EXAMPLE: SECURITY HASHING
' ============================================================================

' User registration
Dim plain_password, password_hash

plain_password = "UserPassword123!"
password_hash = AxHash("sha256", plain_password)

' Store password_hash in database
Response.Write "Password stored as: " & password_hash & "<br>"

' On login, hash input and compare
Dim login_password, login_hash
login_password = "UserPassword123!"
login_hash = AxHash("sha256", login_password)

If login_hash = password_hash Then
    Response.Write "Password is correct!"
End If

' ============================================================================
' PRACTICAL EXAMPLE: API URL BUILDING
' ============================================================================

Dim api_params, api_url
Set api_params = CreateObject("Scripting.Dictionary")

api_params("key") = "abc123"
api_params("query") = "search term"
api_params("limit") = 10
api_params("offset") = 0

api_url = "https://api.example.com/search?" & AxBuildQueryString(api_params)
Response.Write "API URL: " & api_url & "<br>"

%>

<!DOCTYPE html>
<html>
<head>
    <title>Custom Functions Examples</title>
</head>
<body>
    <h1>G3 AxonASP Custom Functions Examples</h1>
    <p>Check the page source and server logs for examples.</p>
</body>
</html>

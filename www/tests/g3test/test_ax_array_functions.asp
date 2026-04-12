<%
Dim t
Dim values
Dim exploded
Dim reversedValues
Dim ranged
Dim words
Dim csvValues

Set t = Server.CreateObject("G3TestSuite")

t.Describe "Ax array-oriented functions"

values = Array("alpha", "beta", "gamma")
t.AssertEqual 3, AxCount(values), "AxCount should return the number of elements in a VBScript array"
t.AssertEqual 0, AxCount("not-an-array"), "AxCount should return 0 for non-array values"

t.AssertEqual "", AxImplode(",", "not-an-array"), "AxImplode should return empty string when input is not an array"
t.AssertEqual "alpha|beta|gamma", AxImplode("|", values), "AxImplode should join array values using glue"

exploded = AxExplode(",", "one,two,three")
t.AssertLength 3, exploded, "AxExplode should split values by delimiter"
t.AssertEqual "one", exploded(0), "AxExplode first item should match"
t.AssertEqual "two", exploded(1), "AxExplode second item should match"
t.AssertEqual "three", exploded(2), "AxExplode third item should match"

exploded = AxExplode(",", "a,b,c,d", 2)
t.AssertLength 2, exploded, "AxExplode should enforce positive limit by truncating the result"
t.AssertEqual "a", exploded(0), "AxExplode limited result first item should match"
t.AssertEqual "b", exploded(1), "AxExplode limited result second item should match"

exploded = AxExplode("", "abc")
t.AssertLength 3, exploded, "AxExplode with empty delimiter should split characters"
t.AssertEqual "a", exploded(0), "AxExplode empty delimiter first character should match"
t.AssertEqual "b", exploded(1), "AxExplode empty delimiter second character should match"
t.AssertEqual "c", exploded(2), "AxExplode empty delimiter third character should match"

reversedValues = AxArrayReverse(Array(10, 20, 30, 40))
t.AssertLength 4, reversedValues, "AxArrayReverse should keep array length"
t.AssertEqual 40, reversedValues(0), "AxArrayReverse first value should be the original last value"
t.AssertEqual 30, reversedValues(1), "AxArrayReverse second value should match"
t.AssertEqual 20, reversedValues(2), "AxArrayReverse third value should match"
t.AssertEqual 10, reversedValues(3), "AxArrayReverse fourth value should match"

t.AssertEqual "stable", AxArrayReverse("stable"), "AxArrayReverse should return the original non-array value"

ranged = AxRange(1, 5)
t.AssertLength 5, ranged, "AxRange should include both start and end bounds with default step"
t.AssertEqual 1, ranged(0), "AxRange default step first value should match"
t.AssertEqual 5, ranged(4), "AxRange default step last value should match"

ranged = AxRange(1, 5, 2)
t.AssertLength 3, ranged, "AxRange with custom positive step should generate expected values"
t.AssertEqual 1, ranged(0), "AxRange positive step first value should match"
t.AssertEqual 3, ranged(1), "AxRange positive step second value should match"
t.AssertEqual 5, ranged(2), "AxRange positive step third value should match"

ranged = AxRange(5, 1, -2)
t.AssertLength 3, ranged, "AxRange with negative step should generate descending values"
t.AssertEqual 5, ranged(0), "AxRange negative step first value should match"
t.AssertEqual 3, ranged(1), "AxRange negative step second value should match"
t.AssertEqual 1, ranged(2), "AxRange negative step third value should match"

ranged = AxRange(2, 4, 0)
t.AssertLength 3, ranged, "AxRange with step 0 should fallback to step 1"
t.AssertEqual 2, ranged(0), "AxRange step 0 fallback first value should match"
t.AssertEqual 3, ranged(1), "AxRange step 0 fallback second value should match"
t.AssertEqual 4, ranged(2), "AxRange step 0 fallback third value should match"

words = AxWordCount("The quick brown fox", 1)
t.AssertLength 4, words, "AxWordCount format 1 should return words as an array"
t.AssertEqual "The", words(0), "AxWordCount first word should match"
t.AssertEqual "fox", words(3), "AxWordCount last word should match"

csvValues = AxStringGetCsv("name,age,city")
t.AssertLength 3, csvValues, "AxStringGetCsv should parse comma-separated values"
t.AssertEqual "name", csvValues(0), "AxStringGetCsv first value should match"
t.AssertEqual "age", csvValues(1), "AxStringGetCsv second value should match"
t.AssertEqual "city", csvValues(2), "AxStringGetCsv third value should match"

csvValues = AxStringGetCsv("a;b;c", ";")
t.AssertLength 3, csvValues, "AxStringGetCsv should support custom delimiters"
t.AssertEqual "a", csvValues(0), "AxStringGetCsv custom delimiter first value should match"
t.AssertEqual "b", csvValues(1), "AxStringGetCsv custom delimiter second value should match"
t.AssertEqual "c", csvValues(2), "AxStringGetCsv custom delimiter third value should match"
%>
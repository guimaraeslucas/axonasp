<%
' JSON object class 3.9.0 Aug, 13th - 2025
' https://github.com/rcdmk/aspJSON
'
' License MIT - see LICENCE.txt for details

Const JSON_ROOT_KEY = "[[JSONroot]]"
Const JSON_DEFAULT_PROPERTY_NAME = "data"
Const JSON_SPECIAL_VALUES_REGEX = "^(?:(?:t(?:r(?:ue?)?)?)|(?:f(?:a(?:l(?:se?)?)?)?)|(?:n(?:u(?:ll?)?)?)|(?:u(?:n(?:d(?:e(?:f(?:i(?:n(?:ed?)?)?)?)?)?)?)?))$"
Const JSON_UNICODE_CHARS_REGEX = "\\u(\d{4})"

Const JSON_ERROR_PARSE = 1
Const JSON_ERROR_PROPERTY_ALREADY_EXISTS = 2
Const JSON_ERROR_PROPERTY_DOES_NOT_EXISTS = 3 ' DEPRECATED
Const JSON_ERROR_NOT_AN_ARRAY = 4
Const JSON_ERROR_NOT_A_STRING = 5
Const JSON_ERROR_INDEX_OUT_OF_BOUNDS = 9 ' Numbered To have the same Error number As the Default "Subscript out of range" exeption

Class JSONobject
    Dim i_debug, i_depth, i_parent
    Dim i_properties, i_version, i_defaultPropertyName
    Dim i_properties_count, i_properties_capacity
    Private vbback

    ' Set to true to show the internals of the parsing mecanism
    Public Property Get debug
        debug = i_debug
    End Property

    Public Property Let debug(value)
        i_debug = value
    End Property

    ' Gets/sets the default property name generated when loading recordsets and arrays (default "data")
    Public Property Get defaultPropertyName
        defaultPropertyName = i_defaultPropertyName
    End Property

    Public Property Let defaultPropertyName(value)
        i_defaultPropertyName = value
    End Property

    ' The depth of the object in the chain, starting with 1
    Public Property Get depth
        depth = i_depth
    End Property

    ' The property pairs ("name": "value" - pairs)
    Public Property Get pairs
        Dim tmp
        tmp = i_properties
        If i_properties_count < i_properties_capacity Then
            ReDim Preserve tmp(i_properties_count - 1)
        End If
        pairs = tmp
    End Property

    ' The parent object
    Public Property Get parent
        Set parent = i_parent
    End Property

    Public Property Set parent(value)
        Set i_parent = value
        i_depth = i_parent.depth + 1
    End Property

    ' The class version
    Public Property Get version
        version = i_version
    End Property

    ' Constructor and destructor
    Private Sub class_initialize()
        i_version = "3.9.0"
        i_depth = 0
        i_debug = False
        i_defaultPropertyName = JSON_DEFAULT_PROPERTY_NAME

        Set i_parent = Nothing
        ReDim i_properties(-1)
        i_properties_capacity = 0
        i_properties_count = 0

        vbback = Chr(8)
    End Sub

    Private Sub class_terminate()
        Dim i
        For i = 0 To UBound(i_properties)
            Set i_properties(i) = Nothing
        Next

        ReDim i_properties(-1)
    End Sub

    ' Parse a JSON string and populate the object
    Public Function parse(ByVal strJson)
        Dim regex, i, size, char, prevchar, quoted
        Dim mode, Item, Key, keyStart, value, openArray, openObject
        Dim actualLCID, tmpArray, tmpObj, addedToArray
        Dim root, currentObject, currentArray

        Log("Load string: """ & strJson & """")

        ' Store the actual LCID and use the en-US to conform with the JSON standard
        actualLCID = Response.LCID
        Response.LCID = 1033

        strJson = Trim(strJson)

        size = Len(strJson)

        ' At least 2 chars to continue
        If size < 2 Then Err.raise JSON_ERROR_PARSE, TypeName(me), "Invalid JSON string to parse"

        ' Init the regex to be used in the loop
        Set regex = New regexp
        regex.global = True
        regex.ignoreCase = True
        regex.pattern = "\w"

        ' setup initial values
        i = 0
        Set root = me
        Key = JSON_ROOT_KEY
        mode = "init"
        quoted = False
        Set currentObject = root

        ' main state machine
        Do While i < size
            i = i + 1
            char = Mid(strJson, i, 1)

            ' root, object or array start
            If mode = "init" Then
                Log("Enter init")

                ' if we are in root, clear previous object properties
                If Key = JSON_ROOT_KEY And TypeName(currentArray) <> "JSONarray" Then
                    ReDim i_properties(-1)
                    i_properties_capacity = 0
                    i_properties_count = 0
                End If

                ' Init object
                If char = "{" Then
                    Log("Create object<ul>")

                    If Key <> JSON_ROOT_KEY Or TypeName(root) = "JSONarray" Then
                        ' creates a new object
                        Set Item = New JSONobject
                        Set Item.parent = currentObject

                        addedToArray = False

                        ' Object is inside an array
                        If TypeName(currentArray) = "JSONarray" Then
                            If currentArray.depth > currentObject.depth Then
                                ' Add it to the array
                                Set Item.parent = currentArray
                                currentArray.Push Item

                                addedToArray = True

                                Log("Added to the array")
                            End If
                        End If

                        If Not addedToArray Then
                            currentObject.Add Key, Item
                            Log("Added to parent object: """ & Key & """")
                        End If

                        Set currentObject = Item
                    End If

                    openObject = openObject + 1
                    mode = "openKey"

                ' Init Array
                ElseIf char = "[" Then
                    Log("Create array<ul>")

                    Set Item = New JSONarray

                    addedToArray = False

                    ' Array is inside an array
                    If IsObject(currentArray) And openArray > 0 Then
                        If currentArray.depth > currentObject.depth Then
                            ' Add it to the array
                            Set Item.parent = currentArray
                            currentArray.Push Item

                            addedToArray = True

                            Log("Added to parent array")
                        End If
                    End If

                    If Not addedToArray Then
                        Set Item.parent = currentObject
                        currentObject.Add Key, Item
                        Log("Added to parent object")
                    End If

                    If Key = JSON_ROOT_KEY And Item.depth = 1 Then
                        Set root = Item
                        Log("Set as root")
                    End If

                    Set currentArray = Item
                    openArray = openArray + 1
                    mode = "openValue"
                End If

            ' Init a key
            ElseIf mode = "openKey" Then
                Key = ""
                If char = """" Then
                    Log("Open key")
                    keyStart = i + 1
                    mode = "closeKey"
                ElseIf char = "}" Then ' Empty objects
                    Log("Empty object")
                    mode = "next"
                    i = i - 1 ' we backup one char To make the Next iteration Get the closing bracket
                End If

            ' Fill in the key until finding a double quote "
            ElseIf mode = "closeKey" Then
                ' If it finds a non scaped quotation, change to value mode
                If char = """" And prevchar <> "\" Then
                    Key = Mid(strJson, keyStart, i - keyStart)
                    Log("Close key: """ & Key & """")
                    mode = "preValue"
                End If

            ' Wait until a colon char (:) to begin the value
            ElseIf mode = "preValue" Then
                If char = ":" Then
                    mode = "openValue"
                    Log("Open value for """ & Key & """")
                End If

            ' Begining of value
            ElseIf mode = "openValue" Then
                value = ""

                ' If the next char is a closing square barcket (]), its closing an empty array
                If char = "]" Then
                    Log("Closing empty array")
                    quoted = False
                    mode = "next"
                    i = i - 1 ' we backup one char To make the Next iteration Get the closing bracket

                ' If it begins with a double quote, its a string value
                ElseIf char = """" Then
                    Log("Open string value")
                    quoted = True
                    keyStart = i + 1
                    mode = "closeValue"

                ' If it begins with open square bracket ([), its an array
                ElseIf char = "[" Then
                    Log("Open array value")
                    quoted = False
                    mode = "init"
                    i = i - 1 ' we backup one char To init With '['

                ' If it begins with open a bracket ({), its an object
                ElseIf char = "{" Then
                    Log("Open object value")
                    quoted = False
                    mode = "init"
                    i = i - 1 ' we backup one char To init With '{'

                Else
                    ' If its a number, start a numeric value
                    If regex.pattern <> "\d" Then regex.pattern = "\d"
                    If regex.test(char) Then
                        Log("Open numeric value")
                        quoted = False
                        value = char
                        mode = "closeValue"
                        If prevchar = "-" Then
                            value = prevchar & char
                        End If

                    ' special values: null, true, false and undefined
                    ElseIf char = "n" Or char = "t" Or char = "f" Or char = "u" Then
                        Log("Open special value")
                        quoted = False
                        value = char
                        mode = "closeValue"
                    End If
                End If

            ' Fill in the value until finish
            ElseIf mode = "closeValue" Then
                If quoted Then
                    If char = """" And prevchar <> "\" Then
                        value = Mid(strJson, keyStart, i - keyStart)

                        value = Replace(value, "\n", vblf)
                        value = Replace(value, "\r", vbcr)
                        value = Replace(value, "\t", vbtab)
                        value = Replace(value, "\b", vbback)
                        value = Replace(value, "\\", "\")

                        regex.pattern = JSON_UNICODE_CHARS_REGEX
                        If regex.test(value) Then
                            Dim match
                            For Each match In regex.Execute(value)
                                value = Replace(value, match.value, ChrW("&H" & match.SubMatches(0)))
                            Next
                        End If

                        Log("Close string value: """ & value & """")
                        mode = "addValue"
                    End If
                Else
                    ' possible boolean and null values
                    If regex.pattern <> JSON_SPECIAL_VALUES_REGEX Then regex.pattern = JSON_SPECIAL_VALUES_REGEX
                    If regex.test(char) Or regex.test(value) Then
                        value = value & char
                        If value = "true" Or value = "false" Or value = "null" Or value = "undefined" Then mode = "addValue"
                    Else
                        char = LCase(char)

                        ' If is a numeric char
                        If regex.pattern <> "\d" Then regex.pattern = "\d"
                        If regex.test(char) Then
                            value = value & char

                        ' If it's not a numeric char, but the prev char was a number
                        ' used to catch separators and special numeric chars
                        ElseIf regex.test(prevchar) Or prevchar = "e" Then
                            If char = "." Or char = "e" Or (prevchar = "e" And (char = "-" Or char = "+")) Then
                                value = value & char
                            Else
                                Log("Close numeric value: " & value)
                                mode = "addValue"
                                i = i - 1
                            End If
                        Else
                            Log("Close numeric value: " & value)
                            mode = "addValue"
                            i = i - 1
                        End If
                    End If
                End If

            ' Add the value to the object or array
            ElseIf mode = "addValue" Then
                If Key <> "" Then
                    Dim useArray
                    useArray = False

                    If Not quoted Then
                        If IsNumeric(value) Then
                            Log("Value converted to number")
                            value = CDbl(value)
                        Else
                            Log("Value converted to " & value)
                            value = Eval(value)
                        End If
                    End If

                    quoted = False

                    ' If it's inside an array
                    If openArray > 0 And IsObject(currentArray) Then
                        useArray = True

                        ' If it's a property of an object that is inside the array
                        ' we add it to the object instead
                        If IsObject(currentObject) Then
                            If currentObject.depth >= currentArray.depth + 1 Then useArray = False
                        End If

                        ' else, we add it to the array
                        If useArray Then
                            currentArray.Push value

                            Log("Value added to array: """ & Key & """: " & value)
                        End If
                    End If

                    If Not useArray Then
                        currentObject.Add Key, value
                        Log("Value added: """ & Key & """")
                    End If
                End If

                mode = "next"
                i = i - 1

            ' Change the current mode according to the current state
            ElseIf mode = "next" Then
                If char = "," Then
                    ' If it's an array
                    If openArray > 0 And IsObject(currentArray) Then
                        ' and the current object is a parent or sibling object
                        If currentArray.depth >= currentObject.depth Then
                            ' start an array value
                            Log("New value")
                            mode = "openValue"
                        Else
                            ' start an object key
                            Log("New key")
                            mode = "openKey"
                        End If
                    Else
                        ' start an object key
                        Log("New key")
                        mode = "openKey"
                    End If

                ElseIf char = "]" Then
                    Log("Close array</ul>")

                    ' If it's and open array, we close it and set the current array as its parent
                    If IsObject(currentArray.parent) Then
                        If TypeName(currentArray.parent) = "JSONarray" Then
                            Set currentArray = currentArray.parent

                        ' if the parent is an object
                        ElseIf TypeName(currentArray.parent) = "JSONobject" Then
                            Set tmpObj = currentArray.parent

                            ' we search for the next parent array to set the current array
                            While IsObject(tmpObj) And TypeName(tmpObj) = "JSONobject"
                                If IsObject(tmpObj.parent) Then
                                    Set tmpObj = tmpObj.parent
                                Else
                                    tmpObj = tmpObj.parent
                                End If
                            Wend

                            Set currentArray = tmpObj
                        End If
                    Else
                        currentArray = currentArray.parent
                    End If

                    openArray = openArray - 1

                    mode = "next"

                ElseIf char = "}" Then
                    Log("Close object</ul>")

                    ' If it's an open object, we close it and set the current object as it's parent
                    If IsObject(currentObject.parent) Then
                        If TypeName(currentObject.parent) = "JSONobject" Then
                            Set currentObject = currentObject.parent

                        ' If the parent is and array
                        ElseIf TypeName(currentObject.parent) = "JSONarray" Then
                            Set tmpObj = currentObject.parent

                            ' we search for the next parent object to set the current object
                            While IsObject(tmpObj) And TypeName(tmpObj) = "JSONarray"
                                Set tmpObj = tmpObj.parent
                            Wend

                            Set currentObject = tmpObj
                        End If
                    Else
                        currentObject = currentObject.parent
                    End If

                    openObject = openObject - 1

                    mode = "next"
                End If
            End If

            prevchar = char
        Loop

        Set regex = Nothing

        Response.LCID = actualLCID

        Set parse = root
    End Function

    ' Add a new property (key-value pair)
    Public Sub Add(ByVal prop, ByVal obj)
        Dim p
        getProperty prop, p

        If GetTypeName(p) = "JSONpair" Then
            Err.raise JSON_ERROR_PROPERTY_ALREADY_EXISTS, TypeName(me), "A property already exists with the name: " & prop & "."
        Else
            Dim Item
            Set Item = New JSONpair
            Item.name = prop
            Set Item.parent = me

            Dim itemType
            itemType = GetTypeName(obj)

            If IsArray(obj) Then
                Dim item2
                Set item2 = New JSONarray
                item2.Items = obj
                Set item2.parent = me

                Set Item.value = item2

            ElseIf itemType = "Field" Then
                Item.value = obj.value
            ElseIf IsObject(obj) And itemType <> "IStringList" Then
                Set Item.value = obj
            Else
                Item.value = obj
            End If

            If i_properties_count >= i_properties_capacity Then
                ReDim Preserve i_properties(i_properties_capacity * 1.2 + 1)
                i_properties_capacity = UBound(i_properties) + 1
            End If

            Set i_properties(i_properties_count) = Item
            i_properties_count = i_properties_count + 1
        End If
    End Sub

    ' Remove a property from the object (key-value pair)
    Public Sub Remove(ByVal prop)
        Dim p, i
        i = getProperty(prop, p)

        ' property exists
        If i > -1 Then ArraySlice i_properties, i
    End Sub

    ' Return the value of a property by its key
    Public Default Function value(ByVal prop)
        Dim p
        getProperty prop, p

        If GetTypeName(p) = "JSONpair" Then
            If IsObject(p.value) Then
                Set value = p.value
            Else
                value = p.value
            End If
        Else
            value = Null
        End If
    End Function

    ' Change the value of a property
    ' Creates the property if it didn't exists
    Public Sub change(ByVal prop, ByVal obj)
        Dim p
        getProperty prop, p

        If GetTypeName(p) = "JSONpair" Then
            If IsArray(obj) Then
                Set Item = New JSONarray
                Item.Items = obj
                Set Item.parent = me

                p.value = Item

            ElseIf IsObject(obj) Then
                Set p.value = obj
            Else
                p.value = obj
            End If
        Else
            Add prop, obj
        End If
    End Sub

    ' Returns the index of a property if it exists, else -1
    ' @param prop as string - the property name
    ' @param out outProp as variant - will be filled with the property value, nothing if not found
    Private Function getProperty(ByVal prop, ByRef outProp)
        Dim i, p, found
        Set outProp = Nothing
        found = False

        i = 0

        Do While i < i_properties_count
            Set p = i_properties(i)

            If p.name = prop Then
                Set outProp = p
                found = True

                Exit Do
                End If

                i = i + 1
            Loop

            If Not found Then
                If prop = i_defaultPropertyName Then
                    i = getProperty(JSON_ROOT_KEY, outProp)
                Else
                    i = -1
                End If
            End If

            getProperty = i
        End Function

        ' Serialize the current object to a JSON formatted string
        Public Function Serialize()
            Dim actualLCID, out
            actualLCID = Response.LCID
            Response.LCID = 1033

            out = serializeObject(me)

            Response.LCID = actualLCID

            Serialize = out
        End Function

        ' Writes the JSON serialized object to the response
        Public Sub Write()
            Response.Write Serialize()
        End Sub

        ' Helpers
        ' Serializes a JSON object to JSON formatted string
        Public Function serializeObject(obj)
            Dim out, prop, value, i, pairs, valueType
            out = "{"

            pairs = obj.pairs

            For i = 0 To UBound(pairs)
                Set prop = pairs(i)

                If out <> "{" Then out = out & ","

                If IsObject(prop.value) Then
                    Set value = prop.value
                Else
                    value = prop.value
                End If

                If prop.name = JSON_ROOT_KEY Then
                    out = out & ("""" & obj.defaultPropertyName & """:")
                Else
                    out = out & ("""" & prop.name & """:")
                End If

                If IsArray(value) Or GetTypeName(value) = "JSONarray" Then
                    out = out & serializeArray(value)

                ElseIf IsObject(value) And GetTypeName(value) = "JSONscript" Then
                    out = out & value.Serialize()

                ElseIf IsObject(value) Then
                    out = out & serializeObject(value)

                Else
                    out = out & serializeValue(value)
                End If
            Next

            out = out & "}"

            serializeObject = out
        End Function

        ' Serializes a value to a valid JSON formatted string representing the value
        ' (quoted for strings, the type name for objects, null for nothing and null values)
        Public Function serializeValue(ByVal value)
            Dim out

            Select Case LCase(GetTypeName(value))
                Case "null"
                    out = "null"

                Case "nothing"
                    out = "undefined"

                Case "boolean"
                    If value Then
                        out = "true"
                    Else
                        out = "false"
                    End If

                Case "byte", "integer", "long", "single", "double", "currency", "decimal"
                    out = value

                Case "date"
                    out = """" & Year(value) & "-" & padZero(Month(value), 2) & "-" & padZero(Day(value), 2) & "T" & padZero(Hour(value), 2) & ":" & padZero(Minute(value), 2) & ":" & padZero(Second(value), 2) & """"

                Case "string", "char", "empty"
                    out = """" & EscapeCharacters(value) & """"

                Case Else
                    out = """" & GetTypeName(value) & """"
            End Select

            serializeValue = out
        End Function

        ' Pads a numeric string with zeros at left
        Private Function padZero(ByVal number, ByVal length)
            Dim result
            result = number

            While Len(result) < length
                result = "0" & result
            Wend

            padZero = result
        End Function

        ' Serializes an array item to JSON formatted string
        Private Function serializeArrayItem(ByRef elm)
            Dim out, val

            If IsObject(elm) Then
                If GetTypeName(elm) = "JSONobject" Then
                    Set val = elm

                ElseIf GetTypeName(elm) = "JSONarray" Then
                    val = elm.Items

                ElseIf IsObject(elm.value) Then
                    Set val = elm.value

                Else
                    val = elm.value
                End If
            Else
                val = elm
            End If

            If IsArray(val) Or GetTypeName(val) = "JSONarray" Then
                out = out & serializeArray(val)

            ElseIf IsObject(val) Then
                out = out & serializeObject(val)

            Else
                out = out & serializeValue(val)
            End If

            serializeArrayItem = out
        End Function

        ' Serializes an array or JSONarray object to JSON formatted string
        Public Function serializeArray(ByRef arr)
            Dim i, j, k, dimensions, out, innerArray, elm, val

            out = "["

            If IsObject(arr) Then
                Log("Serializing jsonArray object")
                innerArray = arr.Items
            Else
                Log("Serializing VB array")
                innerArray = arr
            End If

            dimensions = NumDimensions(innerArray)

            If dimensions > 1 Then
                Log("Multidimensional array")
                For j = 0 To UBound(innerArray, 1)
                    out = out & "["

                    For k = 0 To UBound(innerArray, 2)
                        If k > 0 Then out = out & ","

                        If IsObject(innerArray(j, k)) Then
                            Set elm = innerArray(j, k)
                        Else
                            elm = innerArray(j, k)
                        End If

                        out = out & serializeArrayItem(elm)
                    Next

                    out = out & "]"
                Next
            Else
                For j = 0 To UBound(innerArray)
                    If j > 0 Then out = out & ","

                    If IsObject(innerArray(j)) Then
                        Set elm = innerArray(j)
                    Else
                        elm = innerArray(j)
                    End If

                    out = out & serializeArrayItem(elm)
                Next
            End If

            out = out & "]"

            serializeArray = out
        End Function

        ' Returns the number of dimensions an array has
        Public Function NumDimensions(ByRef arr)
            Dim dimensions
            dimensions = 0

            On Error Resume Next

            Do While Err.number = 0
                dimensions = dimensions + 1
                UBound arr, dimensions
            Loop
            On Error Goto 0

            NumDimensions = dimensions - 1
        End Function

        ' DEPRECATED: Pushes (adds) a value to an array, expanding it
        Public Function ArrayPush(ByRef arr, ByRef value)
            ReDim Preserve arr(UBound(arr) + 1)

            If IsObject(value) Then
                Set arr(UBound(arr)) = value
            Else
                arr(UBound(arr)) = value
            End If

            ArrayPush = arr
        End Function

        ' Removes a value from an array
        Private Function ArraySlice(ByRef arr, ByRef index)
            Dim i
            i = index

            For i = index To i_properties_count - 2
                If IsObject(arr(i)) Then Set arr(i) = Nothing

                If IsObject(arr(i + 1)) Then
                    Set arr(i) = arr(i + 1)
                Else
                    arr(i) = arr(i + 1)
                End If
            Next

            i_properties_count = i_properties_count - 1

            If i_properties_count < i_properties_capacity / 2 Then
                ReDim Preserve arr(i_properties_count * 1.2 + 1)
                i_properties_capacity = UBound(i_properties) + 1
            End If

            ArraySlice = arr
        End Function

        ' Load properties from an ADO RecordSet object into an array
        ' @param rs as ADODB.RecordSet
        Public Sub LoadRecordSet(ByRef rs)
            Dim arr, obj, field

            Set arr = New JSONArray

            While Not rs.eof
                Set obj = New JSONobject

                For Each field In rs.fields
                    obj.Add field.name, field.value
                Next

                arr.Push obj

                rs.MoveNext
            Wend

            Set obj = Nothing

            change i_defaultPropertyName, arr
        End Sub

        ' Load properties from the first record of an ADO RecordSet object
        ' @param rs as ADODB.RecordSet
        Public Sub LoadFirstRecord(ByRef rs)
            Dim field

            For Each field In rs.fields
                Add field.name, field.value
            Next
        End Sub

        ' Returns the value's type name (usefull for types not supported by VBS)
        Public Function GetTypeName(ByVal value)
            Dim valueType

            On Error Resume Next
            valueType = TypeName(value)

            If Err.number <> 0 Then
                Select Case VarType(value)
                    Case 14 ' MySQL Decimal
                        valueType = "Decimal"
                    Case 16 ' MySQL TinyInt
                        valueType = "Integer"
                End Select
            End If
            On Error Goto 0

            GetTypeName = valueType
        End Function

        ' Escapes special characters in the text
        ' @param text as String
        Public Function EscapeCharacters(ByVal text)
            Dim result

            result = text

            If Not IsNull(text) Then
                result = CStr(result)

                result = Replace(result, "\", "\\")
                result = Replace(result, """", "\""")
                result = Replace(result, vbcr, "\r")
                result = Replace(result, vblf, "\n")
                result = Replace(result, vbtab, "\t")
                result = Replace(result, vbback, "\b")
            End If

            EscapeCharacters = result
        End Function

        ' Used to write the log messages to the response on debug mode
        Private Sub Log(ByVal msg)
            If i_debug Then Response.Write "<li>" & msg & "</li>" & vbcrlf
        End Sub
    End Class

    ' JSON array class
    ' Represents an array of JSON objects and values
    Class JSONarray
        Dim i_items, i_depth, i_parent, i_version, i_defaultPropertyName
        Dim i_items_count, i_items_capacity

        ' The class version
        Public Property Get version
            version = i_version
        End Property

        ' The actual array items
        Public Property Get Items
            Dim tmp
            tmp = i_items
            If i_items_count < i_items_capacity Then
                ReDim Preserve tmp(i_items_count - 1)
            End If
            Items = tmp
        End Property

        Public Property Let Items(value)
            If IsArray(value) Then
                i_items = value
                i_items_count = UBound(value) + 1
                i_items_capacity = i_items_count
            Else
                Err.raise JSON_ERROR_NOT_AN_ARRAY, TypeName(me), "The value assigned is not an array."
            End If
        End Property

        ' The capacity of the underlying array
        Public Property Get capacity
            capacity = i_items_capacity
        End Property

        ' The length of the array
        Public Property Get length
            length = i_items_count
        End Property

        ' The depth of the array in the chain (starting with 1)
        Public Property Get depth
            depth = i_depth
        End Property

        ' The parent object or array
        Public Property Get parent
            Set parent = i_parent
        End Property

        Public Property Set parent(value)
            Set i_parent = value
            i_depth = i_parent.depth + 1
            i_defaultPropertyName = i_parent.defaultPropertyName
        End Property

        ' Gets/sets the default property name generated when loading recordsets and arrays (default "data")
        Public Property Get defaultPropertyName
            defaultPropertyName = i_defaultPropertyName
        End Property

        Public Property Let defaultPropertyName(value)
            i_defaultPropertyName = value
        End Property

        ' Constructor and destructor
        Private Sub class_initialize
            i_version = "2.4.0"
            i_defaultPropertyName = JSON_DEFAULT_PROPERTY_NAME
            ReDim i_items(-1)
            i_items_count = 0
            i_items_capacity = 0
            i_depth = 0
        End Sub

        Private Sub class_terminate
            Dim i, j, js, dimensions

            dimensions = 0

            On Error Resume Next

            Do While Err.number = 0
                dimensions = dimensions + 1
                UBound i_items, dimensions
            Loop

            On Error Goto 0

            dimensions = dimensions - 1

            For i = 1 To dimensions
                For j = 0 To UBound(i_items, i)
                    If dimensions = 1 Then
                        Set i_items(j) = Nothing
                    Else
                        Set i_items(i - 1, j) = Nothing
                    End If
                Next
            Next
        End Sub

        ' Adds a value to the array
        Public Sub Push(ByRef value)
            If i_items_count >= i_items_capacity Then
                ReDim Preserve i_items(i_items_capacity * 1.2 + 1)
                i_items_capacity = UBound(i_items) + 1
            End If

            If IsObject(value) Then
                Set i_items(i_items_count) = value
            Else
                i_items(i_items_count) = value
            End If

            i_items_count = i_items_count + 1
        End Sub

        ' Load properties from a ADO RecordSet object
        Public Sub LoadRecordSet(ByRef rs)
            Dim obj, field

            While Not rs.eof
                Set obj = New JSONobject

                For Each field In rs.fields
                    obj.Add field.name, field.value
                Next

                Push obj

                rs.MoveNext
            Wend

            Set obj = Nothing
        End Sub

        ' Returns the item at the specified index
        ' @param index as int - the desired item index
        Public Default Function ItemAt(ByVal index)
            Dim Len
            Len = me.length

            If Len > 0 And index < Len Then
                If IsObject(i_items(index)) Then
                    Set ItemAt = i_items(index)
                Else
                    ItemAt = i_items(index)
                End If
            Else
                Err.raise JSON_ERROR_INDEX_OUT_OF_BOUNDS, TypeName(me), "Index out of bounds."
            End If
        End Function

        ' Serializes this JSONarray object in JSON formatted string value
        ' (uses the JSONobject.SerializeArray method)
        Public Function Serialize()
            Dim js, out, instantiated, actualLCID

            actualLCID = Response.LCID
            Response.LCID = 1033

            If Not IsEmpty(i_parent) Then
                If TypeName(i_parent) = "JSONobject" Then
                    Set js = i_parent
                    i_defaultPropertyName = i_parent.defaultPropertyName
                End If
            End If

            If IsEmpty(js) Then
                Set js = New JSONobject
                js.defaultPropertyName = i_defaultPropertyName
                instantiated = True
            End If

            out = js.SerializeArray(me)

            If instantiated Then Set js = Nothing

            Response.LCID = actualLCID

            Serialize = out
        End Function

        ' Writes the serialized array to the response
        Public Function Write()
            Response.Write Serialize()
        End Function
    End Class

    Class JSONscript
        Dim i_version
        Dim s_value, s_nullString

        ' The value
        Public Property Get value
            value = s_value
        End Property

        Public Property Let value(newValue)
            If (TypeName(newValue) <> "String") Then
                Err.raise JSON_ERROR_NOT_A_STRING, TypeName(me), "The value assigned is not a string."
            End If

            If (Len(newValue) = 0) Then newValue = s_nullString
            s_value = newValue
        End Property

        ' Constructor and destructor
        Private Sub class_initialize()
            i_version = "1.0.0"

            s_nullString = "null"
            s_value = s_nullString
        End Sub

        ' Serializes this object by outputting the raw value
        Public Function Serialize()
            Serialize = s_value
        End Function

        ' Writes the serialized object to the response
        Public Function Write()
            Response.Write Serialize()
        End Function
    End Class

    ' JSON pair class
    ' represents a name/value pair of a JSON object
    Class JSONpair
        Dim i_name, i_value
        Dim i_parent

        ' The name or key of the pair
        Public Property Get name
            name = i_name
        End Property

        Public Property Let name(val)
            i_name = val
        End Property

        ' The value of the pair
        Public Property Get value
            If IsObject(i_value) Then
                Set value = i_value
            Else
                value = i_value
            End If
        End Property

        Public Property Let value(val)
            i_value = val
        End Property

        Public Property Set value(val)
            Set i_value = val
        End Property

        ' The parent object
        Public Property Get parent
            Set parent = i_parent
        End Property

        Public Property Set parent(val)
            Set i_parent = val
        End Property

        ' Constructor and destructor
        Private Sub class_initialize
        End Sub

        Private Sub class_terminate
            If IsObject(value) Then Set value = Nothing
        End Sub
    End Class
%>

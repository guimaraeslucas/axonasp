<%
' FlexForms is a powerful HTML forms generator/builder class to output HTML forms from using a natural arrays approach.
' This is an adapted port of the original PHP version by CubicleSoft adapted to AxonASP Classic ASP.
' License: MIT License
' Git repository: https://github.com/cubiclesoft/php-flexforms/

Dim g_ff_form_handlers
Set g_ff_form_handlers = CreateObject("Scripting.Dictionary")
g_ff_form_handlers.Add "init", CreateObject("Scripting.Dictionary")
g_ff_form_handlers.Add "field_string", CreateObject("Scripting.Dictionary")
g_ff_form_handlers.Add "field_type", CreateObject("Scripting.Dictionary")
g_ff_form_handlers.Add "table_row", CreateObject("Scripting.Dictionary")
g_ff_form_handlers.Add "cleanup", CreateObject("Scripting.Dictionary")
g_ff_form_handlers.Add "finalize", CreateObject("Scripting.Dictionary")

Class FlexForms
    ' Private properties
    Private m_state
    Private m_secretKey
    Private m_extraInfo
    Private m_autoNonce
    Private m_version
    Private axFuncs

    ' Class initializer
    Private Sub Class_Initialize()
        ' Instantiate the native AxonASP functions for maximum performance
        Set axFuncs = Server.CreateObject("G3AXON.FUNCTIONS")

        Set m_state = CreateObject("Scripting.Dictionary")
        m_state.Add "formnum", 0
        m_state.Add "formidbase", "ff_form_"
        m_state.Add "responsive", True
        m_state.Add "formtables", True
        m_state.Add "formwidths", True
        m_state.Add "autofocused", False
        m_state.Add "jqueryuiused", False
        m_state.Add "jqueryuitheme", "smoothness"
        m_state.Add "supporturl", "support"
        m_state.Add "ajax", False
        m_state.Add "action", GetRequestURLBase()
        
        m_state.Add "customfieldtypes", CreateObject("Scripting.Dictionary")
        m_state.Add "js", CreateObject("Scripting.Dictionary")
        m_state.Add "jsoutput", CreateObject("Scripting.Dictionary")
        m_state.Add "css", CreateObject("Scripting.Dictionary")
        m_state.Add "cssoutput", CreateObject("Scripting.Dictionary")
        
        ' Internal execution state
        m_state.Add "hidden", CreateObject("Scripting.Dictionary")
        m_state.Add "insiderow", False
        m_state.Add "insiderowwidth", False
        m_state.Add "firstitem", False

        m_secretKey = False
        m_extraInfo = ""
        m_autoNonce = False
        m_version = ""
    End Sub

    ' Destructor to clean up objects
    Private Sub Class_Terminate()
        Set m_state = Nothing
        Set axFuncs = Nothing
    End Sub

    ' --- Static-like methods (Public but shared state) ---
    Public Sub RegisterFormHandler(mode, callback)
        If g_ff_form_handlers.Exists(mode) Then
            Dim dict
            Set dict = g_ff_form_handlers(mode)
            dict.Add dict.Count, callback
        End If
    End Sub

    ' --- Getters e Setters ---
    Public Function GetState()
        Set GetState = m_state
    End Function

    Public Sub SetState(newstate)
        If IsObject(newstate) Then
            Dim key
            For Each key In newstate.Keys
                If m_state.Exists(key) Then
                    If IsObject(newstate(key)) Then
                        Set m_state(key) = newstate(key)
                    Else
                        m_state(key) = newstate(key)
                    End If
                Else
                    If IsObject(newstate(key)) Then
                        m_state.Add key, newstate(key)
                    Else
                        m_state.Add key, newstate(key)
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub SetAjax(enable)
        If enable Then
            m_state("ajax") = True
            m_state("action") = GetFullRequestURLBase("")
        Else
            m_state("ajax") = False
            m_state("action") = GetRequestURLBase()
        End If
    End Sub

    Public Sub AddJS(name, info)
        If IsObject(m_state("js")) Then
            If m_state("js").Exists(name) Then m_state("js").Remove(name)
            m_state("js").Add name, info
        End If
    End Sub

    Public Sub SetJSOutput(name)
        If IsObject(m_state("jsoutput")) Then
            If m_state("jsoutput").Exists(name) Then m_state("jsoutput").Remove(name)
            m_state("jsoutput").Add name, True
        End If
    End Sub

    Public Sub AddCSS(name, info)
        If IsObject(m_state("css")) Then
            If m_state("css").Exists(name) Then m_state("css").Remove(name)
            m_state("css").Add name, info
        End If
    End Sub

    Public Sub SetCSSOutput(name)
        If IsObject(m_state("cssoutput")) Then
            If m_state("cssoutput").Exists(name) Then m_state("cssoutput").Remove(name)
            m_state("cssoutput").Add name, True
        End If
    End Sub

    Public Sub SetVersion(newversion)
        m_version = Server.URLEncode(CStr(newversion))
    End Sub

    Public Sub SetSecretKey(secretkey)
        m_secretKey = CStr(secretkey)
    End Sub

    Public Sub SetTokenExtraInfo(extrainfo)
        m_extraInfo = CStr(extrainfo)
    End Sub

    ' --- Optimized Cryptography Helpers ---
    Private Function HexToStr(byVal hexStr)
        Dim i, out
        out = ""
        For i = 1 To Len(hexStr) Step 2
            out = out & Chr("&H" & Mid(hexStr, i, 2))
        Next
        HexToStr = out
    End Function

    Private Function HashHMAC(algo, text, key)
        Dim lenBlock, ipad, opad, i, k, innerHash
        lenBlock = 64
        
        If Len(key) > lenBlock Then
            key = HexToStr(axFuncs.AxHash(algo, key))
        End If
        
        If Len(key) < lenBlock Then
            key = key & String(lenBlock - Len(key), Chr(0))
        End If
        
        ipad = ""
        opad = ""
        For i = 1 To lenBlock
            k = Asc(Mid(key, i, 1))
            ipad = ipad & Chr(k Xor &H36)
            opad = opad & Chr(k Xor &H5C)
        Next
        
        innerHash = HexToStr(axFuncs.AxHash(algo, ipad & text))
        HashHMAC = axFuncs.AxHash(algo, opad & innerHash)
    End Function

    ' --- Security Tokens (XSRF Defense) ---
    Public Function CreateSecurityToken(action, extra)
        If m_secretKey = False Then
            Response.Write FFTranslate("Secret key not set for form.")
            Response.End
        End If

        Dim str
        str = action & ":" & m_extraInfo
        
        If IsArray(extra) Then
            Dim val
            For Each val In extra
                str = str & ":" & val
            Next
        ElseIf VarType(extra) = vbString And extra <> "" Then
            Dim keys, key
            keys = axFuncs.AxExplode(",", extra)
            For Each key In keys
                key = Trim(key)
                If key <> "" And Request(key) <> "" Then 
                    str = str & ":" & CStr(Request(key))
                End If
            Next
        End If

        CreateSecurityToken = HashHMAC("sha1", str, m_secretKey)
    End Function

    Public Function IsSecExtraOpt(opt)
        Dim secExtra
        secExtra = Request("sec_extra")
        If secExtra <> "" Then
            IsSecExtraOpt = (InStr("," & secExtra & ",", "," & opt & ",") > 0)
        Else
            IsSecExtraOpt = False
        End If
    End Function

    Public Sub CheckSecurityToken(action)
        If Request(action) <> "" Then
            Dim secT, secExtra, expectedToken
            secT = Request("sec_t")
            secExtra = Request("sec_extra")
            expectedToken = CreateSecurityToken(Request(action), secExtra)
            
            If secT = "" Or secT <> expectedToken Then
                Response.Write FFTranslate("Invalid security token. Cross-site scripting (XSRF attack) attempt averted.")
                Response.End
            Else
                Set m_autoNonce = CreateObject("Scripting.Dictionary")
                m_autoNonce.Add "action", action
                m_autoNonce.Add "value", Request(action)
            End If
        End If
    End Sub

' --- Submit Button Rendering ---
    Private Sub ProcessSubmit(options)
        Response.Write "        <div class=""formsubmit"">" & vbCrLf
        Response.Write "            <div class=""formsubmitinner"">" & vbCrLf

        Dim submitVar, submitName
        submitName = ""
        If options.Exists("submitname") Then submitName = options("submitname")

        submitVar = options("submit")

        ' Se o submit for um Dictionary (múltiplos botões de envio)
        If IsObject(submitVar) Then
            Dim key
            For Each key In submitVar.Keys
                Dim currentName, valStr
                valStr = submitVar(key)
                
                ' Se a chave for numérica, assume o submitname padrão (is_int no PHP)
                If IsNumeric(key) Then
                    currentName = submitName
                Else
                    currentName = key
                End If
                
                Response.Write "                <input class=""submit"" type=""submit"""
                If currentName <> "" Then
                    If options.Exists("hashnames") Then
                        If options("hashnames") Then
                            Response.Write " name=""" & GetHashedFieldName(currentName) & """"
                        Else
                            Response.Write " name=""" & axFuncs.AxHtmlSpecialChars(currentName) & """"
                        End If
                    Else
                        Response.Write " name=""" & axFuncs.AxHtmlSpecialChars(currentName) & """"
                    End If
                End If
                Response.Write " value=""" & axFuncs.AxHtmlSpecialChars(FFTranslate(valStr)) & """ />" & vbCrLf
            Next
        Else
            ' Se o submit for uma String simples (botão único)
            Response.Write "                <input class=""submit"" type=""submit"""
            If submitName <> "" Then
                If options.Exists("hashnames") Then
                    If options("hashnames") Then
                        Response.Write " name=""" & GetHashedFieldName(submitName) & """"
                    Else
                        Response.Write " name=""" & axFuncs.AxHtmlSpecialChars(submitName) & """"
                    End If
                Else
                    Response.Write " name=""" & axFuncs.AxHtmlSpecialChars(submitName) & """"
                End If
            End If
            Response.Write " value=""" & axFuncs.AxHtmlSpecialChars(FFTranslate(submitVar)) & """ />" & vbCrLf
        End If

        Response.Write "            </div>" & vbCrLf
        Response.Write "        </div>" & vbCrLf
    End Sub

    ' --- Component and Message Rendering ---
    Public Sub OutputFormCSS(delaycss)
        If Not m_state("cssoutput").Exists("formcss") Then
            If delaycss Then
                Dim cssInfo
                Set cssInfo = CreateObject("Scripting.Dictionary")
                cssInfo.Add "mode", "link"
                cssInfo.Add "dependency", False
                cssInfo.Add "src", m_state("supporturl") & "/flex_forms.css"
                
                m_state("css").Add "formcss", cssInfo
            Else
                Dim verSuffix
                verSuffix = ""
                If m_version <> "" Then 
                    If InStr(m_state("supporturl") & "/flex_forms.css", "?") > 0 Then
                        verSuffix = "&" & m_version
                    Else
                        verSuffix = "?" & m_version
                    End If
                End If
                Response.Write "<link rel=""stylesheet"" href=""" & axFuncs.AxHtmlSpecialChars(m_state("supporturl") & "/flex_forms.css" & verSuffix) & """ type=""text/css"" media=""all"" />" & vbCrLf
                m_state("cssoutput").Add "formcss", True
            End If
        End If
    End Sub

    Public Sub OutputMessage(msgType, message)
        msgType = LCase(CStr(msgType))
        If msgType = "warn" Then msgType = "warning"

        Call OutputFormCSS(False)

        Response.Write "<div class=""ff_formmessagewrap"">" & vbCrLf
        Response.Write "    <div class=""ff_formmessagewrapinner"">" & vbCrLf
        Response.Write "        <div class=""message message" & axFuncs.AxHtmlSpecialChars(msgType) & """>" & vbCrLf
        Response.Write "            " & FFTranslate(CStr(message)) & vbCrLf
        Response.Write "        </div>" & vbCrLf
        Response.Write "    </div>" & vbCrLf
        Response.Write "</div>" & vbCrLf
    End Sub

    Public Function GetEncodedSignedMessage(msgType, message, prefix)
        Dim translatedMsg
        translatedMsg = FFTranslate(message)
        
        Dim extra
        extra = Array(msgType, translatedMsg)
        
        GetEncodedSignedMessage = Server.URLEncode(prefix & "msgtype") & "=" & Server.URLEncode(msgType) & "&" & _
                                 Server.URLEncode(prefix & "msg") & "=" & Server.URLEncode(translatedMsg) & "&" & _
                                 Server.URLEncode(prefix & "msg_t") & "=" & CreateSecurityToken("forms__message", extra)
    End Function

    Public Sub OutputSignedMessage(prefix)
        Dim msgType, msg, msgT
        msgType = Request(prefix & "msgtype")
        msg = Request(prefix & "msg")
        msgT = Request(prefix & "msg_t")
        
        If msgType <> "" And msg <> "" And msgT <> "" Then
            Dim extra
            extra = Array(msgType, msg)
            If msgT = CreateSecurityToken("forms__message", extra) Then
                Call OutputMessage(msgType, axFuncs.AxHtmlSpecialChars(msg))
            End If
        End If
    End Sub

    Public Sub OutputJQuery(delayjs)
        If Not m_state("jsoutput").Exists("jquery") Then
            If delayjs Then
                Dim jsInfo
                Set jsInfo = CreateObject("Scripting.Dictionary")
                jsInfo.Add "mode", "src"
                jsInfo.Add "dependency", False
                jsInfo.Add "src", m_state("supporturl") & "/jquery-3.5.0.min.js"
                jsInfo.Add "detect", "jQuery"
                m_state("js").Add "jquery", jsInfo
            Else
                Response.Write "<script type=""text/javascript"" src=""" & axFuncs.AxHtmlSpecialChars(m_state("supporturl") & "/jquery-3.5.0.min.js") & """></script>" & vbCrLf
                m_state("jsoutput").Add "jquery", True
            End If
        End If
    End Sub

    Public Sub OutputJQueryUI(delayjs)
        Call OutputJQuery(delayjs)
        
        If Not m_state("jsoutput").Exists("jqueryui") Then
            If delayjs Then
                Dim cssInfo
                Set cssInfo = CreateObject("Scripting.Dictionary")
                cssInfo.Add "mode", "link"
                cssInfo.Add "dependency", False
                cssInfo.Add "src", m_state("supporturl") & "/jquery_ui_themes/" & m_state("jqueryuitheme") & "/jquery-ui-1.12.1.css"
                m_state("css").Add "jqueryui", cssInfo
                
                Dim jsInfo
                Set jsInfo = CreateObject("Scripting.Dictionary")
                jsInfo.Add "mode", "src"
                jsInfo.Add "dependency", "jquery"
                jsInfo.Add "src", m_state("supporturl") & "/jquery-ui-1.12.1.min.js"
                jsInfo.Add "detect", "jQuery.ui"
                m_state("js").Add "jqueryui", jsInfo
            Else
                Response.Write "<link rel=""stylesheet"" href=""" & axFuncs.AxHtmlSpecialChars(m_state("supporturl") & "/jquery_ui_themes/" & m_state("jqueryuitheme") & "/jquery-ui-1.12.1.css") & """ type=""text/css"" media=""all"" />" & vbCrLf
                Response.Write "<script type=""text/javascript"" src=""" & axFuncs.AxHtmlSpecialChars(m_state("supporturl") & "/jquery-ui-1.12.1.min.js") & """></script>" & vbCrLf
                m_state("cssoutput").Add "jqueryui", True
                m_state("jsoutput").Add "jqueryui", True
            End If
        End If
    End Sub

    ' --- Advanced Sanitization ---
    Public Function FilenameSafe(filename)
        Dim regEx
        Set regEx = New RegExp
        regEx.Global = True
        
        regEx.Pattern = "[^A-Za-z0-9_.\-]"
        filename = regEx.Replace(filename, " ")
        
        filename = axFuncs.AxTrim(filename)
        Do While Left(filename, 1) = "."
            filename = Mid(filename, 2)
        Loop
        Do While Right(filename, 1) = "."
            filename = Left(filename, Len(filename) - 1)
        Loop
        
        regEx.Pattern = "\s+"
        FilenameSafe = regEx.Replace(axFuncs.AxTrim(filename), "-")
    End Function

    Public Function NormalizeFiles(key)
        Dim uploader, allFiles, result, entry, i
        Set uploader = Server.CreateObject("G3FILEUPLOADER")
        ' G3FILEUPLOADER.GetAllFilesInfo returns a VBScript array of Dictionaries
        allFiles = uploader.GetAllFilesInfo()
        
        Set result = CreateObject("Scripting.Dictionary")
        
        If IsArray(allFiles) Then
            For i = 0 To UBound(allFiles)
                Set entry = allFiles(i)
                If entry("FieldName") = key Then
                    Dim normalizedEntry
                    Set normalizedEntry = CreateObject("Scripting.Dictionary")
                    If entry("IsSuccess") Then
                        normalizedEntry.Add "success", True
                        normalizedEntry.Add "file", entry("FinalPath")
                        normalizedEntry.Add "name", FilenameSafe(entry("OriginalFileName"))
                        normalizedEntry.Add "ext", entry("Extension")
                        normalizedEntry.Add "type", entry("MimeType")
                        normalizedEntry.Add "size", entry("Size")
                    Else
                        normalizedEntry.Add "success", False
                        normalizedEntry.Add "error", FFTranslate(entry("ErrorMessage"))
                        normalizedEntry.Add "errorcode", "upload_error"
                    End If
                    result.Add result.Count, normalizedEntry
                End If
            Next
        End If
        
        NormalizeFiles = result.Items
    End Function

    Public Function GetValue(key, defaultValue)
        If Request(key) <> "" Then
            GetValue = Request(key)
        Else
            GetValue = defaultValue
        End If
    End Function

    Public Function GetSelectValues(data)
        Dim result
        Set result = CreateObject("Scripting.Dictionary")
        If IsArray(data) Then
            Dim val
            For Each val In data
                result.Add val, True
            Next
        End If
        Set GetSelectValues = result
    End Function

    Public Function ProcessInfoDefaults(info, defaults)
        If IsObject(info) And IsObject(defaults) Then
            Dim key
            For Each key In defaults.Keys
                If Not info.Exists(key) Then
                    If IsObject(defaults(key)) Then
                        Set info(key) = defaults(key)
                    Else
                        info(key) = defaults(key)
                    End If
                End If
            Next
        End If
        Set ProcessInfoDefaults = info
    End Function

    Public Sub SetNestedPathValue(ByRef data, pathparts, val)
        Dim curr, key, i
        Set curr = data
        For i = 0 To UBound(pathparts)
            key = pathparts(i)
            If Not curr.Exists(key) Then
                curr.Add key, CreateObject("Scripting.Dictionary")
            End If
            If i < UBound(pathparts) Then
                Set curr = curr(key)
            Else
                If IsObject(val) Then
                    Set curr(key) = val
                Else
                    curr(key) = val
                End If
            End If
        Next
    End Sub

    Public Function GetIDDiff(origids, newids)
        Dim result, remove, add, id
        Set result = CreateObject("Scripting.Dictionary")
        Set remove = CreateObject("Scripting.Dictionary")
        Set add = CreateObject("Scripting.Dictionary")
        
        For Each id In origids.Keys
            If Not newids.Exists(id) Then remove.Add id, origids(id)
        Next
        
        For Each id In newids.Keys
            If Not origids.Exists(id) Then add.Add id, newids(id)
        Next
        
        result.Add "remove", remove
        result.Add "add", add
        Set GetIDDiff = result
    End Function

    Public Function GetHashedFieldName(name)
        If m_secretKey = False Then
            Response.Write FFTranslate("Secret key not set.")
            Response.End
        End If
        GetHashedFieldName = "f_" & HashHMAC("md5", name, m_secretKey)
    End Function

    Public Function GetHashedFieldValues(nameswithdefaults)
        If m_secretKey = False Then
            Response.Write FFTranslate("Secret key not set.")
            Response.End
        End If
        
        Dim result, name, defaultVal, name2
        Set result = CreateObject("Scripting.Dictionary")
        For Each name In nameswithdefaults.Keys
            defaultVal = nameswithdefaults(name)
            name2 = "f_" & HashHMAC("md5", name, m_secretKey)
            
            If Request(name2) <> "" Then
                result.Add name, Request(name2)
            Else
                result.Add name, defaultVal
            End If
        Next
        Set GetHashedFieldValues = result
    End Function

    ' --- Environment / Server Utilities ---
    Public Function IsSSLRequest()
        Dim httpsVar, serverPort
        httpsVar = LCase(Request.ServerVariables("HTTPS"))
        serverPort = Request.ServerVariables("SERVER_PORT")
        
        If httpsVar = "on" Or httpsVar = "1" Or serverPort = "443" Then
            IsSSLRequest = True
        Else
            IsSSLRequest = False
        End If
    End Function

    Public Function GetRequestURLBase()
        Dim uri
        uri = Request.ServerVariables("SCRIPT_NAME")
        If uri = "" Then uri = "/"
        GetRequestURLBase = uri
    End Function

    Public Function GetRequestHost(protocol)
        Dim host, ssl
        ssl = IsSSLRequest()
        If protocol = "https" Then ssl = True
        
        host = Request.ServerVariables("HTTP_HOST")
        If host = "" Then
            host = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
        End If
        
        If ssl Then
            GetRequestHost = "https://" & host
        Else
            GetRequestHost = "http://" & host
        End If
    End Function

    Public Function GetFullRequestURLBase(protocol)
        GetFullRequestURLBase = GetRequestHost(protocol) & GetRequestURLBase()
    End Function

    Public Function JSSafe(data)
        Dim out
        out = Replace(data, "'", "\'")
        out = Replace(out, vbCr, "\r")
        out = Replace(out, vbLf, "\n")
        JSSafe = out
    End Function

    Public Function FFTranslate(text)
        FFTranslate = text
    End Function

    ' --- Main Form Generation Engine ---
    Private Sub InitFormVars(ByRef options)
        m_state("hidden").RemoveAll
        m_state("insiderow") = False
        m_state("firstitem") = False
        
        Dim callback, i
        Dim initHandlers
        Set initHandlers = g_ff_form_handlers("init")
        For i = 0 To initHandlers.Count - 1
            Set callback = initHandlers(i)
            ' Call the callback with state and options
            ' In VBScript, we use Execute or just call if it's a GetRef
            callback m_state, options
        Next
    End Sub

    Private Sub AlterField(num, ByRef field, id)
        Dim callback, i
        Dim typeHandlers
        Set typeHandlers = g_ff_form_handlers("field_type")
        For i = 0 To typeHandlers.Count - 1
            Set callback = typeHandlers(i)
            callback m_state, num, field, id
        Next
    End Sub

    Private Sub ProcessField(num, ByRef field, id)
        If VarType(field) = vbString Then
            If field = "split" And Not m_state("insiderow") Then
                Response.Write "<hr />" & vbCrLf
            ElseIf field = "startrow" Then
                If m_state("insiderow") Then
                    If m_state("responsive") And m_state("insiderowwidth") Then Response.Write "<td></td>"
                    Response.Write "</tr><tr>" & vbCrLf
                ElseIf m_state("formtables") Then
                    m_state("insiderow") = True
                    m_state("insiderowwidth") = False
                    Response.Write "            <div class=""fieldtablewrap" & (IfThen(m_state("firstitem"), " firstitem", "")) & """><table class=""rowwrap""><tbody><tr>" & vbCrLf
                    m_state("firstitem") = False
                End If
            ElseIf field = "endrow" And m_state("formtables") And m_state("insiderow") Then
                If m_state("responsive") And m_state("insiderowwidth") Then Response.Write "<td></td>"
                Response.Write "            </tr></tbody></table></div>" & vbCrLf
                m_state("insiderow") = False
            ElseIf Left(field, 5) = "html:" Then
                Response.Write Mid(field, 6)
            End If

            Dim callback, i
            Dim stringHandlers
            Set stringHandlers = g_ff_form_handlers("field_string")
            For i = 0 To stringHandlers.Count - 1
                Set callback = stringHandlers(i)
                callback m_state, num, field, id
            Next
        Else
            If Not field.Exists("type") Then Exit Sub
            If field.Exists("use") Then
                If Not field("use") Then Exit Sub
            End If

            If field("type") = "hidden" Then
                Response.Write "                <input type=""hidden"" id=""" & axFuncs.AxHtmlSpecialChars(id) & """ name=""" & axFuncs.AxHtmlSpecialChars(field("name")) & """ value=""" & axFuncs.AxHtmlSpecialChars(field("value")) & """ />" & vbCrLf
            ElseIf m_state("customfieldtypes").Exists(field("type")) Then
                Call AlterField(num, field, id)
            Else
                If m_state("insiderow") Then
                    If Not m_state("responsive") Or Not field.Exists("width") Then
                        Response.Write "<td>"
                    Else
                        Response.Write "<td style=""width: " & axFuncs.AxHtmlSpecialChars(field("width")) & ";"">"
                        m_state("insiderowwidth") = True
                    End If
                End If

                Dim firstItemClass
                firstItemClass = ""
                If (field.Exists("split") And field("split") = False) Or m_state("firstitem") Then
                    firstItemClass = " firstitem"
                End If
                
                Response.Write "            <div class=""formitem" & firstItemClass & """>" & vbCrLf
                m_state("firstitem") = False

                If field.Exists("title") Then
                    Response.Write "            <div class=""formitemtitle"">" & axFuncs.AxHtmlSpecialChars(FFTranslate(field("title"))) & "</div>" & vbCrLf
                ElseIf field.Exists("htmltitle") Then
                    Response.Write "            <div class=""formitemtitle"">" & FFTranslate(field("htmltitle")) & "</div>" & vbCrLf
                ElseIf field("type") = "checkbox" And m_state("insiderow") Then
                    Response.Write "            <div class=""formitemtitle"">&nbsp;</div>" & vbCrLf
                End If

                If field.Exists("width") And Not m_state("formwidths") Then field.Remove "width"

                If field.Exists("name") And field.Exists("default") Then
                    If field("type") = "select" Then
                        If Not field.Exists("select") Then
                            Dim selVal
                            selVal = GetValue(field("name"), field("default"))
                            If IsArray(selVal) Then
                                Set field("select") = GetSelectValues(selVal)
                            Else
                                field("select") = selVal
                            End If
                        End If
                    ElseIf field("type") = "checkbox" Then
                        If Not field.Exists("check") And field.Exists("value") Then
                            If Request(field("name")) = field("value") Then
                                field("check") = True
                            ElseIf Request.ServerVariables("REQUEST_METHOD") = "GET" Then
                                field("check") = field("default")
                            Else
                                field("check") = False
                            End If
                        End If
                    Else
                        If Not field.Exists("value") Then field("value") = GetValue(field("name"), field("default"))
                    End If
                End If

                If field.Exists("focus") Then
                    If field("focus") And m_state("autofocused") = False Then m_state("autofocused") = id
                End If
                
                If field("type") = "select" And field.Exists("mode") Then
                    If field("mode") = "formhandler" Then field.Remove "mode"
                End If

                Call AlterField(num, field, id)

                Dim styleWidth, styleHeight
                styleWidth = ""
                If field.Exists("width") Then
                    styleWidth = " style=""" & (IfThen(m_state("responsive"), "max-", "")) & "width: " & axFuncs.AxHtmlSpecialChars(field("width")) & ";"""
                End If

                Select Case field("type")
                    Case "static"
                        Response.Write "            <div class=""formitemdata"">" & vbCrLf
                        Response.Write "                <div class=""staticwrap""" & styleWidth & ">" & axFuncs.AxHtmlSpecialChars(field("value")) & "</div>" & vbCrLf
                        Response.Write "            </div>" & vbCrLf

                    Case "text", "password"
                        Dim subType
                        subType = field("type")
                        If field.Exists("subtype") Then subType = field("subtype")
                        Response.Write "            <div class=""formitemdata"">" & vbCrLf
                        Response.Write "                <div class=""textitemwrap""" & styleWidth & "><input class=""text"" type=""" & axFuncs.AxHtmlSpecialChars(subType) & """ id=""" & axFuncs.AxHtmlSpecialChars(id) & """ name=""" & axFuncs.AxHtmlSpecialChars(field("name")) & """ value=""" & axFuncs.AxHtmlSpecialChars(field("value")) & """" & (IfThen(m_state("autofocused") = id, " autofocus", "")) & " /></div>" & vbCrLf
                        Response.Write "            </div>" & vbCrLf

                    Case "checkbox"
                        Response.Write "            <div class=""formitemdata"">" & vbCrLf
                        Response.Write "                <div class=""checkboxitemwrap""" & styleWidth & ">" & vbCrLf
                        Response.Write "                    <input class=""checkbox"" type=""checkbox"" id=""" & axFuncs.AxHtmlSpecialChars(id) & """ name=""" & axFuncs.AxHtmlSpecialChars(field("name")) & """ value=""" & axFuncs.AxHtmlSpecialChars(field("value")) & """" & (IfThen(field.Exists("check") And field("check"), " checked", "")) & (IfThen(m_state("autofocused") = id, " autofocus", "")) & " />" & vbCrLf
                        Response.Write "                    <label for=""" & axFuncs.AxHtmlSpecialChars(id) & """>" & axFuncs.AxHtmlSpecialChars(FFTranslate(field("display"))) & "</label>" & vbCrLf
                        Response.Write "                </div>" & vbCrLf
                        Response.Write "            </div>" & vbCrLf

                    Case "select"
                        Dim mode, isMultiple
                        isMultiple = False
                        If field.Exists("multiple") Then If field("multiple") = True Then isMultiple = True
                        
                        If Not isMultiple Then
                            mode = "select"
                            If field.Exists("mode") Then If field("mode") = "radio" Then mode = "radio"
                        Else
                            mode = "checkbox"
                            If field.Exists("mode") Then
                                If field("mode") = "formhandler" Or field("mode") = "select" Then mode = field("mode")
                            End If
                        End If

                        styleHeight = ""
                        If field.Exists("height") And isMultiple Then
                            styleHeight = " style=""height: " & axFuncs.AxHtmlSpecialChars(field("height")) & ";"""
                        End If

                        Dim selectDict
                        If Not field.Exists("select") Then
                            Set selectDict = CreateObject("Scripting.Dictionary")
                        ElseIf VarType(field("select")) = vbString Then
                            Set selectDict = CreateObject("Scripting.Dictionary")
                            selectDict.Add field("select"), True
                        Else
                            Set selectDict = field("select")
                        End If

                        Response.Write "            <div class=""formitemdata"">" & vbCrLf
                        
                        Dim idBase, idNum, optKey, optVal, id2
                        idBase = axFuncs.AxHtmlSpecialChars(id)
                        
                        If mode = "checkbox" Or mode = "radio" Then
                            idNum = 0
                            For Each optKey In field("options").Keys
                                optVal = field("options")(optKey)
                                If IsObject(optVal) Then ' Dictionary of options (optgroup-like)
                                    Dim optKey2, optVal2
                                    For Each optKey2 In optVal.Keys
                                        optVal2 = optVal(optKey2)
                                        id2 = idBase & (IfThen(idNum > 0, "_" & idNum, ""))
                                        Response.Write "                <div class=""" & mode & "itemwrap""" & styleWidth & ">" & vbCrLf
                                        Response.Write "                    <input class=""" & mode & """ type=""" & mode & """ id=""" & id2 & """ name=""" & axFuncs.AxHtmlSpecialChars(field("name")) & (IfThen(mode = "checkbox", "[]", "")) & """ value=""" & axFuncs.AxHtmlSpecialChars(optKey2) & """" & (IfThen(selectDict.Exists(optKey2), " checked", "")) & (IfThen(m_state("autofocused") = id, " autofocus", "")) & " />" & vbCrLf
                                        If m_state("autofocused") = id Then m_state("autofocused") = True
                                        Response.Write "                    <label for=""" & id2 & """>" & axFuncs.AxHtmlSpecialChars(FFTranslate(optKey)) & " - " & (IfThen(optVal2 = "", "&nbsp;", axFuncs.AxHtmlSpecialChars(FFTranslate(optVal2)))) & "</label>" & vbCrLf
                                        Response.Write "                </div>" & vbCrLf
                                        idNum = idNum + 1
                                    Next
                                Else
                                    id2 = idBase & (IfThen(idNum > 0, "_" & idNum, ""))
                                    Response.Write "                <div class=""" & mode & "itemwrap""" & styleWidth & ">" & vbCrLf
                                    Response.Write "                    <input class=""" & mode & """ type=""" & mode & """ id=""" & id2 & """ name=""" & axFuncs.AxHtmlSpecialChars(field("name")) & (IfThen(mode = "checkbox", "[]", "")) & """ value=""" & axFuncs.AxHtmlSpecialChars(optKey) & """" & (IfThen(selectDict.Exists(optKey), " checked", "")) & (IfThen(m_state("autofocused") = id, " autofocus", "")) & " />" & vbCrLf
                                    If m_state("autofocused") = id Then m_state("autofocused") = True
                                    Response.Write "                    <label for=""" & id2 & """>" & (IfThen(optVal = "", "&nbsp;", axFuncs.AxHtmlSpecialChars(FFTranslate(optVal)))) & "</label>" & vbCrLf
                                    Response.Write "                </div>" & vbCrLf
                                    idNum = idNum + 1
                                End If
                            Next
                        Else
                            Response.Write "                <div class=""selectitemwrap""" & styleWidth & ">" & vbCrLf
                            Response.Write "                    <select class=""" & (IfThen(isMultiple, "multi", "single")) & """ id=""" & idBase & """ name=""" & axFuncs.AxHtmlSpecialChars(field("name")) & (IfThen(isMultiple, "[]", "")) & """" & (IfThen(isMultiple, " multiple", "")) & (IfThen(m_state("autofocused") = id, " autofocus", "")) & styleHeight & ">" & vbCrLf
                            
                            For Each optKey In field("options").Keys
                                optVal = field("options")(optKey)
                                If IsObject(optVal) Then
                                    Response.Write "                        <optgroup label=""" & axFuncs.AxHtmlSpecialChars(FFTranslate(optKey)) & """>" & vbCrLf
                                    Dim optKey3, optVal3
                                    For Each optKey3 In optVal.Keys
                                        optVal3 = optVal(optKey3)
                                        Response.Write "                            <option value=""" & axFuncs.AxHtmlSpecialChars(optKey3) & """" & (IfThen(selectDict.Exists(optKey3), " selected", "")) & ">" & (IfThen(optVal3 = "", "&nbsp;", axFuncs.AxHtmlSpecialChars(FFTranslate(optVal3)))) & "</option>" & vbCrLf
                                    Next
                                    Response.Write "                        </optgroup>" & vbCrLf
                                Else
                                    Response.Write "                        <option value=""" & axFuncs.AxHtmlSpecialChars(optKey) & """" & (IfThen(selectDict.Exists(optKey), " selected", "")) & ">" & (IfThen(optVal = "", "&nbsp;", axFuncs.AxHtmlSpecialChars(FFTranslate(optVal)))) & "</option>" & vbCrLf
                                End If
                            Next
                            Response.Write "                    </select>" & vbCrLf
                            Response.Write "                </div>" & vbCrLf
                        End If
                        Response.Write "            </div>" & vbCrLf

                    Case "textarea"
                        Response.Write "            <div class=""formitemdata"">" & vbCrLf
                        Response.Write "                <div class=""textareawrap""" & styleWidth & "><textarea class=""text""" & (IfThen(field.Exists("height"), " style=""height: " & axFuncs.AxHtmlSpecialChars(field("height")) & ";""", "")) & " id=""" & axFuncs.AxHtmlSpecialChars(id) & """ name=""" & axFuncs.AxHtmlSpecialChars(field("name")) & """ rows=""5"" cols=""50""" & (IfThen(m_state("autofocused") = id, " autofocus", "")) & ">" & axFuncs.AxHtmlSpecialChars(field("value")) & "</textarea></div>" & vbCrLf
                        Response.Write "            </div>" & vbCrLf

                    Case "table"
                        Dim idTable
                        idTable = id & "_table"
                        Response.Write "            <div class=""formitemdata"">" & vbCrLf
                        If m_state("formtables") Then
                            Response.Write "                <table id=""" & axFuncs.AxHtmlSpecialChars(idTable) & """ class=""formitemtable" & (IfThen(field.Exists("class"), " " & axFuncs.AxHtmlSpecialChars(field("class")), "")) & """" & (IfThen(field.Exists("width"), " style=""" & (IfThen(m_state("responsive"), "max-", "")) & "width: " & axFuncs.AxHtmlSpecialChars(field("width")) & """", "")) & ">" & vbCrLf
                            Response.Write "                    <thead>" & vbCrLf
                            
                            Dim trAttrs, colAttrs, iCol, tableRowHandlers, keyAttr
                            Set trAttrs = CreateObject("Scripting.Dictionary")
                            trAttrs.Add "class", "head"
                            Set colAttrs = CreateObject("Scripting.Dictionary")
                            If Not field.Exists("cols") Then field.Add "cols", Array()
                            For iCol = 0 To UBound(field("cols"))
                                colAttrs.Add iCol, CreateObject("Scripting.Dictionary")
                            Next
                            
                            Set tableRowHandlers = g_ff_form_handlers("table_row")
                            For i = 0 To tableRowHandlers.Count - 1
                                Set callback = tableRowHandlers(i)
                                callback m_state, num, field, idTable, "head", -1, trAttrs, colAttrs, field("cols")
                            Next
                            
                            Response.Write "                    <tr"
                            For Each keyAttr In trAttrs.Keys
                                Response.Write " " & keyAttr & "=""" & axFuncs.AxHtmlSpecialChars(trAttrs(keyAttr)) & """"
                            Next
                            Response.Write ">" & vbCrLf
                            
                            For iCol = 0 To UBound(field("cols"))
                                Response.Write "                        <th"
                                For Each keyAttr In colAttrs(iCol).Keys
                                    Response.Write " " & keyAttr & "=""" & axFuncs.AxHtmlSpecialChars(colAttrs(iCol)(keyAttr)) & """"
                                Next
                                Response.Write ">" & (IfThen(field.Exists("htmlcols") And field("htmlcols"), FFTranslate(field("cols")(iCol)), axFuncs.AxHtmlSpecialChars(FFTranslate(field("cols")(iCol))))) & "</th>" & vbCrLf
                            Next
                            Response.Write "                    </tr>" & vbCrLf
                            Response.Write "                    </thead>" & vbCrLf
                            Response.Write "                    <tbody>" & vbCrLf
                            
                            Dim baseColAttrs, rowNum, altRow, rows, row, colAttrs2, iRowCol
                            Set baseColAttrs = CreateObject("Scripting.Dictionary")
                            For iCol = 0 To UBound(field("cols"))
                                Dim curCol, isNowrap
                                curCol = field("cols")(iCol)
                                isNowrap = False
                                If field.Exists("nowrap") Then
                                    If VarType(field("nowrap")) = vbString Then
                                        If curCol = field("nowrap") Then isNowrap = True
                                    ElseIf IsArray(field("nowrap")) Then
                                        For i = 0 To UBound(field("nowrap"))
                                            If curCol = field("nowrap")(i) Then isNowrap = True : Exit For
                                        Next
                                    End If
                                End If
                                Set baseColAttrs(iCol) = CreateObject("Scripting.Dictionary")
                                If isNowrap Then baseColAttrs(iCol).Add "class", "nowrap"
                            Next
                            
                            rowNum = 0
                            altRow = False
                            If field.Exists("callback") Then
                                rows = field("callback")(field)
                                field("rows") = rows
                            End If
                            
                            Do While True
                                If Not field.Exists("rows") Then Exit Do
                                rows = field("rows")
                                If Not IsArray(rows) Then Exit Do
                                If UBound(rows) = -1 Then Exit Do
                                
                                For i = 0 To UBound(rows)
                                    row = rows(i)
                                    Set trAttrs = CreateObject("Scripting.Dictionary")
                                    trAttrs.Add "class", "row" & (IfThen(altRow, " altrow", ""))
                                    
                                    Set colAttrs2 = CreateObject("Scripting.Dictionary")
                                    For iCol = 0 To baseColAttrs.Count - 1
                                        Set colAttrs2(iCol) = CreateObject("Scripting.Dictionary")
                                        For Each keyAttr In baseColAttrs(iCol).Keys
                                            colAttrs2(iCol).Add keyAttr, baseColAttrs(iCol)(keyAttr)
                                        Next
                                    Next
                                    
                                    For j = 0 To tableRowHandlers.Count - 1
                                        Set callback = tableRowHandlers(j)
                                        callback m_state, num, field, idTable, "body", rowNum, trAttrs, colAttrs2, row
                                    Next
                                    
                                    If UBound(row) < colAttrs2.Count - 1 Then
                                        colAttrs2(UBound(row))("colspan") = (colAttrs2.Count - UBound(row))
                                    End If
                                    
                                    Response.Write "                    <tr"
                                    For Each keyAttr In trAttrs.Keys
                                        Response.Write " " & keyAttr & "=""" & axFuncs.AxHtmlSpecialChars(trAttrs(keyAttr)) & """"
                                    Next
                                    Response.Write ">" & vbCrLf
                                    
                                    For iRowCol = 0 To UBound(row)
                                        Response.Write "                        <td"
                                        If colAttrs2.Exists(iRowCol) Then
                                            For Each keyAttr In colAttrs2(iRowCol).Keys
                                                Response.Write " " & keyAttr & "=""" & axFuncs.AxHtmlSpecialChars(colAttrs2(iRowCol)(keyAttr)) & """"
                                            Next
                                        End If
                                        Response.Write ">" & row(iRowCol) & "</td>" & vbCrLf
                                    Next
                                    Response.Write "                    </tr>" & vbCrLf
                                    rowNum = rowNum + 1
                                    altRow = Not altRow
                                Next
                                
                                If field.Exists("callback") Then
                                    rows = field("callback")(field)
                                    field("rows") = rows
                                Else
                                    field.Remove "rows"
                                End If
                            Loop
                            
                            Response.Write "                    </tbody>" & vbCrLf
                            Response.Write "                </table>" & vbCrLf
                        Else
                            ' Non-table wrap
                            Response.Write "                <div class=""nontablewrap"" id=""" & axFuncs.AxHtmlSpecialChars(idTable) & """" & (IfThen(field.Exists("width"), " style=""" & (IfThen(m_state("responsive"), "max-", "")) & "width: " & axFuncs.AxHtmlSpecialChars(field("width")) & """", "")) & ">" & vbCrLf
                            
                            Set trAttrs = CreateObject("Scripting.Dictionary")
                            Dim headColAttrs
                            Set headColAttrs = CreateObject("Scripting.Dictionary")
                            If Not field.Exists("cols") Then field.Add "cols", Array()
                            For iCol = 0 To UBound(field("cols"))
                                Set headColAttrs(iCol) = CreateObject("Scripting.Dictionary")
                                headColAttrs(iCol).Add "class", "nontable_th" & (IfThen(iCol = 0, " firstcol", ""))
                            Next
                            
                            Set tableRowHandlers = g_ff_form_handlers("table_row")
                            For i = 0 To tableRowHandlers.Count - 1
                                Set callback = tableRowHandlers(i)
                                callback m_state, num, field, idTable, "head", -1, trAttrs, headColAttrs, field("cols")
                            Next
                            
                            Set baseColAttrs = CreateObject("Scripting.Dictionary")
                            For iCol = 0 To UBound(field("cols"))
                                Set baseColAttrs(iCol) = CreateObject("Scripting.Dictionary")
                                baseColAttrs(iCol).Add "class", "nontable_td"
                            Next
                            
                            rowNum = 0
                            altRow = False
                            If field.Exists("callback") Then
                                rows = field("callback")(field)
                                field("rows") = rows
                            End If
                            
                            Do While True
                                If Not field.Exists("rows") Then Exit Do
                                rows = field("rows")
                                If Not IsArray(rows) Then Exit Do
                                If UBound(rows) = -1 Then Exit Do
                                
                                For i = 0 To UBound(rows)
                                    row = rows(i)
                                    Set trAttrs = CreateObject("Scripting.Dictionary")
                                    trAttrs.Add "class", "nontable_row" & (IfThen(altRow, " altrow", "")) & (IfThen(rowNum = 0, " firstrow", ""))
                                    
                                    Set colAttrs2 = CreateObject("Scripting.Dictionary")
                                    For iCol = 0 To baseColAttrs.Count - 1
                                        Set colAttrs2(iCol) = CreateObject("Scripting.Dictionary")
                                        For Each keyAttr In baseColAttrs(iCol).Keys
                                            colAttrs2(iCol).Add keyAttr, baseColAttrs(iCol)(keyAttr)
                                        Next
                                    Next
                                    
                                    For j = 0 To tableRowHandlers.Count - 1
                                        Set callback = tableRowHandlers(j)
                                        callback m_state, num, field, idTable, "body", rowNum, trAttrs, colAttrs2, row
                                    Next
                                    
                                    Response.Write "                    <div"
                                    For Each keyAttr In trAttrs.Keys
                                        Response.Write " " & keyAttr & "=""" & axFuncs.AxHtmlSpecialChars(trAttrs(keyAttr)) & """"
                                    Next
                                    Response.Write ">" & vbCrLf
                                    
                                    For iRowCol = 0 To UBound(row)
                                        Response.Write "                        <div"
                                        If headColAttrs.Exists(iRowCol) Then
                                            For Each keyAttr In headColAttrs(iRowCol).Keys
                                                Response.Write " " & keyAttr & "=""" & axFuncs.AxHtmlSpecialChars(headColAttrs(iRowCol)(keyAttr)) & """"
                                            Next
                                        End If
                                        Response.Write ">" & (IfThen(field.Exists("htmlcols") And field("htmlcols"), FFTranslate(IfThen(iRowCol <= UBound(field("cols")), field("cols")(iRowCol), "")), axFuncs.AxHtmlSpecialChars(FFTranslate(IfThen(iRowCol <= UBound(field("cols")), field("cols")(iRowCol), ""))))) & "</div>" & vbCrLf
                                        
                                        Response.Write "                        <div"
                                        If colAttrs2.Exists(iRowCol) Then
                                            For Each keyAttr In colAttrs2(iRowCol).Keys
                                                Response.Write " " & keyAttr & "=""" & axFuncs.AxHtmlSpecialChars(colAttrs2(iRowCol)(keyAttr)) & """"
                                            Next
                                        End If
                                        Response.Write ">" & row(iRowCol) & "</div>" & vbCrLf
                                    Next
                                    Response.Write "                    </div>" & vbCrLf
                                    rowNum = rowNum + 1
                                    altRow = Not altRow
                                Next
                                
                                If field.Exists("callback") Then
                                    rows = field("callback")(field)
                                    field("rows") = rows
                                Else
                                    field.Remove "rows"
                                End If
                            Loop
                            Response.Write "                </div>" & vbCrLf
                        End If
                        Response.Write "            </div>" & vbCrLf

                    Case "file"
                        Response.Write "            <div class=""formitemdata"">" & vbCrLf
                        Response.Write "                <div class=""textitemwrap""" & styleWidth & "><input class=""text"" type=""file"" id=""" & axFuncs.AxHtmlSpecialChars(id) & """ name=""" & axFuncs.AxHtmlSpecialChars(field("name")) & (IfThen(field.Exists("multiple") And field("multiple") = True, "[]", "")) & """" & (IfThen(field.Exists("multiple") And field("multiple") = True, " multiple", "")) & (IfThen(field.Exists("accept"), " accept=""" & axFuncs.AxHtmlSpecialChars(field("accept")) & """", "")) & (IfThen(m_state("autofocused") = id, " autofocus", "")) & " /></div>" & vbCrLf
                        Response.Write "            </div>" & vbCrLf

                    Case "custom"
                        Response.Write "            <div class=""formitemdata"">" & vbCrLf
                        Response.Write "                <div id=""" & axFuncs.AxHtmlSpecialChars(id) & """ class=""customitemwrap""" & styleWidth & ">" & vbCrLf
                        Response.Write field("value") & vbCrLf
                        Response.Write "                </div>" & vbCrLf
                        Response.Write "            </div>" & vbCrLf
                End Select

                If field.Exists("desc") Then
                    If field("desc") <> "" Then Response.Write "            <div class=""formitemdesc"">" & axFuncs.AxHtmlSpecialChars(FFTranslate(field("desc"))) & "</div>" & vbCrLf
                ElseIf field.Exists("htmldesc") Then
                    If field("htmldesc") <> "" Then Response.Write "            <div class=""formitemdesc"">" & field("htmldesc") & "</div>" & vbCrLf
                End If

                If field.Exists("error") Then
                    If field("error") <> "" Then
                        Response.Write "            <div class=""formitemresult"">" & vbCrLf
                        Response.Write "                <div class=""formitemerror"">" & axFuncs.AxHtmlSpecialChars(FFTranslate(field("error"))) & "</div>" & vbCrLf
                        Response.Write "            </div>" & vbCrLf
                    End If
                End If

                Response.Write "            </div>" & vbCrLf
                If m_state("insiderow") Then Response.Write "</td>" & vbCrLf
            End If
        End If
    End Sub

    Private Sub CleanupFields()
        If m_state("insiderow") Then
            If m_state("responsive") And m_state("insiderowwidth") Then Response.Write "<td></td>"
            Response.Write "            </tr></tbody></table></div>" & vbCrLf
        End If

        Dim callback, i
        Dim cleanupHandlers
        Set cleanupHandlers = g_ff_form_handlers("cleanup")
        For i = 0 To cleanupHandlers.Count - 1
            Set callback = cleanupHandlers(i)
            callback m_state
        Next
    End Sub

    Public Sub OutputFlexFormsJS(scriptTag)
        If scriptTag Then Response.Write "<script type=""text/javascript"">" & vbCrLf
%>
(function() {
	if (window.hasOwnProperty('FlexForms'))  return;

	// FlexForms base class.
	var FlexFormsInternal = function() {
		if (!(this instanceof FlexFormsInternal))  return new FlexFormsInternal();

		var triggers = {}, version = '', cssoutput = {}, cssleft = 0, jsqueue = {}, ready = false, initialized = false;
		var $this = this;

		// Internal functions.
		var DispatchEvent = function(eventname, params) {
			if (!triggers[eventname])  return;

			triggers[eventname].forEach(function(callback) {
				if (Array.isArray(params))  callback.apply($this, params);
				else  callback.call($this, params);
			});
		};

		// Public DOM-style functions.
		$this.addEventListener = function(eventname, callback) {
			if (!triggers[eventname])  triggers[eventname] = [];

			for (var x in triggers[eventname])
			{
				if (triggers[eventname][x] === callback)  return;
			}

			triggers[eventname].push(callback);
		};

		$this.removeEventListener = function(eventname, callback) {
			if (!triggers[eventname])  return;

			for (var x in triggers[eventname])
			{
				if (triggers[eventname][x] === callback)
				{
					triggers[eventname].splice(x, 1);

					return;
				}
			}
		};

		$this.hasEventListener = function(eventname) {
			return (triggers[eventname] && triggers[eventname].length);
		};

		$this.modules = {};

		$this.SetVersion = function(newver) {
			version = newver;
		};

		$this.GetVersion = function() {
			return version;
		};

		$this.RegisterCSSOutput = function(info) {
			Object.assign(cssoutput, info);
		};

		var CheckEmptyAndNotify = function() {
			if (cssleft)  return;

			for (var x in jsqueue)
			{
				if (jsqueue.hasOwnProperty(x))  return;
			}

			DispatchEvent('done');
		};

		$this.LoadCSS = function(name, url, cssmedia) {
			if (cssoutput[name] !== undefined)
			{
				CheckEmptyAndNotify();

				return cssoutput[name];
			}

			if (version !== '')  url += (url.indexOf('?') > -1 ? '&' : '?') + version;

			var tag = document.createElement('link');

			tag._loaded = false;
			tag.onload = function(e) {
				tag._loaded = true;

				cssleft--;
				CheckEmptyAndNotify();
			};

			cssleft++;

			tag.rel = 'stylesheet';
			tag.type = 'text/css';
			tag.href = url;
			tag.media = (cssmedia != undefined ? cssmedia : 'all');

			document.getElementsByTagName('head')[0].appendChild(tag);

			cssoutput[name] = tag;

			return tag;
		};

		$this.AddCSS = function(name, css, cssmedia) {
			if (cssoutput[name] !== undefined)
			{
				CheckEmptyAndNotify();

				return cssoutput[name];
			}

			var tag = document.createElement('style');
			tag.type = 'text/css';
			tag.media = (cssmedia != undefined ? cssmedia : 'all');

			document.getElementsByTagName('head')[0].appendChild(tag);

			if (tag.styleSheet)  tag.styleSheet.cssText = css;
			else  tag.appendChild(document.createTextNode(css));

			tag._loaded = true;

			cssoutput[name] = tag;

			CheckEmptyAndNotify();

			return tag;
		};

		$this.AddJSQueueItem = function(name, info) {
			jsqueue[name] = info;
		};

		var LoadJSQueueItem = function(name) {
			var done = false;
			var s = document.createElement('script');

			jsqueue[name].loading = true;
			jsqueue[name].retriesleft = jsqueue[name].retriesleft || 3;

			s.onload = function() {
				if (!done)  { done = true;  delete jsqueue[name];  $this.ProcessJSQueue(); }
			};

			s.onreadystatechange = function() {
				if (!done && s.readyState === 'complete')  { done = true;  delete jsqueue[name];  $this.ProcessJSQueue(); }
			};

			s.onerror = function() {
				if (!done)
				{
					done = true;

					jsqueue[name].retriesleft--;
					if (jsqueue[name].retriesleft > 0)
					{
						jsqueue[name].loading = false;

						setTimeout($this.ProcessJSQueue, 250);
					}
				}
			};

			s.src = jsqueue[name].src + (version === '' ? '' : (jsqueue[name].src.indexOf('?') > -1 ? '&' : '?') + version);

			document.body.appendChild(s);
		};

		$this.GetObjectFromPath = function(path) {
			var obj = window;
			path = path.split('.');
			for (var x = 0; x < path.length; x++)
			{
				if (obj[path[x]] === undefined)  return;

				obj = obj[path[x]];
			}

			return obj;
		};

		$this.ProcessJSQueue = function() {
			ready = true;

			for (var name in jsqueue) {
				if (jsqueue.hasOwnProperty(name))
				{
					if ((jsqueue[name].loading === undefined || jsqueue[name].loading === false) && (jsqueue[name].dependency === false || jsqueue[jsqueue[name].dependency] === undefined))
					{
						if (jsqueue[name].detect !== undefined && $this.GetObjectFromPath(jsqueue[name].detect) !== undefined)  delete jsqueue[name];
						else if (jsqueue[name].mode === "src")  LoadJSQueueItem(name);
						else if (jsqueue[name].mode === "inline")
						{
							jsqueue[name].src();

							delete jsqueue[name];
						}
					}
				}
			}

			CheckEmptyAndNotify();
		};

		$this.Init = function() {
			if (ready)  $this.ProcessJSQueue();
			else if (!initialized)
			{
				if (document.addEventListener)
				{
					function regevent(event) {
						document.removeEventListener("DOMContentLoaded", regevent);

						$this.ProcessJSQueue();
					}

					document.addEventListener("DOMContentLoaded", regevent);
				}
				else
				{
					setTimeout($this.ProcessJSQueue(), 250);
				}

				initialized = true;
			}
		};
	};

	window.FlexForms = new FlexFormsInternal();
})();

FlexForms.SetVersion('<%=JSSafe(m_version)%>');
<%
        If scriptTag Then Response.Write "</script>" & vbCrLf
    End Sub

    Private Sub FinalizeForm()
        Call OutputFlexFormsJS(Not m_state("ajax"))
        Call OutputJQuery(True)
        If m_state("jqueryuiused") Then Call OutputJQueryUI(True)
        
        Dim name, info, output, found, cssDict, jsDict, keys, i
        Set output = m_state("cssoutput")
        Set cssDict = m_state("css")
        
        keys = output.Keys
        For i = 0 To output.Count - 1
            name = keys(i)
            If cssDict.Exists(name) Then cssDict.Remove name
        Next
        
        If m_state("ajax") Then Response.Write "<script type=""text/javascript"">" & vbCrLf
        
        Do
            found = False
            keys = cssDict.Keys
            For i = 0 To cssDict.Count - 1
                name = keys(i)
                Set info = cssDict(name)
                If info("mode") = "link" And (info("dependency") = False Or output.Exists(info("dependency"))) Then
                    If m_state("ajax") Then
                        Response.Write "FlexForms.LoadCSS('" & JSSafe(name) & "', '" & JSSafe(info("src")) & "'" & (IfThen(info.Exists("media"), ", '" & JSSafe(info("media")) & "'", "")) & ");" & vbCrLf
                    Else
                        Dim verSuffix
                        verSuffix = ""
                        If m_version <> "" Then
                            If InStr(info("src"), "?") > 0 Then verSuffix = "&" & m_version Else verSuffix = "?" & m_version
                        End If
                        Response.Write "<link rel=""stylesheet"" href=""" & axFuncs.AxHtmlSpecialChars(info("src") & verSuffix) & """ type=""text/css"" media=""" & (IfThen(info.Exists("media"), info("media"), "all")) & """ />" & vbCrLf
                        output.Add name, True
                    End If
                    cssDict.Remove name
                    found = True
                    Exit For
                End If
            Next
        Loop While found
        
        If Not m_state("ajax") Then Response.Write "<script type=""text/javascript"">" & vbCrLf
        
        ' Convert output dictionary to JSON-like string for JS
        Response.Write "FlexForms.RegisterCSSOutput({"
        Dim first
        first = True
        keys = output.Keys
        For i = 0 To output.Count - 1
            name = keys(i)
            If Not first Then Response.Write ","
            Response.Write "'" & JSSafe(name) & "': true"
            first = False
        Next
        Response.Write "});" & vbCrLf
        Response.Write "</script>" & vbCrLf
        
        keys = cssDict.Keys
        For i = 0 To cssDict.Count - 1
            name = keys(i)
            Set info = cssDict(name)
            If info("mode") = "inline" Then Response.Write info("src")
        Next
        cssDict.RemoveAll
        
        ' Javascript processing
        Set output = m_state("jsoutput")
        Set jsDict = m_state("js")
        keys = output.Keys
        For i = 0 To output.Count - 1
            name = keys(i)
            If jsDict.Exists(name) Then jsDict.Remove name
        Next
        
        Do
            found = False
            keys = jsDict.Keys
            For i = 0 To jsDict.Count - 1
                name = keys(i)
                Set info = jsDict(name)
                If info("dependency") = False Or output.Exists(info("dependency")) Then
                    If info("mode") = "src" Then
                        If m_state("ajax") Then
                            Response.Write "FlexForms.AddJSQueueItem('" & JSSafe(name) & "', { mode: 'src', dependency: " & (IfThen(info("dependency") = False, "false", "'" & JSSafe(info("dependency")) & "'")) & ", src: '" & JSSafe(info("src")) & "', detect: '" & JSSafe(info("detect")) & "' });" & vbCrLf
                        Else
                            Dim verSuffix2
                            verSuffix2 = ""
                            If m_version <> "" Then
                                If InStr(info("src"), "?") > 0 Then verSuffix2 = "&" & m_version Else verSuffix2 = "?" & m_version
                            End If
                            Response.Write "<script type=""text/javascript"" src=""" & axFuncs.AxHtmlSpecialChars(info("src") & verSuffix2) & """></script>" & vbCrLf
                        End If
                    ElseIf info("mode") = "inline" Then
                        If Not m_state("ajax") Then
                            Response.Write "<script type=""text/javascript"">" & vbCrLf & info("src") & vbCrLf & "</script>" & vbCrLf
                        Else
                            Response.Write "FlexForms.AddJSQueueItem('" & JSSafe(name) & "', { mode: 'inline', dependency: " & (IfThen(info("dependency") = False, "false", "'" & JSSafe(info("dependency")) & "'")) & ", src: function() {" & vbCrLf & info("src") & vbCrLf & "} });" & vbCrLf
                        End If
                    End If
                    jsDict.Remove name
                    output.Add name, True
                    found = True
                    Exit For
                End If
            Next
        Loop While found
        jsDict.RemoveAll

        Dim finalizeHandlers
        Set finalizeHandlers = g_ff_form_handlers("finalize")
        Dim callback, k
        For k = 0 To finalizeHandlers.Count - 1
            Set callback = finalizeHandlers(k)
            callback m_state
        Next

        Response.Write "<script type=""text/javascript"">FlexForms.Init();</script>" & vbCrLf
    End Sub

    ' Helper for ternary-like operations
    Private Function IfThen(condition, trueVal, falseVal)
        If condition Then IfThen = trueVal Else IfThen = falseVal
    End Function

    ' --- Main Form Generation Engine ---
    Public Sub Generate(options, errors, lastform)
        Call InitFormVars(options)
        Call OutputFormCSS(False)

        m_state("formnum") = m_state("formnum") + 1

        Response.Write "<div class=""ff_formwrap"">" & vbCrLf
        Response.Write "<div class=""ff_formwrapinner"">" & vbCrLf

        Dim hasSubmit, useForm
        hasSubmit = options.Exists("submit")
        useForm = options.Exists("useform")

        If hasSubmit Or useForm Then
            m_state("formid") = m_state("formidbase") & m_state("formnum")
            
            Dim methodStr, hasFileField, fIdx, fieldsList
            hasFileField = False
            
            If options.Exists("fields") Then
                fieldsList = options("fields")
                For fIdx = 0 To UBound(fieldsList)
                    If IsObject(fieldsList(fIdx)) Then
                        If fieldsList(fIdx).Exists("type") Then
                            If fieldsList(fIdx)("type") = "file" Then
                                hasFileField = True
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If

            methodStr = "method=""post"""
            If options.Exists("formmode") Then
                If options("formmode") = "get" Then methodStr = "method=""get"""
            End If

            If methodStr = "method=""post""" And hasFileField Then
                methodStr = methodStr & " enctype=""multipart/form-data"""
            End If

            Response.Write "<form class=""ff_form"" id=""" & m_state("formid") & """ " & methodStr & " action=""" & axFuncs.AxHtmlSpecialChars(m_state("action")) & """>" & vbCrLf

            ' Nonce and Hidden fields logic
            Dim extra, name, value, hiddenDict
            Set extra = CreateObject("Scripting.Dictionary")
            
            If Not options.Exists("nonce") Then
                If IsObject(m_autoNonce) Then
                    If Not m_autoNonce Is Nothing Then
                        options.Add "nonce", m_autoNonce("action")
                        If Not options.Exists("hidden") Then options.Add "hidden", CreateObject("Scripting.Dictionary")
                        If Not options("hidden").Exists(options("nonce")) Then options("hidden").Add options("nonce"), m_autoNonce("value")
                    End If
                End If
            End If

            If options.Exists("hidden") Then
                Set hiddenDict = options("hidden")
                For Each name In hiddenDict.Keys
                    value = hiddenDict(name)
                    m_state("hidden").Add CStr(name), CStr(value)
                    Response.Write "        <input type=""hidden"" name=""" & axFuncs.AxHtmlSpecialChars(name) & """ value=""" & axFuncs.AxHtmlSpecialChars(value) & """ />" & vbCrLf
                    If options.Exists("nonce") Then
                        If options("nonce") <> name Then extra.Add name, value
                    End If
                Next

                If options.Exists("nonce") Then
                    Dim secExtra, secT
                    secExtra = Join(extra.Keys, ",")
                    secT = CreateSecurityToken(hiddenDict(options("nonce")), extra.Items)
                    m_state("hidden").Add "sec_extra", secExtra
                    m_state("hidden").Add "sec_t", secT
                    Response.Write "        <input type=""hidden"" name=""sec_extra"" value=""" & axFuncs.AxHtmlSpecialChars(secExtra) & """ />" & vbCrLf
                    Response.Write "        <input type=""hidden"" name=""sec_t"" value=""" & axFuncs.AxHtmlSpecialChars(secT) & """ />" & vbCrLf
                End If
            End If
        End If

        if options.Exists("fields") Then
            Dim fieldsList2, num, field, fieldId, altClass, responsiveClass
            fieldsList2 = options("fields")
            altClass = ""
            If IsArray(fieldsList2) Then
                If UBound(fieldsList2) = 0 Then
                    If IsObject(fieldsList2(0)) Then
                        If Not fieldsList2(0).Exists("title") And Not fieldsList2(0).Exists("htmltitle") Then altClass = " alt"
                    End If
                End If
            End If
            responsiveClass = ""
            If m_state("responsive") Then responsiveClass = " formfieldsresponsive"
            
            Response.Write "        <div class=""formfields" & altClass & responsiveClass & """>" & vbCrLf
            Response.Write "            <div class=""formfieldsinner"">" & vbCrLf

            If IsArray(fieldsList2) Then
                For num = 0 To UBound(fieldsList2)
                    If Not IsEmpty(fieldsList2(num)) Then
                        If IsObject(fieldsList2(num)) Then
                            Set field = fieldsList2(num)
                        Else
                            field = fieldsList2(num)
                        End If

                        fieldId = "f" & m_state("formnum") & "_" & num
                        
                        If IsObject(field) Then
                            If field.Exists("name") Then
                                If errors.Exists(field("name")) Then field("error") = errors(field("name"))
                                
                                If options.Exists("hashnames") Then
                                    If options("hashnames") Then
                                        field.Add "origname", field("name")
                                        field("name") = GetHashedFieldName(field("name"))
                                    End If
                                End If
                                fieldId = fieldId & "_" & field("name")
                            End If
                        End If
                        
                        Call ProcessField(num, field, fieldId)
                    End If
                Next
            End If

            Call CleanupFields()
            Response.Write "            </div>" & vbCrLf
            Response.Write "        </div>" & vbCrLf
        End If

        If hasSubmit Then Call ProcessSubmit(options)

        If hasSubmit Or useForm Then Response.Write "</form>" & vbCrLf

        Response.Write "</div>" & vbCrLf
        Response.Write "</div>" & vbCrLf

        If lastform Then Call FinalizeForm()
    End Sub
End Class
%>
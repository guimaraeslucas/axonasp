<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Request.ServerVariables Test</title>
    <style>
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; background:#f5f5f5; padding:20px; }
        h1 { margin-bottom:10px; }
        .intro { background:#e3f2fd; border-left:4px solid #2196f3; padding:12px; margin-bottom:16px; }
        table { width:100%; border-collapse:collapse; background:#fff; }
        th, td { border:1px solid #ddd; padding:8px; font-size:13px; }
        th { background:#667eea; color:#fff; text-align:left; }
        tr:nth-child(even) { background:#f9f9f9; }
        code { font-family: Consolas, monospace; }
        .section { margin-top:20px; margin-bottom:8px; font-weight:bold; color:#444; }
    </style>
</head>
<body>
<%
Option Explicit

Dim names, i, key, val
names = Array( _
    "ALL_HTTP", "ALL_RAW", _
    "APPL_MD_PATH", "APPL_PHYSICAL_PATH", _
    "AUTH_PASSWORD", "AUTH_TYPE", "AUTH_USER", _
    "CERT_COOKIE", "CERT_FLAGS", "CERT_ISSUER", "CERT_KEYSIZE", "CERT_SECRETKEYSIZE", "CERT_SERIALNUMBER", "CERT_SERVER_ISSUER", "CERT_SERVER_SUBJECT", "CERT_SUBJECT", _
    "CONTENT_LENGTH", "CONTENT_TYPE", _
    "GATEWAY_INTERFACE", _
    "HTTP_ACCEPT", "HTTP_ACCEPT_LANGUAGE", "HTTP_ACCEPT_ENCODING", "HTTP_COOKIE", "HTTP_REFERER", "HTTP_USER_AGENT", "HTTP_HOST", "HTTP_CONNECTION", "HTTP_CACHE_CONTROL", "HTTP_PRAGMA", "HTTP_UPGRADE_INSECURE_REQUESTS", "HTTP_X_FORWARDED_FOR", _
    "HTTPS", "HTTPS_KEYSIZE", "HTTPS_SECRETKEYSIZE", "HTTPS_SERVER_ISSUER", "HTTPS_SERVER_SUBJECT", _
    "INSTANCE_ID", "INSTANCE_META_PATH", _
    "LOCAL_ADDR", _
    "LOGON_USER", _
    "PATH_INFO", "PATH_TRANSLATED", _
    "QUERY_STRING", _
    "REMOTE_ADDR", "REMOTE_HOST", "REMOTE_USER", _
    "REQUEST_METHOD", _
    "SCRIPT_NAME", _
    "SERVER_NAME", "SERVER_PORT", "SERVER_PORT_SECURE", "SERVER_PROTOCOL", "SERVER_SOFTWARE", _
    "URL" _
)

Response.Write "<h1>Request.ServerVariables Test</h1>"
Response.Write "<div class='intro'>Checks classic ASP-compatible server variables and shows current request values. Also lists all keys available in the collection.</div>"

Response.Write "<div class='section'>Expected Variables</div>"
Response.Write "<table><tr><th>Name</th><th>Value</th></tr>"
For i = 0 To UBound(names)
    val = Request.ServerVariables(names(i))
    Response.Write "<tr><td><code>" & names(i) & "</code></td><td>" & Server.HTMLEncode(CStr(val)) & "</td></tr>"
Next
Response.Write "</table>"

Response.Write "<div class='section'>All Available Keys</div>"
Response.Write "<table><tr><th>Name</th><th>Value</th></tr>"
For Each key In Request.ServerVariables
    val = Request.ServerVariables(key)
    Response.Write "<tr><td><code>" & key & "</code></td><td>" & Server.HTMLEncode(CStr(val)) & "</td></tr>"
Next
Response.Write "</table>"
%>
</body>
</html>

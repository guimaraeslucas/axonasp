$ErrorActionPreference = "Stop"

Write-Host "--- Test 1: JScript Cookie ---"
$session1 = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$res1 = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/jscript_cookie.asp?action=set" -WebSession $session1
Write-Host "Set Response:" $res1.Content
Write-Host "Cookies in session:" $session1.Cookies.GetCookies( (New-Object System.Uri("http://localhost:8801")) )
$res1_get = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/jscript_cookie.asp?action=get" -WebSession $session1
Write-Host "Get Response:" $res1_get.Content
Write-Host ""

Write-Host "--- Test 2: VBScript Cookie ---"
$session2 = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$res2 = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/vbscript_cookie.asp?action=set" -WebSession $session2
Write-Host "Set Response:" $res2.Content
Write-Host "Cookies in session:" $session2.Cookies.GetCookies( (New-Object System.Uri("http://localhost:8801")) )
$res2_get = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/vbscript_cookie.asp?action=get" -WebSession $session2
Write-Host "Get Response:" $res2_get.Content
Write-Host ""

Write-Host "--- Test 3: Classic ASP Compliance Cookie ---"
$session3 = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$res3 = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/compliance_cookie.asp?action=set" -WebSession $session3
Write-Host "Set Response:" $res3.Content
Write-Host "Cookies in session:" $session3.Cookies.GetCookies( (New-Object System.Uri("http://localhost:8801")) )
$res3_get = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/compliance_cookie.asp?action=get" -WebSession $session3
Write-Host "Get Response:" $res3_get.Content
Write-Host ""

Write-Host "--- Test 4: Session State Integrity ---"
$session4 = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$res4 = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/session_test.asp?action=set" -WebSession $session4
Write-Host "Set Response:" $res4.Content
Write-Host "Cookies in session (ASPSESSIONID):" $session4.Cookies.GetCookies( (New-Object System.Uri("http://localhost:8801")) )
$res4_get = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/session_test.asp?action=get" -WebSession $session4
Write-Host "Get Response:" $res4_get.Content
Write-Host ""

Write-Host "--- Test 5: E2E Mimic ---"
$session5 = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$res5 = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/e2e_mimic.asp?lang=pt-BR" -WebSession $session5
Write-Host "Set Response (lang=pt-BR):" $res5.Content
Write-Host "Cookies in session:" $session5.Cookies.GetCookies( (New-Object System.Uri("http://localhost:8801")) )
$res5_get = Invoke-WebRequest -Uri "http://localhost:8801/tests/cookie_bug/e2e_mimic.asp" -WebSession $session5
Write-Host "Get Response (no lang param):" $res5_get.Content
Write-Host "Cookies after Get:" $session5.Cookies.GetCookies( (New-Object System.Uri("http://localhost:8801")) )

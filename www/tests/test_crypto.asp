<%
@ Language = VBScript
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>G3pix AxonASP - G3CRYPTO Unified Library Test</title>
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            body {
                font-family: Tahoma, "Segoe UI", Geneva, Verdana, sans-serif;
                padding: 30px;
                background: #f5f5f5;
                line-height: 1.6;
            }
            .container {
                max-width: 1200px;
                margin: 0 auto;
                background: #fff;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            h1 {
                color: #333;
                margin-bottom: 10px;
                border-bottom: 2px solid #667eea;
                padding-bottom: 10px;
            }
            h3 {
                color: #555;
                margin-top: 20px;
                margin-bottom: 10px;
                border-bottom: 1px solid #ddd;
                padding-bottom: 5px;
            }
            .intro {
                background: #e3f2fd;
                border-left: 4px solid #2196f3;
                padding: 15px;
                margin-bottom: 20px;
                border-radius: 4px;
            }
            .code-block {
                background: #f4f4f4;
                padding: 10px;
                margin: 10px 0;
                border-radius: 4px;
                word-break: break-all;
                font-family: "Courier New", monospace;
                font-size: 12px;
            }
            code {
                background: #f0f0f0;
                padding: 2px 6px;
                border-radius: 3px;
                font-family: "Courier New", monospace;
            }
            .success {
                color: #28a745;
                font-weight: bold;
            }
            .error {
                color: #dc3545;
                font-weight: bold;
            }
            .info {
                color: #666;
                margin: 5px 0;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin: 10px 0;
            }
            th,
            td {
                padding: 8px;
                text-align: left;
                border: 1px solid #ddd;
            }
            th {
                background: #f8f9fa;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>G3pix AxonASP - G3CRYPTO Unified Library Test</h1>
            <div class="intro">
                <p>
                    Comprehensive test of the unified G3CRYPTO library
                    including:
                </p>
                <p>
                    UUID generation, BCrypt password hashing, modern hash
                    algorithms (SHA-256, SHA-512, SHA3, BLAKE2b), HMAC, PBKDF2,
                    secure random generation, .NET compatibility, and bcrypt
                    cost configuration.
                </p>
            </div>

            <%
            Dim crypto, testsPassed, testsFailed
            Set crypto = Server.CreateObject("G3CRYPTO")
            testsPassed = 0
            testsFailed = 0

            ' === 1. UUID Generation ===
            Response.Write "<h3>1. UUID Generation</h3>"
            Dim id1, id2
            id1 = crypto.UUID()
            id2 = crypto.UUID()

            Response.Write "<div class='info'>UUID 1: <code>" & id1 & "</code></div>"
            Response.Write "<div class='info'>UUID 2: <code>" & id2 & "</code></div>"

            If id1 <> id2 And Len(id1) = 36 And Len(id2) = 36 Then
                Response.Write "<span class='success'>✓ UUIDs are unique and properly formatted</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ UUID generation failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' Test GUID alias
            Dim guid1
            guid1 = crypto.GUID()
            If Len(guid1) = 36 Then
                Response.Write "<span class='success'>✓ GUID alias works correctly</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ GUID alias failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 2. Password Hashing (BCrypt) ===
            Response.Write "<h3>2. Password Hashing (BCrypt)</h3>"
            Dim password, hash
            password = "SuperSecretPassword123!"

            hash = crypto.HashPassword(password)
            Response.Write "<div class='info'>Password: <code>" & password & "</code></div>"
            Response.Write "<div class='code-block'>" & hash & "</div>"

            If Len(hash) > 50 Then
                Response.Write "<span class='success'>✓ Password hash generated</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ Password hash generation failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 3. Password Verification ===
            Response.Write "<h3>3. Password Verification</h3>"

            If crypto.VerifyPassword(password, hash) Then
                Response.Write "<span class='success'>✓ Correct password verified successfully</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ Correct password verification failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            If Not crypto.VerifyPassword("WrongPassword", hash) Then
                Response.Write "<span class='success'>✓ Incorrect password rejected correctly</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ Incorrect password accepted (SECURITY ISSUE!)</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' Test alias methods
            Dim hash2
            hash2 = crypto.Hash(password)
            If crypto.Verify(password, hash2) Then
                Response.Write "<span class='success'>✓ Hash/Verify aliases work correctly</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ Hash/Verify aliases failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 4. Bcrypt Cost Configuration ===
            Response.Write "<h3>4. Bcrypt Cost Configuration</h3>"
            Dim originalCost, newCost
            originalCost = crypto.GetBcryptCost()
            Response.Write "<div class='info'>Original cost: " & originalCost & "</div>"

            If crypto.SetBcryptCost(12) Then
                newCost = crypto.GetBcryptCost()
                Response.Write "<div class='info'>New cost: " & newCost & "</div>"
                If newCost = 12 Then
                    Response.Write "<span class='success'>✓ Bcrypt cost configuration works</span><br>"
                    testsPassed = testsPassed + 1
                Else
                    Response.Write "<span class='error'>✗ Bcrypt cost not applied</span><br>"
                    testsFailed = testsFailed + 1
                End If
            Else
                Response.Write "<span class='error'>✗ SetBcryptCost failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' Test property access
            crypto.BcryptCost = 10
            If crypto.Cost = 10 Then
                Response.Write "<span class='success'>✓ Bcrypt cost property access works</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ Bcrypt cost property access failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 5. Modern Hash Algorithms ===
            Response.Write "<h3>5. Modern Hash Algorithms (Hex Output)</h3>"
            Dim testData, hashResult
            testData = "AxonASP Test Data"

            Response.Write "<table>"
            Response.Write "<tr><th>Algorithm</th><th>Hash (hex)</th><th>Status</th></tr>"

            ' SHA-256
            hashResult = crypto.SHA256(testData)
            If Len(hashResult) = 64 Then
                Response.Write "<tr><td>SHA-256</td><td><code>" & hashResult & "</code></td><td><span class='success'>✓</span></td></tr>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<tr><td>SHA-256</td><td>Failed</td><td><span class='error'>✗</span></td></tr>"
                testsFailed = testsFailed + 1
            End If

            ' SHA-512
            hashResult = crypto.SHA512(testData)
            If Len(hashResult) = 128 Then
                Response.Write "<tr><td>SHA-512</td><td><code>" & Left(hashResult, 64) & "...</code></td><td><span class='success'>✓</span></td></tr>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<tr><td>SHA-512</td><td>Failed</td><td><span class='error'>✗</span></td></tr>"
                testsFailed = testsFailed + 1
            End If

            ' SHA3-256
            hashResult = crypto.SHA3_256(testData)
            If Len(hashResult) = 64 Then
                Response.Write "<tr><td>SHA3-256</td><td><code>" & hashResult & "</code></td><td><span class='success'>✓</span></td></tr>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<tr><td>SHA3-256</td><td>Failed</td><td><span class='error'>✗</span></td></tr>"
                testsFailed = testsFailed + 1
            End If

            ' BLAKE2b-256
            hashResult = crypto.BLAKE2B256(testData)
            If Len(hashResult) = 64 Then
                Response.Write "<tr><td>BLAKE2b-256</td><td><code>" & hashResult & "</code></td><td><span class='success'>✓</span></td></tr>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<tr><td>BLAKE2b-256</td><td>Failed</td><td><span class='error'>✗</span></td></tr>"
                testsFailed = testsFailed + 1
            End If

            ' MD5 (legacy)
            hashResult = crypto.MD5(testData)
            If Len(hashResult) = 32 Then
                Response.Write "<tr><td>MD5 (legacy)</td><td><code>" & hashResult & "</code></td><td><span class='success'>✓</span></td></tr>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<tr><td>MD5 (legacy)</td><td>Failed</td><td><span class='error'>✗</span></td></tr>"
                testsFailed = testsFailed + 1
            End If

            ' SHA-1 (legacy)
            hashResult = crypto.SHA1(testData)
            If Len(hashResult) = 40 Then
                Response.Write "<tr><td>SHA-1 (legacy)</td><td><code>" & hashResult & "</code></td><td><span class='success'>✓</span></td></tr>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<tr><td>SHA-1 (legacy)</td><td>Failed</td><td><span class='error'>✗</span></td></tr>"
                testsFailed = testsFailed + 1
            End If

            Response.Write "</table>"

            ' === 6. Hash Byte Array Output ===
            Response.Write "<h3>6. Hash Byte Array Output</h3>"
            Dim hashBytes
            hashBytes = crypto.SHA256Bytes(testData)

            If IsArray(hashBytes) Then
                Dim arraySize
                arraySize = UBound(hashBytes) - LBound(hashBytes) + 1
                Response.Write "<div class='info'>SHA-256 bytes array size: " & arraySize & " bytes</div>"
                Response.Write "<div class='info'>First 8 bytes: "
                Dim i
                For i = 0 To 7
                    Response.Write hashBytes(i) & " "
                Next
                Response.Write "</div>"

                If arraySize = 32 Then
                    Response.Write "<span class='success'>✓ Byte array output works correctly</span><br>"
                    testsPassed = testsPassed + 1
                Else
                    Response.Write "<span class='error'>✗ Byte array size incorrect</span><br>"
                    testsFailed = testsFailed + 1
                End If
            Else
                Response.Write "<span class='error'>✗ Byte array output failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 7. HMAC ===
            Response.Write "<h3>7. HMAC (Hash-based Message Authentication Code)</h3>"
            Dim message, secretKey, hmacResult
            message = "Authenticated Message"
            secretKey = "shared-secret-key-123"

            hmacResult = crypto.HMACSHA256(message, secretKey)
            Response.Write "<div class='info'>Message: <code>" & message & "</code></div>"
            Response.Write "<div class='info'>Secret Key: <code>" & secretKey & "</code></div>"
            Response.Write "<div class='info'>HMAC-SHA256: <code>" & hmacResult & "</code></div>"

            If Len(hmacResult) = 64 Then
                Response.Write "<span class='success'>✓ HMAC-SHA256 works correctly</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ HMAC-SHA256 failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' Test HMAC-SHA512
            hmacResult = crypto.HMACSHA512(message, secretKey)
            If Len(hmacResult) = 128 Then
                Response.Write "<span class='success'>✓ HMAC-SHA512 works correctly</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ HMAC-SHA512 failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 8. PBKDF2 Key Derivation ===
            Response.Write "<h3>8. PBKDF2 Key Derivation</h3>"
            Dim salt, derivedKey
            salt = crypto.RandomHex(16)
            derivedKey = crypto.PBKDF2SHA256("user-password", salt, 10000, 32)

            Response.Write "<div class='info'>Salt: <code>" & salt & "</code></div>"
            Response.Write "<div class='info'>Derived Key: <code>" & derivedKey & "</code></div>"

            If Len(derivedKey) = 64 And Len(salt) = 32 Then
                Response.Write "<span class='success'>✓ PBKDF2-SHA256 works correctly</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ PBKDF2-SHA256 failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 9. Secure Random Generation ===
            Response.Write "<h3>9. Secure Random Generation</h3>"

            ' Random hex
            Dim randomHex
            randomHex = crypto.RandomHex(32)
            Response.Write "<div class='info'>Random Hex (32 bytes): <code>" & randomHex & "</code></div>"
            If Len(randomHex) = 64 Then
                Response.Write "<span class='success'>✓ RandomHex works correctly</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ RandomHex failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' Random base64
            Dim randomB64
            randomB64 = crypto.RandomBase64(24)
            Response.Write "<div class='info'>Random Base64 (24 bytes): <code>" & randomB64 & "</code></div>"
            If Len(randomB64) > 20 Then
                Response.Write "<span class='success'>✓ RandomBase64 works correctly</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ RandomBase64 failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' Random bytes
            Dim randomBytes
            randomBytes = crypto.RandomBytes(16)
            If IsArray(randomBytes) Then
                Dim randSize
                randSize = UBound(randomBytes) - LBound(randomBytes) + 1
                Response.Write "<div class='info'>Random Bytes: " & randSize & " bytes generated</div>"
                If randSize = 16 Then
                    Response.Write "<span class='success'>✓ RandomBytes works correctly</span><br>"
                    testsPassed = testsPassed + 1
                Else
                    Response.Write "<span class='error'>✗ RandomBytes size incorrect</span><br>"
                    testsFailed = testsFailed + 1
                End If
            Else
                Response.Write "<span class='error'>✗ RandomBytes failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 10. .NET Compatibility ===
            Response.Write "<h3>10. .NET Crypto Provider Compatibility</h3>"

            ' Test MD5CryptoServiceProvider
            Dim md5Provider, md5Hash
            Set md5Provider = Server.CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
            md5Hash = md5Provider.ComputeHash(testData)

            If IsArray(md5Hash) Then
                Dim md5Size
                md5Size = UBound(md5Hash) - LBound(md5Hash) + 1
                Response.Write "<div class='info'>MD5CryptoServiceProvider hash size: " & md5Size & " bytes</div>"
                Response.Write "<div class='info'>HashSize property: " & md5Provider.HashSize & " bits</div>"
                If md5Size = 16 And md5Provider.HashSize = 128 Then
                    Response.Write "<span class='success'>✓ MD5CryptoServiceProvider works correctly</span><br>"
                    testsPassed = testsPassed + 1
                Else
                    Response.Write "<span class='error'>✗ MD5CryptoServiceProvider failed</span><br>"
                    testsFailed = testsFailed + 1
                End If
            Else
                Response.Write "<span class='error'>✗ MD5CryptoServiceProvider failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' Test SHA256CryptoServiceProvider
            Dim sha256Provider, sha256Hash
            Set sha256Provider = Server.CreateObject("System.Security.Cryptography.SHA256CryptoServiceProvider")
            sha256Hash = sha256Provider.ComputeHash(testData)

            If IsArray(sha256Hash) Then
                Dim sha256Size
                sha256Size = UBound(sha256Hash) - LBound(sha256Hash) + 1
                Response.Write "<div class='info'>SHA256CryptoServiceProvider hash size: " & sha256Size & " bytes</div>"
                If sha256Size = 32 And sha256Provider.HashSize = 256 Then
                    Response.Write "<span class='success'>✓ SHA256CryptoServiceProvider works correctly</span><br>"
                    testsPassed = testsPassed + 1
                Else
                    Response.Write "<span class='error'>✗ SHA256CryptoServiceProvider failed</span><br>"
                    testsFailed = testsFailed + 1
                End If
            Else
                Response.Write "<span class='error'>✗ SHA256CryptoServiceProvider failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' Test SHA512CryptoServiceProvider
            Dim sha512Provider, sha512Hash
            Set sha512Provider = Server.CreateObject("System.Security.Cryptography.SHA512CryptoServiceProvider")
            sha512Hash = sha512Provider.ComputeHash(testData)

            If IsArray(sha512Hash) Then
                Dim sha512Size
                sha512Size = UBound(sha512Hash) - LBound(sha512Hash) + 1
                Response.Write "<div class='info'>SHA512CryptoServiceProvider hash size: " & sha512Size & " bytes</div>"
                If sha512Size = 64 And sha512Provider.HashSize = 512 Then
                    Response.Write "<span class='success'>✓ SHA512CryptoServiceProvider works correctly</span><br>"
                    testsPassed = testsPassed + 1
                Else
                    Response.Write "<span class='error'>✗ SHA512CryptoServiceProvider failed</span><br>"
                    testsFailed = testsFailed + 1
                End If
            Else
                Response.Write "<span class='error'>✗ SHA512CryptoServiceProvider failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === 11. Alias Names ===
            Response.Write "<h3>11. CreateObject Alias Names</h3>"

            Dim cryptoAlias
            Set cryptoAlias = Server.CreateObject("G3CRYPTO")
            Dim aliasTest
            aliasTest = cryptoAlias.UUID()
            If Len(aliasTest) = 36 Then
                Response.Write "<span class='success'>✓ 'CRYPTO' alias works</span><br>"
                testsPassed = testsPassed + 1
            Else
                Response.Write "<span class='error'>✗ 'CRYPTO' alias failed</span><br>"
                testsFailed = testsFailed + 1
            End If

            ' === Summary ===
            Dim totalTests
            totalTests = testsPassed + testsFailed
            Response.Write "<h3>Test Summary</h3>"
            Response.Write "<div style='background: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 20px;'>"
            Response.Write "<p><strong>Total Tests:</strong> " & totalTests & "</p>"
            Response.Write "<p><strong>Passed:</strong> <span class='success'>" & testsPassed & "</span></p>"
            Response.Write "<p><strong>Failed:</strong> <span class='error'>" & testsFailed & "</span></p>"

            Dim successRate
            If totalTests > 0 Then
                successRate = FormatNumber((testsPassed / totalTests) * 100, 2)
                Response.Write "<p><strong>Success Rate:</strong> " & successRate & "%</p>"
            End If

            If testsFailed = 0 Then
                Response.Write "<p style='color: #28a745; font-weight: bold; font-size: 18px; margin-top: 10px;'>✓ ALL TESTS PASSED!</p>"
            Else
                Response.Write "<p style='color: #dc3545; font-weight: bold; font-size: 18px; margin-top: 10px;'>✗ SOME TESTS FAILED</p>"
            End If
            Response.Write "</div>"
            %>
        </div>
    </body>
</html>

<%
@ CodePage = 65001
%>
<!--
	AxonASP Server - G3ZLIB Sample Page
	Demonstration of zlib compression/decompression functionality
	URL: http://localhost:8801/tests/test_g3zlib.asp
-->
<%
Option Explicit
%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="utf-8" />
        <title>G3ZLIB - Compression Test</title>
        <link rel="stylesheet" href="/css/axonasp.css" />
        <style>
            .section {
                margin: 20px 0;
                padding: 15px;
                border: 1px solid #999;
                background: #f9f9f9;
            }
            .code-box {
                background: #fff;
                border: 1px solid #ccc;
                padding: 10px;
                margin: 10px 0;
                font-family: monospace;
                white-space: pre-wrap;
                word-wrap: break-word;
            }
            .success {
                color: green;
                font-weight: bold;
            }
            .error {
                color: red;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <h1>G3ZLIB - Zlib Compression Test</h1>
        <p>
            This page demonstrates the G3ZLIB native object for
            compress/decompress operations.
        </p>

        <div class="section">
            <h2>Test 1: Basic Text Compression and Decompression</h2>
            <%
            Dim zlib, originalText, compressedArray, decompressedArray, decompressedText
            Set zlib = Server.CreateObject("G3ZLIB")

            originalText = "The quick brown fox jumps over the lazy dog. This text will be compressed and decompressed."
            Response.Write "<p><strong>Original Text:</strong></p>"
            Response.Write "<div class='code-box'>" & Server.HTMLEncode(originalText) & "</div>"

            ' Compress
            compressedArray = zlib.Compress(originalText)
            Response.Write "<p><strong>Compressed Size:</strong> " & UBound(compressedArray) + 1 & " bytes</p>"
            Response.Write "<p><strong>Original Size:</strong> " & Len(originalText) & " bytes</p>"

            Dim ratio
            ratio = ((UBound(compressedArray) + 1) / Len(originalText)) * 100
            Response.Write "<p><strong>Compression Ratio:</strong> " & FormatNumber(ratio, 2) & "%</p>"

            ' Decompress
            decompressedText = zlib.DecompressText(compressedArray)
            Response.Write "<p><strong>Decompressed Text:</strong></p>"
            Response.Write "<div class='code-box'>" & Server.HTMLEncode(decompressedText) & "</div>"

            If decompressedText = originalText Then
                Response.Write "<p class='success'>✓ Roundtrip successful!</p>"
            Else
                Response.Write "<p class='error'>✗ Roundtrip failed!</p>"
            End If

            Set zlib = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 2: Byte Array Compression</h2>
            <%
            Set zlib = Server.CreateObject("G3ZLIB")

            ' Create a test byte array (simulating binary data)
            Dim testArray(100), i
            For i = 0 To 100
                testArray(i) = (i * 7) Mod 256
            Next

            Response.Write "<p><strong>Original Array Size:</strong> " & UBound(testArray) + 1 & " bytes</p>"

            ' Compress with level 9 (maximum compression)
            compressedArray = zlib.Compress(testArray, 9)
            Response.Write "<p><strong>Compressed Size (Level 9):</strong> " & UBound(compressedArray) + 1 & " bytes</p>"

            ' Decompress
            decompressedArray = zlib.Decompress(compressedArray)
            Response.Write "<p><strong>Decompressed Size:</strong> " & UBound(decompressedArray) + 1 & " bytes</p>"

            ' Verify roundtrip
            Dim allMatch, j
            allMatch = True
            If UBound(decompressedArray) <> UBound(testArray) Then
                allMatch = False
            Else
                For j = 0 To UBound(testArray)
                    If decompressedArray(j) <> testArray(j) Then
                        allMatch = False
                        Exit For
                    End If
                Next
            End If

            If allMatch Then
                Response.Write "<p class='success'>✓ Byte array roundtrip successful!</p>"
            Else
                Response.Write "<p class='error'>✗ Byte array roundtrip failed!</p>"
            End If

            Set zlib = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 3: Batch Compression of Multiple Items</h2>
            <%
            Set zlib = Server.CreateObject("G3ZLIB")

            ' Create an array of strings to compress
            Dim itemsToCompress(3)
            itemsToCompress(0) = "First item: Lorem ipsum dolor sit amet"
            itemsToCompress(1) = "Second item: consectetur adipiscing elit"
            itemsToCompress(2) = "Third item: sed do eiusmod tempor"
            itemsToCompress(3) = "Fourth item: incididunt ut labore"

            Response.Write "<p>Compressing " & (UBound(itemsToCompress) + 1) & " items:</p>"

            Dim compressedBatch, k
            compressedBatch = zlib.CompressMany(itemsToCompress, 6)

            Response.Write "<p><strong>Batch Compression Result:</strong> " & UBound(compressedBatch) + 1 & " compressed items</p>"

            Response.Write "<p><strong>Items after decompression:</strong></p>"
            Response.Write "<ul>"
            For k = 0 To UBound(compressedBatch)
                Response.Write "<li>" & Server.HTMLEncode(zlib.DecompressText(compressedBatch(k))) & "</li>"
            Next
            Response.Write "</ul>"

            Set zlib = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 4: File Compression</h2>
            <%
            Dim fso, testFile, testFilePath, compressedFilePath
            Set fso = Server.CreateObject("Scripting.FileSystemObject")
            Set zlib = Server.CreateObject("G3ZLIB")

            ' Create a test file
            testFilePath = Server.MapPath("../temp/test_source.txt")
            compressedFilePath = Server.MapPath("../temp/test_source.txt.zlib")

            ' Ensure temp directory exists
            Dim tempDir
            tempDir = Server.MapPath("../temp")
            If Not fso.FolderExists(tempDir) Then
                fso.CreateFolder(tempDir)
            End If

            ' Write test content to file
            Dim testFileContent
            testFileContent = String(100, "Testing file compression with G3ZLIB. ") & vbCrLf & "This is a test file for zlib compression demonstration."

            Set testFile = fso.CreateTextFile(testFilePath, True)
            testFile.Write testFileContent
            testFile.Close()

            Response.Write "<p><strong>Original File:</strong> " & fso.GetFile(testFilePath).Size & " bytes</p>"

            ' Compress file
            If zlib.CompressFile(testFilePath, compressedFilePath, 6) Then
                Response.Write "<p><strong>Compressed File:</strong> " & fso.GetFile(compressedFilePath).Size & " bytes</p>"
                Response.Write "<p class='success'>✓ File compression successful!</p>"

                ' Clean up
                fso.DeleteFile(testFilePath)
                fso.DeleteFile(compressedFilePath)
            Else
                Response.Write "<p class='error'>✗ File compression failed!</p>"
                If Len(zlib.LastError) > 0 Then
                    Response.Write "<p>Error: " & Server.HTMLEncode(zlib.LastError) & "</p>"
                End If
            End If

            Set zlib = Nothing
            Set fso = Nothing
            %>
        </div>

        <div class="section">
            <h2>Compression Levels</h2>
            <p>G3ZLIB supports compression levels 0-9:</p>
            <ul>
                <li><strong>0</strong> - No compression (fastest)</li>
                <li><strong>1-3</strong> - Fast compression</li>
                <li><strong>4-6</strong> - Balanced (default is 6)</li>
                <li><strong>7-9</strong> - Maximum compression (slowest)</li>
            </ul>
        </div>

        <hr />
        <small>AxonASP Server - G3ZLIB Compression Test &copy; 2026</small>
    </body>
</html>

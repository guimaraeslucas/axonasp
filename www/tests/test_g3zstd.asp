<%
@ CodePage = 65001
%>
<!--
	AxonASP Server - G3ZSTD Sample Page
	Demonstration of Zstandard (zstd) compression functionality
	URL: http://localhost:8801/tests/test_g3zstd.asp
-->
<%
Option Explicit
%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="utf-8" />
        <title>G3ZSTD - Zstandard Compression Test</title>
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
            table {
                border-collapse: collapse;
                width: 100%;
                margin: 10px 0;
            }
            th,
            td {
                border: 1px solid #999;
                padding: 8px;
                text-align: left;
            }
            th {
                background: #e8e8e8;
            }
        </style>
    </head>
    <body>
        <h1>G3ZSTD - Zstandard Compression Test</h1>
        <p>
            This page demonstrates the G3ZSTD native object using Zstandard
            (zstd) compression with configurable levels.
        </p>

        <div class="section">
            <h2>Test 1: Basic Text Compression with Different Levels</h2>
            <%
            Dim zstd, originalText, level, compressedData
            Set zstd = Server.CreateObject("G3ZSTD")

            originalText = "The quick brown fox jumps over the lazy dog. " & _
                           "Zstandard is a real-time compression algorithm providing high compression ratios. " & _
                           "It offers a trade-off between compression ratio and speed."

            Response.Write "<p><strong>Original Text:</strong></p>"
            Response.Write "<div class='code-box'>" & Server.HTMLEncode(originalText) & "</div>"
            Response.Write "<p><strong>Original Size:</strong> " & Len(originalText) & " bytes</p>"

            ' Try different levels
            Response.Write "<table><tr><th>Level</th><th>Compressed Size</th><th>Ratio</th><th>Status</th></tr>"

            Dim testLevels(2), i
            testLevels(0) = 3
            testLevels(1) = 11
            testLevels(2) = 22

            For i = 0 To UBound(testLevels)
                level = testLevels(i)
                If zstd.SetLevel(level) Then
                    compressedData = zstd.Compress(originalText, level)
                    Dim ratio
                    ratio = ((UBound(compressedData) + 1) / Len(originalText)) * 100
                    Response.Write "<tr><td>" & level & "</td><td>" & (UBound(compressedData) + 1) & " bytes</td>" & _
                                   "<td>" & FormatNumber(ratio, 2) & "%</td><td class='success'>✓</td></tr>"
                Else
                    Response.Write "<tr><td>" & level & "</td><td>-</td><td>-</td><td class='error'>✗ Failed</td></tr>"
                End If
            Next
            Response.Write "</table>"

            Set zstd = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 2: Decompression Roundtrip</h2>
            <%
            Set zstd = Server.CreateObject("G3ZSTD")

            Dim testData, compressedZstd, decompressedZstd
            testData = "This is a test string for zstd compression. It contains repeated patterns: " & _
                       "test test test and string string string and compression compression compression."

            Response.Write "<p><strong>Original:</strong></p>"
            Response.Write "<div class='code-box'>" & Server.HTMLEncode(testData) & "</div>"

            ' Compress
            zstd.SetLevel(11)
            compressedZstd = zstd.Compress(testData)
            Response.Write "<p><strong>Compressed Size:</strong> " & (UBound(compressedZstd) + 1) & " bytes</p>"

            ' Decompress
            decompressedZstd = zstd.DecompressText(compressedZstd)
            Response.Write "<p><strong>Decompressed:</strong></p>"
            Response.Write "<div class='code-box'>" & Server.HTMLEncode(decompressedZstd) & "</div>"

            If decompressedZstd = testData Then
                Response.Write "<p class='success'>✓ Roundtrip successful!</p>"
            Else
                Response.Write "<p class='error'>✗ Roundtrip failed!</p>"
            End If

            Set zstd = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 3: Byte Array Compression</h2>
            <%
            Set zstd = Server.CreateObject("G3ZSTD")

            ' Create binary-like test array
            Dim binaryArray(199), j, k
            For j = 0 To 199
                binaryArray(j) = (j * 13) Mod 256
            Next

            Response.Write "<p><strong>Original Array Size:</strong> " & (UBound(binaryArray) + 1) & " bytes</p>"

            ' Compress
            zstd.SetLevel(15)
            compressedZstd = zstd.Compress(binaryArray)
            Response.Write "<p><strong>Compressed Size (Level 15):</strong> " & (UBound(compressedZstd) + 1) & " bytes</p>"

            ' Decompress
            decompressedZstd = zstd.Decompress(compressedZstd)
            Response.Write "<p><strong>Decompressed Size:</strong> " & (UBound(decompressedZstd) + 1) & " bytes</p>"

            ' Verify
            Dim arrayMatch, m
            arrayMatch = True
            If UBound(decompressedZstd) <> UBound(binaryArray) Then
                arrayMatch = False
            Else
                For m = 0 To UBound(binaryArray)
                    If decompressedZstd(m) <> binaryArray(m) Then
                        arrayMatch = False
                        Exit For
                    End If
                Next
            End If

            If arrayMatch Then
                Response.Write "<p class='success'>✓ Byte array roundtrip successful!</p>"
            Else
                Response.Write "<p class='error'>✗ Array mismatch!</p>"
            End If

            Set zstd = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 4: Batch Compression of Multiple Items</h2>
            <%
            Set zstd = Server.CreateObject("G3ZSTD")

            ' Create array of items
            Dim Items(4)
            Items(0) = "Item 1: First batch item for compression"
            Items(1) = "Item 2: Second batch item with data"
            Items(2) = "Item 3: Third item in the compression batch"
            Items(3) = "Item 4: Fourth and final batch item"
            Items(4) = "Item 5: Extra item to complete the set"

            Response.Write "<p>Batch compressing " & (UBound(Items) + 1) & " items at level 12:</p>"

            zstd.SetLevel(12)
            Dim compressedBatch
            compressedBatch = zstd.CompressMany(Items, 12)

            Response.Write "<p><strong>Compressed Items:</strong> " & (UBound(compressedBatch) + 1) & "</p>"

            Response.Write "<p><strong>Decompressed Items:</strong></p>"
            Response.Write "<ol>"
            Dim n
            For n = 0 To UBound(compressedBatch)
                Response.Write "<li>" & Server.HTMLEncode(zstd.DecompressText(compressedBatch(n))) & "</li>"
            Next
            Response.Write "</ol>"

            Set zstd = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 5: File Compression</h2>
            <%
            Dim fso, sourceFile, targetFile, sourceSize, targetSize
            Set fso = Server.CreateObject("Scripting.FileSystemObject")
            Set zstd = Server.CreateObject("G3ZSTD")

            Dim tempDir
            tempDir = Server.MapPath("../temp")
            If Not fso.FolderExists(tempDir) Then
                fso.CreateFolder(tempDir)
            End If

            sourceFile = Server.MapPath("../temp/zstd_source.txt")
            targetFile = Server.MapPath("../temp/zstd_source.txt.zst")

            ' Create source file
            Dim sourceContent
            sourceContent = String(50, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ") & vbCrLf

            Dim tf
            Set tf = fso.CreateTextFile(sourceFile, True)
            tf.Write sourceContent
            tf.Close()

            sourceSize = fso.GetFile(sourceFile).Size
            Response.Write "<p><strong>Original File Size:</strong> " & sourceSize & " bytes</p>"

            ' Compress file
            zstd.SetLevel(18)
            If zstd.CompressFile(sourceFile, targetFile, 18) Then
                targetSize = fso.GetFile(targetFile).Size
                Response.Write "<p><strong>Compressed File Size:</strong> " & targetSize & " bytes</p>"

                Dim compressionRatio
                compressionRatio = (targetSize / sourceSize) * 100
                Response.Write "<p><strong>Compression Ratio:</strong> " & FormatNumber(compressionRatio, 2) & "%</p>"
                Response.Write "<p class='success'>✓ File compression successful!</p>"

                ' Clean up
                fso.DeleteFile(sourceFile)
                fso.DeleteFile(targetFile)
            Else
                Response.Write "<p class='error'>✗ File compression failed!</p>"
                If Len(zstd.LastError) > 0 Then
                    Response.Write "<p>Error: " & Server.HTMLEncode(zstd.LastError) & "</p>"
                End If
            End If

            Set zstd = Nothing
            Set fso = Nothing
            %>
        </div>
        <div class="section">
            <h2>Compression Level Guide</h2>
            <p>G3ZSTD supports compression levels from -5 to 22:</p>
            <table>
                <tr>
                    <th>Range</th>
                    <th>Description</th>
                    <th>Use Case</th>
                </tr>
                <tr>
                    <td>-5 to 1</td>
                    <td>Very Fast</td>
                    <td>Real-time applications, streaming</td>
                </tr>
                <tr>
                    <td>2 to 8</td>
                    <td>Fast</td>
                    <td>General purpose compression</td>
                </tr>
                <tr>
                    <td>9 to 15</td>
                    <td>Balanced (default 11)</td>
                    <td>Good balance of speed and ratio</td>
                </tr>
                <tr>
                    <td>16 to 22</td>
                    <td>Maximum Compression</td>
                    <td>Archive, long-term storage</td>
                </tr>
            </table>
        </div>

        <hr />
        <small
            >AxonASP Server - G3ZSTD Zstandard Compression Test &copy;
            2026</small
        >
    </body>
</html>

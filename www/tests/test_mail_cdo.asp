<%@ Language=JScript %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>CDO Mail Capabilities Test</title>
    <link rel="stylesheet" href="../css/axonasp.css">
    <style>
        .test-card {
            background-color: var(--bg-elevated);
            border: 1px solid var(--border);
            padding: 1.5rem;
            margin-bottom: 1.5rem;
            border-radius: var(--radius-md);
        }
        .status-pass {
            color: var(--success);
            font-weight: bold;
        }
        .status-fail {
            color: var(--danger);
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div id="main-container" class="manual-page">
        <header id="header">
            <h1>CDO Mail Capabilities Test</h1>
            <p class="text-muted">Verifying HTML body parsing, attachment adding, and inline related part formatting using the internal JScript Engine.</p>
        </header>

        <div id="content">
            <%
            function getRegExpSafe(str) {
                return str.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
            }

            var fso = Server.CreateObject("Scripting.FileSystemObject");
            
            var tempAttachPath = Server.MapPath("temp_attach.txt");
            var tempImagePath = Server.MapPath("temp_image.png");
            
            var f1 = fso.CreateTextFile(tempAttachPath, true);
            f1.WriteLine("attachment dummy content");
            f1.Close();

            var f2 = fso.CreateTextFile(tempImagePath, true);
            f2.WriteLine("image dummy content");
            f2.Close();

            var passCount = 0;
            var testCount = 0;

            function runTest(name, fn) {
                testCount++;
                try {
                    var res = fn();
                    if (res === true) {
                        passCount++;
                        Response.Write("<div class='test-card'><h3>Test: " + name + "</h3><p>Result: <span class='status-pass'>PASS</span></p></div>");
                    } else {
                        Response.Write("<div class='test-card'><h3>Test: " + name + "</h3><p>Result: <span class='status-fail'>FAIL</span> - " + res + "</p></div>");
                    }
                } catch(e) {
                    Response.Write("<div class='test-card'><h3>Test: " + name + "</h3><p>Result: <span class='status-fail'>FAIL (Exception)</span> - " + e.message + "</p></div>");
                }
            }

            var objCDOMsg = Server.CreateObject("G3MAIL");

            // Test 1: HTMLBody property get/set and JScript string method capabilities
            runTest("HTMLBody get, set, and string manipulation (.match, .replace)", function() {
                objCDOMsg.HTMLBody = "<html><body><img src=\"/images/logo.png\"></body></html>";
                var html = objCDOMsg.HTMLBody;
                if (html != "<html><body><img src=\"/images/logo.png\"></body></html>") {
                    return "Initial HTMLBody mismatch: " + html;
                }

                var arrImageElemStrings = objCDOMsg.HTMLBody.match(/<img\s+[^>]+>/ig);
                if (!arrImageElemStrings || arrImageElemStrings.length !== 1 || arrImageElemStrings[0] !== '<img src="/images/logo.png">') {
                    return "Regex match failed on HTMLBody: " + arrImageElemStrings;
                }

                var strImageSrc = "/images/logo.png";
                var strCID = "logo_cid";
                objCDOMsg.HTMLBody = objCDOMsg.HTMLBody.replace(
                    new RegExp(" src=\"[^\"]*" + getRegExpSafe(strImageSrc), "ig"),
                    " src=\"cid:" + strCID
                );

                var expectedHtml = "<html><body><img src=\"cid:logo_cid\"></body></html>";
                if (objCDOMsg.HTMLBody !== expectedHtml) {
                    return "Regex replace on HTMLBody failed. Got: " + objCDOMsg.HTMLBody;
                }

                return true;
            });

            // Test 2: AddRelatedBodyPart and modifying returned BodyPart's Fields.Item / Update
            runTest("AddRelatedBodyPart with Fields.Item and Update", function() {
                var strCID = "inline_logo";
                var objBP = objCDOMsg.AddRelatedBodyPart(tempImagePath, strCID);
                
                if (!objBP) {
                    return "AddRelatedBodyPart returned null/undefined";
                }

                objBP.Fields.Item("urn:schemas:mailheader:Content-ID") = "<" + strCID + ">";
                objBP.Fields.Update();

                var gottenCID = objBP.Fields.Item("urn:schemas:mailheader:Content-ID");
                if (gottenCID !== "<" + strCID + ">") {
                    return "Fields.Item get returned mismatch: " + gottenCID;
                }

                return true;
            });

            // Test 3: AddAttachment
            runTest("AddAttachment", function() {
                var res = objCDOMsg.AddAttachment(tempAttachPath);
                if (res !== true) {
                    return "AddAttachment returned: " + res;
                }
                return true;
            });

            try {
                fso.DeleteFile(tempAttachPath);
                fso.DeleteFile(tempImagePath);
            } catch(e) {}

            Response.Write("<div class='test-card'><h2>Summary: " + passCount + "/" + testCount + " tests passed.</h2></div>");
            %>
        </div>
    </div>
</body>
</html>
